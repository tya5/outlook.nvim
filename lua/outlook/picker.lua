-- Mail list UI: snacks.picker when available, vim.ui.select fallback.
--
-- Latency handling (see docs/DESIGN.md 6.1 for the full rationale):
--  * both UI paths open only once the result (from cache or the
--    helper) is ready; a cache hit resolves synchronously with no
--    perceptible wait, a miss shows a "loading" notification first.
--  * identical requests (same folder/filter) are cached briefly and
--    in-flight duplicates are coalesced, so repeated keypresses don't
--    each re-hit Outlook COM.

local helper = require("outlook.helper")
local notify = require("outlook.notify")
local preview = require("outlook.preview")
local config = require("outlook.config")

local M = {}

-- "Bold" is not a highlight group that exists by default; define our
-- own so unread messages render distinctly regardless of colorscheme.
vim.api.nvim_set_hl(0, "OutlookUnread", { bold = true, default = true })

---@type table<string, {items: table[], ts: integer}>
local cache = {}
---@type table<string, fun(ok:boolean, result:table)[]>
local inflight = {}

local function now_ms()
  return vim.uv.now()
end

--- Build a deterministic cache/inflight key. Doesn't rely on
--- vim.json.encode's table iteration order (unspecified for string
--- keys, even if stable in practice for a fixed Lua/table-hashing
--- build): params here are always flat scalars (folder/limit/
--- unread_only/query), so sorting the keys and joining is sufficient.
local function cache_key(method, params)
  params = params or {}
  local keys = {}
  for k in pairs(params) do
    keys[#keys + 1] = k
  end
  table.sort(keys)
  local parts = {}
  for _, k in ipairs(keys) do
    parts[#parts + 1] = k .. "=" .. tostring(params[k])
  end
  return method .. ":" .. table.concat(parts, "&")
end

--- Fetch items for `method`/`params`, deduping concurrent identical
--- requests and reusing a recent result when within cache_ttl_ms.
--- @param force boolean? bypass the cache (used by :OutlookRefresh)
--- @return boolean served_from_cache true if `callback` was already
--- invoked synchronously from a warm cache entry; false means a real
--- helper round-trip is in flight (either just started, or joined an
--- existing one) and callers may want to show a loading indicator.
local function fetch(method, params, force, callback)
  local key = cache_key(method, params)

  if not force then
    local cached = cache[key]
    if cached and (now_ms() - cached.ts) < config.options.cache_ttl_ms then
      callback(true, cached.items)
      return true
    end
  end

  if inflight[key] then
    table.insert(inflight[key], callback)
    return false
  end
  inflight[key] = { callback }

  helper.request(method, params, function(ok, result)
    local waiters = inflight[key]
    inflight[key] = nil
    local items = ok and (result.items or {}) or nil
    if ok then
      cache[key] = { items = items, ts = now_ms() }
    end
    for _, cb in ipairs(waiters) do
      if ok then
        cb(true, items)
      else
        cb(false, result)
      end
    end
  end)
  return false
end

--- Drop cached list_messages/search_messages results so the next
--- picker.list()/search picks up fresh unread state instead of serving
--- a stale snapshot for up to cache_ttl_ms after a mutation.
local function invalidate_lists()
  for key in pairs(cache) do
    if key:find("^list_messages:") or key:find("^search_messages:") then
      cache[key] = nil
    end
  end
end

local function has_snacks()
  return pcall(require, "snacks")
end

local function format_item(msg)
  local status = msg.unread and "●" or " "
  local flag = (msg.flag_status == "flagged") and "🚩" or " "
  return string.format("%s%s %-20s %s", status, flag, msg.from or "", msg.subject or "(no subject)")
end

--- list_messages/search_messages intentionally omit the message body
--- (see docs/DESIGN.md and helper/outlook-helper.ps1's
--- ConvertTo-MessageSummary): fetching Body for every row in a folder
--- listing is slow and touches an Object Model Guard-sensitive
--- property. The preview pane is header-only until the user explicitly
--- asks to load the body (see body_cache/render_preview below) — never
--- automatically on cursor movement, so browsing the list never fires a
--- get_message per row.
local function preview_lines(msg)
  local lines = {
    ("Subject : %s"):format(msg.subject or ""),
    ("From    : %s"):format(msg.from or ""),
    ("Date    : %s"):format(msg.received or ""),
  }
  if msg.flag_status and msg.flag_status ~= "none" then
    table.insert(lines, ("Flag    : %s"):format(msg.flag_status))
  end
  return lines
end

---@type table<string, string> entry_id -> fetched body, so re-viewing an
--- already-loaded message's preview doesn't re-fetch it.
local body_cache = {}

--- The ctx from the most recent snacks preview() call, so the explicit
--- "load body" action (bound to <C-l>, see M.show) can write into the
--- already-visible preview pane instead of needing some "force a
--- redraw" API from snacks that may or may not exist.
local current_preview_ctx = nil

--- Compose the preview pane's lines: header, plus either the cached
--- body or a hint to load it.
local function render_preview(item)
  local lines = preview_lines(item)
  local body = body_cache[item.entry_id]
  table.insert(lines, "")
  if body then
    table.insert(lines, string.rep("-", 40))
    table.insert(lines, "")
    vim.list_extend(lines, vim.split(body, "\n", { plain = true }))
  else
    table.insert(lines, "(<C-l> で本文を読み込む — 読み込むと既読になります)")
  end
  return lines
end

--- Push freshly rendered lines into the preview pane if it's still
--- showing this same item (the user may have moved the cursor to a
--- different item between triggering the load and the response
--- arriving; in that case there's nothing to update).
local function refresh_preview_if_current(item)
  if current_preview_ctx and current_preview_ctx.item and current_preview_ctx.item.entry_id == item.entry_id then
    -- The picker may have been closed between triggering the load and
    -- the response arriving; set_lines on a destroyed preview buffer
    -- would error, so this is best-effort.
    pcall(function()
      current_preview_ctx.preview:set_lines(render_preview(item))
    end)
  end
end

--- Fetch a message's full body via get_message, then mark it read (if
--- it was unread) — matching common mail-client UX where viewing a
--- message is what marks it read, whether that view is the full
--- floating window (open_message) or the picker's own preview pane
--- (load_body). Both are explicit user actions; nothing here fires on
--- its own.
local function fetch_and_mark_read(item, on_body)
  helper.request("get_message", { entry_id = item.entry_id, store_id = item.store_id }, function(ok, result)
    if not ok then
      return vim.schedule(function()
        notify.error(result)
      end)
    end
    vim.schedule(function()
      on_body(result)
    end)
    if item.unread then
      helper.request("mark_read", { entry_id = item.entry_id, store_id = item.store_id }, function(ok2)
        if ok2 then
          item.unread = false
          invalidate_lists()
        end
      end)
    end
  end)
end

--- Explicit action (picker's <C-l>): load the current item's body into
--- the picker's own preview pane, without leaving the picker (unlike
--- <CR>, which opens the full floating window). Only ever runs when
--- the user presses the key — never automatically.
local function load_body(item)
  if body_cache[item.entry_id] then
    refresh_preview_if_current(item)
    return
  end
  fetch_and_mark_read(item, function(result)
    body_cache[item.entry_id] = result.body or ""
    refresh_preview_if_current(item)
  end)
end

local function to_picker_item(msg)
  return {
    text = format_item(msg),
    subject = msg.subject,
    from = msg.from,
    received = msg.received,
    unread = msg.unread,
    flag_status = msg.flag_status,
    entry_id = msg.entry_id,
    store_id = msg.store_id,
  }
end

function M.open_message(item)
  fetch_and_mark_read(item, function(result)
    body_cache[item.entry_id] = result.body or ""
    preview.open(result)
  end)
end

function M.toggle_read(item)
  local method = item.unread and "mark_read" or "mark_unread"
  helper.request(method, { entry_id = item.entry_id, store_id = item.store_id }, function(ok, result)
    if ok then
      invalidate_lists()
    else
      vim.schedule(function()
        notify.error(result)
      end)
    end
  end)
end

--- v1 only toggles between "flagged" and "none" (see
--- helper/outlook-helper.ps1's Invoke-SetFlag); "complete" isn't
--- reachable from here yet.
function M.toggle_flag(item)
  local method = (item.flag_status == "flagged") and "clear_flag" or "set_flag"
  helper.request(method, { entry_id = item.entry_id, store_id = item.store_id }, function(ok, result)
    if ok then
      invalidate_lists()
    else
      vim.schedule(function()
        notify.error(result)
      end)
    end
  end)
end

--- Render an already-fetched message list (used by :OutlookSearch, which
--- fetches via its own helper call and hands results straight to the UI).
--- @param messages table[] raw items from the helper (see docs/DESIGN.md)
--- @param opts table? { title }
function M.show(messages, opts)
  opts = opts or {}

  if has_snacks() then
    local Snacks = require("snacks")
    Snacks.picker.pick({
      source = "outlook_messages",
      title = opts.title or "Outlook",
      items = vim.tbl_map(to_picker_item, messages),
      format = function(item)
        return { { item.text, item.unread and "OutlookUnread" or "Normal" } }
      end,
      preview = function(ctx)
        current_preview_ctx = ctx
        ctx.preview:set_lines(render_preview(ctx.item))
        return true
      end,
      confirm = function(picker, item)
        picker:close()
        M.open_message(item)
      end,
      actions = {
        toggle_read = function(_, item)
          M.toggle_read(item)
        end,
        toggle_flag = function(_, item)
          M.toggle_flag(item)
        end,
        load_body = function(_, item)
          load_body(item)
        end,
      },
      win = {
        input = {
          keys = {
            ["<C-r>"] = { "toggle_read", mode = { "i", "n" } },
            ["<C-f>"] = { "toggle_flag", mode = { "i", "n" } },
            ["<C-l>"] = { "load_body", mode = { "i", "n" } },
          },
        },
      },
    })
    return
  end

  local labels = {}
  local items = {}
  for _, msg in ipairs(messages) do
    items[#items + 1] = msg
    labels[#labels + 1] = format_item(msg)
  end
  vim.ui.select(labels, { prompt = opts.title or "Outlook" }, function(_, idx)
    if not idx then
      return
    end
    M.open_message(items[idx])
  end)
end

--- Fetch and render a folder listing.
--- @param opts table? { folder, limit, unread_only, title, force }
function M.list(opts)
  opts = opts or {}
  local params = {
    folder = opts.folder or "inbox",
    limit = opts.limit or 50,
    unread_only = opts.unread_only or false,
  }

  local served_from_cache = fetch("list_messages", params, opts.force, function(ok, result)
    if not ok then
      return vim.schedule(function()
        notify.error(result)
      end)
    end
    vim.schedule(function()
      M.show(result, { title = opts.title })
    end)
  end)

  -- Only show a "loading" indicator when a real helper round-trip is
  -- actually happening; a cache hit already invoked the callback above
  -- synchronously and needs no such feedback (previously this fired
  -- unconditionally on every non-snacks call, including cache hits).
  if not served_from_cache then
    notify.info("Outlook: 読み込み中…")
  end
end

return M
