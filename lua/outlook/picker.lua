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

--- Keep a picker row in sync with its item table after a mutation
--- (toggle_read/toggle_flag/mark-read-on-view all update `item` in
--- place, but the row snacks already rendered for it was built from a
--- one-time snapshot at list-open time — see docs/DESIGN.md and
--- docs/HANDOFF.md: without this, the change is only visible after
--- closing and reopening the picker, even though the underlying
--- Outlook COM write already succeeded).
--- @param picker table? the picker instance, if available (actions get
--- one; the vim.ui.select fallback path doesn't)
local function refresh_row(picker, item)
  item.text = format_item(item)
  if not picker then
    return
  end
  -- Confirmed against snacks.nvim's source (lua/snacks/picker/core/picker.lua):
  -- Picker:refresh() clears the selection, sets the target back to the
  -- current item, and re-runs the finder/matcher. Since our finder is
  -- the same `items` table we mutated above (see M.show), this should
  -- redraw the row with the fresh text. Still pcall-guarded since it's
  -- untested against a real install from this environment.
  pcall(function()
    picker:refresh()
  end)
end

--- Fetch a message's full body via get_message, then mark it read (if
--- it was unread) — matching common mail-client UX where viewing a
--- message is what marks it read, whether that view is the full
--- floating window (open_message) or the picker's own preview pane
--- (load_body). Both are explicit user actions; nothing here fires on
--- its own.
--- @param picker table? see refresh_row
local function fetch_and_mark_read(item, on_body, picker)
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
      helper.request("mark_read", { entry_id = item.entry_id, store_id = item.store_id }, function(ok2, mark_result)
        if ok2 then
          item.unread = mark_result.unread
          refresh_row(picker, item)
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
local function load_body(item, picker)
  if body_cache[item.entry_id] then
    refresh_preview_if_current(item)
    return
  end
  fetch_and_mark_read(item, function(result)
    body_cache[item.entry_id] = result.body or ""
    refresh_preview_if_current(item)
  end, picker)
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

--- "Load more" state for the currently-open snacks picker (see
--- M.load_more below), or nil if no open picker supports it (the
--- vim.ui.select fallback, or M.show() called without method/params).
--- @type { method: string, params: table, items: table[], limit: integer, loading: boolean }?
local current_list_state = nil

local LOAD_MORE_PAGE_SIZE = 50

--- @param picker table? see refresh_row. The confirm handler passes
--- none (the picker is closing anyway); the vim.ui.select fallback has
--- none either way.
function M.open_message(item, picker)
  fetch_and_mark_read(item, function(result)
    body_cache[item.entry_id] = result.body or ""
    preview.open(result)
  end, picker)
end

--- @param picker table? see refresh_row
function M.toggle_read(item, picker)
  local method = item.unread and "mark_read" or "mark_unread"
  helper.request(method, { entry_id = item.entry_id, store_id = item.store_id }, function(ok, result)
    if ok then
      item.unread = result.unread
      refresh_row(picker, item)
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
--- @param picker table? see refresh_row
function M.toggle_flag(item, picker)
  local method = (item.flag_status == "flagged") and "clear_flag" or "set_flag"
  helper.request(method, { entry_id = item.entry_id, store_id = item.store_id }, function(ok, result)
    if ok then
      item.flag_status = result.flag_status
      refresh_row(picker, item)
      invalidate_lists()
    else
      vim.schedule(function()
        notify.error(result)
      end)
    end
  end)
end

--- Explicit action (picker's <C-e>): fetch a bigger page of the same
--- list/search and replace the picker's items with it, without closing
--- the picker.
---
--- Deliberately NOT <C-n>/<C-p> (snacks.picker's default list_down/
--- list_up navigation — confirmed against snacks.nvim's source) or the
--- other Ctrl-key defaults it already binds (<C-f>=preview_scroll_down,
--- <C-d>/<C-u>=scroll, etc., see docs/DESIGN.md 6.2): those are
--- pressed constantly while browsing, so overriding one for an
--- occasional-use action would be a worse trade than <C-r>/<C-f>/<C-l>
--- already are. <C-e> is confirmed unbound by default.
---
--- This is NOT true incremental pagination: Outlook COM's Items
--- collection has no offset/cursor, and snacks.picker's finder is
--- re-run from scratch rather than appended to (confirmed against
--- snacks.nvim's source — see docs/DESIGN.md 6.1), so "loading more"
--- means re-fetching from the start with a larger `limit` and
--- replacing the whole item set. Guarded against overlapping requests
--- via state.loading; only ever runs when the user presses the key.
function M.load_more(picker)
  local state = current_list_state
  if not state or state.loading then
    return
  end
  state.loading = true

  local new_limit = state.limit + LOAD_MORE_PAGE_SIZE
  local params = vim.tbl_extend("force", {}, state.params, { limit = new_limit })

  notify.info("Outlook: もっと読み込み中…")
  helper.request(state.method, params, function(ok, result)
    state.loading = false
    if not ok then
      return vim.schedule(function()
        notify.error(result)
      end)
    end
    vim.schedule(function()
      if current_list_state ~= state then
        return -- superseded by a newer list()/search in the meantime
      end
      local reached_end = #result.items <= #state.items
      state.limit = new_limit
      state.params = params

      -- Mutate the same table object handed to Snacks.picker.pick (see
      -- M.show) in place, rather than assigning a new table, so the
      -- picker's finder reference stays valid.
      local new_items = vim.tbl_map(to_picker_item, result.items)
      for i = #state.items, 1, -1 do
        state.items[i] = nil
      end
      for i, it in ipairs(new_items) do
        state.items[i] = it
      end

      pcall(function()
        picker:refresh()
      end)

      if reached_end then
        notify.info("Outlook: これ以上のメールはありません")
      end
    end)
  end)
end

--- Render an already-fetched message list.
--- @param messages table[] raw items from the helper (see docs/DESIGN.md)
--- @param opts table? { title, method, params } — method/params (e.g.
--- {folder=,limit=,unread_only=} or {query=,limit=}) enable <C-e>
--- "load more"; omit them to disable it for this listing.
function M.show(messages, opts)
  opts = opts or {}

  if has_snacks() then
    local Snacks = require("snacks")
    local items = vim.tbl_map(to_picker_item, messages)

    current_list_state = opts.method
        and {
          method = opts.method,
          params = opts.params or {},
          items = items,
          limit = (opts.params and opts.params.limit) or #items,
          loading = false,
        }
      or nil

    Snacks.picker.pick({
      source = "outlook_messages",
      title = opts.title or "Outlook",
      items = items,
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
        toggle_read = function(picker, item)
          M.toggle_read(item, picker)
        end,
        toggle_flag = function(picker, item)
          M.toggle_flag(item, picker)
        end,
        load_body = function(picker, item)
          load_body(item, picker)
        end,
        load_more = function(picker, _)
          M.load_more(picker)
        end,
      },
      win = {
        input = {
          keys = {
            ["<C-r>"] = { "toggle_read", mode = { "i", "n" } },
            ["<C-f>"] = { "toggle_flag", mode = { "i", "n" } },
            ["<C-l>"] = { "load_body", mode = { "i", "n" } },
            ["<C-e>"] = { "load_more", mode = { "i", "n" } },
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
      M.show(result, { title = opts.title, method = "list_messages", params = params })
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
