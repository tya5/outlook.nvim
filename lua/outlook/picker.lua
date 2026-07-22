-- Mail list UI: snacks.picker when available, vim.ui.select fallback.
--
-- Latency handling (see docs/DESIGN.md and project decisions):
--  * snacks.picker path uses an async finder, so the picker window opens
--    immediately with its built-in loading state instead of blocking on
--    the helper round-trip.
--  * vim.ui.select path has no such affordance, so we notify "loading"
--    immediately and populate once data arrives.
--  * identical requests (same folder/filter) are cached briefly and
--    in-flight duplicates are coalesced, so repeated keypresses don't
--    each re-hit Outlook COM.

local helper = require("outlook.helper")
local notify = require("outlook.notify")
local preview = require("outlook.preview")
local config = require("outlook.config")

local M = {}

---@type table<string, {items: table[], ts: integer}>
local cache = {}
---@type table<string, fun(ok:boolean, result:table)[]>
local inflight = {}

local function now_ms()
  return vim.uv.now()
end

local function cache_key(method, params)
  return method .. ":" .. vim.json.encode(params or {})
end

--- Fetch items for `method`/`params`, deduping concurrent identical
--- requests and reusing a recent result when within cache_ttl_ms.
--- @param force boolean? bypass the cache (used by :OutlookRefresh)
local function fetch(method, params, force, callback)
  local key = cache_key(method, params)

  if not force then
    local cached = cache[key]
    if cached and (now_ms() - cached.ts) < config.options.cache_ttl_ms then
      callback(true, cached.items)
      return
    end
  end

  if inflight[key] then
    table.insert(inflight[key], callback)
    return
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
  return string.format("%s %-20s %s", status, msg.from or "", msg.subject or "(no subject)")
end

local function to_picker_item(msg)
  return {
    text = format_item(msg),
    subject = msg.subject,
    from = msg.from,
    received = msg.received,
    unread = msg.unread,
    entry_id = msg.entry_id,
    store_id = msg.store_id,
    preview = { text = msg.preview or "", ft = "text" },
  }
end

function M.open_message(item)
  helper.request("get_message", { entry_id = item.entry_id, store_id = item.store_id }, function(ok, result)
    if not ok then
      return vim.schedule(function()
        notify.error(result)
      end)
    end
    vim.schedule(function()
      preview.open(result)
    end)
    if item.unread then
      -- Opening a message marks it read, matching common mail-client UX.
      helper.request("mark_read", { entry_id = item.entry_id, store_id = item.store_id }, function(ok)
        if ok then
          invalidate_lists()
        end
      end)
    end
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
        return { { item.text, item.unread and "Bold" or "Normal" } }
      end,
      preview = function(ctx)
        ctx.preview:set_lines(vim.split(ctx.item.preview.text, "\n", { plain = true }))
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
      },
      win = {
        input = {
          keys = {
            ["<C-r>"] = { "toggle_read", mode = { "i", "n" } },
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

  if not has_snacks() then
    notify.info("Outlook: 読み込み中…")
  end

  fetch("list_messages", params, opts.force, function(ok, result)
    if not ok then
      return vim.schedule(function()
        notify.error(result)
      end)
    end
    vim.schedule(function()
      M.show(result, { title = opts.title })
    end)
  end)
end

return M
