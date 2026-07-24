-- Bridges Outlook COM calendar events (via the PowerShell helper's
-- list_events method) into almanac.nvim's generic Event/EventProvider
-- contract (see almanac.nvim's docs/DESIGN.md, sections 3.2 and 5).
--
-- almanac.nvim is an optional dependency: outlook.nvim itself has no
-- calendar UI of its own and never will (see docs/DESIGN.md) — without
-- almanac.nvim installed, :OutlookCalendar just notifies and does
-- nothing, matching the snacks.nvim-optional precedent already used
-- throughout this plugin.

local helper = require("outlook.helper")
local notify = require("outlook.notify")
local preview = require("outlook.preview")
local config = require("outlook.config")

local M = {}

local function has_almanac()
  return pcall(require, "almanac")
end

--- @param msg table raw event summary from the helper (entry_id, store_id, subject, start, stop, all_day, location, busy)
--- @return almanac.Event
local function to_almanac_event(msg)
  return {
    id = msg.entry_id,
    title = msg.subject,
    start = msg.start,
    stop = msg.stop,
    all_day = msg.all_day,
    location = msg.location,
    busy = msg.busy,
    data = { entry_id = msg.entry_id, store_id = msg.store_id },
  }
end

-- Cache + in-flight dedup for list_events, mirroring picker.lua's fetch()
-- for list_messages/search_messages (see docs/DESIGN.md 6.1 and
-- docs/HANDOFF.md 11). Without this, every view switch and every
-- <C-f>/<C-b> page (month view especially — its range is 5-6 weeks,
-- wide enough that recurring-meeting expansion over it is real COM
-- work) re-hits Outlook COM from scratch, even when revisiting a
-- range already fetched moments ago.
---@type table<string, {events: almanac.Event[], ts: integer}>
local cache = {}
---@type table<string, fun(ok:boolean, events_or_err:any)[]>
local inflight = {}

local function now_ms()
  return vim.uv.now()
end

local function range_key(range)
  return ("%d:%d"):format(range.from, range.to)
end

--- @return boolean served_from_cache true if `cb` was already invoked
--- synchronously from a warm cache entry; false means a real helper
--- round-trip is in flight (just started, or joined an existing one).
local function fetch(range, cb)
  local key = range_key(range)

  local cached = cache[key]
  if cached and (now_ms() - cached.ts) < config.options.cache_ttl_ms then
    cb(true, cached.events)
    return true
  end

  if inflight[key] then
    table.insert(inflight[key], cb)
    return false
  end
  inflight[key] = { cb }

  helper.request("list_events", { from = range.from, to = range.to }, function(ok, result)
    local waiters = inflight[key]
    inflight[key] = nil
    local events = ok and vim.tbl_map(to_almanac_event, result.items) or nil
    if ok then
      cache[key] = { events = events, ts = now_ms() }
    end
    for _, waiter in ipairs(waiters) do
      waiter(ok, ok and events or result)
    end
  end)
  return false
end

--- almanac.EventProvider: fetches Outlook calendar events for the given
--- range via the helper's list_events method (async function(range, cb) form).
--- @param range almanac.Range
--- @param cb fun(events: almanac.Event[])
function M.provider(range, cb)
  local served_from_cache = fetch(range, function(ok, events_or_err)
    vim.schedule(function()
      if not ok then
        notify.error(events_or_err)
        cb({})
        return
      end
      cb(events_or_err)
    end)
  end)

  -- Only show a "loading" indicator when a real helper round-trip is
  -- actually happening; a cache hit already invoked cb() above
  -- synchronously (well, on the next scheduled tick) and needs no such
  -- feedback — matches picker.lua's M.list().
  if not served_from_cache then
    notify.info("Outlook: loading calendar…")
  end
end

--- The single shared almanac.Calendar instance backing :OutlookCalendar
--- (created lazily on first open).
M._cal = nil

--- Open (creating on first use) the almanac.nvim calendar wired to
--- Outlook. No-ops with an error notification if almanac.nvim isn't
--- installed.
--- @return table? the almanac.Calendar instance, or nil if unavailable
function M.open()
  if not has_almanac() then
    notify.error("almanac.nvim is not installed — :OutlookCalendar requires it (see README)")
    return nil
  end

  if not M._cal then
    local Almanac = require("almanac")
    M._cal = Almanac({ events = M.provider })
    -- Mirrors outlook.picker's open_message: fetch full details on
    -- selection and open them the same way a mail message is opened
    -- (lua/outlook/preview.lua's M.open_event, "buffer" or "float" per
    -- opts.message_window). occurrence_start (the specific occurrence's
    -- own start epoch, already known from list_events) lets the helper
    -- resolve the exact occurrence of a recurring meeting instead of
    -- always returning the series master — see Invoke-GetAppointment.
    M._cal:on("event_selected", function(_, event)
      helper.request("get_appointment", {
        entry_id = event.data.entry_id,
        store_id = event.data.store_id,
        occurrence_start = event.start,
      }, function(ok, result)
        vim.schedule(function()
          if not ok then
            notify.error(result)
            return
          end
          preview.open_event(result)
        end)
      end)
    end)
  end

  M._cal:show()
  return M._cal
end

return M
