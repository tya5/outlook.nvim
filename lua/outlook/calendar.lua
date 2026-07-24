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

--- almanac.EventProvider: fetches Outlook calendar events for the given
--- range via the helper's list_events method (async function(range, cb) form).
--- @param range almanac.Range
--- @param cb fun(events: almanac.Event[])
function M.provider(range, cb)
  helper.request("list_events", { from = range.from, to = range.to }, function(ok, result)
    vim.schedule(function()
      if not ok then
        notify.error(result)
        cb({})
        return
      end
      cb(vim.tbl_map(to_almanac_event, result.items))
    end)
  end)
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
