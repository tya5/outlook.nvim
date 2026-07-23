-- Tests for lua/outlook/calendar.lua: the Outlook COM <-> almanac.nvim
-- Event/EventProvider bridge. outlook.helper is faked (no real IPC);
-- almanac.nvim is faked via package.loaded injection (same technique as
-- tests/outlook/picker_spec.lua's fake `snacks` module), since it isn't
-- installed in this environment either.

describe("outlook.calendar", function()
  local calls

  before_each(function()
    calls = {}

    package.loaded["outlook.calendar"] = nil
    package.loaded["outlook.helper"] = nil
    package.loaded["outlook.notify"] = nil
    package.loaded["almanac"] = nil

    package.loaded["outlook.helper"] = {
      request = function(method, params, cb)
        table.insert(calls, { method = method, params = params, cb = cb })
      end,
      is_running = function()
        return true
      end,
      start = function() end,
      prewarm = function() end,
    }
  end)

  after_each(function()
    package.loaded["almanac"] = nil
  end)

  describe("provider", function()
    it("requests list_events with the range and maps results to almanac.Event", function()
      local calendar = require("outlook.calendar")
      local got_events

      calendar.provider({ from = 1000, to = 2000 }, function(events)
        got_events = events
      end)

      assert.equals(1, #calls)
      assert.equals("list_events", calls[1].method)
      assert.same({ from = 1000, to = 2000 }, calls[1].params)

      calls[1].cb(true, {
        items = {
          {
            entry_id = "e1",
            store_id = "s1",
            subject = "Team sync",
            start = 1100,
            stop = 1200,
            all_day = false,
            location = "Zoom",
            busy = "busy",
          },
        },
      })
      vim.wait(20)

      assert.is_not_nil(got_events)
      assert.equals(1, #got_events)
      local ev = got_events[1]
      assert.equals("e1", ev.id)
      assert.equals("Team sync", ev.title)
      assert.equals(1100, ev.start)
      assert.equals(1200, ev.stop)
      assert.equals("Zoom", ev.location)
      assert.equals("busy", ev.busy)
      assert.same({ entry_id = "e1", store_id = "s1" }, ev.data)
    end)

    it("resolves to an empty list and notifies on error", function()
      local notify_errors = 0
      package.loaded["outlook.notify"] = {
        error = function()
          notify_errors = notify_errors + 1
        end,
        info = function() end,
      }
      package.loaded["outlook.calendar"] = nil
      local calendar = require("outlook.calendar")

      local got_events
      calendar.provider({ from = 1000, to = 2000 }, function(events)
        got_events = events
      end)
      calls[1].cb(false, { code = "OUTLOOK_NOT_RUNNING", message = "boom" })
      vim.wait(20)

      assert.same({}, got_events)
      assert.equals(1, notify_errors)
    end)
  end)

  describe("open", function()
    it("notifies and returns nil when almanac.nvim is not installed", function()
      local notify_errors = 0
      package.loaded["outlook.notify"] = {
        error = function()
          notify_errors = notify_errors + 1
        end,
        info = function() end,
      }
      package.loaded["outlook.calendar"] = nil
      local calendar = require("outlook.calendar")

      local result = calendar.open()

      assert.is_nil(result)
      assert.equals(1, notify_errors)
    end)

    it("constructs an almanac.Calendar wired to provider() and reuses it on repeated opens", function()
      local new_calls = 0
      local fake_cal = {
        show = function(self)
          return self
        end,
        on = function(self)
          return self
        end,
      }
      package.loaded["almanac"] = setmetatable({}, {
        __call = function(_, opts)
          new_calls = new_calls + 1
          assert.equals("function", type(opts.events))
          return fake_cal
        end,
      })
      package.loaded["outlook.calendar"] = nil
      local calendar = require("outlook.calendar")

      local first = calendar.open()
      local second = calendar.open()

      assert.equals(fake_cal, first)
      assert.equals(fake_cal, second)
      assert.equals(1, new_calls) -- constructed once, shown twice
    end)
  end)
end)
