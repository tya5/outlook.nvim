-- Tests for lua/outlook/preview.lua: the full-item detail views (mail
-- message and calendar event). Unlike helper/outlook-helper.ps1, this is
-- plain Lua manipulating real buffers/windows, so it's tested directly
-- against a real headless Neovim instance (no mocking needed).

describe("outlook.preview", function()
  before_each(function()
    vim.cmd("silent! only")
    package.loaded["outlook.preview"] = nil
    package.loaded["outlook.config"] = nil
    package.loaded["outlook.notify"] = nil
  end)

  after_each(function()
    vim.cmd("silent! only")
  end)

  describe("open_event", function()
    it("renders subject/organizer/time/location/status metadata and the body", function()
      local preview = require("outlook.preview")

      preview.open_event({
        subject = "Team sync",
        organizer = "Alice",
        start = "2026-08-24 08:00",
        stop = "2026-08-24 08:30",
        location = "Zoom",
        busy = "busy",
        body = "Agenda line 1\nAgenda line 2",
      })

      local lines = vim.api.nvim_buf_get_lines(0, 0, -1, false)
      assert.equals("Subject   : Team sync", lines[1])
      assert.equals("Organizer : Alice", lines[2])
      assert.equals("Start     : 2026-08-24 08:00", lines[3])
      assert.equals("End       : 2026-08-24 08:30", lines[4])
      assert.equals("Location  : Zoom", lines[5])
      assert.equals("Status    : busy", lines[6])

      local found_body = false
      for _, line in ipairs(lines) do
        if line == "Agenda line 1" then
          found_body = true
        end
      end
      assert.is_true(found_body)
      assert.equals("nofile", vim.bo.buftype)
      assert.is_false(vim.bo.modifiable)
    end)

    it("marks an all-day event without a separate start/end line", function()
      local preview = require("outlook.preview")

      preview.open_event({ subject = "Offsite", all_day = true, start = "2026-08-24", body = "" })

      local lines = vim.api.nvim_buf_get_lines(0, 0, -1, false)
      assert.equals("Date      : 2026-08-24 (all day)", lines[2])
    end)

    it("gm opens the first recognized meeting-provider URL in the body via vim.ui.open", function()
      local opened
      vim.ui.open = function(url)
        opened = url
        return {}
      end

      local preview = require("outlook.preview")
      preview.open_event({
        subject = "Team sync",
        body = "Join: https://example.com/not-a-meeting-link\n"
          .. "Join Microsoft Teams Meeting https://teams.microsoft.com/l/meetup-join/abc123",
      })

      vim.api.nvim_feedkeys(vim.api.nvim_replace_termcodes("gm", true, false, true), "x", false)

      assert.equals("https://teams.microsoft.com/l/meetup-join/abc123", opened)
    end)

    it("gm notifies when no URL is found in the body", function()
      local notified = false
      package.loaded["outlook.notify"] = {
        info = function()
          notified = true
        end,
        error = function() end,
      }

      local preview = require("outlook.preview")
      preview.open_event({ subject = "Team sync", body = "No links here." })

      vim.api.nvim_feedkeys(vim.api.nvim_replace_termcodes("gm", true, false, true), "x", false)

      assert.is_true(notified)
    end)
  end)

  describe("open (mail)", function()
    it("renders subject/from/date metadata and the body", function()
      local preview = require("outlook.preview")

      preview.open({ subject = "Hello", from = "Bob", received = "2026-08-24 09:00", body = "Hi there" })

      local lines = vim.api.nvim_buf_get_lines(0, 0, -1, false)
      assert.equals("Subject : Hello", lines[1])
      assert.equals("From    : Bob", lines[2])
      assert.equals("Date    : 2026-08-24 09:00", lines[3])
      assert.equals("mail", vim.bo.filetype)
    end)
  end)
end)
