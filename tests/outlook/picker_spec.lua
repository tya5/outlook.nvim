-- Tests for the latency-mitigation logic in lua/outlook/picker.lua:
-- short-lived result caching and coalescing of concurrent identical
-- requests. outlook.helper is replaced with a fake (package.loaded
-- injection) so no real IPC/process is involved.

describe("outlook.picker (cache/dedupe)", function()
  local calls
  local orig_ui_select

  before_each(function()
    calls = {}

    package.loaded["outlook.picker"] = nil
    package.loaded["outlook.helper"] = nil
    package.loaded["outlook.config"] = nil
    package.loaded["outlook.notify"] = nil
    package.loaded["outlook.preview"] = nil

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

    orig_ui_select = vim.ui.select
    vim.ui.select = function(_, _, on_choice)
      on_choice(nil, nil)
    end
  end)

  after_each(function()
    -- picker.list()'s resolution path always finishes with a
    -- vim.schedule()'d M.show() call (even on a cache hit); drain the
    -- event loop before restoring vim.ui.select so a stray scheduled
    -- callback never fires later against a different test's stubs.
    vim.wait(50)
    vim.ui.select = orig_ui_select
  end)

  it("reuses a cached result for an identical request within the TTL", function()
    local picker = require("outlook.picker")

    picker.list({ folder = "inbox" })
    assert.equals(1, #calls)
    calls[1].cb(true, { items = {} })

    picker.list({ folder = "inbox" })
    assert.equals(1, #calls) -- served from cache: no second helper.request
  end)

  it("bypasses the cache when force = true", function()
    local picker = require("outlook.picker")

    picker.list({ folder = "inbox" })
    calls[1].cb(true, { items = {} })

    picker.list({ folder = "inbox", force = true })
    assert.equals(2, #calls)
  end)

  it("coalesces concurrent identical requests into a single helper call", function()
    local picker = require("outlook.picker")

    picker.list({ folder = "inbox" })
    picker.list({ folder = "inbox" })
    assert.equals(1, #calls) -- second call joined the in-flight request instead of firing its own

    calls[1].cb(true, { items = {} })
  end)

  it("treats different folders as separate cache entries", function()
    local picker = require("outlook.picker")

    picker.list({ folder = "inbox" })
    calls[1].cb(true, { items = {} })

    picker.list({ folder = "inbox", unread_only = true })
    assert.equals(2, #calls)
  end)
end)
