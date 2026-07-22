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

  it("clears the in-flight slot on error, so a later call re-hits the helper", function()
    local picker = require("outlook.picker")

    picker.list({ folder = "inbox" })
    assert.equals(1, #calls)
    calls[1].cb(false, { code = "TIMEOUT", message = "boom" })

    picker.list({ folder = "inbox" })
    assert.equals(2, #calls) -- previous in-flight entry must not still be occupying the key
  end)

  it("invalidates cached list results after a successful mark_read/mark_unread", function()
    local picker = require("outlook.picker")

    picker.list({ folder = "inbox" })
    calls[1].cb(true, { items = {} })
    assert.equals(1, #calls)

    -- Cache is warm: a repeat list() would normally be served from it.
    picker.toggle_read({ entry_id = "e1", store_id = "s1", unread = true })
    assert.equals(2, #calls) -- the mark_read request itself
    calls[2].cb(true, { entry_id = "e1", unread = false })

    picker.list({ folder = "inbox" })
    assert.equals(3, #calls) -- cache was invalidated by the mutation, not served stale
  end)

  it("only shows a loading notification on a real fetch, not on a cache hit", function()
    local info_calls = 0
    package.loaded["outlook.notify"] = {
      info = function()
        info_calls = info_calls + 1
      end,
      error = function() end,
    }
    package.loaded["outlook.picker"] = nil
    local picker = require("outlook.picker")

    picker.list({ folder = "inbox" }) -- cache miss: one loading notification
    assert.equals(1, info_calls)
    calls[1].cb(true, { items = {} })

    picker.list({ folder = "inbox" }) -- cache hit: no additional notification
    assert.equals(1, info_calls)

    picker.list({ folder = "inbox", force = true }) -- forced bypass: notifies again
    assert.equals(2, info_calls)
  end)

  it("treats different folders as separate cache entries", function()
    local picker = require("outlook.picker")

    picker.list({ folder = "inbox" })
    calls[1].cb(true, { items = {} })

    picker.list({ folder = "inbox", unread_only = true })
    assert.equals(2, #calls)
  end)
end)
