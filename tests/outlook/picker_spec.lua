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

  it("invalidates the list cache and updates the item in place after mark_read/mark_unread", function()
    local picker = require("outlook.picker")

    picker.list({ folder = "inbox" })
    calls[1].cb(true, { items = {} })
    assert.equals(1, #calls)

    -- Cache is warm: a repeat list() would normally be served from it.
    local item = { entry_id = "e1", store_id = "s1", unread = true, flag_status = "none" }
    picker.toggle_read(item)
    assert.equals(2, #calls) -- the mark_read request itself
    calls[2].cb(true, { entry_id = "e1", unread = false })

    -- The already-open picker's row must reflect the change immediately
    -- (via the mutated item + recomputed text), not only after the
    -- picker is closed and reopened against the invalidated cache.
    assert.is_false(item.unread)
    assert.is_nil(item.text:find("●", 1, true))

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

  it("toggles the flag (set_flag/clear_flag) and invalidates the list cache", function()
    local picker = require("outlook.picker")

    picker.list({ folder = "inbox" })
    calls[1].cb(true, { items = {} })
    assert.equals(1, #calls)

    local item = { entry_id = "e1", store_id = "s1", flag_status = "none" }
    picker.toggle_flag(item)
    assert.equals(2, #calls)
    assert.equals("set_flag", calls[2].method)
    calls[2].cb(true, { entry_id = "e1", flag_status = "flagged" })

    -- Same immediacy requirement as toggle_read: the open picker's row
    -- must show the flag without needing to reopen.
    assert.equals("flagged", item.flag_status)
    assert.is_not_nil(item.text:find("🚩", 1, true))

    picker.list({ folder = "inbox" })
    assert.equals(3, #calls) -- cache invalidated by the flag change

    picker.toggle_flag(item)
    assert.equals(4, #calls)
    assert.equals("clear_flag", calls[4].method)
    calls[4].cb(true, { entry_id = "e1", flag_status = "none" })

    assert.equals("none", item.flag_status)
    assert.is_nil(item.text:find("🚩", 1, true))
  end)

  it("open_message fetches the body, marks read, and invalidates the list cache", function()
    local picker = require("outlook.picker")

    picker.list({ folder = "inbox" })
    calls[1].cb(true, { items = {} })
    assert.equals(1, #calls)

    picker.open_message({ entry_id = "e1", store_id = "s1", unread = true })
    assert.equals(2, #calls) -- the get_message request
    assert.equals("get_message", calls[2].method)
    calls[2].cb(true, { subject = "hi", from = "a@b.com", received = "2026-01-01 00:00", body = "hello" })

    assert.equals(3, #calls) -- mark_read fired because the item was unread
    assert.equals("mark_read", calls[3].method)
    calls[3].cb(true, { entry_id = "e1", unread = false })

    picker.list({ folder = "inbox" })
    assert.equals(4, #calls) -- cache invalidated by the mark_read, not served stale
  end)

  it("open_message does not mark read when the item is already read", function()
    local picker = require("outlook.picker")

    picker.open_message({ entry_id = "e2", store_id = "s1", unread = false })
    assert.equals(1, #calls) -- only get_message
    calls[1].cb(true, { subject = "hi", from = "a@b.com", received = "2026-01-01 00:00", body = "hello" })

    assert.equals(1, #calls) -- no mark_read follow-up
  end)

  it("treats different folders as separate cache entries", function()
    local picker = require("outlook.picker")

    picker.list({ folder = "inbox" })
    calls[1].cb(true, { items = {} })

    picker.list({ folder = "inbox", unread_only = true })
    assert.equals(2, #calls)
  end)
end)

-- M.load_more() only runs on the snacks.picker path, so these tests
-- inject a fake `snacks` module (package.loaded["snacks"] = {...}) to
-- exercise it without a real snacks.nvim install (unavailable in this
-- environment — see docs/HANDOFF.md for what's confirmed vs. not).
describe("outlook.picker load_more (snacks path)", function()
  local calls
  local pick_calls

  local function msg(entry_id)
    return {
      entry_id = entry_id,
      subject = "s",
      from = "f",
      received = "t",
      unread = false,
      flag_status = "none",
    }
  end

  before_each(function()
    calls = {}
    pick_calls = {}

    package.loaded["outlook.picker"] = nil
    package.loaded["outlook.helper"] = nil
    package.loaded["outlook.config"] = nil
    package.loaded["outlook.notify"] = nil
    package.loaded["outlook.preview"] = nil
    package.loaded["snacks"] = nil

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

    package.loaded["snacks"] = {
      picker = {
        pick = function(opts)
          table.insert(pick_calls, opts)
        end,
      },
    }
  end)

  after_each(function()
    vim.wait(50)
    package.loaded["snacks"] = nil
  end)

  it("fetches a bigger page and grows the picker's items table in place", function()
    local picker = require("outlook.picker")

    picker.list({ folder = "inbox", limit = 50 })
    calls[1].cb(true, { items = { msg("e1") } })
    vim.wait(20) -- M.list()'s success path calls M.show() via vim.schedule

    assert.equals(1, #pick_calls)
    local items_ref = pick_calls[1].items
    assert.equals(1, #items_ref)

    local fake_picker = {
      refresh = function() end,
    }
    pick_calls[1].actions.load_more(fake_picker, nil)

    assert.equals(2, #calls)
    assert.equals("list_messages", calls[2].method)
    assert.equals(100, calls[2].params.limit) -- 50 + the load-more page size

    calls[2].cb(true, { items = { msg("e1"), msg("e2") } })
    vim.wait(20)

    assert.equals(2, #items_ref) -- same table object grew in place
    assert.equals("e2", items_ref[2].entry_id)
  end)

  it("does not fire overlapping load_more requests while one is in flight", function()
    local picker = require("outlook.picker")

    picker.list({ folder = "inbox", limit = 50 })
    calls[1].cb(true, { items = {} })
    vim.wait(20)

    local fake_picker = {
      refresh = function() end,
    }
    pick_calls[1].actions.load_more(fake_picker, nil)
    assert.equals(2, #calls)

    pick_calls[1].actions.load_more(fake_picker, nil) -- still loading: no-op
    assert.equals(2, #calls)

    calls[2].cb(true, { items = {} })
  end)

  it("supports load_more for search results too (method/params from :OutlookSearch)", function()
    local picker = require("outlook.picker")

    picker.show({ msg("e1") }, {
      title = "Outlook: Search: foo",
      method = "search_messages",
      params = { query = "foo", limit = 50 },
    })

    local fake_picker = {
      refresh = function() end,
    }
    pick_calls[1].actions.load_more(fake_picker, nil)

    assert.equals(1, #calls)
    assert.equals("search_messages", calls[1].method)
    assert.equals("foo", calls[1].params.query)
    assert.equals(100, calls[1].params.limit)
  end)
end)
