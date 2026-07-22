-- Tests for the Neovim <-> helper-process IPC layer (lua/outlook/helper.lua).
--
-- There is no real PowerShell/Outlook in this environment, so `vim.fn.jobstart`
-- and `vim.fn.chansend` are stubbed to simulate the helper process: the fake
-- jobstart captures the on_stdout/on_exit callbacks the module registers, and
-- tests drive them directly to simulate bytes arriving from the (fake) child
-- process. This covers the framing/dispatch logic that real PowerShell
-- integration can't be exercised for here.

local function fresh_helper()
  package.loaded["outlook.helper"] = nil
  package.loaded["outlook.notify"] = nil
  return require("outlook.helper")
end

describe("outlook.helper", function()
  local orig = {}

  before_each(function()
    orig.jobstart = vim.fn.jobstart
    orig.chansend = vim.fn.chansend
    orig.executable = vim.fn.executable
    orig.has = vim.fn.has
    orig.filereadable = vim.fn.filereadable

    -- Reset config to pure defaults before every test so a test that
    -- overrides e.g. request_timeout_ms can't leak into later tests
    -- (in this file or, since PlenaryBustedDirectory may run multiple
    -- spec files in one process, other spec files too).
    package.loaded["outlook.config"] = nil
    require("outlook.config")
  end)

  after_each(function()
    vim.fn.jobstart = orig.jobstart
    vim.fn.chansend = orig.chansend
    vim.fn.executable = orig.executable
    vim.fn.has = orig.has
    vim.fn.filereadable = orig.filereadable
  end)

  local function stub_env()
    local real_has = vim.fn.has
    vim.fn.has = function(feat)
      if feat == "win32" then
        return 1
      end
      return real_has(feat)
    end
    vim.fn.executable = function()
      return 1
    end
    vim.fn.filereadable = function()
      return 1
    end
  end

  it("sends a request and resolves the matching response", function()
    stub_env()
    local sent = {}
    local on_stdout_cb
    vim.fn.jobstart = function(_, opts)
      on_stdout_cb = opts.on_stdout
      return 1
    end
    vim.fn.chansend = function(_, data)
      table.insert(sent, data)
      return #data
    end

    local helper = fresh_helper()
    local got_ok, got_result
    helper.request("ping", {}, function(ok, result)
      got_ok, got_result = ok, result
    end)

    assert.equals(1, #sent)
    local req = vim.json.decode(vim.trim(sent[1]))
    assert.equals("ping", req.method)

    on_stdout_cb(1, { vim.json.encode({ id = req.id, ok = true, result = { pong = true } }), "" })

    assert.is_true(got_ok)
    assert.is_true(got_result.pong)
  end)

  it("buffers a response split across multiple stdout chunks", function()
    stub_env()
    local on_stdout_cb
    vim.fn.jobstart = function(_, opts)
      on_stdout_cb = opts.on_stdout
      return 1
    end
    vim.fn.chansend = function()
      return 0
    end

    local helper = fresh_helper()
    local got_ok, got_result
    helper.request("list_messages", {}, function(ok, result)
      got_ok, got_result = ok, result
    end)

    local full = vim.json.encode({ id = 1, ok = true, result = { items = {} } })
    local mid = math.floor(#full / 2)

    on_stdout_cb(1, { full:sub(1, mid) })
    assert.is_nil(got_ok) -- line not complete yet: no callback fired

    on_stdout_cb(1, { full:sub(mid + 1), "" })
    assert.is_true(got_ok)
    assert.same({}, got_result.items)
  end)

  it("rejects pending requests when the helper process exits", function()
    stub_env()
    local on_exit_cb
    vim.fn.jobstart = function(_, opts)
      on_exit_cb = opts.on_exit
      return 1
    end
    vim.fn.chansend = function()
      return 0
    end

    local helper = fresh_helper()
    local got_ok, got_err
    helper.request("ping", {}, function(ok, err)
      got_ok, got_err = ok, err
    end)

    on_exit_cb(1, 1)

    assert.is_false(got_ok)
    assert.equals("HELPER_EXITED", got_err.code)
  end)

  it("rejects a request with code=TIMEOUT if no response arrives in time", function()
    stub_env()
    vim.fn.jobstart = function()
      return 1
    end
    vim.fn.chansend = function()
      return 0
    end

    require("outlook.config").extend({ request_timeout_ms = 20 })

    local helper = fresh_helper()
    local got_ok, got_err
    helper.request("ping", {}, function(ok, err)
      got_ok, got_err = ok, err
    end)

    vim.wait(200, function()
      return got_ok ~= nil
    end, 10)

    assert.is_false(got_ok)
    assert.equals("TIMEOUT", got_err.code)
  end)

  it("does not fire a stale timeout after the request already resolved", function()
    stub_env()
    local on_stdout_cb
    vim.fn.jobstart = function(_, opts)
      on_stdout_cb = opts.on_stdout
      return 1
    end
    vim.fn.chansend = function()
      return 0
    end

    require("outlook.config").extend({ request_timeout_ms = 30 })

    local helper = fresh_helper()
    local resolutions = 0
    helper.request("ping", {}, function()
      resolutions = resolutions + 1
    end)

    on_stdout_cb(1, { vim.json.encode({ id = 1, ok = true, result = { pong = true } }), "" })

    -- Wait past the (now-cancelled) timeout window; the callback must
    -- not fire a second time with a TIMEOUT error.
    vim.wait(60)

    assert.equals(1, resolutions)
  end)

  it("fails fast without spawning when powershell.exe is missing", function()
    stub_env()
    vim.fn.executable = function()
      return 0
    end
    local jobstart_called = false
    vim.fn.jobstart = function()
      jobstart_called = true
      return 1
    end

    local helper = fresh_helper()
    local got_ok, got_err
    helper.request("ping", {}, function(ok, err)
      got_ok, got_err = ok, err
    end)

    assert.is_false(jobstart_called)
    assert.is_false(got_ok)
    assert.is_not_nil(got_err)
  end)
end)
