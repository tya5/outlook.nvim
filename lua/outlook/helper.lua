-- Manages the long-lived PowerShell COM helper process and speaks the
-- newline-delimited JSON protocol described in docs/DESIGN.md (section 3).
--
-- All public entry points are non-blocking: starting the process and
-- sending/receiving requests happen over libuv via jobstart, so callers
-- never block the UI thread waiting on Outlook COM latency.

local notify = require("outlook.notify")
local config = require("outlook.config")

local M = {}

local plugin_root = vim.fn.fnamemodify(debug.getinfo(1, "S").source:sub(2), ":p:h:h:h")
local helper_script = plugin_root .. "/helper/outlook-helper.ps1"

local state = {
  job_id = nil,
  starting = false,
  next_id = 0,
  pending = {}, -- id -> { callback = fun(ok, result_or_error), timer = uv_timer }
  stdout_buf = "",
}

local function stop_timer(entry)
  if entry.timer then
    pcall(function()
      entry.timer:stop()
      entry.timer:close()
    end)
  end
end

--- Look up and remove a pending request, stopping its timeout timer, then
--- invoke its callback. Safe to call at most once per id (no-op if the id
--- is unknown, e.g. already resolved or already timed out).
local function resolve(id, ok, result_or_err)
  local entry = state.pending[id]
  if not entry then
    return
  end
  state.pending[id] = nil
  stop_timer(entry)
  entry.callback(ok, result_or_err)
end

local function reject_all(reason)
  local pending = state.pending
  state.pending = {}
  for _, entry in pairs(pending) do
    stop_timer(entry)
    entry.callback(false, { code = "HELPER_EXITED", message = reason })
  end
end

local function handle_line(line)
  if line == "" then
    return
  end
  local ok, msg = pcall(vim.json.decode, line)
  if not ok or type(msg) ~= "table" then
    notify.error("helperからの応答をパースできませんでした: " .. line)
    return
  end
  if msg.event then
    -- Reserved for future push notifications (see docs/DESIGN.md 3.1).
    -- No event types are emitted in v1 (polling only).
    return
  end
  if msg.ok then
    resolve(msg.id, true, msg.result)
  else
    resolve(msg.id, false, msg.error)
  end
end

local function on_stdout(_, data)
  if not data then
    return
  end
  state.stdout_buf = state.stdout_buf .. table.concat(data, "\n")
  while true do
    local nl = state.stdout_buf:find("\n")
    if not nl then
      break
    end
    local line = state.stdout_buf:sub(1, nl - 1)
    state.stdout_buf = state.stdout_buf:sub(nl + 1)
    handle_line(vim.trim(line))
  end
end

local function on_stderr(_, data)
  if not data then
    return
  end
  local text = vim.trim(table.concat(data, "\n"))
  if text ~= "" then
    notify.error("helper stderr: " .. text)
  end
end

local function on_exit(_, code)
  state.job_id = nil
  state.starting = false
  if code ~= 0 then
    notify.error(("helperプロセスが終了しました (exit code %d)"):format(code))
  end
  reject_all("helper process exited")
end

function M.is_running()
  return state.job_id ~= nil
end

--- Start the helper process in the background if it isn't running yet.
--- Safe to call repeatedly (no-op once started); never blocks.
function M.start()
  if M.is_running() or state.starting then
    return
  end

  if vim.fn.has("win32") == 0 then
    notify.error("outlook.nvim はWindows専用です")
    return
  end
  if vim.fn.executable("powershell.exe") == 0 then
    notify.error("powershell.exe が見つかりません")
    return
  end
  if vim.fn.filereadable(helper_script) == 0 then
    notify.error("helperスクリプトが見つかりません: " .. helper_script)
    return
  end

  state.starting = true
  state.stdout_buf = ""
  state.job_id = vim.fn.jobstart({
    "powershell.exe",
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    helper_script,
  }, {
    rpc = false,
    stdout_buffered = false,
    on_stdout = on_stdout,
    on_stderr = on_stderr,
    on_exit = on_exit,
  })

  if state.job_id <= 0 then
    notify.error("helperプロセスの起動に失敗しました")
    state.job_id = nil
  end
  state.starting = false
end

--- Fire-and-forget warm start; identical to start() but named for intent
--- at call sites (see config.lua's `prewarm` option).
M.prewarm = M.start

function M.stop()
  if state.job_id then
    vim.fn.jobstop(state.job_id)
    state.job_id = nil
  end
  reject_all("helper stopped")
end

function M.restart()
  M.stop()
  M.start()
end

--- @param method string
--- @param params table?
--- @param callback fun(ok: boolean, result_or_error: table)
function M.request(method, params, callback)
  M.start()
  if not M.is_running() then
    callback(false, { code = "HELPER_NOT_RUNNING", message = "helperプロセスを起動できませんでした" })
    return
  end

  state.next_id = state.next_id + 1
  local id = state.next_id

  local ok, line = pcall(vim.json.encode, { id = id, method = method, params = params or vim.empty_dict() })
  if not ok then
    callback(false, { code = "ENCODE_ERROR", message = line })
    return
  end

  local timeout_ms = config.options.request_timeout_ms
  local timer = nil
  if timeout_ms and timeout_ms > 0 then
    timer = vim.defer_fn(function()
      resolve(id, false, {
        code = "TIMEOUT",
        message = ("helperの応答がありませんでした (%dms)"):format(timeout_ms),
      })
    end, timeout_ms)
  end
  state.pending[id] = { callback = callback, timer = timer }

  vim.fn.chansend(state.job_id, line .. "\n")
end

return M
