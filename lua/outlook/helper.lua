-- Manages the long-lived PowerShell COM helper process and speaks the
-- newline-delimited JSON protocol described in docs/DESIGN.md (section 3).
--
-- All public entry points are non-blocking: starting the process and
-- sending/receiving requests happen over libuv via jobstart, so callers
-- never block the UI thread waiting on Outlook COM latency.

local notify = require("outlook.notify")

local M = {}

local plugin_root = vim.fn.fnamemodify(debug.getinfo(1, "S").source:sub(2), ":p:h:h:h")
local helper_script = plugin_root .. "/helper/outlook-helper.ps1"

local state = {
  job_id = nil,
  starting = false,
  next_id = 0,
  pending = {}, -- id -> callback(ok, result_or_error)
  stdout_buf = "",
}

local function reject_all(reason)
  local pending = state.pending
  state.pending = {}
  for _, cb in pairs(pending) do
    cb(false, { code = "HELPER_EXITED", message = reason })
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
  local cb = state.pending[msg.id]
  if not cb then
    return
  end
  state.pending[msg.id] = nil
  if msg.ok then
    cb(true, msg.result)
  else
    cb(false, msg.error)
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
  state.pending[id] = callback

  local ok, line = pcall(vim.json.encode, { id = id, method = method, params = params or vim.empty_dict() })
  if not ok then
    state.pending[id] = nil
    callback(false, { code = "ENCODE_ERROR", message = line })
    return
  end
  vim.fn.chansend(state.job_id, line .. "\n")
end

return M
