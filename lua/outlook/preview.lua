-- Full message view opened from the picker's <CR> (M.open_message).
--
-- Default is a normal, listed buffer opened in the current window
-- (opts.message_window = "buffer"): not a floating window, not a forced
-- split — it behaves like opening any other file, so it sits alongside
-- your other buffers and is reachable with normal buffer navigation
-- (<C-^>, :bnext/:bprevious, a bufferline, etc.). Set opts.message_window
-- = "float" to get a floating window (snacks.win-backed when available)
-- instead.

local M = {}

local function has_snacks()
  return pcall(require, "snacks")
end

local function lines_for(message)
  local header = {
    ("Subject : %s"):format(message.subject or ""),
    ("From    : %s"):format(message.from or ""),
    ("Date    : %s"):format(message.received or ""),
  }
  if message.flag_status and message.flag_status ~= "none" then
    table.insert(header, ("Flag    : %s"):format(message.flag_status))
  end
  table.insert(header, string.rep("-", 40))
  local body = vim.split(message.body or "", "\n", { plain = true })
  return vim.list_extend(header, body)
end

local function set_readonly(buf)
  vim.bo[buf].modifiable = false
  vim.bo[buf].buftype = "nofile"
  vim.bo[buf].filetype = "mail"
end

--- Open in the current window as a normal, listed buffer — the same
--- footing as any other open file, not a special split/float.
local function open_buffer(lines, message)
  local buf = vim.api.nvim_create_buf(true, false)
  vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)
  set_readonly(buf)
  vim.bo[buf].bufhidden = "hide" -- stays open in the buffer list, unlike the floating/scratch views

  -- Buffer names are resolved like filenames, so strip path separators
  -- out of the subject first; names must also be unique (two messages
  -- with the same sanitized subject would otherwise fail
  -- nvim_buf_set_name), and falling back to an unnamed buffer is
  -- harmless, so this is best-effort.
  local safe_subject = (message.subject or "message"):gsub("[/\\]", "_")
  pcall(vim.api.nvim_buf_set_name, buf, ("outlook/%s"):format(safe_subject))

  vim.api.nvim_win_set_buf(0, buf)
  vim.wo[0].wrap = true
  return buf
end

local function open_float(lines, message)
  if has_snacks() then
    local Snacks = require("snacks")
    local win = Snacks.win({
      text = lines,
      width = 0.8,
      height = 0.8,
      border = "rounded",
      title = message.subject or "Outlook",
      wo = { wrap = true },
    })
    set_readonly(win.buf)
    vim.bo[win.buf].bufhidden = "wipe"
    win:map("q", function()
      win:close()
    end, { desc = "Close" })
    return win
  end

  local buf = vim.api.nvim_create_buf(false, true)
  vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)
  set_readonly(buf)
  vim.bo[buf].bufhidden = "wipe"
  local width = math.floor(vim.o.columns * 0.8)
  local height = math.floor(vim.o.lines * 0.8)
  local win_id = vim.api.nvim_open_win(buf, true, {
    relative = "editor",
    width = width,
    height = height,
    row = math.floor((vim.o.lines - height) / 2),
    col = math.floor((vim.o.columns - width) / 2),
    border = "rounded",
    title = message.subject or "Outlook",
  })
  vim.wo[win_id].wrap = true
  vim.keymap.set("n", "q", "<cmd>close<cr>", { buffer = buf, silent = true, desc = "Close" })
  return win_id
end

--- @param message table result of the `get_message` request
function M.open(message)
  local lines = lines_for(message)

  if require("outlook.config").options.message_window == "float" then
    return open_float(lines, message)
  end
  return open_buffer(lines, message)
end

return M
