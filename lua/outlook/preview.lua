-- Read-only floating window for a full message body.
-- Uses snacks.win when available, a plain nvim_open_win otherwise.

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
  vim.bo[buf].bufhidden = "wipe"
  vim.bo[buf].filetype = "mail"
end

--- @param message table result of the `get_message` request
function M.open(message)
  local lines = lines_for(message)

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
    win:map("q", function()
      win:close()
    end, { desc = "Close" })
    return win
  end

  local buf = vim.api.nvim_create_buf(false, true)
  vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)
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
  set_readonly(buf)
  vim.wo[win_id].wrap = true
  vim.keymap.set("n", "q", "<cmd>close<cr>", { buffer = buf, silent = true, desc = "Close" })
  return win_id
end

return M
