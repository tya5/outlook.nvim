-- Thin notify wrapper: snacks.nvim if present, vim.notify otherwise.
local M = {}

local function backend()
  local ok, snacks = pcall(require, "snacks")
  if ok and snacks.notify then
    return snacks.notify
  end
  return nil
end

local function message_of(err)
  if type(err) == "table" then
    return err.message or err.code or vim.inspect(err)
  end
  return tostring(err)
end

function M.error(err)
  local msg = message_of(err)
  local notify = backend()
  if notify then
    notify.error(msg, { title = "outlook.nvim" })
  else
    vim.notify(msg, vim.log.levels.ERROR, { title = "outlook.nvim" })
  end
end

function M.info(msg, opts)
  local notify = backend()
  if notify then
    notify.info(msg, vim.tbl_extend("force", { title = "outlook.nvim" }, opts or {}))
  else
    vim.notify(msg, vim.log.levels.INFO, { title = "outlook.nvim" })
  end
end

return M
