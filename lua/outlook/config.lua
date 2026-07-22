-- Config module, following the folke/lazy.nvim ecosystem convention
-- (lazy.nvim, snacks.nvim, noice.nvim, ...): a `defaults` table, a merged
-- `options` table, and an `extend()` used by setup().

local M = {}

---@class outlook.Config
M.defaults = {}

---@type outlook.Config
M.options = vim.deepcopy(M.defaults)

---@param opts? outlook.Config
function M.extend(opts)
  M.options = vim.tbl_deep_extend("force", vim.deepcopy(M.defaults), opts or {})
  return M.options
end

return M
