-- outlook.nvim entrypoint
-- Architecture and module responsibilities: see docs/DESIGN.md

local M = {}

---@class outlook.Config
local defaults = {}

M.config = vim.deepcopy(defaults)

---@param opts? outlook.Config
function M.setup(opts)
  M.config = vim.tbl_deep_extend("force", vim.deepcopy(defaults), opts or {})
end

return M
