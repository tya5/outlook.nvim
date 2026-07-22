-- outlook.nvim entrypoint
-- Architecture and module responsibilities: see docs/DESIGN.md

local M = {}

---@param opts? outlook.Config
function M.setup(opts)
  require("outlook.config").extend(opts)
end

return M
