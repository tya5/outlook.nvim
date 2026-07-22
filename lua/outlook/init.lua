-- outlook.nvim entrypoint
-- Architecture and module responsibilities: see docs/DESIGN.md

local M = {}

---@param opts? outlook.Config
function M.setup(opts)
  local config = require("outlook.config").extend(opts)
  require("outlook.commands").setup()
  require("outlook.keymaps").setup()
  if config.prewarm then
    require("outlook.helper").prewarm()
  end
end

return M
