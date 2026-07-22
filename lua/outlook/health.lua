-- :checkhealth outlook
-- See https://neovim.io/doc/user/health.html

local M = {}

function M.check()
  vim.health.start("outlook.nvim")

  if vim.fn.has("win32") == 1 then
    vim.health.ok("Running on Windows")
  else
    vim.health.error("outlook.nvim requires Windows (Outlook COM automation is Windows-only)")
  end

  if vim.fn.executable("powershell.exe") == 1 then
    vim.health.ok("powershell.exe found on PATH")
  else
    vim.health.error("powershell.exe not found on PATH", {
      "outlook.nvim's helper process requires Windows PowerShell 5.1 (powershell.exe)",
    })
  end

  if vim.fn.has("nvim-0.10") == 1 then
    vim.health.ok("Neovim >= 0.10")
  else
    vim.health.warn("Neovim >= 0.10 recommended (vim.system, vim.json)")
  end

  if pcall(require, "snacks") then
    vim.health.ok("snacks.nvim detected (picker/win/notify UI will be used)")
  else
    vim.health.info("snacks.nvim not found — falling back to vim.ui.select for lists")
  end
end

return M
