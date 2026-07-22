-- Minimal init for running the test suite headlessly with plenary.nvim.
-- See docs/DESIGN.md and tests/README.md for how to run these.

vim.opt.runtimepath:append(vim.fn.getcwd())

local plenary_dir = vim.fn.stdpath("data") .. "/site/pack/vendor/start/plenary.nvim"
if vim.fn.isdirectory(plenary_dir) == 0 then
  vim.fn.system({
    "git",
    "clone",
    "--depth=1",
    "https://github.com/nvim-lua/plenary.nvim",
    plenary_dir,
  })
end
vim.opt.runtimepath:append(plenary_dir)
vim.cmd("runtime! plugin/plenary.vim")
