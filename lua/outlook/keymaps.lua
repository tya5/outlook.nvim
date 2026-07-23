local M = {}

function M.setup()
  if not require("outlook.config").options.keys then
    return
  end

  local function map(lhs, rhs, desc)
    vim.keymap.set("n", lhs, rhs, { desc = desc, silent = true })
  end

  map("<leader>mm", "<cmd>OutlookOpen<cr>", "Mail: open inbox")
  map("<leader>mu", "<cmd>OutlookUnread<cr>", "Mail: unread only")
  map("<leader>ms", "<cmd>OutlookSearch<cr>", "Mail: search")
  map("<leader>mc", "<cmd>OutlookCalendar<cr>", "Mail: calendar (requires almanac.nvim)")

  local ok, wk = pcall(require, "which-key")
  if ok then
    wk.add({ { "<leader>m", group = "mail" } })
  end
end

return M
