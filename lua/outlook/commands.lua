local M = {}

function M.setup()
  vim.api.nvim_create_user_command("OutlookOpen", function()
    require("outlook.picker").list({ title = "Outlook: Inbox" })
  end, { desc = "Open the Outlook inbox" })

  vim.api.nvim_create_user_command("OutlookRefresh", function()
    require("outlook.picker").list({ title = "Outlook: Inbox", force = true })
  end, { desc = "Refresh the Outlook inbox, bypassing the cache" })

  vim.api.nvim_create_user_command("OutlookUnread", function()
    require("outlook.picker").list({ title = "Outlook: Unread", unread_only = true })
  end, { desc = "Open unread Outlook messages" })

  vim.api.nvim_create_user_command("OutlookSearch", function()
    vim.ui.input({ prompt = "Outlook 検索 (件名/差出人): " }, function(query)
      if not query or query == "" then
        return
      end
      require("outlook.notify").info("Outlook: 検索中…")
      require("outlook.helper").request("search_messages", { query = query, limit = 50 }, function(ok, result)
        if not ok then
          return vim.schedule(function()
            require("outlook.notify").error(result)
          end)
        end
        vim.schedule(function()
          require("outlook.picker").show(result.items, { title = "Outlook: Search: " .. query })
        end)
      end)
    end)
  end, { desc = "Search Outlook messages" })

  vim.api.nvim_create_user_command("OutlookHelperRestart", function()
    require("outlook.helper").restart()
  end, { desc = "Restart the outlook.nvim PowerShell helper process" })
end

return M
