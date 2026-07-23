-- Config module, following the folke/lazy.nvim ecosystem convention
-- (lazy.nvim, snacks.nvim, noice.nvim, ...): a `defaults` table, a merged
-- `options` table, and an `extend()` used by setup().

local M = {}

---@class outlook.Config
M.defaults = {
  -- Register the default <leader>m keymaps (see keymaps.lua). Set to
  -- false if you'd rather wire your own via the `keys` field of your
  -- lazy.nvim plugin spec.
  keys = true,
  -- Spawn the PowerShell helper (and connect to Outlook) as soon as
  -- setup() runs, instead of waiting for the first request. Only useful
  -- if your lazy.nvim spec loads the plugin eagerly (e.g. `event =
  -- "VeryLazy"`) rather than on `cmd`/`keys` — in that case it trades a
  -- small background cost at startup for a near-instant first open.
  -- No effect if setup() itself only runs on first command/keypress.
  prewarm = false,
  -- How long a list_messages/search_messages result is reused for
  -- identical requests before hitting the helper again (milliseconds).
  cache_ttl_ms = 15000,
  -- How long to wait for a helper response before giving up on a
  -- request (milliseconds). Guards against a wedged Outlook COM call
  -- (or a malformed response line the helper silently drops) leaving a
  -- request pending forever.
  request_timeout_ms = 30000,
  -- Where the full message view (opened via the picker's <CR>) shows
  -- up: "buffer" (a normal listed buffer in the current window — the
  -- same footing as any other open file) or "float" (a floating
  -- window, snacks.win-backed when available). See lua/outlook/preview.lua.
  message_window = "buffer",
}

---@type outlook.Config
M.options = vim.deepcopy(M.defaults)

---@param opts? outlook.Config
function M.extend(opts)
  M.options = vim.tbl_deep_extend("force", vim.deepcopy(M.defaults), opts or {})
  return M.options
end

return M
