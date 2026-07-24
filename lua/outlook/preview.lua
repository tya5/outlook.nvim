-- Full-item detail views: the mail message view opened from the
-- picker's <CR> (M.open) and the calendar event view opened from
-- almanac.nvim's event_selected (M.open_event). Both share the same
-- generic buffer/float window plumbing below, just with different
-- metadata lines and (for events) an extra keymap.
--
-- Default is a normal, listed buffer opened in the current window
-- (opts.message_window = "buffer"): not a floating window, not a forced
-- split — it behaves like opening any other file, so it sits alongside
-- your other buffers and is reachable with normal buffer navigation
-- (<C-^>, :bnext/:bprevious, a bufferline, etc.). Set opts.message_window
-- = "float" to get a floating window (snacks.win-backed when available)
-- instead.

local M = {}

local function has_snacks()
  return pcall(require, "snacks")
end

local function lines_for_message(message)
  local header = {
    ("Subject : %s"):format(message.subject or ""),
    ("From    : %s"):format(message.from or ""),
    ("Date    : %s"):format(message.received or ""),
  }
  if message.flag_status and message.flag_status ~= "none" then
    table.insert(header, ("Flag    : %s"):format(message.flag_status))
  end
  table.insert(header, string.rep("-", 40))
  local body = vim.split(message.body or "", "\n", { plain = true })
  return vim.list_extend(header, body)
end

local function lines_for_event(event)
  local header = { ("Subject   : %s"):format(event.subject or "") }
  if event.organizer and event.organizer ~= "" then
    table.insert(header, ("Organizer : %s"):format(event.organizer))
  end
  if event.all_day then
    table.insert(header, ("Date      : %s (all day)"):format(event.start or ""))
  else
    table.insert(header, ("Start     : %s"):format(event.start or ""))
    table.insert(header, ("End       : %s"):format(event.stop or ""))
  end
  if event.location and event.location ~= "" then
    table.insert(header, ("Location  : %s"):format(event.location))
  end
  if event.busy then
    table.insert(header, ("Status    : %s"):format(event.busy))
  end
  table.insert(header, string.rep("-", 40))
  local body = vim.split(event.body or "", "\n", { plain = true })
  return vim.list_extend(header, body)
end

-- Classic Outlook COM automation has no reliable "join URL" property
-- for online meetings (unlike the Graph API's onlineMeeting.joinUrl) —
-- Teams/Zoom/etc. links are just plain text/hyperlinks inside the
-- appointment body — so scanning the body for a recognized
-- meeting-provider URL is the practical way to support a direct
-- "open the meeting link" action instead of making the user hunt for
-- it with the cursor and `gx`.
local MEETING_URL_HOSTS = {
  "teams%.microsoft%.com",
  "teams%.live%.com",
  "zoom%.us",
  "meet%.google%.com",
  "webex%.com",
  "gotomeeting%.com",
  "chime%.aws",
  "whereby%.com",
}

local function find_meeting_url(body)
  local urls = {}
  for url in (body or ""):gmatch("https?://[%w%-%._~:/?#%[%]@!$&'()*+,;=%%]+") do
    urls[#urls + 1] = url
  end
  for _, host_pat in ipairs(MEETING_URL_HOSTS) do
    for _, url in ipairs(urls) do
      if url:match(host_pat) then
        return url
      end
    end
  end
  return urls[1]
end

local function set_readonly(buf, filetype)
  vim.bo[buf].modifiable = false
  vim.bo[buf].buftype = "nofile"
  vim.bo[buf].filetype = filetype
end

local function apply_extra_keys(buf, extra_keys)
  for _, k in ipairs(extra_keys or {}) do
    vim.keymap.set("n", k[1], k[2], { buffer = buf, silent = true, desc = k.desc })
  end
end

--- Open in the current window as a normal, listed buffer — the same
--- footing as any other open file, not a special split/float.
--- @param opts { name: string?, filetype: string, extra_keys: table[]? }
local function open_buffer(lines, opts)
  local buf = vim.api.nvim_create_buf(true, false)
  vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)
  set_readonly(buf, opts.filetype)
  vim.bo[buf].bufhidden = "hide" -- stays open in the buffer list, unlike the floating/scratch views

  -- Buffer names are resolved like filenames, so strip path separators
  -- out of the name first; names must also be unique (two items with
  -- the same sanitized name would otherwise fail nvim_buf_set_name),
  -- and falling back to an unnamed buffer is harmless, so this is
  -- best-effort.
  local safe_name = (opts.name or "outlook"):gsub("[/\\]", "_")
  pcall(vim.api.nvim_buf_set_name, buf, ("outlook/%s"):format(safe_name))

  vim.api.nvim_win_set_buf(0, buf)
  vim.wo[0].wrap = true
  apply_extra_keys(buf, opts.extra_keys)
  return buf
end

--- @param opts { name: string?, filetype: string, extra_keys: table[]? }
local function open_float(lines, opts)
  if has_snacks() then
    local Snacks = require("snacks")
    local win = Snacks.win({
      text = lines,
      width = 0.8,
      height = 0.8,
      border = "rounded",
      title = opts.name or "Outlook",
      wo = { wrap = true },
    })
    set_readonly(win.buf, opts.filetype)
    vim.bo[win.buf].bufhidden = "wipe"
    win:map("q", function()
      win:close()
    end, { desc = "Close" })
    apply_extra_keys(win.buf, opts.extra_keys)
    return win
  end

  local buf = vim.api.nvim_create_buf(false, true)
  vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)
  set_readonly(buf, opts.filetype)
  vim.bo[buf].bufhidden = "wipe"
  local width = math.floor(vim.o.columns * 0.8)
  local height = math.floor(vim.o.lines * 0.8)
  local win_id = vim.api.nvim_open_win(buf, true, {
    relative = "editor",
    width = width,
    height = height,
    row = math.floor((vim.o.lines - height) / 2),
    col = math.floor((vim.o.columns - width) / 2),
    border = "rounded",
    title = opts.name or "Outlook",
  })
  vim.wo[win_id].wrap = true
  vim.keymap.set("n", "q", "<cmd>close<cr>", { buffer = buf, silent = true, desc = "Close" })
  apply_extra_keys(buf, opts.extra_keys)
  return win_id
end

--- @param message table result of the `get_message` request
function M.open(message)
  local lines = lines_for_message(message)
  local opts = { name = message.subject, filetype = "mail" }

  if require("outlook.config").options.message_window == "float" then
    return open_float(lines, opts)
  end
  return open_buffer(lines, opts)
end

--- @param event table result of the `get_appointment` request
function M.open_event(event)
  local lines = lines_for_event(event)
  local extra_keys = {
    {
      "gm",
      function()
        local url = find_meeting_url(event.body)
        if not url then
          require("outlook.notify").info("No meeting link found in this event")
          return
        end
        vim.ui.open(url)
      end,
      desc = "Open online meeting link",
    },
  }
  local opts = { name = event.subject, filetype = "outlook-event", extra_keys = extra_keys }

  if require("outlook.config").options.message_window == "float" then
    return open_float(lines, opts)
  end
  return open_buffer(lines, opts)
end

return M
