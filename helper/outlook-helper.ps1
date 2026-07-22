#Requires -Version 5.1
<#
  outlook.nvim helper process.

  Runs as a long-lived child process of the Neovim plugin, talking Outlook
  COM (New-Object -ComObject Outlook.Application) on one side and
  newline-delimited JSON over stdin/stdout on the other.

  Protocol and method list: see docs/DESIGN.md ("3. IPC プロトコル").
  Not implemented yet — this is a repository scaffold placeholder.
#>

# TODO: read one JSON request per line from stdin, dispatch by `method`,
# write one JSON response per line to stdout. Keep a single persistent
# Outlook.Application/Namespace handle for the lifetime of the process.
