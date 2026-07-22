#Requires -Version 5.1
<#
  outlook.nvim helper process.

  Long-lived child process of the Neovim plugin. Talks Outlook COM
  (New-Object / GetActiveObject against Outlook.Application) on one side
  and newline-delimited JSON over stdin/stdout on the other.

  Protocol and method list: see docs/DESIGN.md, section 3 ("IPCプロトコル").

  Design notes (see docs/DESIGN.md for the full rationale):
  - Windows PowerShell 5.1's default console host apartment is STA, which
    is what Outlook COM automation requires; this must be run via
    powershell.exe (not pwsh, which defaults to MTA).
  - We never auto-launch Outlook: GetActiveObject only attaches to an
    already-running instance, so a missing Outlook process surfaces as a
    clear OUTLOOK_NOT_RUNNING error instead of silently spawning a new
    session mid-request.
  - The Outlook.Application/Namespace handles are cached at script scope
    for the life of the process and re-resolved lazily if a call fails,
    so an Outlook restart is recovered from on the next request rather
    than requiring a helper restart.
#>

# [Console]::OutputEncoding/InputEncoding setters throw IOException when
# stdio isn't a real console (exactly our case: Neovim's jobstart spawns
# us with piped stdio), which would kill this process before it ever
# writes a line. Read/write UTF-8 explicitly via our own stream
# wrappers instead of depending on the Console class' encoding at all.
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:StdIn = New-Object System.IO.StreamReader([Console]::OpenStandardInput(), $Utf8NoBom)
$script:StdOut = New-Object System.IO.StreamWriter([Console]::OpenStandardOutput(), $Utf8NoBom)
$script:StdOut.AutoFlush = $true

$script:OutlookApp = $null
$script:Namespace = $null

# Outlook well-known folder constants (OlDefaultFolders).
$script:FolderMap = @{
  inbox  = 6  # olFolderInbox
  sent   = 5  # olFolderSentMail
  drafts = 16 # olFolderDrafts
}
$OL_MAIL_ITEM = 43 # olMail (MailItem.Class)

function Write-Response {
  param($Object)
  $json = $Object | ConvertTo-Json -Depth 10 -Compress
  $script:StdOut.WriteLine($json)
}

function Send-Ok {
  param($Id, $Result)
  Write-Response @{ id = $Id; ok = $true; result = $Result }
}

function Send-Err {
  param($Id, $Code, $Message)
  Write-Response @{ id = $Id; ok = $false; error = @{ code = $Code; message = $Message } }
}

function Connect-Outlook {
  # Returns $true/$false; never launches Outlook.
  if ($script:Namespace) {
    return $true
  }
  try {
    $script:OutlookApp = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    $script:Namespace = $script:OutlookApp.GetNamespace("MAPI")
    return $true
  } catch {
    $script:OutlookApp = $null
    $script:Namespace = $null
    return $false
  }
}

function Get-FolderByName {
  param([string]$Name)
  if (-not (Connect-Outlook)) {
    return $null
  }
  if (-not $Name) {
    $Name = "inbox"
  }
  $olFolder = $script:FolderMap[$Name.ToLower()]
  if (-not $olFolder) {
    $olFolder = $script:FolderMap["inbox"]
  }
  return $script:Namespace.GetDefaultFolder($olFolder)
}

function ConvertTo-MessageSummary {
  # Deliberately does not touch $Item.Body or $Item.Parent: Body is slow
  # to fetch per-row across a whole folder listing and is one of the
  # properties Object Model Guard may prompt on; Parent.StoreID resolves
  # a COM object per item instead of once per folder. Callers pass the
  # already-resolved folder StoreID in instead. Full body is fetched
  # on demand by get_message when a single message is opened.
  param($Item, [string]$StoreId)
  return @{
    entry_id = $Item.EntryID
    store_id = $StoreId
    subject  = $Item.Subject
    from     = $Item.SenderName
    received = $Item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
    unread   = [bool]$Item.UnRead
  }
}

function Invoke-ListFolders {
  if (-not (Connect-Outlook)) {
    return $null
  }
  $items = New-Object System.Collections.Generic.List[object]
  $root = $script:Namespace.DefaultStore.GetRootFolder()
  foreach ($f in $root.Folders) {
    $items.Add(@{ name = $f.Name; path = $f.Name }) | Out-Null
    foreach ($sub in $f.Folders) {
      $items.Add(@{ name = $sub.Name; path = "$($f.Name)/$($sub.Name)" }) | Out-Null
    }
  }
  return @{ items = $items }
}

function Invoke-ListMessages {
  param($Params)
  if (-not (Connect-Outlook)) {
    return $null
  }

  $folderName = if ($Params.folder) { $Params.folder } else { "inbox" }
  $limit = if ($Params.limit) { [int]$Params.limit } else { 50 }
  $unreadOnly = [bool]$Params.unread_only

  $folder = Get-FolderByName -Name $folderName
  if (-not $folder) {
    return $null
  }

  $storeId = $folder.StoreID
  $items = $folder.Items
  $items.Sort("[ReceivedTime]", $true)
  if ($unreadOnly) {
    $items = $items.Restrict("[UnRead] = True")
  }

  $out = New-Object System.Collections.Generic.List[object]
  $count = 0
  foreach ($it in $items) {
    if ($count -ge $limit) {
      break
    }
    if ($it.Class -ne $OL_MAIL_ITEM) {
      continue
    }
    $out.Add((ConvertTo-MessageSummary -Item $it -StoreId $storeId)) | Out-Null
    $count++
  }
  return @{ items = $out }
}

function Invoke-GetMessage {
  param($Params)
  if (-not (Connect-Outlook)) {
    return $null
  }
  $item = $script:Namespace.GetItemFromID($Params.entry_id, $Params.store_id)
  if (-not $item) {
    return $null
  }
  return @{
    subject  = $item.Subject
    from     = $item.SenderName
    received = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
    body     = $item.Body
  }
}

function Invoke-SetRead {
  param($Params, [bool]$Unread)
  if (-not (Connect-Outlook)) {
    return $null
  }
  $item = $script:Namespace.GetItemFromID($Params.entry_id, $Params.store_id)
  if (-not $item) {
    return $null
  }
  $item.UnRead = $Unread
  $item.Save()
  return @{ entry_id = $item.EntryID; unread = [bool]$item.UnRead }
}

function Invoke-SearchMessages {
  param($Params)
  if (-not (Connect-Outlook)) {
    return $null
  }
  $folderName = if ($Params.folder) { $Params.folder } else { "inbox" }
  $limit = if ($Params.limit) { [int]$Params.limit } else { 50 }
  $folder = Get-FolderByName -Name $folderName
  if (-not $folder) {
    return $null
  }

  $escaped = $Params.query -replace "'", "''"
  $filter = "@SQL=" + '"urn:schemas:httpmail:subject" LIKE ' + "'%$escaped%'" `
    + " OR " + '"urn:schemas:httpmail:fromemail" LIKE ' + "'%$escaped%'"

  $storeId = $folder.StoreID
  $items = $folder.Items
  $items.Sort("[ReceivedTime]", $true)
  $matched = $items.Restrict($filter)

  $out = New-Object System.Collections.Generic.List[object]
  $count = 0
  foreach ($it in $matched) {
    if ($count -ge $limit) {
      break
    }
    if ($it.Class -ne $OL_MAIL_ITEM) {
      continue
    }
    $out.Add((ConvertTo-MessageSummary -Item $it -StoreId $storeId)) | Out-Null
    $count++
  }
  return @{ items = $out }
}

function Invoke-Method {
  param([string]$Method, $Params)
  switch ($Method) {
    "ping"            { return @{ pong = $true; outlook_connected = (Connect-Outlook) } }
    "list_folders"    { return Invoke-ListFolders }
    "list_messages"   { return Invoke-ListMessages -Params $Params }
    "get_message"     { return Invoke-GetMessage -Params $Params }
    "mark_read"       { return Invoke-SetRead -Params $Params -Unread $false }
    "mark_unread"     { return Invoke-SetRead -Params $Params -Unread $true }
    "search_messages" { return Invoke-SearchMessages -Params $Params }
    default           { throw "unknown method: $Method" }
  }
}

# Main loop: one JSON request per line in, one JSON response per line out.
while ($true) {
  $line = $script:StdIn.ReadLine()
  if ($null -eq $line) {
    break # stdin closed (Neovim stopped the job) -> exit
  }
  $line = $line.Trim()
  if ($line -eq "") {
    continue
  }

  try {
    $req = $line | ConvertFrom-Json
  } catch {
    continue # malformed request line; nothing we can address a response to
  }

  $id = $req.id
  try {
    $result = Invoke-Method -Method $req.method -Params $req.params
    if ($null -eq $result) {
      # Every Invoke-* returns $null only when Connect-Outlook failed or
      # GetItemFromID found nothing; both collapse to this generic code
      # in v1 rather than threading a distinct not-found case through.
      Send-Err -Id $id -Code "OUTLOOK_NOT_RUNNING" -Message "Outlook に接続できませんでした (未起動か、対象が見つかりません)"
    } else {
      Send-Ok -Id $id -Result $result
    }
  } catch {
    Send-Err -Id $id -Code "INTERNAL_ERROR" -Message $_.Exception.Message
  }
}
