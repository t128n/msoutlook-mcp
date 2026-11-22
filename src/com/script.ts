/**
 * The persistent Windows PowerShell helper that drives classic Outlook through
 * COM. The host side is in `powershell.ts`.
 *
 * Protocol: one JSON request per line on stdin, one JSON answer per line on
 * stdout. Both directions are pure ASCII; non-ASCII characters are written as
 * \uXXXX so the console code page never matters.
 *
 *   {"id":1,"op":"ping"}
 *   {"id":2,"op":"listMail","folder":"inbox","limit":20}
 *   -> {"id":2,"ok":true,"result":{...}}
 *   -> {"id":2,"ok":false,"error":{"code":-2147221005,"message":"..."}}
 *
 * The pattern follows the helper in t128n/torbit PR #19 (MIT). The op set is
 * extended for mail, calendar, rules and contacts. The original Python project,
 * Astral0/outlook-com-mcp (MIT), is the source of the COM behaviour and the
 * documented gotchas; see NOTICE.
 */
import {
  DEFAULT_MAX_ATTACHMENT_BYTES,
  DEFAULT_MAX_ATTACHMENT_TOTAL_BYTES,
  MAX_TRANSPORT_HEADER_CHARS,
} from "./limits.ts";

export const OUTLOOK_HELPER_SCRIPT = String.raw`
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch { }
$script:app = $null
$script:ns = $null

function Send-Json($obj) {
  $json = $obj | ConvertTo-Json -Compress -Depth 12
  $ascii = [regex]::Replace($json, '[^\x00-\x7E]', { param($m) '\u{0:x4}' -f [int][char]$m.Value })
  [Console]::Out.WriteLine($ascii)
  [Console]::Out.Flush()
}

function Get-App([bool]$launch) {
  if ($null -ne $script:app) { return $script:app }
  if ($launch) {
    $script:app = New-Object -ComObject Outlook.Application
  } else {
    # GetActiveObject never starts Outlook. Auto-launch is opt-in (#6).
    $script:app = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
  }
  return $script:app
}

function Get-Ns([bool]$launch) {
  if ($null -ne $script:ns) { return $script:ns }
  $app = Get-App $launch
  $script:ns = $app.GetNamespace('MAPI')
  return $script:ns
}

function Iso($value) {
  if ($null -eq $value) { return $null }
  try {
    $dto = [DateTimeOffset]$value
    return $dto.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
  } catch {
    return [string]$value
  }
}

function Get-Prop($obj, [string]$tag) {
  try {
    $v = $obj.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/' + $tag)
    if ($null -eq $v) { return $null }
    return [string]$v
  } catch {
    return $null
  }
}

function Str-Array($value) {
  $out = New-Object System.Collections.ArrayList
  if ($null -ne $value) {
    foreach ($v in @($value)) {
      if ($null -ne $v -and [string]$v -ne '') { [void]$out.Add([string]$v) }
    }
  }
  return @($out)
}

function Sender-Smtp($item) {
  $raw = ''
  try { $raw = [string]$item.SenderEmailAddress } catch { }
  if ($raw -and $raw.Contains('@') -and -not $raw.StartsWith('/')) { return $raw }
  $smtp = Get-Prop $item '0x5D01001E'
  if (-not $smtp) { $smtp = Get-Prop $item '0x5D02001E' }
  if (-not $smtp -and $raw) { $smtp = Get-Prop $item '0x39FE001E' }
  return $smtp
}

function Recipient-Smtp($recipient) {
  $addr = ''
  try { $addr = [string]$recipient.Address } catch { }
  if ($addr -and $addr.Contains('@') -and -not $addr.StartsWith('/')) { return $addr }
  $smtp = Get-Prop $recipient '0x39FE001E'
  if ($smtp) { return $smtp }
  return $addr
}

function Summary-Mail($item) {
  $received = $null
  $sent = $null
  try { $received = $item.ReceivedTime } catch { }
  try { $sent = $item.SentOn } catch { }
  $folder = $null
  try { $folder = [string]$item.Parent.FolderPath } catch { }
  $isRead = $false
  try { $isRead = -not [bool]$item.UnRead } catch { }
  $hasAttachments = $false
  try { $hasAttachments = (([int]$item.Attachments.Count) -gt 0) } catch { }
  $size = $null
  try { $size = [int64]$item.Size } catch { }
  $importance = $null
  try { $importance = [int]$item.Importance } catch { }
  $entryId = ''
  try { $entryId = [string]$item.EntryID } catch { }
  $to = $null
  try { $to = [string]$item.To } catch { }
  $cc = $null
  try { $cc = [string]$item.CC } catch { }
  $sender = $null
  try { $sender = [string]$item.SenderName } catch { }
  $senderAddress = $null
  try { $senderAddress = [string]$item.SenderEmailAddress } catch { }
  $subject = ''
  try { $subject = [string]$item.Subject } catch { }
  return [ordered]@{
    entry_id = $entryId
    received_at = Iso $received
    sent_at = Iso $sent
    sender = $sender
    sender_address = $senderAddress
    sender_smtp = Sender-Smtp $item
    to = $to
    cc = $cc
    subject = $subject
    is_read = $isRead
    has_attachments = $hasAttachments
    size_bytes = $size
    importance = $importance
    folder = $folder
    internet_message_id = Get-Prop $item '0x1035001F'
  }
}

function Summary-Event($item) {
  $entryId = ''
  try { $entryId = [string]$item.EntryID } catch { }
  $start = $null
  $end = $null
  try { $start = $item.Start } catch { }
  try { $end = $item.End } catch { }
  $subject = ''
  try { $subject = [string]$item.Subject } catch { }
  $duration = $null
  try { $duration = [int]$item.Duration } catch { }
  $location = ''
  try { $location = [string]$item.Location } catch { }
  $organizer = $null
  try { $organizer = [string]$item.Organizer } catch { }
  $isRecurring = $false
  try { $isRecurring = [bool]$item.IsRecurring } catch { }
  $isAllDay = $false
  try { $isAllDay = [bool]$item.AllDayEvent } catch { }
  $meetingStatus = $null
  try { $meetingStatus = [int]$item.MeetingStatus } catch { }
  $responseStatus = $null
  try { $responseStatus = [int]$item.ResponseStatus } catch { }
  $required = ''
  try { $required = [string]$item.RequiredAttendees } catch { }
  $optional = ''
  try { $optional = [string]$item.OptionalAttendees } catch { }
  return [ordered]@{
    entry_id = $entryId
    subject = $subject
    start = Iso $start
    end = Iso $end
    duration_minutes = $duration
    location = $location
    organizer = $organizer
    is_recurring = $isRecurring
    is_all_day = $isAllDay
    meeting_status = $meetingStatus
    response_status = $responseStatus
    required = $required
    optional = $optional
  }
}

function Recipient-List($item) {
  $out = New-Object System.Collections.ArrayList
  try {
    foreach ($recipient in $item.Recipients) {
      $type = $null
      try { $type = [int]$recipient.Type } catch { }
      [void]$out.Add([ordered]@{
        name = [string]$recipient.Name
        address = [string]$recipient.Address
        smtp = Recipient-Smtp $recipient
        type = $type
      })
    }
  } catch { }
  return @($out)
}

function Attachment-Data($attachment, [int64]$maxBytes) {
  # Save to a private temp file so the bytes are exactly what download_attachment
  # writes, including embedded messages that have no PR_ATTACH_DATA_BIN value.
  $tmp = [IO.Path]::Combine([IO.Path]::GetTempPath(), 'msoutlook-mcp-' + [Guid]::NewGuid().ToString('N') + '.bin')
  try {
    $attachment.SaveAsFile($tmp)
    $length = [int64](New-Object IO.FileInfo($tmp)).Length
    if ($length -gt $maxBytes) { return [ordered]@{ skipped_reason = 'too_large' } }
    $bytes = [IO.File]::ReadAllBytes($tmp)
    return [ordered]@{
      data_base64 = [Convert]::ToBase64String($bytes)
      bytes = $length
    }
  } catch {
    return [ordered]@{ skipped_reason = 'unavailable' }
  } finally {
    try { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force } } catch { }
  }
}

function Attachment-List($item, [bool]$includeBytes = $false, [int64]$maxBytes = ${DEFAULT_MAX_ATTACHMENT_BYTES}, [int64]$maxTotal = ${DEFAULT_MAX_ATTACHMENT_TOTAL_BYTES}) {
  $out = New-Object System.Collections.ArrayList
  $total = [int64]0
  $count = 0
  try { $count = [int]$item.Attachments.Count } catch { return @($out) }
  for ($i = 1; $i -le $count; $i++) {
    $entry = [ordered]@{
      index = $i
      filename = $null
      size = $null
      content_id = ''
      mime_type = ''
      is_inline = $false
    }
    if ($includeBytes) { $entry['data_base64'] = '' }
    try {
      $a = $item.Attachments.Item($i)
      $size = $null
      try { $size = [int64]$a.Size } catch { }
      $contentId = ''
      try { $contentId = [string](Get-Prop $a '0x3712001F') } catch { }
      $mimeType = ''
      try { $mimeType = [string](Get-Prop $a '0x370E001F') } catch { }
      $entry['filename'] = [string]$a.FileName
      $entry['size'] = $size
      $entry['content_id'] = $contentId
      $entry['mime_type'] = $mimeType
      $entry['is_inline'] = (-not [string]::IsNullOrWhiteSpace($contentId))
      if ($includeBytes) {
        if ($null -ne $size -and $size -gt $maxBytes) {
          $entry['skipped_reason'] = 'too_large'
        } elseif ($null -ne $size -and ($total + $size) -gt $maxTotal) {
          $entry['skipped_reason'] = 'total_cap'
        } else {
          $read = Attachment-Data $a $maxBytes
          if ($read.Contains('data_base64')) {
            if (($total + [int64]$read['bytes']) -gt $maxTotal) {
              $entry['skipped_reason'] = 'total_cap'
            } else {
              $entry['data_base64'] = [string]$read['data_base64']
              $total = $total + [int64]$read['bytes']
            }
          } else {
            $entry['skipped_reason'] = [string]$read['skipped_reason']
          }
        }
      }
    } catch {
      if ($includeBytes) { $entry['skipped_reason'] = 'unavailable' }
    }
    [void]$out.Add($entry)
  }
  return @($out)
}

function Folder-Tree($folder, [int]$depth, [int]$maxDepth) {
  $name = $null
  $path = $null
  $itemCount = 0
  $unread = 0
  try { $name = [string]$folder.Name } catch { }
  try { $path = [string]$folder.FolderPath } catch { }
  try { $itemCount = [int]$folder.Items.Count } catch { }
  try { $unread = [int]$folder.UnReadItemCount } catch { }
  $children = New-Object System.Collections.ArrayList
  if ($depth -lt $maxDepth) {
    try {
      foreach ($sub in $folder.Folders) {
        [void]$children.Add((Folder-Tree $sub ($depth + 1) $maxDepth))
      }
    } catch { }
  }
  return [ordered]@{
    name = $name
    path = $path
    item_count = $itemCount
    unread_count = $unread
    children = @($children)
  }
}

function Resolve-Store($ns, [string]$mailbox) {
  if ([string]::IsNullOrWhiteSpace($mailbox)) { return $ns.DefaultStore }
  $target = $mailbox.ToLowerInvariant()
  foreach ($store in $ns.Stores) {
    try {
      if (([string]$store.DisplayName).ToLowerInvariant() -eq $target) { return $store }
    } catch { }
  }
  foreach ($account in $ns.Accounts) {
    try {
      $smtp = [string]$account.SmtpAddress
      if ($smtp -and $smtp.ToLowerInvariant() -eq $target) { return $account.DeliveryStore }
    } catch { }
  }
  throw ("MAILBOX_NOT_FOUND: '" + $mailbox + "' is not mounted in this Outlook profile")
}

function Default-Folder($store, [int]$type) {
  try {
    $folder = $store.GetDefaultFolder($type)
    if ($null -ne $folder) { return $folder }
  } catch { }
  $names = @()
  if ($type -eq 6) { $names = @('inbox', 'boite de reception') }
  if ($type -eq 5) { $names = @('sent items', 'elements envoyes') }
  if ($type -eq 16) { $names = @('drafts', 'brouillons') }
  if ($type -eq 9) { $names = @('calendar', 'calendrier') }
  if ($type -eq 10) { $names = @('contacts') }
  $root = $store.GetRootFolder()
  foreach ($sub in $root.Folders) {
    if ($names -contains ([string]$sub.Name).ToLowerInvariant()) { return $sub }
  }
  throw ("FOLDER_NOT_FOUND: no default folder of type " + $type + " in store '" + [string]$store.DisplayName + "'")
}

function Resolve-Folder($ns, [string]$pathOrName, [string]$mailbox) {
  if ([string]::IsNullOrWhiteSpace($pathOrName)) { $pathOrName = 'inbox' }
  if ($pathOrName.StartsWith('\\')) {
    $parts = @($pathOrName.Split('\') | Where-Object { $_ -ne '' })
    if ($parts.Count -lt 1) { throw 'FOLDER_NOT_FOUND: empty full path' }
    $storeName = $parts[0]
    $targetStore = $null
    foreach ($store in $ns.Stores) {
      try {
        if ([string]$store.DisplayName -eq $storeName) { $targetStore = $store; break }
      } catch { }
    }
    if ($null -eq $targetStore) { throw ("FOLDER_NOT_FOUND: store '" + $storeName + "'") }
    $current = $targetStore.GetRootFolder()
    for ($i = 1; $i -lt $parts.Count; $i++) {
      $current = $current.Folders.Item($parts[$i])
    }
    return $current
  }

  $useStoreDefaults = -not [string]::IsNullOrWhiteSpace($mailbox)
  $store = Resolve-Store $ns $mailbox
  $lower = $pathOrName.ToLowerInvariant()
  if ($lower -eq 'inbox') {
    if ($useStoreDefaults) { return Default-Folder $store 6 }
    return $ns.GetDefaultFolder(6)
  }
  if ($lower -eq 'sent') {
    if ($useStoreDefaults) { return Default-Folder $store 5 }
    return $ns.GetDefaultFolder(5)
  }
  if ($lower -eq 'drafts') {
    if ($useStoreDefaults) { return Default-Folder $store 16 }
    return $ns.GetDefaultFolder(16)
  }
  $root = $store.GetRootFolder()
  foreach ($sub in $root.Folders) {
    try {
      if (([string]$sub.Name).ToLowerInvariant() -eq $lower) { return $sub }
    } catch { }
  }
  throw ("FOLDER_NOT_FOUND: " + $pathOrName)
}

function Get-Item($ns, [string]$entryId) {
  try {
    return $ns.GetItemFromID($entryId)
  } catch {
    throw ("ENTRY_ID_INVALID: " + $entryId)
  }
}

function Op-Ping($req) {
  $launch = $false
  if ($req.launch) { $launch = $true }
  $app = Get-App $launch
  $ns = Get-Ns $launch
  $version = $null
  $profile = $null
  $defaultStore = $null
  $userSmtp = $null
  $inboxCount = $null
  $inboxUnread = $null
  $calendarCount = $null
  try { $version = [string]$app.Version } catch { }
  try { $profile = [string]$app.DefaultProfileName } catch { }
  try { $defaultStore = [string]$ns.DefaultStore.DisplayName } catch { }
  try { $userSmtp = [string]$ns.Accounts.Item(1).SmtpAddress } catch { }
  try { $inbox = $ns.GetDefaultFolder(6); $inboxCount = [int]$inbox.Items.Count; $inboxUnread = [int]$inbox.UnReadItemCount } catch { }
  try { $calendar = $ns.GetDefaultFolder(9); $calendarCount = [int]$calendar.Items.Count } catch { }
  $accounts = New-Object System.Collections.ArrayList
  try {
    foreach ($account in $ns.Accounts) {
      [void]$accounts.Add([ordered]@{
        display = [string]$account.DisplayName
        smtp = [string]$account.SmtpAddress
      })
    }
  } catch { }
  return [ordered]@{
    version = $version
    profile = $profile
    default_store = $defaultStore
    user_smtp = $userSmtp
    inbox_count = $inboxCount
    inbox_unread = $inboxUnread
    calendar_count = $calendarCount
    accounts = @($accounts)
  }
}

function Op-ListFolders($req) {
  $ns = Get-Ns $false
  $maxDepth = 3
  if ($req.maxDepth) { $maxDepth = [int]$req.maxDepth }
  $stores = @()
  if ($req.mailbox) {
    $stores = @(Resolve-Store $ns ([string]$req.mailbox))
  } else {
    $stores = @($ns.Stores)
  }
  $out = New-Object System.Collections.ArrayList
  foreach ($store in $stores) {
    try {
      $root = $store.GetRootFolder()
      $tree = Folder-Tree $root 0 $maxDepth
      $tree['store'] = [string]$store.DisplayName
      [void]$out.Add($tree)
    } catch { }
  }
  return @($out)
}

function Read-Cursor($req) {
  $at = $null
  if ($null -ne $req.cursorAt -and [string]$req.cursorAt -ne '') {
    try { $at = [DateTimeOffset]::Parse([string]$req.cursorAt).ToUniversalTime() } catch { $at = $null }
  }
  $ids = @{}
  try {
    foreach ($cursorId in @($req.cursorIds)) {
      if ($null -ne $cursorId -and [string]$cursorId -ne '') { $ids[[string]$cursorId] = $true }
    }
  } catch { }
  return [pscustomobject]@{ at = $at; ids = $ids }
}

function Skip-By-Cursor($item, $cursor, [bool]$ascending) {
  if ($null -eq $cursor -or $null -eq $cursor.at) { return $false }
  $received = $null
  try { $received = $item.ReceivedTime } catch { }
  if ($null -eq $received) { return $false }
  $at = $null
  try { $at = [DateTimeOffset]$received } catch { }
  if ($null -eq $at) { return $false }
  # received_at has second precision, so a cursor boundary is a whole second.
  $cmp = $at.ToUnixTimeSeconds().CompareTo($cursor.at.ToUnixTimeSeconds())
  if ($ascending) {
    if ($cmp -lt 0) { return $true }
  } else {
    if ($cmp -gt 0) { return $true }
  }
  if ($cmp -eq 0) {
    $entryId = ''
    try { $entryId = [string]$item.EntryID } catch { }
    if ($cursor.ids.ContainsKey($entryId)) { return $true }
  }
  return $false
}

function Op-ListMail($req) {
  $ns = Get-Ns $false
  $folder = Resolve-Folder $ns ([string]$req.folder) ([string]$req.mailbox)
  $items = $folder.Items
  $ascending = $false
  if ($req.sortAscending) { $ascending = $true }
  $items.Sort('[ReceivedTime]', (-not $ascending))
  $filter = [string]$req.filter
  if ($filter) { $items = $items.Restrict('@SQL=' + $filter) }
  $cursor = Read-Cursor $req
  $limit = 20
  if ($req.limit) { $limit = [int]$req.limit }
  $out = New-Object System.Collections.ArrayList
  $item = $items.GetFirst()
  while ($null -ne $item -and $out.Count -lt $limit) {
    if (-not (Skip-By-Cursor $item $cursor $ascending)) {
      $class = 0
      try { $class = [int]$item.Class } catch { }
      if ($class -eq 43) { [void]$out.Add((Summary-Mail $item)) }
    }
    $item = $items.GetNext()
  }
  $truncated = $false
  if ($out.Count -ge $limit -and $null -ne $item) {
    while ($null -ne $item) {
      $class = 0
      try { $class = [int]$item.Class } catch { }
      if ($class -eq 43) { $truncated = $true; break }
      $item = $items.GetNext()
    }
  }
  $folderPath = $null
  $mailboxUsed = $null
  try { $folderPath = [string]$folder.FolderPath } catch { }
  try { $mailboxUsed = [string]$folder.Store.DisplayName } catch { }
  return [ordered]@{
    items = @($out)
    truncated = $truncated
    folder = $folderPath
    mailbox_used = $mailboxUsed
  }
}

function Op-SearchMail($req) {
  $ns = Get-Ns $false
  $folder = Resolve-Folder $ns ([string]$req.folder) ([string]$req.mailbox)
  $items = $folder.Items
  $items.Sort('[ReceivedTime]', $true)
  $items = $items.Restrict('@SQL=' + [string]$req.filter)
  $cursor = Read-Cursor $req
  $limit = 50
  if ($req.limit) { $limit = [int]$req.limit }
  $out = New-Object System.Collections.ArrayList
  $item = $items.GetFirst()
  while ($null -ne $item -and $out.Count -lt $limit) {
    if (-not (Skip-By-Cursor $item $cursor $false)) {
      $class = 0
      try { $class = [int]$item.Class } catch { }
      if ($class -eq 43) { [void]$out.Add((Summary-Mail $item)) }
    }
    $item = $items.GetNext()
  }
  $truncated = $false
  if ($out.Count -ge $limit -and $null -ne $item) {
    while ($null -ne $item) {
      $class = 0
      try { $class = [int]$item.Class } catch { }
      if ($class -eq 43) { $truncated = $true; break }
      $item = $items.GetNext()
    }
  }
  $folderPath = $null
  $mailboxUsed = $null
  try { $folderPath = [string]$folder.FolderPath } catch { }
  try { $mailboxUsed = [string]$folder.Store.DisplayName } catch { }
  return [ordered]@{
    items = @($out)
    truncated = $truncated
    folder = $folderPath
    mailbox_used = $mailboxUsed
  }
}

function Op-ReadMail($req) {
  $ns = Get-Ns $false
  $item = Get-Item $ns ([string]$req.entryId)
  $out = Summary-Mail $item
  $body = ''
  try { $body = [string]$item.Body } catch { }
  $maxBody = 50000
  if ($null -ne $req.maxBodyBytes) { $maxBody = [int]$req.maxBodyBytes }
  $truncated = $false
  if ($maxBody -gt 0 -and $body.Length -gt $maxBody) {
    $body = $body.Substring(0, $maxBody)
    $truncated = $true
  }
  $out['body'] = $body
  $out['body_truncated'] = $truncated
  if ($req.includeHtml) {
    $html = ''
    try { $html = [string]$item.HTMLBody } catch { }
    $htmlTruncated = $false
    if ($maxBody -gt 0 -and $html.Length -gt $maxBody) {
      $html = $html.Substring(0, $maxBody)
      $htmlTruncated = $true
    }
    $out['html'] = $html
    $out['html_truncated'] = $htmlTruncated
  }
  $out['recipients'] = Recipient-List $item
  $includeAttachmentBytes = $false
  if ($req.includeAttachmentBytes) { $includeAttachmentBytes = $true }
  $maxAttachmentBytes = [int64]${DEFAULT_MAX_ATTACHMENT_BYTES}
  if ($null -ne $req.maxAttachmentBytes) { $maxAttachmentBytes = [int64]$req.maxAttachmentBytes }
  $maxAttachmentTotalBytes = [int64]${DEFAULT_MAX_ATTACHMENT_TOTAL_BYTES}
  if ($null -ne $req.maxAttachmentTotalBytes) { $maxAttachmentTotalBytes = [int64]$req.maxAttachmentTotalBytes }
  $out['attachments'] = Attachment-List -item $item -includeBytes $includeAttachmentBytes -maxBytes $maxAttachmentBytes -maxTotal $maxAttachmentTotalBytes
  if ($req.includeTransportHeaders) {
    $headers = ''
    try { $headers = [string](Get-Prop $item '0x007D001F') } catch { }
    $headersTruncated = $false
    # Guard the helper/stdout path; the host enforces the exact UTF-8 byte cap.
    if ($headers.Length -gt ${MAX_TRANSPORT_HEADER_CHARS}) {
      $cut = ${MAX_TRANSPORT_HEADER_CHARS}
      $last = [int][char]$headers[$cut - 1]
      # Do not split a surrogate pair at the cap.
      if ($last -ge 0xD800 -and $last -le 0xDBFF) { $cut = $cut - 1 }
      $headers = $headers.Substring(0, $cut)
      $headersTruncated = $true
    }
    $out['transport_headers'] = $headers
    $out['transport_headers_truncated'] = $headersTruncated
  }
  return $out
}

function Op-DownloadAttachment($req) {
  $ns = Get-Ns $false
  $item = Get-Item $ns ([string]$req.entryId)
  $index = [int]$req.index
  $count = [int]$item.Attachments.Count
  if ($index -lt 1 -or $index -gt $count) {
    throw ("ATTACHMENT_INDEX_OUT_OF_RANGE: " + $index + " (1.." + $count + ")")
  }
  $attachment = $item.Attachments.Item($index)
  $base = [string]$req.destDir
  if ([string]::IsNullOrWhiteSpace($base)) {
    $local = [string]$env:LOCALAPPDATA
    if ([string]::IsNullOrWhiteSpace($local)) { $local = [IO.Path]::GetTempPath() }
    $base = Join-Path $local 'msoutlook-mcp\downloads'
  }
  [void][IO.Directory]::CreateDirectory($base)
  $fileName = [string]$attachment.FileName
  if ([string]::IsNullOrWhiteSpace($fileName)) { $fileName = 'attachment_' + $index + '.bin' }
  $fileName = [IO.Path]::GetFileName($fileName)
  $target = Join-Path $base $fileName
  $stem = [IO.Path]::GetFileNameWithoutExtension($fileName)
  $ext = [IO.Path]::GetExtension($fileName)
  $n = 1
  while (Test-Path -LiteralPath $target) {
    $target = Join-Path $base ($stem + '_' + $n + $ext)
    $n = $n + 1
  }
  $attachment.SaveAsFile($target)
  $size = $null
  try { $size = [int64]$attachment.Size } catch { }
  return [ordered]@{ saved_to = [string]$target; size = $size }
}

function Op-CreateDraft($req) {
  $launch = $false
  $app = Get-App $launch
  $ns = Get-Ns $launch
  $mail = $app.CreateItem(0)
  $targetDrafts = $null
  if ($req.from_mailbox) {
    $store = Resolve-Store $ns ([string]$req.from_mailbox)
    $targetDrafts = Default-Folder $store 16
    $mail.SentOnBehalfOfName = [string]$req.from_mailbox
  }
  $to = Str-Array $req.to
  $cc = Str-Array $req.cc
  $bcc = Str-Array $req.bcc
  $mail.To = ($to -join '; ')
  if ($cc.Count -gt 0) { $mail.CC = ($cc -join '; ') }
  if ($bcc.Count -gt 0) { $mail.BCC = ($bcc -join '; ') }
  $mail.Subject = [string]$req.subject
  if ($req.html) {
    $mail.BodyFormat = 2
    $mail.HTMLBody = [string]$req.body
  } else {
    $mail.BodyFormat = 1
    $mail.Body = [string]$req.body
  }
  $attached = New-Object System.Collections.ArrayList
  foreach ($path in (Str-Array $req.attachments)) {
    $full = [IO.Path]::GetFullPath($path)
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) {
      throw ("ATTACHMENT_NOT_FOUND: " + $full)
    }
    [void]$mail.Attachments.Add($full)
    [void]$attached.Add($full)
  }
  $mail.Save()
  $entryId = [string]$mail.EntryID
  $draftsFolder = $null
  try { $draftsFolder = [string]$mail.Parent.FolderPath } catch { }
  if ($null -ne $targetDrafts) {
    try {
      $moved = $mail.Move($targetDrafts)
      $entryId = [string]$moved.EntryID
      $draftsFolder = [string]$moved.Parent.FolderPath
    } catch { }
  }
  return [ordered]@{
    entry_id = $entryId
    drafts_folder = $draftsFolder
    attachments = @($attached)
  }
}

function Op-GetRecipients($req) {
  $ns = Get-Ns $false
  $item = Get-Item $ns ([string]$req.entryId)
  return Recipient-List $item
}

function Op-SendDraft($req) {
  $ns = Get-Ns $false
  $item = Get-Item $ns ([string]$req.entryId)
  $item.Send()
  return [ordered]@{ ok = $true }
}

function Op-PrepareReply($req) {
  $ns = Get-Ns $false
  $original = Get-Item $ns ([string]$req.entry_id)
  $targetDrafts = $null
  if ($req.from_mailbox) {
    $store = Resolve-Store $ns ([string]$req.from_mailbox)
    $targetDrafts = Default-Folder $store 16
  }
  if ($req.reply_all) { $reply = $original.ReplyAll() } else { $reply = $original.Reply() }
  if ($null -eq $reply) { throw 'REPLY_FAILED: Outlook did not create a reply' }
  if ($req.from_mailbox) { $reply.SentOnBehalfOfName = [string]$req.from_mailbox }
  if ($req.html) {
    $reply.BodyFormat = 2
    $existing = ''
    try { $existing = [string]$reply.HTMLBody } catch { }
    $reply.HTMLBody = [string]$req.body + '<br><br>' + $existing
  } else {
    $reply.BodyFormat = 1
    $existing = ''
    try { $existing = [string]$reply.Body } catch { }
    $reply.Body = [string]$req.body + [Environment]::NewLine + [Environment]::NewLine + $existing
  }
  $recipients = Recipient-List $reply
  $reply.Save()
  $entryId = [string]$reply.EntryID
  $draftsFolder = $null
  try { $draftsFolder = [string]$reply.Parent.FolderPath } catch { }
  if ($null -ne $targetDrafts) {
    try {
      $moved = $reply.Move($targetDrafts)
      $entryId = [string]$moved.EntryID
      $draftsFolder = [string]$moved.Parent.FolderPath
    } catch { }
  }
  return [ordered]@{
    entry_id = $entryId
    drafts_folder = $draftsFolder
    attachments = @()
    recipients = $recipients
  }
}

function Op-MoveMail($req) {
  $ns = Get-Ns $false
  $item = Get-Item $ns ([string]$req.entryId)
  $target = Resolve-Folder $ns ([string]$req.targetFolder) ([string]$req.mailbox)
  $moved = $item.Move($target)
  $newEntryId = $null
  $targetPath = $null
  try { $newEntryId = [string]$moved.EntryID } catch { }
  try { $targetPath = [string]$target.FolderPath } catch { }
  return [ordered]@{ new_entry_id = $newEntryId; target_folder = $targetPath }
}

function Op-MarkRead($req) {
  $ns = Get-Ns $false
  $item = Get-Item $ns ([string]$req.entryId)
  $item.UnRead = -not [bool]$req.read
  $item.Save()
  return [ordered]@{ ok = $true }
}

function Op-FlagMail($req) {
  $ns = Get-Ns $false
  $item = Get-Item $ns ([string]$req.entryId)
  if ($req.flag) { $item.FlagStatus = 2 } else { $item.FlagStatus = 0 }
  $item.Save()
  return [ordered]@{ ok = $true }
}

function Op-ListEvents($req) {
  $ns = Get-Ns $false
  if ($req.calendar) {
    $calendar = Resolve-Folder $ns ([string]$req.calendar) ''
  } else {
    $calendar = $ns.GetDefaultFolder(9)
  }
  $daysAhead = 7
  $daysBack = 0
  $limit = 50
  if ($null -ne $req.daysAhead) { $daysAhead = [int]$req.daysAhead }
  if ($null -ne $req.daysBack) { $daysBack = [int]$req.daysBack }
  if ($null -ne $req.limit) { $limit = [int]$req.limit }
  $now = Get-Date
  $start = $now.AddDays(-$daysBack)
  $end = $now.AddDays($daysAhead)
  $items = $calendar.Items
  $items.IncludeRecurrences = $true
  $items.Sort('[Start]')
  $out = New-Object System.Collections.ArrayList
  $item = $items.GetFirst()
  while ($null -ne $item -and $out.Count -lt $limit) {
    $s = $null
    try { $s = [DateTime]$item.Start } catch { }
    if ($null -ne $s) {
      if ($s -gt $end) { break }
      if ($s -ge $start) { [void]$out.Add((Summary-Event $item)) }
    }
    $item = $items.GetNext()
  }
  return @($out)
}

function Op-ReadEvent($req) {
  $ns = Get-Ns $false
  $item = Get-Item $ns ([string]$req.entryId)
  $out = Summary-Event $item
  $body = ''
  try { $body = [string]$item.Body } catch { }
  $maxBody = 50000
  if ($null -ne $req.maxBodyBytes) { $maxBody = [int]$req.maxBodyBytes }
  $truncated = $false
  if ($maxBody -gt 0 -and $body.Length -gt $maxBody) {
    $body = $body.Substring(0, $maxBody)
    $truncated = $true
  }
  $out['body'] = $body
  $out['body_truncated'] = $truncated
  $out['recipients'] = Recipient-List $item
  return $out
}

function Op-SelfBusy($req) {
  $ns = Get-Ns $false
  $calendar = $ns.GetDefaultFolder(9)
  $start = [DateTime]$req.startIso
  $daysAhead = 5
  if ($null -ne $req.daysAhead) { $daysAhead = [int]$req.daysAhead }
  $endWindow = $start.AddDays($daysAhead + 1)
  $items = $calendar.Items
  $items.IncludeRecurrences = $true
  $items.Sort('[Start]')
  $out = New-Object System.Collections.ArrayList
  $item = $items.GetFirst()
  while ($null -ne $item) {
    $s = $null
    $e = $null
    try { $s = [DateTime]$item.Start } catch { }
    try { $e = [DateTime]$item.End } catch { }
    if ($null -ne $s -and $null -ne $e) {
      if ($s -gt $endWindow) { break }
      if ($e -ge $start) {
        $skip = $false
        try { if ([int]$item.MeetingStatus -eq 5) { $skip = $true } } catch { }
        try { if ([int]$item.BusyStatus -eq 0) { $skip = $true } } catch { }
        try { if ([bool]$item.AllDayEvent) { $skip = $true } } catch { }
        if (-not $skip) {
          [void]$out.Add([ordered]@{ start = Iso $s; end = Iso $e })
        }
      }
    }
    $item = $items.GetNext()
  }
  return @($out)
}

function Op-AttendeeBusy($req) {
  $ns = Get-Ns $false
  $smtp = [string]$req.smtp
  $slot = 30
  if ($null -ne $req.slotMinutes) { $slot = [int]$req.slotMinutes }
  try {
    $recipient = $ns.CreateRecipient($smtp)
    if (-not $recipient.Resolve()) {
      return [ordered]@{ busy_string = $null; error = 'unresolved (not found in the GAL)' }
    }
    $start = [DateTime]$req.startIso
    $result = $recipient.FreeBusy($start, $slot, $true)
    if (-not $result) {
      return [ordered]@{ busy_string = $null; error = 'FreeBusy returned empty (no Exchange data or access denied)' }
    }
    return [ordered]@{ busy_string = [string]$result; error = $null }
  } catch {
    return [ordered]@{ busy_string = $null; error = 'COM error: ' + $_.Exception.Message }
  }
}

function Op-CreateEventDraft($req) {
  $app = Get-App $false
  $appt = $app.CreateItem(1)
  $start = [DateTime]$req.start_iso
  $appt.Subject = [string]$req.subject
  $appt.Start = $start
  $appt.Duration = [int]$req.duration_minutes
  if ($req.location) { $appt.Location = [string]$req.location }
  if ($req.body) { $appt.Body = [string]$req.body }
  $required = Str-Array $req.required_attendees
  $optional = Str-Array $req.optional_attendees
  if ($required.Count -gt 0 -or $optional.Count -gt 0) {
    $appt.MeetingStatus = 1
    foreach ($address in $required) {
      $r = $appt.Recipients.Add($address)
      $r.Type = 1
    }
    foreach ($address in $optional) {
      $r = $appt.Recipients.Add($address)
      $r.Type = 2
    }
    [void]$appt.Recipients.ResolveAll()
  }
  $appt.Save()
  return [ordered]@{
    entry_id = [string]$appt.EntryID
    subject = [string]$req.subject
    start = Iso $start
    duration_minutes = [int]$req.duration_minutes
    required = @($required)
    optional = @($optional)
    is_meeting = (($required.Count + $optional.Count) -gt 0)
  }
}

function Op-GetEventRecipients($req) {
  $ns = Get-Ns $false
  $item = Get-Item $ns ([string]$req.entryId)
  return Recipient-List $item
}

function Op-SendMeetingInvites($req) {
  $ns = Get-Ns $false
  $item = Get-Item $ns ([string]$req.entryId)
  $status = 0
  try { $status = [int]$item.MeetingStatus } catch { }
  if ($status -eq 0) { throw 'NOT_A_MEETING: this item has no attendees to invite' }
  $item.Send()
  return [ordered]@{ ok = $true }
}

function Op-RespondMeeting($req) {
  $ns = Get-Ns $false
  $item = Get-Item $ns ([string]$req.entryId)
  $status = 3
  $response = [string]$req.response
  if ($response -eq 'tentative') { $status = 2 }
  if ($response -eq 'declined') { $status = 4 }
  $responseItem = $item.Respond($status, $true)
  if ($req.comment) {
    try { $responseItem.Body = [string]$req.comment + [Environment]::NewLine + [Environment]::NewLine + [string]$responseItem.Body } catch { }
  }
  if ($req.sendResponse) {
    $responseItem.Send()
    return [ordered]@{ ok = $true; action = 'sent' }
  }
  $responseItem.Save()
  return [ordered]@{ ok = $true; action = 'saved' }
}

function Rule-Summary($rule) {
  $conditions = New-Object System.Collections.ArrayList
  $actions = New-Object System.Collections.ArrayList
  try {
    $c = $rule.Conditions
    try {
      if ($c.From.Enabled) {
        $addresses = New-Object System.Collections.ArrayList
        foreach ($r in $c.From.Recipients) {
          $value = Recipient-Smtp $r
          if (-not $value) { $value = [string]$r.Name }
          if ($value) { [void]$addresses.Add($value) }
        }
        [void]$conditions.Add('From IN [' + ($addresses -join ', ') + ']')
      }
    } catch { }
    try {
      if ($c.SentTo.Enabled) {
        $addresses = New-Object System.Collections.ArrayList
        foreach ($r in $c.SentTo.Recipients) {
          $value = Recipient-Smtp $r
          if (-not $value) { $value = [string]$r.Name }
          if ($value) { [void]$addresses.Add($value) }
        }
        [void]$conditions.Add('SentTo IN [' + ($addresses -join ', ') + ']')
      }
    } catch { }
    try { if ($c.Subject.Enabled) { [void]$conditions.Add('Subject CONTAINS ' + (@($c.Subject.Text) -join ', ')) } } catch { }
    try { if ($c.BodyOrSubject.Enabled) { [void]$conditions.Add('BodyOrSubject CONTAINS ' + (@($c.BodyOrSubject.Text) -join ', ')) } } catch { }
    try { if ($c.HasAttachment.Enabled) { [void]$conditions.Add('HasAttachment') } } catch { }
    try { if ($c.Importance.Enabled) { [void]$conditions.Add('Importance = ' + [string]$c.Importance.Importance) } } catch { }
  } catch { }
  try {
    $a = $rule.Actions
    try { if ($a.MoveToFolder.Enabled) { [void]$actions.Add('Move -> ' + [string]$a.MoveToFolder.Folder.FolderPath) } } catch { }
    try { if ($a.CopyToFolder.Enabled) { [void]$actions.Add('Copy -> ' + [string]$a.CopyToFolder.Folder.FolderPath) } } catch { }
    try { if ($a.MarkAsTask.Enabled) { [void]$actions.Add('MarkAsTask') } } catch { }
    try { if ($a.Delete.Enabled) { [void]$actions.Add('Delete') } } catch { }
    try { if ($a.DeletePermanently.Enabled) { [void]$actions.Add('DeletePermanently') } } catch { }
    try { if ($a.AssignToCategory.Enabled) { [void]$actions.Add('AssignToCategory ' + (@($a.AssignToCategory.Categories) -join ', ')) } } catch { }
    try {
      if ($a.Forward.Enabled) {
        $addresses = New-Object System.Collections.ArrayList
        foreach ($r in $a.Forward.Recipients) {
          $value = Recipient-Smtp $r
          if (-not $value) { $value = [string]$r.Name }
          if ($value) { [void]$addresses.Add($value) }
        }
        [void]$actions.Add('Forward -> [' + ($addresses -join ', ') + ']')
      }
    } catch { }
    try { if ($a.Stop.Enabled) { [void]$actions.Add('StopProcessingMoreRules') } } catch { }
  } catch { }
  $executionOrder = $null
  try { $executionOrder = [int]$rule.ExecutionOrder } catch { }
  $ruleType = $null
  try { $ruleType = [int]$rule.RuleType } catch { }
  $isLocal = $false
  try { $isLocal = [bool]$rule.IsLocalRule } catch { }
  return [ordered]@{
    name = [string]$rule.Name
    enabled = [bool]$rule.Enabled
    execution_order = $executionOrder
    is_local_rule = $isLocal
    rule_type = $ruleType
    conditions = @($conditions)
    actions = @($actions)
  }
}

function Op-ListRules($req) {
  $ns = Get-Ns $false
  $rules = $ns.DefaultStore.GetRules()
  $out = New-Object System.Collections.ArrayList
  for ($i = 1; $i -le [int]$rules.Count; $i++) {
    try { [void]$out.Add((Rule-Summary $rules.Item($i))) } catch { }
  }
  return @($out)
}

function Op-ToggleRule($req) {
  $ns = Get-Ns $false
  $rules = $ns.DefaultStore.GetRules()
  $target = $null
  for ($i = 1; $i -le [int]$rules.Count; $i++) {
    if ([string]$rules.Item($i).Name -eq [string]$req.name) { $target = $rules.Item($i); break }
  }
  if ($null -eq $target) { throw ('RULE_NOT_FOUND: ' + [string]$req.name) }
  $target.Enabled = [bool]$req.enabled
  $rules.Save($false)
  return [ordered]@{ ok = $true }
}

function Op-CreateRule($req) {
  $ns = Get-Ns $false
  $rules = $ns.DefaultStore.GetRules()
  for ($i = 1; $i -le [int]$rules.Count; $i++) {
    if ([string]$rules.Item($i).Name -eq [string]$req.name) { throw ('RULE_NAME_EXISTS: ' + [string]$req.name) }
  }
  $rule = $rules.Create([string]$req.name, 0)
  $from = Str-Array $req.from_addresses
  if ($from.Count -gt 0) {
    foreach ($address in $from) { [void]$rule.Conditions.From.Recipients.Add($address) }
    [void]$rule.Conditions.From.Recipients.ResolveAll()
    $rule.Conditions.From.Enabled = $true
  }
  $sentTo = Str-Array $req.sent_to
  if ($sentTo.Count -gt 0) {
    foreach ($address in $sentTo) { [void]$rule.Conditions.SentTo.Recipients.Add($address) }
    [void]$rule.Conditions.SentTo.Recipients.ResolveAll()
    $rule.Conditions.SentTo.Enabled = $true
  }
  $subject = Str-Array $req.subject_contains
  if ($subject.Count -gt 0) {
    $rule.Conditions.Subject.Text = [string[]]$subject
    $rule.Conditions.Subject.Enabled = $true
  }
  $bodySubject = Str-Array $req.body_or_subject_contains
  if ($bodySubject.Count -gt 0) {
    $rule.Conditions.BodyOrSubject.Text = [string[]]$bodySubject
    $rule.Conditions.BodyOrSubject.Enabled = $true
  }
  if ($req.has_attachment) { $rule.Conditions.HasAttachment.Enabled = $true }
  if ($req.importance_high) {
    $rule.Conditions.Importance.Importance = 2
    $rule.Conditions.Importance.Enabled = $true
  }
  if ($req.move_to_folder) {
    $target = Resolve-Folder $ns ([string]$req.move_to_folder) ''
    $rule.Actions.MoveToFolder.Folder = $target
    $rule.Actions.MoveToFolder.Enabled = $true
  }
  if ($req.copy_to_folder) {
    $target = Resolve-Folder $ns ([string]$req.copy_to_folder) ''
    $rule.Actions.CopyToFolder.Folder = $target
    $rule.Actions.CopyToFolder.Enabled = $true
  }
  $categories = Str-Array $req.assign_categories
  if ($categories.Count -gt 0) {
    $rule.Actions.AssignToCategory.Categories = [string[]]$categories
    $rule.Actions.AssignToCategory.Enabled = $true
  }
  $forward = Str-Array $req.forward_to
  if ($forward.Count -gt 0) {
    foreach ($address in $forward) { [void]$rule.Actions.Forward.Recipients.Add($address) }
    [void]$rule.Actions.Forward.Recipients.ResolveAll()
    $rule.Actions.Forward.Enabled = $true
  }
  if ($req.stop_processing) {
    try { $rule.Actions.Stop.Enabled = $true } catch { }
  }
  $rule.Enabled = [bool]$req.enabled
  $rules.Save($false)
  if ($req.apply_now) {
    try {
      $inbox = $ns.GetDefaultFolder(6)
      $rule.Execute($false, $inbox, $false, 0)
    } catch { }
  }
  return [ordered]@{ rule = (Rule-Summary $rule) }
}

function Op-DeleteRule($req) {
  $ns = Get-Ns $false
  $rules = $ns.DefaultStore.GetRules()
  $found = $false
  for ($i = 1; $i -le [int]$rules.Count; $i++) {
    if ([string]$rules.Item($i).Name -eq [string]$req.name) {
      $rules.Remove($i)
      $found = $true
      break
    }
  }
  if (-not $found) { throw ('RULE_NOT_FOUND: ' + [string]$req.name) }
  $rules.Save($false)
  return [ordered]@{ ok = $true }
}

function Exchange-User($eu) {
  $out = [ordered]@{
    name = $null
    smtp_address = $null
    alias = $null
    job_title = $null
    department = $null
    office_location = $null
    company_name = $null
    business_phone = $null
    mobile_phone = $null
  }
  try { $out['name'] = [string]$eu.Name } catch { }
  try { $out['smtp_address'] = [string]$eu.PrimarySmtpAddress } catch { }
  try { $out['alias'] = [string]$eu.Alias } catch { }
  try { $out['job_title'] = [string]$eu.JobTitle } catch { }
  try { $out['department'] = [string]$eu.Department } catch { }
  try { $out['office_location'] = [string]$eu.OfficeLocation } catch { }
  try { $out['company_name'] = [string]$eu.CompanyName } catch { }
  try { $out['business_phone'] = [string]$eu.BusinessTelephoneNumber } catch { }
  try { $out['mobile_phone'] = [string]$eu.MobileTelephoneNumber } catch { }
  return $out
}

function Op-SearchContacts($req) {
  $ns = Get-Ns $false
  $limit = 20
  if ($null -ne $req.limit) { $limit = [int]$req.limit }
  $query = [string]$req.query
  try {
    $conn = New-Object -ComObject ADODB.Connection
    $conn.Provider = 'ADsDSOObject'
    $conn.Open('Active Directory Provider')
    $cmd = New-Object -ComObject ADODB.Command
    $cmd.ActiveConnection = $conn
    $escaped = $query.Replace('\', '\5c').Replace('(', '\28').Replace(')', '\29').Replace('*', '\2a')
    $cmd.CommandText = '<LDAP://>;(&(objectCategory=person)(objectClass=user)(anr=' + $escaped + '));cn,mail,title,department,physicalDeliveryOfficeName,telephoneNumber,mobile,company;subtree'
    $cmd.Properties.Item('Page Size').Value = $limit
    $cmd.Properties.Item('Size Limit').Value = $limit
    $rs = $cmd.Execute()
    $out = New-Object System.Collections.ArrayList
    while (-not $rs.EOF -and $out.Count -lt $limit) {
      $fields = $rs.Fields
      $contact = [ordered]@{}
      try { $contact['name'] = [string]$fields.Item('cn').Value } catch { $contact['name'] = $null }
      try { $contact['smtp_address'] = [string]$fields.Item('mail').Value } catch { $contact['smtp_address'] = $null }
      try { $contact['job_title'] = [string]$fields.Item('title').Value } catch { $contact['job_title'] = $null }
      try { $contact['department'] = [string]$fields.Item('department').Value } catch { $contact['department'] = $null }
      try { $contact['office_location'] = [string]$fields.Item('physicalDeliveryOfficeName').Value } catch { $contact['office_location'] = $null }
      try { $contact['business_phone'] = [string]$fields.Item('telephoneNumber').Value } catch { $contact['business_phone'] = $null }
      try { $contact['mobile_phone'] = [string]$fields.Item('mobile').Value } catch { $contact['mobile_phone'] = $null }
      try { $contact['company_name'] = [string]$fields.Item('company').Value } catch { $contact['company_name'] = $null }
      [void]$out.Add($contact)
      $rs.MoveNext()
    }
    try { $rs.Close() } catch { }
    try { $conn.Close() } catch { }
    return [ordered]@{ contacts = @($out); hint = $null }
  } catch { }

  try {
    $recipient = $ns.CreateRecipient($query)
    $resolved = $recipient.Resolve()
    if (-not $resolved) { return [ordered]@{ contacts = @(); hint = $null } }
    $addressEntry = $recipient.AddressEntry
    if ($null -eq $addressEntry) { return [ordered]@{ contacts = @(); hint = $null } }
    $eu = $null
    try { $eu = $addressEntry.GetExchangeUser() } catch { }
    if ($null -eq $eu) {
      return [ordered]@{ contacts = @(); hint = 'Resolved, but not an Exchange user.' }
    }
    return [ordered]@{
      contacts = @((Exchange-User $eu))
      hint = 'CreateRecipient fallback (1 result max). LDAP unavailable.'
    }
  } catch {
    throw ('CONTACT_SEARCH_FAILED: ' + $_.Exception.Message)
  }
}

function Op-GetContactDetails($req) {
  $ns = Get-Ns $false
  $smtp = [string]$req.smtpAddress
  try {
    $recipient = $ns.CreateRecipient($smtp)
    if (-not $recipient.Resolve()) { throw ('CONTACT_NOT_FOUND: ' + $smtp) }
  } catch {
    throw ('CONTACT_NOT_FOUND: ' + $smtp)
  }
  $addressEntry = $recipient.AddressEntry
  if ($null -eq $addressEntry) { throw ('ADDRESS_ENTRY_NULL: ' + $smtp) }
  $eu = $null
  try { $eu = $addressEntry.GetExchangeUser() } catch { }
  if ($null -eq $eu) { throw ('NOT_EXCHANGE_USER: ' + $smtp) }
  $out = Exchange-User $eu
  try { $out['street_address'] = [string]$eu.StreetAddress } catch { $out['street_address'] = $null }
  try { $out['city'] = [string]$eu.City } catch { $out['city'] = $null }
  try { $out['state'] = [string]$eu.StateOrProvince } catch { $out['state'] = $null }
  try { $out['postal_code'] = [string]$eu.PostalCode } catch { $out['postal_code'] = $null }
  $manager = $null
  try {
    $mgr = $eu.GetExchangeUserManager()
    if ($null -ne $mgr) {
      $manager = [ordered]@{ name = [string]$mgr.Name; smtp_address = [string]$mgr.PrimarySmtpAddress }
    }
  } catch { }
  $out['manager'] = $manager
  return $out
}

while ($true) {
  $line = [Console]::In.ReadLine()
  if ($null -eq $line) { break }
  if ([string]::IsNullOrWhiteSpace($line)) { continue }
  $id = 0
  try {
    $req = $line | ConvertFrom-Json
    $id = $req.id
    switch ($req.op) {
      'ping' { $result = Op-Ping $req }
      'listFolders' { $result = Op-ListFolders $req }
      'listMail' { $result = Op-ListMail $req }
      'searchMail' { $result = Op-SearchMail $req }
      'readMail' { $result = Op-ReadMail $req }
      'downloadAttachment' { $result = Op-DownloadAttachment $req }
      'createDraft' { $result = Op-CreateDraft $req }
      'getRecipients' { $result = Op-GetRecipients $req }
      'sendDraft' { $result = Op-SendDraft $req }
      'prepareReply' { $result = Op-PrepareReply $req }
      'moveMail' { $result = Op-MoveMail $req }
      'markRead' { $result = Op-MarkRead $req }
      'flagMail' { $result = Op-FlagMail $req }
      'listEvents' { $result = Op-ListEvents $req }
      'readEvent' { $result = Op-ReadEvent $req }
      'selfBusy' { $result = Op-SelfBusy $req }
      'attendeeBusy' { $result = Op-AttendeeBusy $req }
      'createEventDraft' { $result = Op-CreateEventDraft $req }
      'getEventRecipients' { $result = Op-GetEventRecipients $req }
      'sendMeetingInvites' { $result = Op-SendMeetingInvites $req }
      'respondMeeting' { $result = Op-RespondMeeting $req }
      'listRules' { $result = Op-ListRules $req }
      'toggleRule' { $result = Op-ToggleRule $req }
      'createRule' { $result = Op-CreateRule $req }
      'deleteRule' { $result = Op-DeleteRule $req }
      'searchContacts' { $result = Op-SearchContacts $req }
      'getContactDetails' { $result = Op-GetContactDetails $req }
      default { throw ('UNKNOWN_OP: ' + [string]$req.op) }
    }
    Send-Json ([ordered]@{ id = $id; ok = $true; result = $result })
  } catch {
    $ex = $_.Exception
    $hr = 0
    try { $hr = [int]$ex.HResult } catch { }
    $inner = $ex
    while ($null -ne $inner.InnerException) {
      $inner = $inner.InnerException
      try { if ($inner.HResult) { $hr = [int]$inner.HResult } } catch { }
    }
    # A closed Outlook leaves stale objects behind. Drop everything and reconnect next call.
    $script:app = $null
    $script:ns = $null
    Send-Json ([ordered]@{
      id = $id
      ok = $false
      error = [ordered]@{ code = $hr; message = [string]$ex.Message }
    })
  }
}
`;
