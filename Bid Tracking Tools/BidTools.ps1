<#
.SYNOPSIS
  Menu-driven bid tools for creating bid folders and updating the bid list workbook.
.NOTES
  If PowerShell blocks script execution, run one of the following first:
    - Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
    - Unblock-File -Path .\BidTools.ps1
#>

param(
  [string]$BidRoot = "S:\Bid Documents 2026",
  [string]$TemplateRoot = "S:\Bid Documents 2026\26000 Proposal Templates\15 - Folder Structure",
  [string]$WorkbookPath = "S:\Bid Documents 2026\26000 Proposal Templates\Bid List.xlsx",
  [string]$WorksheetName = "Bid List"
)

$ErrorActionPreference = "Stop"

$HeaderDefaults = @(
  "Bid Folder",
  "Bid Number",
  "Estimator",
  "Bid Due Date",
  "Customer/GC",
  "Bid Name",
  "Proposal Date",
  "Proposal Amount",
  "Bid Status"
)

function Sanitize-Name([string]$s) {
  if ($null -eq $s) { return "" }
  $s = $s -replace '[\\/:*?"<>|]', ' '
  $s = ($s -replace '\s+', ' ').Trim()
  return $s
}

function Read-YesNoDefaultNo([string]$prompt) {
  while ($true) {
    $raw = (Read-Host "$prompt (Y/N) [N]").Trim()
    if ([string]::IsNullOrWhiteSpace($raw)) { return $false }
    switch -Regex ($raw) {
      '^(y|yes)$' { return $true }
      '^(n|no)$'  { return $false }
      default     { Write-Host "Please enter Y or N (or press Enter for N)." -ForegroundColor Yellow }
    }
  }
}

function Read-NonEmpty([string]$prompt) {
  while ($true) {
    $value = Sanitize-Name (Read-Host $prompt)
    if (![string]::IsNullOrWhiteSpace($value)) { return $value }
    Write-Host "Value is required." -ForegroundColor Yellow
  }
}

function Assert-Paths {
  if (!(Test-Path $BidRoot)) { throw "BidRoot not found: $BidRoot" }
  if (!(Test-Path $TemplateRoot)) { throw "TemplateRoot not found: $TemplateRoot" }
}

function Get-BidFolders {
  Get-ChildItem -Path $BidRoot -Directory -ErrorAction Stop |
    Sort-Object Name
}

function Get-NextBidNumber {
  $max = 0
  foreach ($f in (Get-BidFolders)) {
    if ($f.Name -match '^\s*(\d+)\b') {
      $n = [int]$Matches[1]
      if ($n -gt $max) { $max = $n }
    }
  }
  if ($max -eq 0) { throw "No existing bid-number folders found in: $BidRoot" }
  return ($max + 1)
}

function Normalize-BidDate([string]$bidDateRaw) {
  if ($bidDateRaw -notmatch '^(0?[1-9]|1[0-2])-(0?[1-9]|[12]\d|3[01])$') {
    throw "Bid Date must be in MM-DD format (ex: 12-5 or 12-05). You entered: $bidDateRaw"
  }
  $parts = $bidDateRaw.Split('-')
  $mm = "{0:D2}" -f [int]$parts[0]
  $dd = "{0:D2}" -f [int]$parts[1]
  return "$mm-$dd"
}

function Build-BidFolderName([int]$bidNumber, [string]$initials, [string]$bidDate, [string]$customer, [string]$bidName) {
  $name = "{0} - {1} - {2} - {3} - {4}" -f $bidNumber, $initials, $bidDate, $customer, $bidName
  return (Sanitize-Name $name)
}

function Parse-BidFolderName([string]$folderName) {
  $parts = $folderName -split '\s-\s', 5
  if ($parts.Count -lt 5) { return $null }
  return [pscustomobject]@{
    BidNumber = ($parts[0]).Trim()
    Initials  = ($parts[1]).Trim()
    BidDate   = ($parts[2]).Trim()
    Customer  = ($parts[3]).Trim()
    BidName   = ($parts[4]).Trim()
    Folder    = $folderName
  }
}

function Get-PendingSavePath([string]$path) {
  $directory = Split-Path $path
  $baseName = [System.IO.Path]::GetFileNameWithoutExtension($path)
  $extension = [System.IO.Path]::GetExtension($path)
  $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
  return (Join-Path $directory ("{0} - Pending Update {1}{2}" -f $baseName, $stamp, $extension))
}

function New-ExcelContext([string]$path) {
  if (!(Test-Path $path)) { throw "Workbook not found: $path" }
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $readOnly = $false
  try {
    $workbook = $excel.Workbooks.Open($path, $null, $false)
  }
  catch {
    $readOnly = $true
    $workbook = $excel.Workbooks.Open($path, $null, $true)
  }
  $worksheet = $null
  foreach ($sheet in $workbook.Worksheets) {
    if ($sheet.Name -eq $WorksheetName) {
      $worksheet = $sheet
      break
    }
  }
  if ($null -eq $worksheet) {
    $worksheet = $workbook.Worksheets.Add()
    $worksheet.Name = $WorksheetName
  }
  return [pscustomobject]@{
    Excel = $excel
    Workbook = $workbook
    Worksheet = $worksheet
    ReadOnly = $readOnly
    PendingSavePath = $null
  }
}

function Close-ExcelContext($ctx) {
  if ($null -eq $ctx) { return }
  if ($ctx.ReadOnly) {
    if ($null -ne $ctx.PendingSavePath) {
      $ctx.Workbook.SaveAs($ctx.PendingSavePath)
    }
  }
  else {
    $ctx.Workbook.Save()
  }
  $ctx.Workbook.Close($false)
  $ctx.Excel.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ctx.Worksheet) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ctx.Workbook) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ctx.Excel) | Out-Null
}

function Ensure-Headers($worksheet) {
  $legacyHeaderMap = @{
    "Folder Name" = "Bid Folder"
    "Bid#" = "Bid Number"
    "GC/Owner" = "Customer/GC"
    "Description" = "Bid Name"
    "Due Date" = "Bid Due Date"
    "Status" = "Bid Status"
  }
  $headers = @{}
  $hasAny = $false
  for ($col = 1; $col -le 30; $col++) {
    $value = $worksheet.Cells.Item(1, $col).Text
    if (![string]::IsNullOrWhiteSpace($value)) {
      $canonical = if ($legacyHeaderMap.ContainsKey($value)) { $legacyHeaderMap[$value] } else { $value }
      if ($canonical -ne $value) {
        $worksheet.Cells.Item(1, $col).Value2 = $canonical
      }
      if (-not $headers.ContainsKey($canonical)) {
        $headers[$canonical] = $col
      }
      $hasAny = $true
    }
  }
  if (-not $hasAny) {
    for ($i = 0; $i -lt $HeaderDefaults.Count; $i++) {
      $worksheet.Cells.Item(1, $i + 1).Value2 = $HeaderDefaults[$i]
      $headers[$HeaderDefaults[$i]] = $i + 1
    }
    return $headers
  }

  $needsReorder = $false
  for ($i = 0; $i -lt $HeaderDefaults.Count; $i++) {
    $current = $worksheet.Cells.Item(1, $i + 1).Text
    if ($current -ne $HeaderDefaults[$i]) {
      $needsReorder = $true
      break
    }
  }
  if ($needsReorder) {
    for ($i = 0; $i -lt $HeaderDefaults.Count; $i++) {
      $worksheet.Cells.Item(1, $i + 1).Value2 = $HeaderDefaults[$i]
    }
  }
  $headers = @{}
  for ($i = 0; $i -lt $HeaderDefaults.Count; $i++) {
    $headers[$HeaderDefaults[$i]] = $i + 1
  }
  return $headers
}

function Get-LastRow($worksheet) {
  $used = $worksheet.UsedRange
  if ($null -eq $used) { return 1 }
  return $used.Rows.Count
}

function Get-CellText($worksheet, [int]$row, [int]$col) {
  $value = $worksheet.Cells.Item($row, $col).Text
  if ($null -eq $value) { return "" }
  return $value.ToString().Trim()
}

function Get-RowIndexByBidNumber($worksheet, $headers, [string]$bidNumber) {
  $col = $headers["Bid Number"]
  $lastRow = Get-LastRow $worksheet
  for ($row = 2; $row -le $lastRow; $row++) {
    $value = $worksheet.Cells.Item($row, $col).Text
    if ($value -eq $bidNumber) { return $row }
  }
  return $null
}

function Get-RowIndexByFolder($worksheet, $headers, [string]$folderName) {
  $col = $headers["Bid Folder"]
  $lastRow = Get-LastRow $worksheet
  for ($row = 2; $row -le $lastRow; $row++) {
    $value = $worksheet.Cells.Item($row, $col).Text
    if ($value -eq $folderName) { return $row }
  }
  return $null
}

function Write-Row($worksheet, $headers, $row, $bidInfo) {
  $worksheet.Cells.Item($row, $headers["Bid Folder"]).Value2 = $bidInfo.Folder
  $worksheet.Cells.Item($row, $headers["Bid Number"]).Value2 = $bidInfo.BidNumber
  $worksheet.Cells.Item($row, $headers["Estimator"]).Value2 = $bidInfo.Initials
  $worksheet.Cells.Item($row, $headers["Bid Due Date"]).Value2 = $bidInfo.BidDate
  $worksheet.Cells.Item($row, $headers["Customer/GC"]).Value2 = $bidInfo.Customer
  $worksheet.Cells.Item($row, $headers["Bid Name"]).Value2 = $bidInfo.BidName
}

function Sync-BidWorkbook {
  if (!(Test-Path $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
  $ctx = New-ExcelContext -path $WorkbookPath
  try {
    if ($ctx.ReadOnly) {
      $ctx.PendingSavePath = Get-PendingSavePath -path $WorkbookPath
    }
    $worksheet = $ctx.Worksheet
    $headers = Ensure-Headers $worksheet
    $lastRow = Get-LastRow $worksheet

    foreach ($folder in (Get-BidFolders)) {
      $info = Parse-BidFolderName $folder.Name
      if ($null -eq $info) { continue }

      $row = Get-RowIndexByBidNumber $worksheet $headers $info.BidNumber
      if ($null -eq $row) {
        $row = Get-RowIndexByFolder $worksheet $headers $info.Folder
      }
      if ($null -eq $row) {
        $lastRow++
        $row = $lastRow
      }
      Write-Row $worksheet $headers $row $info
    }
  }
  finally {
    Close-ExcelContext $ctx
  }

  if ($ctx.ReadOnly -and $null -ne $ctx.PendingSavePath) {
    Write-Host "Workbook is open by another user; saved updates to:" -ForegroundColor Yellow
    Write-Host $ctx.PendingSavePath -ForegroundColor Yellow
  }
  else {
    Write-Host "Workbook updated with current bid folders." -ForegroundColor Green
  }
}

function Update-BidStatus {
  if (!(Test-Path $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
  $bidNumber = Read-NonEmpty "Enter bid number to update"
  $status = (Read-Host "Status (leave blank to keep current)").Trim()
  $award = ""
  $proposalAmount = ""
  $proposalDate = ""

  $ctx = New-ExcelContext -path $WorkbookPath
  try {
    if ($ctx.ReadOnly) {
      $ctx.PendingSavePath = Get-PendingSavePath -path $WorkbookPath
    }
    $worksheet = $ctx.Worksheet
    $headers = Ensure-Headers $worksheet
    $row = Get-RowIndexByBidNumber $worksheet $headers $bidNumber
    if ($null -eq $row) { throw "Bid number not found in workbook: $bidNumber" }

    if ($headers.ContainsKey("Proposal Amount")) {
      $currentAmount = Get-CellText $worksheet $row $headers["Proposal Amount"]
      if (![string]::IsNullOrWhiteSpace($currentAmount)) {
        if (Read-YesNoDefaultNo ("Proposal Amount is '{0}'. Update it?" -f $currentAmount)) {
          $proposalAmount = (Read-Host "Proposal Amount").Trim()
        }
      }
      else {
        $proposalAmount = (Read-Host "Proposal Amount (leave blank to skip)").Trim()
      }
    }

    if ($headers.ContainsKey("Proposal Date")) {
      $currentDate = Get-CellText $worksheet $row $headers["Proposal Date"]
      if (![string]::IsNullOrWhiteSpace($currentDate)) {
        if (Read-YesNoDefaultNo ("Proposal Date is '{0}'. Update it?" -f $currentDate)) {
          $proposalDate = (Read-Host "Proposal Date (MM-DD or date string)").Trim()
        }
      }
      else {
        $proposalDate = (Read-Host "Proposal Date (leave blank to skip)").Trim()
      }
    }

    if ($headers.ContainsKey("Award")) {
      $award = (Read-Host "Award (leave blank to keep current)").Trim()
    }
    if (-not [string]::IsNullOrWhiteSpace($proposalAmount)) {
      $worksheet.Cells.Item($row, $headers["Proposal Amount"]).Value2 = $proposalAmount
    }
    if (-not [string]::IsNullOrWhiteSpace($proposalAmount) -and $headers.ContainsKey("Proposal Amount")) {
      $worksheet.Cells.Item($row, $headers["Proposal Amount"]).Value2 = $proposalAmount
    }
    if (-not [string]::IsNullOrWhiteSpace($proposalDate) -and $headers.ContainsKey("Proposal Date")) {
      $worksheet.Cells.Item($row, $headers["Proposal Date"]).Value2 = $proposalDate
    }
    if (-not [string]::IsNullOrWhiteSpace($award) -and $headers.ContainsKey("Award")) {
      $worksheet.Cells.Item($row, $headers["Award"]).Value2 = $award
    }
  }
  finally {
    Close-ExcelContext $ctx
  }

  if ($ctx.ReadOnly -and $null -ne $ctx.PendingSavePath) {
    Write-Host "Workbook is open by another user; saved updates to:" -ForegroundColor Yellow
    Write-Host $ctx.PendingSavePath -ForegroundColor Yellow
  }
  else {
    Write-Host "Workbook status updated." -ForegroundColor Green
  }
}

function New-BidFolder {
  Assert-Paths

  Sync-BidWorkbook

  $initials   = Read-NonEmpty "Estimator initials (ex: MD)"
  $bidDateRaw = Read-NonEmpty "Bid Due Date (MM-DD, ex: 12-5)"
  $customer   = Read-NonEmpty "Customer/GC"
  $bidName    = Read-NonEmpty "Bid Name"

  $bidDate = Normalize-BidDate $bidDateRaw
  $newNum = Get-NextBidNumber

  $newFolderName = Build-BidFolderName $newNum $initials $bidDate $customer $bidName
  $dest = Join-Path $BidRoot $newFolderName
  if (Test-Path $dest) { throw "Destination already exists: $dest" }

  New-Item -Path $dest -ItemType Directory | Out-Null

  $copyTemplate = Read-YesNoDefaultNo "Copy subfolder structure from the template?"
  if ($copyTemplate) {
    Copy-Item -Path (Join-Path $TemplateRoot '*') -Destination $dest -Recurse -Force -Exclude 'Thumbs.db'
    Get-ChildItem -Path $dest -Filter "Thumbs.db" -Recurse -Force -ErrorAction SilentlyContinue |
      Remove-Item -Force -ErrorAction SilentlyContinue
  }

  Write-Host "" 
  Write-Host "Created new bid folder:" -ForegroundColor Green
  Write-Host $dest -ForegroundColor Green
  Write-Host ""

  if (Read-YesNoDefaultNo "Update the bid list workbook now?") {
    Sync-BidWorkbook
  }

  Start-Process explorer.exe $dest
}

function Show-Menu {
  Write-Host "" 
  Write-Host "Bid Tools" -ForegroundColor Cyan
  Write-Host "1) Create new bid folder"
  Write-Host "2) Sync bid list workbook with folders"
  Write-Host "3) Update bid status in workbook"
  Write-Host "4) Exit"
}

:main while ($true) {
  Show-Menu
  $choice = (Read-Host "Choose an option (1-4)").Trim()
  switch ($choice) {
    '1' { New-BidFolder }
    '2' { Sync-BidWorkbook }
    '3' { Update-BidStatus }
    '4' { break main }
    default { Write-Host "Invalid option. Choose 1-4." -ForegroundColor Yellow }
  }
}
