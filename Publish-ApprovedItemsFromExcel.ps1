<#
.SYNOPSIS
    Publishes approved Sitecore items from an Excel workbook.
.DESCRIPTION
    Run this script in the Sitecore PowerShell ISE. The workbook must contain
    headers named "Path" and "Publish Subitems". The Publish Subitems value
    must be Yes or No (case-insensitive) to enable descendant publishing.

    The script publishes approved items only. A workflow state is considered
    approved when it has its Final workflow flag enabled, or its name/display
    name is exactly "Approved".

    English (en) language items are processed and published only.
    Publishing is performed with Smart publish mode.
#>

# This script uses Sitecore PowerShell Extensions dialogs and must be run
# inside the Sitecore PowerShell ISE. Windows Forms is not used because it
# is not supported by the PowerShell web console.

# Windows Forms is intentionally not used because this script runs in the
# Sitecore PowerShell web console.

function Get-NormalizedSitecorePath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $normalized = $Path.Trim()

    if ($normalized -notmatch '^master:/') {
        if ($normalized -notmatch '^master\\:') {
            $normalized = "master:" + $normalized
        }
    }

    return $normalized -replace '\\', '/'
}

function Get-WorkflowStateName {
    param([Sitecore.Data.Items.Item]$Item)

    if ($null -eq $Item) { return "" }
    $stateField = $Item.Fields["__Workflow state"]
    if ($null -eq $stateField -or [string]::IsNullOrWhiteSpace($stateField.Value)) { return "No workflow state" }

    try {
        $stateItem = $Item.Database.GetItem([Sitecore.Data.ID]::Parse($stateField.Value))
        if ($null -eq $stateItem) { return $stateField.Value }
        if (-not [string]::IsNullOrWhiteSpace($stateItem.DisplayName)) { return $stateItem.DisplayName }
        return $stateItem.Name
    }
    catch { return $stateField.Value }
}

function Test-IncludeSubitems {
    param($Value)
    if ($null -eq $Value) { return $false }
    return $Value.ToString().Trim() -match '^(yes|y|true|1)$'
}

function Test-ApprovedWorkflowState {
    param([Sitecore.Data.Items.Item]$Item)

    if ($null -eq $Item) { return $false }

    $stateField = $Item.Fields["__Workflow state"]
    if ($null -eq $stateField -or [string]::IsNullOrWhiteSpace($stateField.Value)) {
        return $false
    }

    try {
        $stateId = [Sitecore.Data.ID]::Parse($stateField.Value)
        $stateItem = $Item.Database.GetItem($stateId)
        if ($null -eq $stateItem) { return $false }

        $finalValue = $stateItem["Final"]
        if ($finalValue -eq "1" -or $finalValue -match '(?i)^true$|^yes$') {
            return $true
        }

        $name = if ($stateItem.DisplayName) { $stateItem.DisplayName } else { $stateItem.Name }
        return $name.Trim() -match '^(?i:approved)$'
    }
    catch {
        return $false
    }
}

function Read-ExcelRows {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "The workbook does not exist: $Path"
    }

    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($Path, 0, $true)
        $sheet = $workbook.Worksheets.Item(1)
        $usedRange = $sheet.UsedRange
        $lastRow = $usedRange.Rows.Count
        $lastColumn = $usedRange.Columns.Count

        if ($lastRow -lt 2) {
            return @()
        }

        $headers = @{}
        for ($column = 1; $column -le $lastColumn; $column++) {
            $header = $sheet.Cells.Item(1, $column).Text
            if (-not [string]::IsNullOrWhiteSpace($header)) {
                $headers[$header.Trim().ToLowerInvariant()] = $column
            }
        }

        if (-not $headers.ContainsKey("path") -or -not $headers.ContainsKey("publish subitems")) {
            throw "Excel headers must include exactly named columns: Path and Publish Subitems."
        }

        $rows = New-Object System.Collections.Generic.List[object]
        for ($row = 2; $row -le $lastRow; $row++) {
            $pathValue = $sheet.Cells.Item($row, $headers["path"]).Text
            $subitemsValue = $sheet.Cells.Item($row, $headers["publish subitems"]).Text
            if ([string]::IsNullOrWhiteSpace($pathValue)) { continue }

            $rows.Add([pscustomobject]@{
                Path = $pathValue.Trim()
                Subitems = $subitemsValue
                Row = $row
            }) | Out-Null
        }

        return $rows
    }
    finally {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) | Out-Null } catch { }
        }
        if ($null -ne $excel) {
            try { $excel.Quit() | Out-Null } catch { }
        }
        if ($null -ne $workbook) {
            [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
        }
        if ($null -ne $excel) {
            [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Read-UploadedExcelRows {
    param([Sitecore.Data.Items.Item]$MediaItem)

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $tempPath = [System.IO.Path]::GetTempFileName() + ".xlsx"

    try {
        $stream = $MediaItem.Fields["Blob"].GetBlobStream()
        try {
            $fileStream = [System.IO.File]::Create($tempPath)
            try { $stream.CopyTo($fileStream) } finally { $fileStream.Close() }
        } finally { $stream.Close() }

        $archive = [System.IO.Compression.ZipFile]::OpenRead($tempPath)
        try {
            $sharedStrings = @()
            $sharedEntry = $archive.GetEntry("xl/sharedStrings.xml")
            if ($null -ne $sharedEntry) {
                $sharedXml = New-Object System.Xml.XmlDocument
                $sharedXml.Load($sharedEntry.Open())
                foreach ($si in $sharedXml.SelectNodes("//*[local-name()='si']")) {
                    $text = ""
                    foreach ($textNode in $si.SelectNodes(".//*[local-name()='t']")) { $text += $textNode.InnerText }
                    $sharedStrings += $text
                }
            }

            $sheetEntry = $archive.Entries | Where-Object { $_.FullName -match '^xl/worksheets/sheet[0-9]+\.xml$' } | Select-Object -First 1
            if ($null -eq $sheetEntry) { throw "The workbook does not contain a readable worksheet." }

            $sheetXml = New-Object System.Xml.XmlDocument
            $sheetXml.Load($sheetEntry.Open())
            $sheetRows = $sheetXml.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")
            if ($sheetRows.Count -lt 2) { throw "The workbook does not contain a header and data rows." }

            $headers = @{}
            $headerRowIndex = -1
            for ($candidateIndex = 0; $candidateIndex -lt [Math]::Min(10, $sheetRows.Count); $candidateIndex++) {
                foreach ($cell in $sheetRows[$candidateIndex].SelectNodes("./*[local-name()='c']")) {
                    if ($cell.GetAttribute("r") -match '^([A-Za-z]+)') {
                        $columnIndex = Get-XlsxColumnIndex -Reference $Matches[1]
                        $headersValue = Get-XlsxCellText -Cell $cell -SharedStrings $sharedStrings
                        if (-not [string]::IsNullOrWhiteSpace([string]$headersValue)) {
                            $headers[$headersValue.Trim().ToLowerInvariant()] = $columnIndex
                        }
                    }
                }
                if ($headers.ContainsKey("path") -and $headers.ContainsKey("publish subitems")) {
                    $headerRowIndex = $candidateIndex
                    break
                }
                $headers.Clear()
            }
            if ($headerRowIndex -lt 0) {
                throw "Excel headers must include Path and Publish Subitems. The headers were not found in the first 10 rows."
            }

            $rows = New-Object System.Collections.Generic.List[object]
            for ($rowIndex = $headerRowIndex + 1; $rowIndex -lt $sheetRows.Count; $rowIndex++) {
                $row = $sheetRows[$rowIndex]
                $values = @{}
                foreach ($cell in $row.SelectNodes("./*[local-name()='c']")) {
                    if ($cell.GetAttribute("r") -match '^([A-Za-z]+)') {
                        $columnIndex = Get-XlsxColumnIndex -Reference $Matches[1]
                        $values[$columnIndex] = Get-XlsxCellText -Cell $cell -SharedStrings $sharedStrings
                    }
                }
                $pathValue = if ($values.ContainsKey($headers["path"])) { $values[$headers["path"]] } else { $null }
                if ([string]::IsNullOrWhiteSpace([string]$pathValue)) { continue }
                $subitemsValue = if ($values.ContainsKey($headers["publish subitems"])) { $values[$headers["publish subitems"]] } else { $null }
                $rowNumber = if ($row.GetAttribute("r")) { $row.GetAttribute("r") } else { $rowIndex + 1 }
                $rows.Add([pscustomobject]@{ Path = $pathValue.Trim(); Subitems = $subitemsValue; Row = $rowNumber }) | Out-Null
            }
            return $rows
        } finally { $archive.Dispose() }
    } finally {
        Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
    }
}

function Get-XlsxColumnIndex {
    param([string]$Reference)
    $index = 0
    foreach ($ch in $Reference.ToUpperInvariant().ToCharArray()) {
        $index = ($index * 26) + ([int][char]$ch - 64)
    }
    return $index
}

function Get-XlsxCellText {
    param($Cell, [string[]]$SharedStrings)

    $type = $Cell.GetAttribute("t")
    $valueNode = $Cell.SelectSingleNode("./*[local-name()='v']")
    $inlineNode = $Cell.SelectSingleNode("./*[local-name()='is']")

    if ($null -ne $inlineNode) {
        $text = ""
        foreach ($textNode in $inlineNode.SelectNodes(".//*[local-name()='t']")) { $text += $textNode.InnerText }
        return $text
    }
    if ($null -eq $valueNode) { return $null }
    if ($type -eq "s") {
        $index = [int]$valueNode.InnerText
        if ($index -ge 0 -and $index -lt $SharedStrings.Count) { return $SharedStrings[$index] }
        return $null
    }
    return $valueNode.InnerText
}

$uploadFolder = Get-Item "master:/sitecore/media library/PowershellFileUpload" -ErrorAction SilentlyContinue
if ($null -eq $uploadFolder) {
    Write-Host "Upload folder missing: create master:/sitecore/media library/PowershellFileUpload and try again." -ForegroundColor Red
    return
}

$excelMedia = Receive-File -ParentItem $uploadFolder -Overwrite
if ($null -eq $excelMedia -or $excelMedia -eq "cancel") {
    Write-Host "Operation cancelled." -ForegroundColor Yellow
    return
}

$targetOptions = [ordered]@{
    web = "Web"
    pub = "Pub"
}
$dialogResult = Read-Variable `
    -Title "Publish approved Sitecore items" `
    -Description "Select one or both publishing target databases." `
    -OkButtonName "Continue" `
    -CancelButtonName "Cancel" `
    -Parameters @(
        @{ Name = "PublishTargets"; Title = "Publishing targets"; Editor = "checklist"; Options = $targetOptions; Tooltip = "Select Web, Pub, or both." }
    )

if ($dialogResult -ne "ok") {
    Write-Host "Operation cancelled." -ForegroundColor Yellow
    return
}

if ($null -eq $PublishTargets -or @($PublishTargets).Count -eq 0) {
    Write-Host "Publishing target required: select Web, Pub, or both." -ForegroundColor Red
    return
}

$selectedTargetNames = @($PublishTargets | ForEach-Object { $_.ToString().Trim().ToLowerInvariant() })
$targets = New-Object System.Collections.Generic.List[object]
$selection = [pscustomobject]@{
    FilePath = [string]$excelMedia.Paths.FullPath
    DryRun = $false
}

foreach ($targetName in $selectedTargetNames) {
    try {
        $database = Get-Database -Name $targetName -ErrorAction Stop
        if ($null -eq $database) { throw "Database not found." }
        $targets.Add([pscustomobject]@{ Name = $targetName; Database = $database }) | Out-Null
    }
    catch {
        Write-Host "Publishing target unavailable: '$targetName' is not configured on this Sitecore instance." -ForegroundColor Red
        return
    }
}

try {
    $rows = @(Read-UploadedExcelRows -MediaItem $excelMedia)
}
catch {
    Write-Host "Unable to read workbook: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}

Write-Host "Workbook rows found: $($rows.Count). Dry run: $($selection.DryRun)" -ForegroundColor Cyan

$processed = 0
$published = 0
$skippedNotApproved = 0
$skippedMissing = 0
$errors = 0
$script:publishReport = New-Object System.Collections.ArrayList

function Add-PublishReport {
    param(
        [int]$ExcelRow,
        [string]$Path,
        [string]$Subitems,
        [string]$Target,
        [string]$Status,
        [string]$Language = "en",
        [string]$WorkflowState = "",
        [string]$Details = ""
    )

    $reportRow = [pscustomobject][ordered]@{
        "Excel Row" = $ExcelRow
        Path = $Path
        "Publish Subitems" = $Subitems
        Target = $Target
        Status = $Status
        Language = $Language
        "Workflow State" = $WorkflowState
        Details = $Details
        "Processed At" = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    [void]$script:publishReport.Add($reportRow)
}

foreach ($row in $rows) {
    $sitecorePath = Get-NormalizedSitecorePath -Path $row.Path
    $item = Get-Item -Path $sitecorePath -Language "en" -ErrorAction SilentlyContinue

    if ($null -eq $item) {
        Write-Host "[ROW $($row.Row) - MISSING] $($row.Path)" -ForegroundColor Red
        $skippedMissing++
        foreach ($target in $targets) {
            Add-PublishReport -ExcelRow $row.Row -Path $row.Path -Subitems ([string]$row.Subitems) -Target $target.Name -Status "Missing" -Details "English item was not found."
        }
        continue
    }

    if (-not (Test-ApprovedWorkflowState -Item $item)) {
        Write-Host "[ROW $($row.Row) - NOT APPROVED] $($item.Paths.FullPath)" -ForegroundColor DarkYellow
        $skippedNotApproved++
        foreach ($target in $targets) {
            Add-PublishReport -ExcelRow $row.Row -Path $row.Path -Subitems ([string]$row.Subitems) -Target $target.Name -Status "Not approved" -WorkflowState (Get-WorkflowStateName -Item $item) -Details "Item was not published because its workflow state is not approved."
        }
        continue
    }

    $includeSubitems = Test-IncludeSubitems -Value $row.Subitems
    $processed++
    Write-Host "[APPROVED] $($item.Paths.FullPath) | Publish subitems: $includeSubitems" -ForegroundColor Green

    foreach ($target in $targets) {
        $targetName = $target.Name
        if ($selection.DryRun) {
            Write-Host "  [DRY RUN] Publish smart to '$targetName'" -ForegroundColor Yellow
        }
        else {
            try {
                Publish-Item -Item $item -Target $target.Database -Language ([Sitecore.Globalization.Language]::Parse("en")) -Recurse:$includeSubitems -PublishMode Smart
                $published++
                Add-PublishReport -ExcelRow $row.Row -Path $item.Paths.FullPath -Subitems ([string]$row.Subitems) -Target $targetName -Status "Published" -WorkflowState (Get-WorkflowStateName -Item $item) -Details "Smart publish completed."
                Write-Host "  [PUBLISHED] Target '$targetName'" -ForegroundColor Green
            }
            catch {
                $errors++
                Add-PublishReport -ExcelRow $row.Row -Path $item.Paths.FullPath -Subitems ([string]$row.Subitems) -Target $targetName -Status "Error" -WorkflowState (Get-WorkflowStateName -Item $item) -Details $_.Exception.Message
                Write-Host "  [ERROR] Target '$targetName': $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}

Write-Host "" 
Write-Host "Completed. Approved processed: $processed | Published operations: $published | Missing: $skippedMissing | Not approved: $skippedNotApproved | Errors: $errors | Dry run: $($selection.DryRun)" -ForegroundColor Cyan

if (@($script:publishReport).Count -gt 0) {
    # Re-selecting the properties avoids SPE/PowerShell treating a single
    # resulting row as a scalar value.
    $reportData = @($script:publishReport | Select-Object "Excel Row", Path, "Publish Subitems", Target, Status, Language, "Workflow State", Details, "Processed At")
    Write-Host "Publishing report rows: $($reportData.Count)" -ForegroundColor Cyan
    $reportData | Show-ListView -Title "Publishing report" -InfoTitle "Publishing results" -InfoDescription "Use the Excel button at the top to download all results." -Property "Excel Row", Path, "Publish Subitems", Target, Status, Language, "Workflow State", Details, "Processed At" -PageSize 25

    $dataFolder = [Sitecore.Configuration.Settings]::DataFolder
    if ([string]::IsNullOrWhiteSpace($dataFolder)) {
        $dataFolder = [System.IO.Path]::GetTempPath()
    }
    $reportFolder = Join-Path $dataFolder "PublishingReports"
    if (-not (Test-Path -LiteralPath $reportFolder)) { New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null }
    $reportPath = Join-Path $reportFolder ("PublishingReport-{0}.csv" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
    try {
        $script:publishReport | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8
        Write-Host "Publishing report created: $reportPath" -ForegroundColor Cyan
        # Show-ListView provides the Excel export button; Send-File is not
        # needed and would display an additional download prompt.
    }
    catch {
        Write-Host "The publishing report could not be downloaded: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Report location: $reportPath" -ForegroundColor Yellow
    }
}
else {
    Write-Host "No publishing report rows were generated." -ForegroundColor Yellow
}
