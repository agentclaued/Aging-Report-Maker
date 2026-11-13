from __future__ import annotations

import shutil
import subprocess
import tempfile
import time
from pathlib import Path

import pandas as pd

from app import progress


_REQUIRED_FIELDS = (
    "Order#",
    "Name",
    "Ref Type",
    "Total",
    "Current",
    "1-30 Days",
    "31-60 Days",
    "61-90 Days",
    "Over 90 Days",
)


def _column_key(value: object) -> str:
    return " ".join(str(value).strip().casefold().split())


def _powershell_literal(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def pivot_skipped_types(ageing_rows: pd.DataFrame) -> list[tuple[str, int]]:
    """Ref Types (IN, CR, DB) whose rows are too few for Excel to build a PivotTable (needs 2+)."""
    ref_column = next((column for column in ageing_rows.columns if _column_key(column) == "ref type"), None)
    if ref_column is None:
        return []
    counts = ageing_rows[ref_column].astype(str).str.strip().str.upper().value_counts()
    return [(ref_type, int(counts.get(ref_type, 0))) for ref_type in ("IN", "CR", "DB") if counts.get(ref_type, 0) < 2]


def run_regular_accounts_pivot_process(workbook_path: Path, ageing_rows: pd.DataFrame, token: str = "") -> None:
    """Finish a saved Regular Accounts workbook in a separate hidden Excel process.

    This runs after report generation and before the browser download. The app-host's
    Excel opens the temporary XLSX, adds the INV/DB summary and DB reason native
    PivotTables, saves the same file, then closes.
    """
    header_lookup = {_column_key(column): str(column) for column in ageing_rows.columns}
    missing_fields = [field for field in _REQUIRED_FIELDS if _column_key(field) not in header_lookup]
    if missing_fields:
        raise ValueError(
            "Regular Accounts PivotTables require these AGEING SHEET columns: "
            + ", ".join(missing_fields)
            + "."
        )

    source_address = f"'AGEING SHEET'!R1C1:R{len(ageing_rows) + 1}C{len(ageing_rows.columns)}"
    pivot_output_max_rows = len(ageing_rows) + 8
    skipped_types = {ref_type for ref_type, _ in pivot_skipped_types(ageing_rows)}
    skip_in, skip_cr, skip_db = ("$true" if ref_type in skipped_types else "$false" for ref_type in ("IN", "CR", "DB"))
    # Excel works on a private local copy (short unique name, local temp folder) so a locked,
    # scanned, shared or still-downloading final file cannot make Workbooks.Open fail.
    work_dir = Path(tempfile.mkdtemp(prefix="regular_pivot_"))
    work_path = work_dir / "report.xlsx"
    shutil.copyfile(workbook_path, work_path)
    workbook_literal = _powershell_literal(str(work_path.resolve()))
    source_address_literal = _powershell_literal(source_address)
    script = f"""
$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
$pivotCache = $null
$pivotOutputMaxRows = {pivot_output_max_rows}
$skipIn = {skip_in}
$skipCr = {skip_cr}
$skipDb = {skip_db}

function Add-RegularAccountsPivot {{
    param(
        [object]$worksheet,
        [string]$destination,
        [string]$pivotName,
        [string]$rowFieldName,
        [string]$refTypeValue,
        [string]$title = '',
        [bool]$includeName = $false,
        [bool]$showRefTypeInRows = $false,
        [bool]$useRefTypeFilter = $true
    )
    if ($title) {{
        $worksheet.Range($destination).Offset(-2, 0).Value2 = $title
        $worksheet.Range($destination).Offset(-2, 0).Font.Bold = $true
        $worksheet.Range($destination).Offset(-2, 0).Font.Size = 14
    }}
    $destinationRange = $worksheet.Range($destination)
    # Office 2021 rejects a PivotTable when any cell in its eventual output
    # area is merged, not only when the anchor itself is merged. The table can
    # have at most one row per source row, so unmerge its complete reserved
    # output area without affecting title rows above the destination.
    $pivotEndRow = [Math]::Min(
        $worksheet.Rows.Count,
        $destinationRange.Row + $pivotOutputMaxRows
    )
    $pivotEndColumn = [Math]::Min(
        $worksheet.Columns.Count,
        $destinationRange.Column + 11
    )
    $worksheet.Range(
        $destinationRange,
        $worksheet.Cells($pivotEndRow, $pivotEndColumn)
    ).UnMerge()
    [void]$pivotCache.CreatePivotTable($destinationRange, $pivotName)
    $pivot = $worksheet.PivotTables().Item($pivotName)
    $pivot.ManualUpdate = $true
    $rowField = $pivot.PivotFields().Item($rowFieldName)
    $rowField.Orientation = 1
    $rowField.Position = 1
    $rowField.LayoutForm = 1
    $rowField.RepeatLabels = $true
    $rowField.Subtotals = @($false) * 12
    if ($includeName) {{
        $nameField = $pivot.PivotFields().Item('Name')
        $nameField.Orientation = 1
        $nameField.Position = 2
        $nameField.LayoutForm = 1
        $nameField.RepeatLabels = $true
        $nameField.Subtotals = @($false) * 12
    }}
    $refTypeField = $pivot.PivotFields().Item('Ref Type')
    if ($showRefTypeInRows) {{
        $refTypeField.Orientation = 1
        $refTypeField.Position = if ($includeName) {{ 3 }} else {{ 2 }}
        $refTypeField.LayoutForm = 1
        $refTypeField.RepeatLabels = $true
        $refTypeField.Subtotals = @($false) * 12
    }} elseif ($useRefTypeFilter) {{
        $refTypeField.Orientation = 3
        $refTypeField.Position = 1
        $refTypeField.CurrentPage = $refTypeValue
    }}
    foreach ($valueFieldName in @('Current', '1-30 Days', '31-60 Days', '61-90 Days', 'Over 90 Days', 'Total')) {{
        $dataField = $pivot.AddDataField($pivot.PivotFields().Item($valueFieldName), "Sum of $valueFieldName", -4157)
        $dataField.NumberFormat = '$#,##0.00;[Red]($#,##0.00)'
    }}
    $pivot.RowGrand = $true
    $pivot.ColumnGrand = $true
    $pivot.RowAxisLayout(1)
    $pivot.ShowDrillIndicators = $true
    $pivot.EnableDrilldown = $true
    $pivot.TableStyle2 = 'PivotStyleMedium9'
    $pivot.ManualUpdate = $false
    # Remove Excel's separate Values row from the PivotTable layout.
    $pivot.ShowValuesRow = $false
    $tableRange = $pivot.TableRange1
    $pivotHeaderRange = $tableRange.Rows.Item(1)
    $pivotGrandTotalRange = $tableRange.Rows.Item($tableRange.Rows.Count)
    $pivotGrandTotalRange.Interior.Color = $pivotHeaderRange.Interior.Color
    $pivotGrandTotalRange.Font.Color = $pivotHeaderRange.Font.Color
    $pivotGrandTotalRange.Font.Bold = $true
    if ($showRefTypeInRows) {{
        foreach ($pivotItem in @($refTypeField.PivotItems())) {{
            if ([string]$pivotItem.Name -ne $refTypeValue) {{
                $pivotItem.Visible = $false
            }}
        }}
    }}
}}

function New-RegularAccountsPivotSource {{
    param([object]$sourceSheet, [string]$sheetName, [string]$refTypeValue)
    $sourceLastRow = $sourceSheet.Cells($sourceSheet.Rows.Count, 1).End(-4162).Row
    $sourceLastColumn = $sourceSheet.Cells(1, $sourceSheet.Columns.Count).End(-4159).Column
    $refTypeColumn = 0
    for ($columnNumber = 1; $columnNumber -le $sourceLastColumn; $columnNumber++) {{
        if ([string]$sourceSheet.Cells(1, $columnNumber).Value2 -eq 'Ref Type') {{
            $refTypeColumn = $columnNumber
            break
        }}
    }}
    if ($refTypeColumn -eq 0) {{ throw "AGEING SHEET is missing Ref Type." }}
    $targetSheet = $workbook.Worksheets.Add()
    $targetSheet.Name = $sheetName
    $targetSheet.Range($targetSheet.Cells(1, 1), $targetSheet.Cells(1, $sourceLastColumn)).Value2 =
        $sourceSheet.Range($sourceSheet.Cells(1, 1), $sourceSheet.Cells(1, $sourceLastColumn)).Value2
    $targetRow = 2
    if ($sourceLastRow -ge 2) {{
        # One Excel read, filter in memory, one Excel write (row-by-row Excel calls were the slowest step).
        $allValues = $sourceSheet.Range($sourceSheet.Cells(1, 1), $sourceSheet.Cells($sourceLastRow, $sourceLastColumn)).Value2
        $matchRows = New-Object System.Collections.Generic.List[int]
        for ($sourceRow = 2; $sourceRow -le $sourceLastRow; $sourceRow++) {{
            if ([string]$allValues[$sourceRow, $refTypeColumn] -eq $refTypeValue) {{ $matchRows.Add($sourceRow) }}
        }}
        $matchCount = $matchRows.Count
        if ($matchCount -gt 0) {{
            $outValues = New-Object 'object[,]' $matchCount, $sourceLastColumn
            for ($i = 0; $i -lt $matchCount; $i++) {{
                $sourceRow = $matchRows[$i]
                for ($column = 1; $column -le $sourceLastColumn; $column++) {{
                    $outValues[$i, ($column - 1)] = $allValues[$sourceRow, $column]
                }}
            }}
            $targetSheet.Range($targetSheet.Cells(2, 1), $targetSheet.Cells($matchCount + 1, $sourceLastColumn)).Value2 = $outValues
            $targetRow = $matchCount + 2
        }}
    }}
    $targetSheet.Visible = 2
    return "'$sheetName'!R1C1:R$($targetRow - 1)C$sourceLastColumn"
}}

function Set-RegularAccountsPivotGrandTotalStyle {{
    param([object]$worksheet, [object]$pivot, [int]$lastColumn = 0)
    # Excel's PivotTable style can reset TableRange formatting. Its final row
    # is the visible Grand Total row, even when Excel leaves its text label blank.
    $tableRange = $pivot.TableRange2
    $firstColumn = $tableRange.Column
    $finalColumn = $tableRange.Column + $tableRange.Columns.Count - 1
    if ($lastColumn -gt $finalColumn) {{ $finalColumn = $lastColumn }}
    $grandTotalRow = $tableRange.Row + $tableRange.Rows.Count - 1
    $grandTotalRange = $worksheet.Range(
        $worksheet.Cells($grandTotalRow, $firstColumn),
        $worksheet.Cells($grandTotalRow, $finalColumn)
    )
    $grandTotalRange.Interior.ThemeColor = 5  # Blue, Accent 1.
    $grandTotalRange.Interior.TintAndShade = 0
    $grandTotalRange.Font.Color = 16777215  # White.
    $grandTotalRange.Font.Bold = $true
}}

function Get-RegularAccountsInvExtras {{
    param([object]$worksheet)
    $extras = @{{}}
    $lastRow = $worksheet.Cells($worksheet.Rows.Count, 1).End(-4162).Row
    for ($rowNumber = 8; $rowNumber -le $lastRow; $rowNumber++) {{
        $account = [string]$worksheet.Cells($rowNumber, 1).Value2
        if (-not $account -or $account -eq 'Grand Total') {{ continue }}
        $extras[$account] = @{{
            Terms = $worksheet.Cells($rowNumber, 10).Value2
            Salesman = $worksheet.Cells($rowNumber, 11).Value2
        }}
    }}
    return $extras
}}

function Restore-RegularAccountsInvExtras {{
    param([object]$worksheet, [object]$pivot, [hashtable]$extras)
    # With ShowValuesRow disabled, TableRange1 starts on the visible
    # Acct/Name/Ref Type header.
    # Companion columns begin immediately after the PivotTable data.
    $tableRange = $pivot.TableRange1
    $headerRow = $tableRange.Row
    $lastPivotRow = $tableRange.Row + $tableRange.Rows.Count - 1
    $companionStartColumn = $tableRange.Column + $tableRange.Columns.Count
    $termsColumn = $companionStartColumn
    $salesmanColumn = $companionStartColumn + 1
    $pivotHeaderStyleCell = $worksheet.Cells($headerRow, $companionStartColumn - 1)
    foreach ($columnNumber in $termsColumn..$salesmanColumn) {{
        $headerCell = $worksheet.Cells($headerRow, $columnNumber)
        $pivotHeaderStyleCell.Copy($headerCell)
    }}
    $worksheet.Cells($headerRow, $termsColumn).Value2 = 'TERMS#'
    $worksheet.Cells($headerRow, $salesmanColumn).Value2 = 'SALEMAN'
    $worksheet.Columns($termsColumn).AutoFit() | Out-Null
    $worksheet.Columns($salesmanColumn).AutoFit() | Out-Null

    for ($rowNumber = $headerRow + 1; $rowNumber -le $lastPivotRow; $rowNumber++) {{
        $account = [string]$worksheet.Cells($rowNumber, 1).Value2
        if ($account -eq 'Grand Total') {{
            $grandTotalExtrasRange = $worksheet.Range($worksheet.Cells($rowNumber, $termsColumn), $worksheet.Cells($rowNumber, $salesmanColumn))
            $grandTotalExtrasRange.Interior.ThemeColor = 5  # Blue, Accent 1.
            $grandTotalExtrasRange.Interior.TintAndShade = 0
            $grandTotalExtrasRange.Font.Color = 16777215  # White.
            $grandTotalExtrasRange.Font.Bold = $true
            continue
        }}
        if (-not $extras.ContainsKey($account)) {{ continue }}
        $entry = $extras[$account]
        $worksheet.Cells($rowNumber, $termsColumn).Value2 = $entry.Terms
        $worksheet.Cells($rowNumber, $salesmanColumn).Value2 = $entry.Salesman
    }}
}}

$sysDesktops = @("$env:windir\\System32\\config\\systemprofile\\Desktop", "$env:windir\\SysWOW64\\config\\systemprofile\\Desktop")
try {{
    # Excel started without a desktop (session 0: boot task / service) fails Workbooks.Open with
    # "cannot access the file" unless these Desktop folders exist. Create them when missing.
    foreach ($sysDesktop in $sysDesktops) {{
        try {{
            if (-not (Test-Path -LiteralPath $sysDesktop)) {{ [void](New-Item -ItemType Directory -Path $sysDesktop -Force) }}
        }} catch {{ }}
    }}
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    for ($attempt = 1; $attempt -le 5; $attempt++) {{
        try {{
            $workbook = $excel.Workbooks.Open({workbook_literal}, 0, $false)
            break
        }} catch {{
            if ($attempt -eq 5) {{ throw }}
            Start-Sleep -Seconds 2
        }}
    }}

    $invSheet = $workbook.Worksheets.Item('INV SUMMARY')
    $invExtras = Get-RegularAccountsInvExtras $invSheet
    $invLastRow = $invSheet.Cells($invSheet.Rows.Count, 1).End(-4162).Row
    if ($invLastRow -ge 6) {{
        $invSheet.Range("A6:L$invLastRow").Clear()
    }}

    $crSheet = $workbook.Worksheets.Item("CR'S SUMMARY")
    $crLastRow = $crSheet.Cells($crSheet.Rows.Count, 1).End(-4162).Row
    if ($crLastRow -ge 6) {{
        $crSheet.Range("A6:I$crLastRow").Clear()
    }}

    $dbSheet = $workbook.Worksheets.Item("DB'S SUMMARY")
    $dbReasonTitleRow = 0
    $dbUsedLastRow = $dbSheet.Cells($dbSheet.Rows.Count, 2).End(-4162).Row
    for ($rowNumber = 1; $rowNumber -le $dbUsedLastRow; $rowNumber++) {{
        if ($dbSheet.Cells($rowNumber, 2).Value2 -eq "DB'S BY REASON") {{
            $dbReasonTitleRow = $rowNumber
            break
        }}
    }}
    $dbClearLastRow = if ($dbReasonTitleRow -gt 0) {{ $dbReasonTitleRow - 1 }} else {{ $dbSheet.Cells($dbSheet.Rows.Count, 1).End(-4162).Row }}
    if ($dbClearLastRow -ge 7) {{
        $dbSheet.Range("A7:I$dbClearLastRow").Clear()
    }}
    if ($dbReasonTitleRow -gt 0) {{
        # Preserve the existing title and spacer row, then replace only the
        # former static DB'S BY REASON table with its native PivotTable.
        $dbReasonHeaderRow = $dbReasonTitleRow + 3
        $dbReasonLastRow = $dbSheet.Cells($dbSheet.Rows.Count, 1).End(-4162).Row
        if ($dbReasonLastRow -ge $dbReasonHeaderRow) {{
            $dbSheet.Range("A$($dbReasonHeaderRow):I$($dbReasonLastRow)").Clear()
        }}
    }}

    $ageingSourceSheet = $workbook.Worksheets.Item('AGEING SHEET')
    $invSourceAddress = New-RegularAccountsPivotSource $ageingSourceSheet '_RegularAccounts_IN_Pivot' 'IN'
    $crSourceAddress = New-RegularAccountsPivotSource $ageingSourceSheet '_RegularAccounts_CR_Pivot' 'CR'
    $dbSourceAddress = New-RegularAccountsPivotSource $ageingSourceSheet '_RegularAccounts_DB_Pivot' 'DB'

    if (-not $skipIn) {{
        $pivotCache = $workbook.PivotCaches().Create(1, $invSourceAddress)
        $pivotCache.RefreshOnFileOpen = $true
        Add-RegularAccountsPivot $invSheet 'A7' 'InvSummaryPivot' 'Acct' 'IN' '' $true $true $false
        $invPivot = $invSheet.PivotTables().Item('InvSummaryPivot')
        Restore-RegularAccountsInvExtras $invSheet $invPivot $invExtras
        Set-RegularAccountsPivotGrandTotalStyle $invSheet $invPivot 11
        $invSheet.Activate()
        $excel.ActiveWindow.FreezePanes = $false
        $excel.ActiveWindow.SplitColumn = 0
        $excel.ActiveWindow.SplitRow = 0
    }}
    if (-not $skipCr) {{
        $pivotCache = $workbook.PivotCaches().Create(1, $crSourceAddress)
        $pivotCache.RefreshOnFileOpen = $true
        Add-RegularAccountsPivot $crSheet 'A7' 'CrSummaryPivot' 'Acct' 'CR' '' $true $true $false
        Set-RegularAccountsPivotGrandTotalStyle $crSheet ($crSheet.PivotTables().Item('CrSummaryPivot'))
        $crSheet.Activate()
        $excel.ActiveWindow.FreezePanes = $false
        $excel.ActiveWindow.SplitColumn = 0
        $excel.ActiveWindow.SplitRow = 0
    }}
    if (-not $skipDb) {{
        $pivotCache = $workbook.PivotCaches().Create(1, $dbSourceAddress)
        $pivotCache.RefreshOnFileOpen = $true
        Add-RegularAccountsPivot $dbSheet 'A7' 'DbSummaryPivot' 'Acct' 'DB' '' $true $true $false
        Set-RegularAccountsPivotGrandTotalStyle $dbSheet ($dbSheet.PivotTables().Item('DbSummaryPivot'))
        $dbSheet.Activate()
        $excel.ActiveWindow.FreezePanes = $false
        $excel.ActiveWindow.SplitColumn = 0
        $excel.ActiveWindow.SplitRow = 0
    }}
    if ($dbReasonTitleRow -gt 0 -and -not $skipDb) {{
        $pivotCache = $workbook.PivotCaches().Create(1, {source_address_literal})
        $pivotCache.RefreshOnFileOpen = $true
        Add-RegularAccountsPivot $dbSheet "A$($dbReasonTitleRow + 3)" 'DBsByReasonPivot' 'Order#' 'DB' '' $false $true
        $dbReasonPivot = $dbSheet.PivotTables().Item('DBsByReasonPivot')
        $dbReasonRange = $dbReasonPivot.TableRange1
        $dbSheet.Range(
            $dbSheet.Cells($dbReasonRange.Row + 1, $dbReasonRange.Column),
            $dbSheet.Cells($dbReasonRange.Row + $dbReasonRange.Rows.Count - 1, $dbReasonRange.Column)
        ).Font.Bold = $true
        Set-RegularAccountsPivotGrandTotalStyle $dbSheet $dbReasonPivot
        # Match only the DB'S BY REASON title band to DB'S BY AGEING.
        $dbAgeingTitleRange = $dbSheet.Cells(5, 2).MergeArea
        $dbReasonTitleRange = $dbSheet.Cells($dbReasonTitleRow, 2).Resize(1, $dbAgeingTitleRange.Columns.Count)
        $dbReasonTitleRange.UnMerge()
        [void]$dbReasonTitleRange.Merge()
        $dbAgeingTitleRange.Copy($dbReasonTitleRange)
        $dbSheet.Cells($dbReasonTitleRow, 2).Value2 = "DB'S BY REASON"
        foreach ($borderIndex in 7..12) {{
            $dbReasonTitleRange.Borders.Item($borderIndex).LineStyle = -4142
        }}
    }}
    # Remove only the remaining black outline from the B5 merged title bands.
    foreach ($summarySheet in @($invSheet, $crSheet, $dbSheet)) {{
        $summaryTitleRange = $summarySheet.Cells(5, 2).MergeArea
        foreach ($borderIndex in 7..12) {{
            $summaryTitleRange.Borders.Item($borderIndex).LineStyle = -4142
        }}
    }}
    $workbook.Save()
}}
catch {{
    $diagUser = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $diagSession = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
    $diagAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $diagExists = Test-Path -LiteralPath {workbook_literal}
    $diagExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count
    $diagDesktops = @($sysDesktops | Where-Object {{ Test-Path -LiteralPath $_ }}).Count
    throw ("[DIAG user=$diagUser session=$diagSession admin=$diagAdmin fileExists=$diagExists excelProcs=$diagExcel sysDesktops=$diagDesktops/2] " + $_.Exception.Message)
}}
finally {{
    if ($workbook -ne $null) {{ $workbook.Close($false) }}
    if ($excel -ne $null) {{ $excel.Quit() }}
}}
"""
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".ps1", prefix="regular_accounts_pivot_", encoding="utf-8-sig", delete=False
    ) as script_file:
        script_file.write(script)
        script_path = Path(script_file.name)
    process = subprocess.Popen(
        [
            r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
            "-NoLogo",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_path),
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        creationflags=creationflags,
    )
    progress.register_process(token, process)
    try:
        try:
            stdout_text, stderr_text = process.communicate(timeout=300)
        except subprocess.TimeoutExpired:
            process.kill()
            process.communicate()
            raise RuntimeError("Excel PivotTable process timed out after 300 seconds.")
    finally:
        progress.release_process(token)
        script_path.unlink(missing_ok=True)
    try:
        if progress.is_cancelled(token):
            raise progress.JobCancelled("Cancelled by user")
        if process.returncode != 0:
            message = (stderr_text or stdout_text or "Unknown Excel automation error.").strip()
            raise RuntimeError(f"Excel could not complete the Regular Accounts PivotTable process: {message}")
        # Copy the finished workbook back over the final file, retrying if it is briefly locked.
        for attempt in range(1, 8):
            try:
                shutil.copyfile(work_path, workbook_path)
                break
            except PermissionError:
                if attempt == 7:
                    raise RuntimeError(
                        f"The finished report could not be saved because '{workbook_path.name}' is locked. "
                        "Close it on every PC and run again."
                    )
                time.sleep(2)
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
