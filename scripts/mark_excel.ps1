# Fill and comment cells in an Excel working paper, and add an OPEN ITEMS index.
#
#   powershell -File mark_excel.ps1 "book.xlsx" marks.json
#   powershell -File mark_excel.ps1 "book.xlsx" marks.json -Out "book - OPEN ITEMS MARKED.xlsx"
#   powershell -File mark_excel.ps1 "book.xlsx" marks.json -WhatIfOnly
#
# marks.json is a list of:
#   {"sheet": "BS", "cell": "D34", "sev": "OPEN", "ref": "SA-01",
#    "note": "Re footed to 13,059.28. Still not corrected."}
#
# sev takes C H M L (severity) or DONE OPEN NEW (status).
#
# Drives Excel through COM rather than a library. openpyxl reads these workbooks
# correctly and writes them back missing parts: a live three megabyte financial
# reporting set lost 86 relationship parts on a bare load and save, and the file
# still opened. See references/file-integrity.md.
#
# This is slow. Minutes, not seconds, on a workbook of a few megabytes with sixty
# sheets. Run it in the background and do not conclude it has hung. If it fails
# part way it can leave an orphaned EXCEL.EXE holding the file.
#
# -WhatIfOnly resolves every sheet and cell reference and reports what would be
# marked, without opening the workbook for writing. Run it first.

param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$MarksJson,
    [string]$Out,
    [switch]$NoIndex,
    [switch]$WhatIfOnly
)

$ErrorActionPreference = "Stop"

# RGB, matching the palette in mark_pdf.py. Excel wants BGR packed.
$PAL = @{
    "C"    = @{ r = 255; g = 107; b = 107; label = "CRITICAL"; desc = "must be corrected before signing" }
    "H"    = @{ r = 255; g = 168; b = 51;  label = "HIGH";     desc = "correct before circulation" }
    "M"    = @{ r = 255; g = 230; b = 89;  label = "MEDIUM";   desc = "disclosure or presentation" }
    "L"    = @{ r = 158; g = 204; b = 255; label = "LOW";      desc = "typographical" }
    "DONE" = @{ r = 140; g = 217; b = 140; label = "DONE";     desc = "verified as incorporated" }
    "OPEN" = @{ r = 255; g = 192; b = 0;   label = "OPEN";     desc = "previously raised, still not done" }
    "NEW"  = @{ r = 255; g = 107; b = 107; label = "NEW";      desc = "new issue on this version" }
}
function BGR($c) { return ($c.b * 65536) + ($c.g * 256) + $c.r }

$src = (Resolve-Path -LiteralPath $Path).Path
$marks = Get-Content -LiteralPath $MarksJson -Raw -Encoding UTF8 | ConvertFrom-Json
if ($marks -isnot [System.Array]) { $marks = @($marks) }

$bad = $marks | Where-Object { -not $PAL.ContainsKey($_.sev) }
if ($bad) {
    Write-Host "UNKNOWN SEVERITY CODES. Valid: $($PAL.Keys -join ', ')"
    $bad | ForEach-Object { Write-Host ("   {0}  sev={1}" -f $_.ref, $_.sev) }
    exit 2
}

if (-not $Out) {
    $Out = [System.IO.Path]::Combine(
        [System.IO.Path]::GetDirectoryName($src),
        [System.IO.Path]::GetFileNameWithoutExtension($src) + " - OPEN ITEMS MARKED" +
        [System.IO.Path]::GetExtension($src))
}
if ($Out -eq $src) { Write-Host "Refusing to write over the source workbook."; exit 2 }

$xl = $null; $wb = $null; $failed = $false
$applied = @(); $missed = @()

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.ScreenUpdating = $false

    $wb = $xl.Workbooks.Open($src)

    # Calculation cannot be set before a workbook is open: Excel raises
    # 0x800A03EC. Set it here, and set it back before saving.
    $xl.Calculation = -4135            # xlCalculationManual

    # resolve sheets once; names in these workbooks carry stray spaces
    $byName = @{}
    foreach ($ws in $wb.Worksheets) {
        $byName[$ws.Name] = $ws
        $byName[$ws.Name.Trim()] = $ws
    }

    foreach ($m in $marks) {
        $ws = $byName[$m.sheet]
        if (-not $ws) { $ws = $byName[$m.sheet.Trim()] }
        if (-not $ws) { $missed += , @($m, "no sheet by that name"); continue }
        try { $rng = $ws.Range($m.cell) }
        catch { $missed += , @($m, "cell reference not valid"); continue }

        if ($WhatIfOnly) {
            $v = $rng.Text
            Write-Host ("  {0,-22} {1,-8} [{2,-4}] {3,-10} current value: {4}" -f `
                $ws.Name, $m.cell, $m.sev, $m.ref, $v)
            $applied += $m
            continue
        }

        $c = $PAL[$m.sev]
        if ($rng.Interior.ColorIndex -ne -4142 -and $rng.Interior.Color -ne 16777215) {
            # the preparer's own fill means something; say so rather than hiding it
            $missed += , @($m, "cell already filled, left alone")
            continue
        }
        $rng.Interior.Color = (BGR $c)
        $rng.ClearComments() | Out-Null
        $txt = "[{0}] {1}`n{2}`n`n{3}" -f $m.sev, $c.label, $c.desc, $m.note
        if ($m.ref) { $txt = "[{0}] {1} - {2}`n`n{3}" -f $m.sev, $m.ref, $c.label, $m.note }
        $cm = $rng.AddComment($txt)
        $cm.Shape.TextFrame.AutoSize = $true
        $applied += $m
    }

    if ($WhatIfOnly) {
        Write-Host ""
        Write-Host ("  {0} of {1} marks resolve to a real cell" -f $applied.Count, $marks.Count)
    }
    elseif (-not $NoIndex) {
        $idx = $wb.Worksheets.Add($wb.Worksheets.Item(1))
        $idx.Name = "OPEN ITEMS"
        $hdr = @("Ref", "Sheet", "Cell", "Status", "Observation")
        for ($i = 0; $i -lt $hdr.Count; $i++) {
            $idx.Cells.Item(1, $i + 1).Value2 = $hdr[$i]
            $idx.Cells.Item(1, $i + 1).Font.Bold = $true
        }
        $r = 2
        foreach ($m in $applied) {
            $c = $PAL[$m.sev]
            $idx.Cells.Item($r, 1).Value2 = $m.ref
            $idx.Cells.Item($r, 2).Value2 = $m.sheet
            $idx.Cells.Item($r, 3).Value2 = $m.cell
            $idx.Cells.Item($r, 4).Value2 = $c.label
            $idx.Cells.Item($r, 4).Interior.Color = (BGR $c)
            $idx.Cells.Item($r, 5).Value2 = $m.note
            # quoting matters: sheet names here contain spaces
            $target = "'" + $m.sheet + "'!" + $m.cell
            $idx.Hyperlinks.Add($idx.Cells.Item($r, 1), "", $target) | Out-Null
            $r++
        }
        $idx.Range("A1:E1").AutoFilter() | Out-Null
        $idx.Columns.Item(1).ColumnWidth = 10
        $idx.Columns.Item(2).ColumnWidth = 24
        $idx.Columns.Item(3).ColumnWidth = 8
        $idx.Columns.Item(4).ColumnWidth = 11
        $idx.Columns.Item(5).ColumnWidth = 90
        $xl.ActiveWindow.SplitRow = 1
        $xl.ActiveWindow.FreezePanes = $true
    }

    if (-not $WhatIfOnly) {
        $xl.Calculation = -4105        # back to automatic, or the recipient opens a dead workbook
        $wb.SaveAs($Out)
        Write-Host ""
        Write-Host ("  marked   {0} cells across {1} sheets" -f $applied.Count,
            (($applied | ForEach-Object { $_.sheet } | Sort-Object -Unique).Count))
        foreach ($k in @("C", "H", "M", "L", "DONE", "OPEN", "NEW")) {
            $n = ($applied | Where-Object { $_.sev -eq $k }).Count
            if ($n) { Write-Host ("  {0,-9} {1}" -f $PAL[$k].label, $n) }
        }
        Write-Host ("  wrote    {0}" -f $Out)
    }

    if ($missed) {
        Write-Host ""
        Write-Host "NOT MARKED. Tell the user: a cell already carrying the preparer's own"
        Write-Host "fill was left alone, and a missing sheet usually means the workbook"
        Write-Host "was restructured between versions."
        foreach ($x in $missed) {
            Write-Host ("   {0,-10} {1,-22} {2,-8} {3}" -f $x[0].ref, $x[0].sheet, $x[0].cell, $x[1])
        }
    }
}
catch {
    Write-Host ""
    Write-Host ("  FAILED: {0}" -f $_.Exception.Message)
    Write-Host "  Check for an orphaned Excel holding the file before retrying:"
    Write-Host "     Get-Process EXCEL -ErrorAction SilentlyContinue"
    Write-Host "  Only kill it once you are sure the user has no Excel window of their own open."
    $failed = $true
}
finally {
    if ($wb) {
        $wb.Close($false) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        $wb = $null
    }
    if ($xl) {
        $xl.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
        $xl = $null
    }
    # Excel can linger for a few seconds through COM teardown even on a clean
    # run. That is not an orphan. Only treat it as one if it outlives the script.
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
}

if ($failed) { exit 1 }
exit 0
