# Open a redlined document in Word and report what Word itself sees.
#
#   powershell -File verify_docx.ps1 "in - REDLINED.docx"
#   powershell -File verify_docx.ps1 "in - REDLINED.docx" -ExpectRevisions 7 -ExpectComments 3
#
# Never verify a redline by reading the XML you just wrote. A malformed revision
# opens cleanly and shows nothing: w:t instead of w:delText inside a deletion is
# the classic, and it produces a document that reports no findings at all.
#
# Also exports a PDF with the markup rendered (Item = 7,
# wdExportDocumentWithMarkup). The default export shows the document as if every
# change were accepted, which looks clean and tells the reader nothing.

param(
    [Parameter(Mandatory = $true)][string]$Path,
    [int]$ExpectRevisions = -1,
    [int]$ExpectComments = -1,
    [switch]$NoPdf
)

$ErrorActionPreference = "Stop"
$full = (Resolve-Path -LiteralPath $Path).Path
$word = $null
$doc = $null
$fail = $false

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    # ReadOnly, AddToRecentFiles off, no conversion prompt
    $doc = $word.Documents.Open($full, $false, $true, $false)

    $rev = $doc.Revisions.Count
    $cmt = $doc.Comments.Count
    Write-Host ""
    Write-Host ("  file           {0}" -f (Split-Path $full -Leaf))
    Write-Host ("  pages          {0}" -f $doc.ComputeStatistics(2))
    Write-Host ("  revisions      {0}" -f $rev)
    Write-Host ("  comments       {0}" -f $cmt)

    if ($rev -gt 0) {
        $byAuthor = @{}
        foreach ($r in $doc.Revisions) {
            $a = $r.Author
            if (-not $byAuthor.ContainsKey($a)) { $byAuthor[$a] = 0 }
            $byAuthor[$a]++
        }
        Write-Host "  revision authors"
        foreach ($k in $byAuthor.Keys) { Write-Host ("     {0,-32} {1}" -f $k, $byAuthor[$k]) }
    }

    if ($ExpectRevisions -ge 0 -and $rev -lt $ExpectRevisions) {
        Write-Host ""
        Write-Host ("  FEWER REVISIONS THAN EDITS SUPPLIED ({0} against {1})." -f $rev, $ExpectRevisions)
        Write-Host "  Usual cause: w:t instead of w:delText inside a w:del. Do not send this file."
        $fail = $true
    }
    if ($ExpectComments -ge 0 -and $cmt -lt $ExpectComments) {
        Write-Host ""
        Write-Host ("  FEWER COMMENTS THAN SUPPLIED ({0} against {1})." -f $cmt, $ExpectComments)
        Write-Host "  Usual cause: the anchor spans a run boundary. Run merge_runs.py first."
        $fail = $true
    }

    if (-not $NoPdf) {
        $pdf = [System.IO.Path]::ChangeExtension($full, ".pdf")
        $pdf = $pdf -replace "\.pdf$", " - markup.pdf"
        # 17 = wdExportFormatPDF, Item 7 = wdExportDocumentWithMarkup
        $doc.ExportAsFixedFormat($pdf, 17, $false, 0, 0, 0, 0, 7, $true, $true, 0)
        Write-Host ""
        Write-Host ("  markup pdf     {0}" -f (Split-Path $pdf -Leaf))
    }
}
catch {
    Write-Host ""
    Write-Host ("  WORD COULD NOT OPEN OR PROCESS THE FILE: {0}" -f $_.Exception.Message)
    Write-Host "  If Word offered to repair it, the package is malformed. The usual"
    Write-Host "  cause is a comments part wired into the document but missing its"
    Write-Host "  relationship or its content type override. Do not send this file."
    $fail = $true
}
finally {
    if ($doc) { $doc.Close($false) | Out-Null }
    if ($word) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
}

if ($fail) { exit 1 }
Write-Host ""
Write-Host "  Word opened it without repair."
exit 0
