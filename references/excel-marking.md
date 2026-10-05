# Marking up an Excel working paper

The workbook is the live working paper. It is the file the preparer will open
tomorrow morning to make the corrections, and it is frequently the only copy. The
risk here is not a bad mark. It is handing back a file that no longer opens
properly, which costs the client a day and costs you the engagement.

Read `file-integrity.md` before writing anything.

---

## Use Excel through COM

`openpyxl` reads these workbooks accurately and must not be used to write them.
On a live financial reporting set of a few megabytes, a read and save round trip
lost 86 relationship parts: charts, drawings, printer settings, defined name
scopes. The file still opens, which is why this is dangerous rather than merely
broken.

Drive the real application instead:

```powershell
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open($src)
```

COM preserves everything, because Excel is doing the writing.

Always close in a `finally`, and always quit the application:

```powershell
$wb.SaveAs($dst)
$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
```

A run that fails before `Quit` leaves an `EXCEL.EXE` holding a lock on the file.
The symptom is a second attempt failing immediately with a file in use error, or
the script exiting with a non zero status and no output. Kill the orphan before
retrying.

---

## Performance

COM is slow. Setting an interior colour and a comment on each of thirty cells
across a sixty sheet, three megabyte workbook takes minutes, not seconds.

Run it in the background. Do not conclude it has hung, and do not start a second
run alongside the first: two Excel automation clients on one file will fight.

Reduce the work where you can. Fetch each worksheet object once and reuse it
rather than indexing into `Worksheets` per cell, and leave screen updating and
calculation off for the duration:

```powershell
$xl.ScreenUpdating = $false
$xl.Calculation = -4135      # xlCalculationManual
```

Set calculation back to automatic before saving, or the recipient opens a
workbook that does not recalculate and may not notice.

---

## The marks

```json
[
  {
    "sheet": "BS",
    "cell": "D34",
    "sev": "OPEN",
    "ref": "SA-01",
    "note": "Total current liabilities 13,059.27. Re footed to 13,059.28. Still not corrected in this version."
  }
]
```

Fill the cell interior with the palette colour and attach a comment carrying the
reference and the note. Excel colours are BGR packed integers, not RGB:

```powershell
# RGB(255,192,0), the status orange
$bgr = 0 * 65536 + 192 * 256 + 255      # 49407
$ws.Range($cell).Interior.Color = $bgr
```

Getting the byte order wrong gives you a plausible looking wrong colour rather
than an error, so check one cell visually before running the batch.

Where a cell already carries a fill that means something in the preparer's own
scheme, do not overwrite it silently. Mark the adjacent cell in the first empty
column of the row instead and say so in the comment.

---

## The index sheet

A workbook with thirty coloured cells scattered across sixty sheets is not
actionable. Add a sheet, first in the tab order, named `OPEN ITEMS`:

| Ref | Sheet | Cell | Status | Observation |
|---|---|---|---|---|

One row per mark, the status cell filled in the same palette colour as the mark
itself, and an autofilter on the header row:

```powershell
$idx = $wb.Worksheets.Add($wb.Worksheets.Item(1))
$idx.Name = "OPEN ITEMS"
$idx.Range("A1:E1").AutoFilter() | Out-Null
```

Add a hyperlink from each Ref cell to the marked cell, so the sheet works as
navigation:

```powershell
$idx.Hyperlinks.Add($idx.Range("A$r"), "", "'$sheet'!$cell") | Out-Null
```

The quoting matters: sheet names in these workbooks contain spaces.

---

## When the workbook cannot be written

Some workbooks defeat automation: a crashed COM run that will not complete, a
file the client has open, a protected structure, a workbook so large that Excel
exhausts memory part way through.

Do not force it and do not fall back to a library that will damage it. Deliver
the marks as a separate single sheet workbook, keyed by sheet name and cell
reference, in the same format as the index sheet above. The preparer can work
from that in minutes, and their file is untouched.

Say which route you took, in the covering note. A client who thinks their
workbook was marked and finds it clean will assume the review found nothing.

---

## Do not mark the Excel instead of the PDF

The printed statements are the document that gets signed and filed. The Excel is
a working paper and is frequently not the source of what printed: findings turn
up where the Excel is right and the typesetting is wrong. Marking only the
workbook leaves the error on the page that goes out.

Where the user asks for the workbook specifically, usually because they want to
know which corrections were actually made, mark it. But confirm separately that
the same point is marked in the statements.
