# Not damaging the client's file

Every file in a markup job is a live working paper arriving without version
control. There is no commit to revert to. Assume every write is permanent and
test before making one.

---

## The rule

**Never write to the input path.** Output to a new file whose name carries the
marking, and leave the source untouched:

```
<company> AR 2025-26 Standalone 31082026 - OPEN ITEMS.pdf
<company> Standalone FS FY 2025-26 - OPEN ITEMS MARKED.xlsx
<company> Standalone Audit Report - REDLINED.docx
```

This also gives the client the comparison they will want, and it survives the
case where your marks turn out to be wrong.

---

## openpyxl destroys workbooks on save

It reads accurately. It writes back a file that only contains the parts it
understands, and silently drops the rest.

On a live financial reporting workbook of around three megabytes, a bare
`load_workbook` then `save` lost **86 parts** of the OOXML package: chart XML,
drawing relationships, printer settings, defined name scoping, custom views. The
resulting file opens in Excel without a warning. The damage surfaces weeks later
when someone prints a schedule and the layout is gone, or a chart in the investor
deck is empty.

Test it on the actual file before deciding anything:

```python
import zipfile, openpyxl, os
before = set(zipfile.ZipFile(src).namelist())
wb = openpyxl.load_workbook(src); wb.save(tmp); wb.close()
after = set(zipfile.ZipFile(tmp).namelist())
print(len(before), len(after), sorted(before - after)[:10])
```

`xldump.py --roundtrip` in the `annual-report-review` skill runs exactly this and
prints a verdict.

Any loss at all means use Excel COM, or deliver a separate workbook. There is no
threshold below which it is acceptable: a lost printer setting is a lost printer
setting in a document that is about to be filed.

Reading with `openpyxl` is fine and is the right tool for it. Use `data_only=True`
to get values, which is what printed, rather than formulas.

---

## PDF

`pymupdf` is safe for annotation. It rewrites the file, but annotations are a
first class PDF object and nothing else is disturbed.

Two cautions.

**Incremental save against a new file.** Saving to a new path performs a full
write, which is what you want. Saving back over the source incrementally is
faster and is exactly the operation the rule above forbids.

**Check the page count afterwards.** A mismatch between input and output page
count means something went wrong in a way that will not otherwise announce
itself.

Do not use tools that rasterise. A marked up PDF whose text layer has been
replaced by images cannot be searched, and the client searches it constantly.

---

## Word

Editing `word/document.xml` inside the zip is safe as long as you rewrite the
whole package rather than patching bytes in place, and as long as the three way
agreement between the document, its relationships and `[Content_Types].xml` is
maintained whenever a part is added.

Where it goes wrong, Word announces it: on open it reports unreadable content and
offers to repair. That repair usually discards the comments. So always open the
output before sending it, which `verify_docx.ps1` does.

Legacy `.doc` is not a zip and none of this applies. Convert to `.docx` through
Word first, and tell the user you did, because the conversion itself can shift
pagination.

---

## Orphaned COM processes

A failed Excel or Word automation run leaves the application running headless
with the file locked. Symptoms: the next run fails instantly with a sharing
violation, or the script exits non zero having printed nothing.

Check and clear:

```powershell
Get-Process EXCEL, WINWORD -ErrorAction SilentlyContinue | Format-Table Id, StartTime
Stop-Process -Name EXCEL -Force
```

Only do this when the user has no Excel or Word window of their own open. Killing
the process takes their unsaved work with it. Ask first, or check for a visible
window before acting.

---

## Before you report success

Reopen the output and count what is actually in it: annotations, revisions,
comments, filled cells. Compare against the number of marks you supplied.

The failure mode of this entire skill is a file that looks marked and is not. A
highlight created without `update()`, a deletion wired with `w:t`, a comment part
missing its content type, an anchor that matched nothing. None of these raise an
error. All of them produce a deliverable that tells the client there were no
findings.
