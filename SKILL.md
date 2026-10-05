---
name: review-markup
description: >
  Marks review findings into the client's own documents rather than writing them up
  separately: colour coded highlights with severity graded popup comments in PDF
  financial statements, tracked changes and margin comments in Word audit reports and
  board reports, and cell highlighting with comments and an index sheet in Excel working
  papers. Carries the severity and status colour conventions, the anchoring techniques
  that actually find text in a typeset financial table, and the file integrity rules that
  stop a markup job from silently destroying a live client workbook. Use whenever the
  user asks for findings to be highlighted, marked up, redlined, tracked, commented or
  colour coded in the same file they supplied, or asks which previously raised points are
  still open shown in the document itself. Pairs with the annual-report-review and
  ipo-offer-document-review skills, which produce the findings this skill delivers.
---

# Marking findings into the client's own documents

A review report is read once. A marked up document is read at the desk, next to
the figure, by the person who has to change it. When a client asks for the
findings in their own file, that is usually the deliverable that gets actioned,
and the review report becomes the supporting paper.

Three rules govern the whole job.

**Do not damage the file.** The client's workbook and document are live working
papers. Several of the obvious libraries read these files correctly and write
them back subtly broken. `references/file-integrity.md` sets out which, and what
to do instead. Test before you write, every time, on the actual file.

**Anchor on what is printed, not on what you think is printed.** Typeset
financial tables are full of non breaking spaces, narrow spaces inside thousands
groups, en dashes standing in for minus signs and figures split across text runs.
An anchor that looks right and matches nothing produces a document with no marks
on the page that matters. Always run the dry pass first and read the unmatched
list.

**Every mark carries its severity and its reference.** A yellow highlight with no
comment tells the reader that something is wrong somewhere near here. The mark is
worth having only when it names the finding, its grade, and the identifier that
ties it back to the review report.

## Start with the dry pass

Never write marks blind.

```bash
python <skill>/scripts/mark_pdf.py in.pdf marks.json --dry-run
```

It reports, per mark, how many times the anchor was found on the page named, and
lists every anchor that matched nothing. Fix the anchors, then run for real.

Each of the three has one. `redline_docx.py --dry-run` tells you which edits
would apply, and names the ones whose target spans a run boundary.
`mark_excel.ps1 -WhatIfOnly` resolves every sheet name and cell reference and
prints what is currently in each cell, which catches the case where the workbook
was restructured between versions and your references now point at the wrong
rows.

An anchor that matches nothing is nearly always one of four things, in order of
frequency: the figure is on a different page than your note says, the spaces are
not the spaces you typed, the text is split across runs, or the number was
already corrected in this version and the finding is closed. The fourth is a
finding in itself. Check it before you rewrite the anchor.

## The colour conventions

Two palettes, used for two different jobs. Keep them separate and state in the
legend which one the document uses, because a reader who assumes severity when
you meant status will read a closed item as critical.

**Severity**, for a first markup of a document under review:

| Code | Colour | Meaning |
|---|---|---|
| `C` | red | Critical. Must be corrected before signing |
| `H` | orange | High. Correct before circulation |
| `M` | yellow | Medium. Disclosure or presentation |
| `L` | blue | Low. Typographical |

**Status**, for a later cycle where the question is what was actioned:

| Code | Colour | Meaning |
|---|---|---|
| `DONE` | green | Verified as incorporated |
| `OPEN` | orange | Previously raised, still not done |
| `NEW` | red | New issue, including one created by a correction |

The status palette is the one clients ask for by the third cycle, usually in the
words "highlight what is still not done". Use green generously in that pass. A
document showing eleven green marks and three orange is a far more useful
instrument than one showing only the three, because it proves the other eleven
were checked.

## PDF

The workhorse. `references/pdf-annotation.md` carries the mechanics; the short
version is that `mark_pdf.py` takes a JSON list of marks and produces a new file
with highlights, popup comments and a legend.

```bash
python <skill>/scripts/mark_pdf.py "in.pdf" marks.json --out "in - MARKED UP.pdf"
```

Put the severity in the comment subject as well as the colour. Acrobat and most
readers let the user sort and filter the comment list by subject, so a subject of
`[C] SA-01` gives the client a working issue list in the sidebar without any
further document. That sidebar is the feature the deliverable is really for.

Never write over the input. Output to a new name carrying the marking in it.

## Word

Tracked changes, not highlights. An auditor's report marked with insertions and
deletions can be accepted or rejected clause by clause in Word, which is exactly
what the recipient wants to do with it.

```bash
python <skill>/scripts/redline_docx.py "in.docx" edits.json --out "in - REDLINED.docx"
```

`references/word-tracked-changes.md` has the detail. Two things cause almost all
the failures: the text you are editing is split across runs, so run
`merge_runs.py` first and let it rewrite the paragraphs into single runs; and a
comment anchor must sit inside one contiguous run, so keep anchors short.

Verify by opening the result, not by reading the XML you just wrote:

```powershell
powershell -File <skill>/scripts/verify_docx.ps1 "in - REDLINED.docx"
```

It reports the revision and comment counts Word itself sees, and exports a PDF
with the markup visible so the client can read the redline without Word.

## Excel

The dangerous one, because the workbook is the live working paper and because
the normal Python route damages it. Drive Excel through COM:

```powershell
powershell -File <skill>/scripts/mark_excel.ps1 "book.xlsx" marks.json
```

It fills the cells, attaches a comment to each, adds an `OPEN ITEMS` index sheet
with an autofilter, and saves to a new file. Expect it to take minutes on a
workbook of a few megabytes with sixty sheets, so run it in the background and do
not conclude it has hung. If it fails part way it can leave an orphaned Excel
process holding the file; `references/file-integrity.md` says how to recover.

Where the workbook cannot be written at all, deliver the marks as a separate
single sheet workbook keyed by sheet name and cell reference. A client can act on
that. A corrupted working paper costs them a day.

## Withdrawing a mark

When a later pass shows a mark was wrong, or the finding has been overtaken,
remove it and say so. Do not leave a stale red highlight in a document the client
is circulating, and do not quietly drop it either: name the withdrawn reference in
the covering note so the reader can reconcile against the earlier marked copy.

## Before handing anything over

`references/handover.md` is short and all of it matters. The item that matters
most: do not send Word files outside the engagement. They carry author names,
revision history and any unresolved comments the preparer left behind, and that
travels. Convert to PDF with the markup rendered and send that.

## Reference files

| File | Read it |
|---|---|
| `references/pdf-annotation.md` | Before marking a PDF. Anchoring, colours, legend placement, the pitfalls |
| `references/word-tracked-changes.md` | Before redlining a docx. The OOXML, run splitting, comment anchoring, verification |
| `references/excel-marking.md` | Before touching a workbook. COM technique, the index sheet, performance |
| `references/file-integrity.md` | Before writing to any client file. Which libraries damage what, and the tests |
| `references/handover.md` | Before sending. Naming, legend, covering note, metadata |

## Scripts

| Script | Use |
|---|---|
| `scripts/mark_pdf.py` | Highlights and popup comments in a PDF, with legend. Run `--dry-run` first |
| `scripts/merge_runs.py` | Rewrites docx paragraphs into single runs. Run before redlining |
| `scripts/redline_docx.py` | Tracked insertions and deletions, and comments, in a docx |
| `scripts/verify_docx.ps1` | Opens the result in Word, reports revision and comment counts, exports a markup PDF |
| `scripts/mark_excel.ps1` | Cell fills, comments and an index sheet, through Excel COM |
