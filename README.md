# review-markup

A Claude skill for putting review findings into the client's own documents
rather than writing them up separately: colour coded highlights with severity
graded popup comments in PDF, tracked changes and margin comments in Word, and
cell fills with comments and an index sheet in Excel.

It is the delivery layer for
[annual-report-review](https://github.com/parinshahAIbuilder/annual-report-review)
and
[ipo-offer-document-review](https://github.com/parinshahAIbuilder/ipo-offer-document-review),
and it stands on its own for any review that ends in a marked up file.

A review report is read once. A marked up document is read at the desk, next to
the figure, by the person who has to change it.

## What it does

```bash
python scripts/mark_pdf.py "statements.pdf" marks.json --dry-run
python scripts/mark_pdf.py "statements.pdf" marks.json

python scripts/merge_runs.py "audit report.docx" --out merged.docx
python scripts/redline_docx.py merged.docx edits.json
powershell -File scripts/verify_docx.ps1 "audit report - REDLINED.docx"

powershell -File scripts/mark_excel.ps1 "working papers.xlsx" marks.json -WhatIfOnly
powershell -File scripts/mark_excel.ps1 "working papers.xlsx" marks.json
```

Two colour palettes, kept apart because a reader who assumes severity when you
meant status will read a closed item as critical:

| | |
|---|---|
| **Severity**, for a first markup | critical, high, medium, low |
| **Status**, for a later cycle | verified as done, previously raised and still open, new issue |

The status palette is the one clients ask for by the third cycle, usually in the
words "highlight what is still not done".

## The three rules it exists to enforce

**Do not damage the file.** `openpyxl` reads these workbooks correctly and writes
them back missing parts. On a live three megabyte financial reporting set, a bare
load and save lost **86 relationship parts**: chart XML, drawing relationships,
printer settings, defined name scoping. The file still opened, which is why this
is dangerous rather than merely broken. The skill drives Excel through COM
instead, and ships the round trip test so you can settle it on your own file.

**Anchor on what is printed, not on what you think is printed.** Typeset
financial tables are full of narrow no break spaces inside thousands groups, en
dashes standing in for minus signs and figures split across text runs. Every
script has a dry pass, and an anchor that matches nothing is reported rather than
quietly dropped. It usually means one of four things, and the fourth is that the
point is already fixed, which is a finding in itself.

**Verify by opening the result, never by reading the XML you just wrote.** The
failure mode of this entire area is a file that looks marked and is not, and
nothing raises an error:

- a PDF highlight created without `update()`
- a Word deletion wired with `w:t` instead of `w:delText`, which opens cleanly
  and shows no revision at all
- a comment part missing its content type override
- a root element whose namespace declarations ElementTree dropped as unused
  while `mc:Ignorable` still names them, which Word reports only as "the file
  appears to be corrupted"

Each of those cost time to find. They are written into the reference notes so
they cost nobody else any.

## Installing

Copy the folder into your personal skills directory:

```
~/.claude/skills/review-markup/
```

Then start a new Claude Code session. Skills are discovered at session start, so
a session already running will not see it.

## Requirements

Python with `pymupdf`:

```
pip install pymupdf
```

The Word and Excel scripts are PowerShell and drive the real applications
through COM, so they need Microsoft Office on Windows. The PDF side needs
neither and runs anywhere.

## What is in here

```
SKILL.md                              the workflow
scripts/mark_pdf.py                   highlights, popup comments, legend
scripts/merge_runs.py                 collapses split runs so find and replace works
scripts/redline_docx.py               tracked insertions and deletions, and comments
scripts/verify_docx.ps1               opens it in Word, counts what Word sees, exports markup
scripts/mark_excel.ps1                cell fills, comments, OPEN ITEMS index sheet
references/pdf-annotation.md          anchoring, colours, legend placement
references/word-tracked-changes.md    the OOXML, run splitting, comment anchoring
references/excel-marking.md           COM technique, the index sheet, performance
references/file-integrity.md          which libraries damage what, and the tests
references/handover.md                naming, legend, covering note, metadata
```

### mark_pdf.py

Writes real PDF annotations, not coloured rectangles drawn on the page. The
difference matters: a real annotation appears in the reader's comment sidebar,
and can be sorted, filtered and replied to.

The severity goes in the annotation's `subject` field as well as in the colour,
so a subject of `[C] SA-01` turns the sidebar into a working issue list without
any further document. That one workflow is what the deliverable is really for.

### verify_docx.ps1

The only test that counts. It opens the redlined document in Word and reports
`Revisions.Count` and `Comments.Count`, which is what Word itself sees, and fails
loudly if either is short of what you supplied.

It also exports a PDF with `Item = 7`, `wdExportDocumentWithMarkup`. The default
export renders the document as if every change were accepted, which looks clean
and tells the reader there were no findings.

### mark_excel.ps1

Fills and comments the cells, then adds an `OPEN ITEMS` sheet at the front: one
row per mark, status cell filled in the same colour, autofilter on the header,
and a hyperlink from each reference to the marked cell. A workbook with thirty
coloured cells scattered across sixty sheets is not actionable without it.

It leaves alone any cell already carrying the preparer's own fill, and says so,
rather than overwriting something that meant something.

Expect minutes rather than seconds on a workbook of a few megabytes. That is COM,
not a hang.

## A note on handover

Do not send Word files outside the engagement. They carry author names, editing
history and every comment anyone left in them, including the internal queries the
preparer never meant the company to see. Convert to PDF with the markup rendered
and send that. The redlined document goes to the person who will accept and
reject the revisions, and nobody else.

The same applies to the Excel working paper, which carries the preparer's open
review notes in cells. Those are often more revealing than the findings.

## Licence

MIT. See [LICENSE](LICENSE).
