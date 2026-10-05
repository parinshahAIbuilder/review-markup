# Handing the marked documents over

Short, and all of it has gone wrong at least once.

---

## Do not send Word files outside the engagement

A `.docx` carries the author and last modified by names, the editing history, the
total editing time, and every comment anyone left in it including the ones the
preparer thought were private. On an audit report, those comments are frequently
internal queries that were never meant for the company, and on a board report
they are frequently the opposite.

Check what is in there before it moves:

```python
import zipfile
z = zipfile.ZipFile(path)
print(z.read("docProps/core.xml").decode())
print("comments part present:", "word/comments.xml" in z.namelist())
```

Then convert to PDF with the markup rendered and send the PDF. The redlined
`.docx` goes to the person who will accept and reject the revisions, and nobody
else.

The same applies to the Excel working paper, which carries the preparer's open
review notes in cells. Those notes are often more revealing than the findings.

---

## Naming

Put the marking in the file name, and keep the original name intact in front of
it so the two sort together:

```
<original name> - MARKED UP.pdf        severity palette, first pass
<original name> - OPEN ITEMS.pdf       status palette, later pass
<original name> - REDLINED.docx        tracked changes
<original name> - OPEN ITEMS MARKED.xlsx
```

Do not add dates to the name unless the source has one. The source name usually
already carries the version date, and a second date invites confusion about which
is which.

---

## The covering note

One page, with the marked files. It carries what the document itself cannot.

- **Which file and which modification date was marked.** Not the title, the file
  name, so the recipient can be certain they are looking at the same one.
- **Which palette the colours mean**, in words.
- **The count by code.** Four critical, nine high, eleven verified as done.
- **The unmatched anchors.** Every point you could not place in the document, with
  the page you expected it on. These are the ones where either your reference is
  wrong or the point is already fixed, and only the recipient can say which.
- **Anything withdrawn** since the previous marked copy, by reference.
- **Whether the workbook was marked in place or delivered separately**, if Excel
  was in scope.

Where a point is marked in one document and not the other, say so. A reader who
finds a finding in the standalone and not the consolidated will assume it does
not apply there.

---

## What the client actually does with these

They open the PDF, open the comment sidebar, sort by subject, and work down the
list. Everything in `pdf-annotation.md` about putting the severity and reference
in `subject` exists to make that one workflow work.

So before sending, open the output and look at the sidebar yourself. If the list
does not read as a usable issue list in that pane, with each row naming its grade
and reference, the markup is not finished, however good the highlights look on the
page.

---

## Do not restate the review report in the comments

A popup comment should be two or three sentences: what is wrong, what it should
be, and the reference. The reasoning, the citation and the recommended correction
belong in the review report, which the comment points to.

Comments long enough to need scrolling in the popup do not get read.
