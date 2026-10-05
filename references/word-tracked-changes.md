# Redlining a Word document

The deliverable is a document the recipient can walk through in Word, accepting
and rejecting clause by clause. That means real tracked changes in the OOXML, not
coloured text and not strikethrough formatting.

There is no library that does this well. Write the XML.

---

## The edits file

```json
[
  {
    "find": "refer note 32",
    "replace": "refer note 41",
    "ref": "AR-03",
    "note": "Notes renumbered to 1 to 63 in the final set. Cross reference is stale."
  },
  {
    "find": " and the Company has not been sanctioned any working capital limits",
    "replace": "",
    "ref": "AR-07",
    "note": "Contradicts CARO clause (ii)(b), which reports a sanctioned limit."
  },
  {
    "find": "Provisions (current)",
    "comment": "Re footed to 65.11. See observation SA-01.",
    "ref": "SA-01"
  }
]
```

An empty `replace` is a deletion. A `comment` with no `replace` is a margin
comment with no edit, which is the right instrument where the point is a question
for the auditor rather than a wording change.

---

## Run splitting

Word stores a paragraph as a sequence of runs, and splits them wherever
formatting, spell check state or an editing session boundary changed. The phrase
you want to edit is very often spread across three runs for no visible reason,
and a find and replace over run text finds nothing.

Run `merge_runs.py` first. It rewrites each paragraph's runs into a single run
carrying the first run's properties, which is safe for an audit report and a
board report because the body text is uniform. Do not run it on a document with
meaningful inline formatting, such as a table where single words are bold: it
will flatten them.

Where that is a risk, redline only the paragraphs that need it, and leave the
rest alone.

---

## Insertions and deletions

An insertion is a run wrapped in `w:ins`:

```xml
<w:ins w:id="101" w:author="Reviewer" w:date="2026-08-18T00:00:00Z">
  <w:r><w:t xml:space="preserve">refer note 41</w:t></w:r>
</w:ins>
```

A deletion is a run wrapped in `w:del`, and the text element inside it must be
`w:delText`, not `w:t`:

```xml
<w:del w:id="102" w:author="Reviewer" w:date="2026-08-18T00:00:00Z">
  <w:r><w:delText xml:space="preserve">refer note 32</w:delText></w:r>
</w:del>
```

Using `w:t` inside `w:del` is the single commonest mistake. Word opens the file
without complaint, shows no revision, and the deleted text simply appears as
normal text. Nothing warns you.

`w:id` must be unique across the document. Start high, at 1000, to avoid
colliding with revisions already in the file.

`xml:space="preserve"` on every `w:t` and `w:delText`. Without it Word strips
leading and trailing spaces and the redline reads as two words run together.

To replace, emit the deletion and the insertion adjacently, in that order, inside
the paragraph at the position the original text occupied. That means splitting the
containing run into three: the text before, the change, the text after.

---

## Comments

Three places have to agree, and a comment that is wired into only two of them
makes Word declare the file corrupt and offer to repair it.

1. **`word/comments.xml`** holds the comment bodies. Create the part if it is
   absent, with the `w:comments` root and the same namespace declarations the
   document uses.
2. **`word/_rels/document.xml.rels`** needs a relationship to it, type
   `.../officeDocument/2006/relationships/comments`.
3. **`[Content_Types].xml`** needs the override for
   `/word/comments.xml`, content type
   `application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml`.

In the document body, the anchor is three elements:

```xml
<w:commentRangeStart w:id="5"/>
  ... the anchored run ...
<w:commentRangeEnd w:id="5"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
  <w:commentReference w:id="5"/></w:r>
```

**The anchor must sit inside a single contiguous run.** An anchor spanning a run
boundary produces a comment that attaches to nothing, and Word shows it in the
pane with no connecting line. Keep anchors to a few words, and run `merge_runs.py`
first.

Comment ids and revision ids are separate numbering spaces. They may overlap
without harm, but keeping them apart makes the XML readable when something goes
wrong.

---

## Verification

Do not verify by reading the XML you just wrote. Open the document.

```powershell
powershell -File scripts/verify_docx.ps1 "out.docx"
```

It reports what Word itself sees: `Revisions.Count` and `Comments.Count`. If
either is lower than the number of edits you supplied, something silently failed,
and the usual cause is `w:t` inside `w:del` or an anchor spanning runs.

It also exports a PDF with the markup rendered:

```
ExportAsFixedFormat(path, 17, ..., Item:=7)
```

`Item = 7` is `wdExportDocumentWithMarkup`. The default exports the document as
if all changes were accepted, which looks like a clean document and tells the
reader nothing. Getting this wrong produces a deliverable that appears to show no
findings at all.

---

## Withdrawing a redline

When a later version of the document makes one of your edits wrong, for example
because the notes were renumbered and the cross reference is now correct, say so
explicitly and tell the recipient to reject that specific revision by its
reference. Advice to reject your own earlier redline is worth more than silence,
and it is the thing that establishes that the second pass was a real review
rather than a re run of the first.
