# Marking up a PDF

Use `pymupdf`. It is the only library in reach that writes real PDF annotations
rather than drawing coloured rectangles on top of the page, and the difference
matters: a real annotation appears in the reader's comment sidebar, can be
filtered and sorted, and can be replied to. A drawn rectangle is just ink.

---

## The marks file

```json
[
  {
    "page": 7,
    "anchor": "13,059.27",
    "sev": "C",
    "ref": "SA-01",
    "note": "Total current liabilities. Components re footed to 13,059.28. Printed figure understates by 0.01 and the balance sheet does not balance as printed. Correct by rounding Provisions (current) 65.10 to 65.11."
  },
  {
    "page": 7,
    "anchor": "Total equity and liabilities",
    "occurrence": 2,
    "sev": "H",
    "ref": "SA-02",
    "note": "Carries the mis cast forward."
  }
]
```

`page` is the one based PDF page, which is frequently not the printed folio
number. Resolve that once and write PDF pages throughout, or every anchor will be
one or two pages out.

`occurrence` selects which match on the page when the anchor appears more than
once, one based, default all of them. Highlighting all occurrences of a label
like "Total" is almost never what you want.

`sev` takes either palette: `C` `H` `M` `L` for severity, `DONE` `OPEN` `NEW` for
status. Do not mix palettes within one document.

---

## The core sequence

```python
page = d[pg - 1]                      # hold the page object
rects = page.search_for(anchor)
a = page.add_highlight_annot(rects)   # from the SAME page object
a.set_colors(stroke=rgb)
a.set_info(title=author,
           subject="[%s] %s" % (sev, ref),
           content="SEVERITY: %s\n\n%s" % (label, note))
a.set_opacity(0.45)
a.update()                            # nothing is written until update()
```

Four things in that block are not optional.

**The page object must be held.** Creating the annotation from a page reference
obtained separately, or from an index into the document on the same line, raises
`FzErrorArgument: annotation not bound to any page`. Assign `page` first, then
search and annotate from it.

**`update()` writes the annotation.** Without it the colours and the comment are
set on an object that never reaches the file, and the output looks like the
highlight worked but carries no text.

**The severity belongs in `subject`.** That is the field readers expose as a
sortable, filterable column in the comment list. A subject of `[C] SA-01` turns
the sidebar into the issue list. Putting it only in the body wastes the feature.

**Opacity around 0.45.** A full strength highlight over small type in a financial
table makes the figures unreadable, which defeats the purpose.

---

## Anchoring

`search_for` matches the text layer exactly. Typeset financial statements defeat
exact matching routinely.

Try, in order:

1. The anchor as written.
2. The anchor with every run of whitespace treated as flexible. Thousands groups
   in typeset tables often carry a narrow no break space rather than a comma or a
   normal space.
3. The anchor with `-` swapped for the en dash `–` and the minus sign `−`, and
   the reverse. Brackets around negatives are also common: `(1,234.56)`.
4. A shorter distinctive fragment. The last six characters of a figure are more
   reliable than the whole figure with its separators.
5. A neighbouring label on the same row rather than the number itself.

`mark_pdf.py` does one through three automatically and reports which variant
matched. Where it still fails, fall back to five: highlighting the row label with
the note attached is a perfectly good mark, and much better than no mark.

Never widen an anchor until it matches something. An anchor that drifts onto the
wrong row produces a mark that is actively misleading.

---

## The legend

Put it in the bottom margin of the first page, as a drawn box rather than an
annotation, so it prints. A reader who opens the file on a phone with no comment
sidebar still needs to know what the colours mean.

```python
page = d[0]
r = pymupdf.Rect(36, page.rect.height - 86, page.rect.width - 36,
                 page.rect.height - 28)
page.draw_rect(r, color=(0.4, 0.4, 0.4), fill=(1, 1, 1), width=0.6)
```

Then one swatch and label per code actually used in this document. Do not print
the codes you did not use: a legend listing Critical when nothing is critical
invites the reader to go looking for it.

State on the legend which palette the document uses, in words. "Colours indicate
severity" or "Colours indicate status of points raised on the previous draft".

---

## Reporting back

Print, at the end of the run:

- the count of marks applied, by code
- every anchor that matched nothing, with its page and reference
- every anchor that matched more times than expected

Hand the unmatched list to the user with the file. An unmatched anchor means
either your page reference is wrong or the point is already fixed, and the user
is the one who can tell you which.

---

## What not to do

**Do not redact or cover text.** A highlight sits under the glyphs. Anything that
sits over them is a different operation with a different purpose and it will be
read as an attempt to hide something.

**Do not write over the input file.** Output to a new name. A markup pass that
destroys the only clean copy of the statements is unrecoverable, and these
documents arrive without version control.

**Do not add a summary page to the front of the client's document.** It shifts
every page number in every reference anyone has already written. Deliver the
summary as its own file.
