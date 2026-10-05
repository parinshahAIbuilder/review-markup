#!/usr/bin/env python
"""Rewrite each paragraph's runs into a single run, so find and replace works.

    python merge_runs.py in.docx --out merged.docx
    python merge_runs.py in.docx --report

Word splits a paragraph into runs wherever formatting, spell check state or an
editing session boundary changed, which is routinely mid phrase and for no
visible reason. A phrase spread across three runs cannot be found by a text
search over run content, and that is the reason most redline attempts silently
apply nothing.

Merging takes the first run's properties for the whole paragraph. That is safe
for audit report and board report body text, which is uniform. It is NOT safe
where single words inside a paragraph are bold or italic: those will be
flattened. Run --report first and look at what would be lost.

Paragraphs already carrying tracked changes or comment anchors are left alone.
"""
import argparse, os, re, shutil, sys, zipfile
from xml.etree import ElementTree as ET

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
SKIP_IF = (W + "ins", W + "del", W + "commentRangeStart", W + "commentReference",
           W + "hyperlink", W + "fldChar", W + "instrText")


def register(raw):
    head = raw[:raw.find(">", raw.find("<w:document"))] if "<w:document" in raw else raw[:2000]
    for m in re.finditer(r'xmlns:([A-Za-z0-9]+)="([^"]+)"', head):
        ET.register_namespace(m.group(1), m.group(2))


_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + chr(13) + chr(10)


def serialize(raw, root, tag="w:document"):
    """Serialize back, keeping the original root start tag.

    ElementTree emits declarations only for namespaces it sees in use, but the
    root carries mc:Ignorable naming prefixes that appear nowhere else. Dropping
    those declarations leaves mc:Ignorable pointing at undeclared prefixes, and
    Word reports that only as "the file appears to be corrupted". Splice the
    original start tag, with all its declarations and attributes, back in.
    """
    i = raw.find("<" + tag)
    orig = raw[i:raw.find(">", i) + 1]
    decl = raw[:i] if raw[:i].lstrip().startswith("<?xml") else _DECL
    s = ET.tostring(root, encoding="unicode")
    a = s.find("<" + tag)
    b = s.find(">", a) + 1
    return (decl + orig + s[b:]).encode("utf-8")


def rpr_sig(r):
    p = r.find(W + "rPr")
    return ET.tostring(p) if p is not None else b""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--out")
    ap.add_argument("--report", action="store_true")
    a = ap.parse_args()

    z = zipfile.ZipFile(a.docx)
    raw = z.read("word/document.xml").decode("utf-8")
    register(raw)
    root = ET.fromstring(raw)

    merged = skipped = mixed = 0
    mixed_examples = []
    for p in root.iter(W + "p"):
        if any(p.find(t) is not None for t in SKIP_IF) or \
           any(c.tag in SKIP_IF for c in p):
            skipped += 1
            continue
        runs = [c for c in p if c.tag == W + "r"]
        if len(runs) < 2:
            continue
        if any(c.find(W + "br") is not None or c.find(W + "tab") is not None
               or c.find(W + "drawing") is not None for c in runs):
            skipped += 1
            continue
        text = "".join("".join(t.text or "" for t in r.iter(W + "t")) for r in runs)
        if not text.strip():
            continue
        sigs = {rpr_sig(r) for r in runs}
        if len(sigs) > 1:
            mixed += 1
            if len(mixed_examples) < 8:
                mixed_examples.append(text[:84])
        merged += 1
        if a.report:
            continue

        keep = runs[0]
        for t in list(keep.findall(W + "t")):
            keep.remove(t)
        t = ET.SubElement(keep, W + "t")
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        for r in runs[1:]:
            p.remove(r)

    verb = "would be merged" if a.report else "merged         "
    print(f"  paragraphs {verb} {merged}")
    print(f"  paragraphs left alone  {skipped}   (tracked changes, comments, "
          f"hyperlinks, breaks, images)")
    print(f"  of the merged, mixed formatting flattened in {mixed}")
    if mixed_examples:
        print("\n  FORMATTING LOST IN THESE PARAGRAPHS. Check none of them needs it:")
        for e in mixed_examples:
            print(f"     {e!r}")

    if a.report:
        print("\nReport only. Nothing written.")
        return 0

    out = a.out or os.path.splitext(a.docx)[0] + " - merged.docx"
    shutil.copyfile(a.docx, out)
    body = serialize(raw, root)

    tmp = out + ".tmp"
    zin = zipfile.ZipFile(out)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for i in zin.infolist():
            zo.writestr(i, body if i.filename == "word/document.xml" else zin.read(i.filename))
    zin.close()
    os.replace(tmp, out)
    print(f"\nwrote {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
