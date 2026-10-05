#!/usr/bin/env python
"""Tracked insertions and deletions, and margin comments, in a Word document.

    python redline_docx.py in.docx edits.json --dry-run
    python redline_docx.py in.docx edits.json --out "in - REDLINED.docx"

Run merge_runs.py first. Word splits paragraphs into runs mid phrase, and an
edit whose target spans a run boundary cannot be applied.

edits.json is a list of:

    {"find": "refer note 32", "replace": "refer note 41",
     "ref": "AR-03", "note": "Notes renumbered. Cross reference is stale."}

    {"find": " and no limits were sanctioned", "replace": "",
     "ref": "AR-07", "note": "Contradicts CARO clause (ii)(b)."}

    {"find": "Provisions (current)", "comment": "Re footed to 65.11. See SA-01.",
     "ref": "SA-01"}

An empty "replace" is a deletion. A "comment" with no "replace" is a margin
comment with no edit. Supplying both puts a comment on the change.

Add "occurrence": n to pick which match when the text appears more than once.
Without it, only the first match is edited and the rest are reported.

Always verify the output with verify_docx.ps1, which reports what Word itself
sees. The failure mode here is silent: a malformed revision opens cleanly and
simply shows nothing.
"""
import argparse, json, os, re, sys, zipfile
from datetime import datetime, timezone
from xml.etree import ElementTree as ET

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % WNS
XMLSP = "{http://www.w3.org/XML/1998/namespace}space"
RELNS = "http://schemas.openxmlformats.org/package/2006/relationships"
CTNS = "http://schemas.openxmlformats.org/package/2006/content-types"
CT_COMMENTS = ("application/vnd.openxmlformats-officedocument."
               "wordprocessingml.comments+xml")
REL_COMMENTS = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                "relationships/comments")


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


def run_text(r):
    return "".join(t.text or "" for t in r.findall(W + "t"))


def mkrun(template, text, delete=False):
    """A run copying template's properties, carrying text."""
    r = ET.Element(W + "r")
    rpr = template.find(W + "rPr")
    if rpr is not None:
        r.append(ET.fromstring(ET.tostring(rpr)))
    t = ET.SubElement(r, W + ("delText" if delete else "t"))
    t.text = text
    t.set(XMLSP, "preserve")
    return r


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("edits")
    ap.add_argument("--out")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--author", default="Financial Reporting Review")
    ap.add_argument("--initials", default="FR")
    a = ap.parse_args()

    edits = json.load(open(a.edits, encoding="utf-8"))
    when = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    zin = zipfile.ZipFile(a.docx)
    raw = zin.read("word/document.xml").decode("utf-8")
    register(raw)
    doc = ET.fromstring(raw)

    # parent map, because ElementTree has none
    parent = {c: p for p in doc.iter() for c in p}

    # existing ids so we never collide
    used_rev = {int(e.get(W + "id")) for e in doc.iter()
                if e.tag in (W + "ins", W + "del") and (e.get(W + "id") or "").isdigit()}
    rev_id = max(used_rev or {999}) + 1

    comments_raw = None
    if "word/comments.xml" in zin.namelist():
        comments_raw = zin.read("word/comments.xml").decode("utf-8")
        croot = ET.fromstring(comments_raw)
        used_c = {int(c.get(W + "id")) for c in croot.findall(W + "comment")
                  if (c.get(W + "id") or "").isdigit()}
    else:
        croot = ET.Element(W + "comments")
        used_c = set()
    cmt_id = max(used_c or {0}) + 1

    applied, skipped, extra = [], [], []

    for e in edits:
        find = e["find"]
        want = e.get("occurrence", 1)
        hits = []
        for r in doc.iter(W + "r"):
            if parent.get(r) is None:
                continue
            if find in run_text(r):
                hits.append(r)

        if not hits:
            # is it findable at paragraph level? then it spans runs
            spans = any(find in "".join(run_text(x) for x in p.findall(W + "r"))
                        for p in doc.iter(W + "p"))
            skipped.append((e, "spans runs, run merge_runs.py first" if spans
                            else "text not found"))
            continue
        if want > len(hits):
            skipped.append((e, f"occurrence {want} requested, {len(hits)} found"))
            continue
        if len(hits) > 1 and "occurrence" not in e:
            extra.append((e, len(hits)))

        r = hits[want - 1]
        if a.dry_run:
            applied.append(e)
            continue

        p = parent[r]
        kids = list(p)
        at = kids.index(r)
        txt = run_text(r)
        i = txt.index(find)
        before, after = txt[:i], txt[i + len(find):]

        new = []
        if before:
            new.append(mkrun(r, before))

        has_edit = "replace" in e
        cmt = e.get("comment")
        cid = None
        if cmt:
            cid = cmt_id
            cmt_id += 1
            s = ET.Element(W + "commentRangeStart")
            s.set(W + "id", str(cid))
            new.append(s)

        if has_edit:
            d = ET.Element(W + "del")
            d.set(W + "id", str(rev_id)); rev_id += 1
            d.set(W + "author", a.author)
            d.set(W + "date", when)
            d.append(mkrun(r, find, delete=True))
            new.append(d)
            if e["replace"]:
                ins = ET.Element(W + "ins")
                ins.set(W + "id", str(rev_id)); rev_id += 1
                ins.set(W + "author", a.author)
                ins.set(W + "date", when)
                ins.append(mkrun(r, e["replace"]))
                new.append(ins)
        else:
            new.append(mkrun(r, find))

        if cmt:
            end = ET.Element(W + "commentRangeEnd")
            end.set(W + "id", str(cid))
            new.append(end)
            ref = ET.Element(W + "r")
            rpr = ET.SubElement(ref, W + "rPr")
            st = ET.SubElement(rpr, W + "rStyle")
            st.set(W + "val", "CommentReference")
            cr = ET.SubElement(ref, W + "commentReference")
            cr.set(W + "id", str(cid))
            new.append(ref)

            c = ET.SubElement(croot, W + "comment")
            c.set(W + "id", str(cid))
            c.set(W + "author", a.author)
            c.set(W + "initials", a.initials)
            c.set(W + "date", when)
            cp = ET.SubElement(c, W + "p")
            crun = ET.SubElement(cp, W + "r")
            ct = ET.SubElement(crun, W + "t")
            ct.text = (f"[{e['ref']}] " if e.get("ref") else "") + cmt
            ct.set(XMLSP, "preserve")

        if after:
            new.append(mkrun(r, after))

        p.remove(r)
        for off, el in enumerate(new):
            p.insert(at + off, el)
            parent[el] = p
        applied.append(e)

    n_ins = sum(1 for e in applied if e.get("replace"))
    n_del = sum(1 for e in applied if "replace" in e)
    n_cmt = sum(1 for e in applied if e.get("comment"))
    print(f"  deletions  {n_del}")
    print(f"  insertions {n_ins}")
    print(f"  comments   {n_cmt}")
    print(f"  applied    {len(applied)} of {len(edits)} edits")

    if extra:
        print("\nTEXT APPEARS MORE THAN ONCE, the first was used. Add \"occurrence\":")
        for e, n in extra:
            print(f"   {e.get('ref', ''):<10} {n} matches  {e['find'][:56]!r}")
    if skipped:
        print("\nNOT APPLIED:")
        for e, why in skipped:
            print(f"   {e.get('ref', ''):<10} {why:<40} {e['find'][:46]!r}")

    if a.dry_run:
        print("\nDry run. Nothing written.")
        return 0

    out = a.out or os.path.splitext(a.docx)[0] + " - REDLINED.docx"
    body = serialize(raw, doc)
    names = zin.namelist()
    need_comments = n_cmt > 0 or "word/comments.xml" in names

    patched = {"word/document.xml": body}

    if need_comments:
        # do NOT set xmlns:w by hand here: ET emits it from the registered
        # prefix and the result is a duplicate attribute, which Word reports
        # only as "the file appears to be corrupted"
        patched["word/comments.xml"] = (
            serialize(comments_raw, croot, "w:comments") if comments_raw
            else ET.tostring(croot, encoding="UTF-8", xml_declaration=True))
        # relationship
        ET.register_namespace("", RELNS)
        rels = ET.fromstring(zin.read("word/_rels/document.xml.rels").decode("utf-8"))
        if not any(r.get("Type") == REL_COMMENTS for r in rels):
            ids = [r.get("Id", "") for r in rels]
            n = 1
            while f"rId{n}" in ids:
                n += 1
            e = ET.SubElement(rels, "{%s}Relationship" % RELNS)
            e.set("Id", f"rId{n}")
            e.set("Type", REL_COMMENTS)
            e.set("Target", "comments.xml")
        patched["word/_rels/document.xml.rels"] = ET.tostring(
            rels, encoding="UTF-8", xml_declaration=True)
        del_unused = None
        # content type override
        ET.register_namespace("", CTNS)
        ct = ET.fromstring(zin.read("[Content_Types].xml").decode("utf-8"))
        if not any(o.get("PartName") == "/word/comments.xml" for o in ct):
            o = ET.SubElement(ct, "{%s}Override" % CTNS)
            o.set("PartName", "/word/comments.xml")
            o.set("ContentType", CT_COMMENTS)
        patched["[Content_Types].xml"] = ET.tostring(ct, encoding="UTF-8",
                                                     xml_declaration=True)

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for i in zin.infolist():
            zo.writestr(i, patched.pop(i.filename, None) or zin.read(i.filename))
        for name, data in patched.items():
            zo.writestr(name, data)
    zin.close()

    print(f"\nwrote {out}")
    print("\nVerify before sending. The failure mode here is silent:")
    print(f'   powershell -File verify_docx.ps1 "{out}"')
    return 0


if __name__ == "__main__":
    sys.exit(main())
