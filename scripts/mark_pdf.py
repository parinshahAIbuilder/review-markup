#!/usr/bin/env python
"""Highlight findings in a PDF with severity coloured popup comments and a legend.

    python mark_pdf.py in.pdf marks.json --dry-run
    python mark_pdf.py in.pdf marks.json --out "in - MARKED UP.pdf"

Run --dry-run first, always. It reports how many times each anchor was found and
lists every anchor that matched nothing, without writing a file.

marks.json is a list of:

    {"page": 7, "anchor": "13,059.27", "sev": "C", "ref": "SA-01",
     "note": "Total current liabilities re foots to 13,059.28.",
     "occurrence": 1}

page is the one based PDF page, which is usually not the printed folio number.
occurrence is optional and one based; omit it to mark every match on the page.
sev takes C H M L (severity) or DONE OPEN NEW (status). Do not mix the two
palettes in one document.
"""
import argparse, json, os, re, sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PAL = {
    "C":    ((1.00, 0.42, 0.42), "CRITICAL", "must be corrected before signing"),
    "H":    ((1.00, 0.66, 0.20), "HIGH", "correct before circulation"),
    "M":    ((1.00, 0.90, 0.35), "MEDIUM", "disclosure or presentation"),
    "L":    ((0.62, 0.80, 1.00), "LOW", "typographical"),
    "DONE": ((0.55, 0.85, 0.55), "DONE", "verified as incorporated"),
    "OPEN": ((1.00, 0.75, 0.00), "OPEN", "previously raised, still not done"),
    "NEW":  ((1.00, 0.42, 0.42), "NEW", "new issue on this version"),
}
SEVERITY = ("C", "H", "M", "L")
STATUS = ("DONE", "OPEN", "NEW")

# what typesetters substitute for the characters you typed
SPACES = (" ", " ", " ", " ")
DASHES = ("-", "–", "−", "—")


def variants(anchor):
    """Yield (text, description) candidates. search_for matches literal text
    only, so flexibility has to come from substituting the characters, not from
    a pattern."""
    seen = set()

    def emit(t, how):
        if t and t not in seen:
            seen.add(t)
            return t, how
        return None

    cands = [(anchor, "exact")]
    if " " in anchor:
        for sp in SPACES:
            cands.append((anchor.replace(" ", sp), f"space as U+{ord(sp):04X}"))
        cands.append((anchor.replace(" ", ""), "spaces removed"))
    for d in DASHES:
        if any(x in anchor for x in DASHES if x != d):
            for x in DASHES:
                if x != d and x in anchor:
                    cands.append((anchor.replace(x, d), f"U+{ord(x):04X} as U+{ord(d):04X}"))
    for t, how in cands:
        r = emit(t, how)
        if r:
            yield r


def search(page, anchor):
    for text, how in variants(anchor):
        try:
            r = page.search_for(text)
        except Exception:
            r = []
        if r:
            return r, how
    return [], None


def legend(page, codes, palette_name, pymupdf):
    used = [c for c in (SEVERITY + STATUS) if c in codes]
    if not used:
        return
    h = 20 + 13 * len(used)
    r = pymupdf.Rect(36, page.rect.height - h - 24, min(300, page.rect.width - 36),
                     page.rect.height - 24)
    page.draw_rect(r, color=(0.45, 0.45, 0.45), fill=(1, 1, 1), width=0.6)
    y = r.y0 + 12
    page.insert_text((r.x0 + 7, y), f"Colours indicate {palette_name}.",
                     fontsize=6.5, fontname="hebo", color=(0.2, 0.2, 0.2))
    for c in used:
        y += 13
        rgb, lab, desc = PAL[c]
        sw = pymupdf.Rect(r.x0 + 7, y - 6, r.x0 + 19, y + 1)
        page.draw_rect(sw, color=(0.5, 0.5, 0.5), fill=rgb, width=0.3)
        page.insert_text((r.x0 + 24, y), f"{lab} - {desc}".replace(" - ", "   "),
                         fontsize=6.5, color=(0.2, 0.2, 0.2))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("marks")
    ap.add_argument("--out")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--author", default="Financial Reporting Review")
    ap.add_argument("--opacity", type=float, default=0.45)
    ap.add_argument("--no-legend", action="store_true")
    a = ap.parse_args()

    marks = json.load(open(a.marks, encoding="utf-8"))
    import pymupdf
    d = pymupdf.open(a.pdf)
    pages_in = len(d)

    bad = [m for m in marks if m.get("sev") not in PAL]
    if bad:
        print("UNKNOWN SEVERITY CODES. Valid: " + ", ".join(PAL))
        for m in bad:
            print(f"   {m.get('ref', '?')}  sev={m.get('sev')!r}")
        return 2

    codes = {m["sev"] for m in marks}
    mixed = codes & set(SEVERITY) and codes & set(STATUS)
    if mixed:
        print("WARNING: severity and status codes both present. The legend cannot")
        print("say what the colours mean. Split into two documents.\n")
    palette_name = "status of points raised on the previous draft" \
        if codes & set(STATUS) else "severity"

    applied, unmatched, over = {}, [], []
    for m in marks:
        pg = int(m["page"])
        if not 1 <= pg <= pages_in:
            unmatched.append((m, f"page {pg} outside 1 to {pages_in}"))
            continue
        page = d[pg - 1]                       # hold the page object
        rects, how = search(page, m["anchor"])
        if not rects:
            unmatched.append((m, "anchor not found on page"))
            continue

        occ = m.get("occurrence")
        if occ:
            if occ > len(rects):
                unmatched.append((m, f"occurrence {occ} requested, {len(rects)} found"))
                continue
            rects = [rects[occ - 1]]
        elif len(rects) > 1:
            over.append((m, len(rects)))

        if a.dry_run:
            applied[m["sev"]] = applied.get(m["sev"], 0) + 1
            print(f"  p{pg:<4} {m.get('ref', ''):<10} [{m['sev']:<4}] "
                  f"{len(rects)} match{'es' if len(rects) > 1 else ' '}  "
                  f"({how})  {m['anchor'][:46]!r}")
            continue

        rgb, lab, desc = PAL[m["sev"]]
        an = page.add_highlight_annot(rects)
        an.set_colors(stroke=rgb)
        an.set_info(title=a.author,
                    subject=f"[{m['sev']}] {m.get('ref', '')}".strip(),
                    content=f"{lab} - {desc}\n\n{m.get('note', '')}".replace(" - ", "   "))
        an.set_opacity(a.opacity)
        an.update()
        applied[m["sev"]] = applied.get(m["sev"], 0) + 1

    if not a.dry_run and not a.no_legend:
        legend(d[0], codes, palette_name, pymupdf)

    print()
    for c in SEVERITY + STATUS:
        if c in applied:
            print(f"  {PAL[c][1]:<9} {applied[c]}")
    print(f"  {'TOTAL':<9} {sum(applied.values())} of {len(marks)} marks")

    if over:
        print("\nMULTIPLE MATCHES, all were marked. Add \"occurrence\" to narrow:")
        for m, n in over:
            print(f"   p{m['page']:<4} {m.get('ref', ''):<10} {n} matches  {m['anchor'][:50]!r}")

    if unmatched:
        print("\nANCHORS THAT MATCHED NOTHING. Hand this list to the user with the")
        print("file: either the page reference is wrong or the point is already fixed,")
        print("and only they can say which.")
        for m, why in unmatched:
            print(f"   p{m['page']:<4} {m.get('ref', ''):<10} {why:<42} {m['anchor'][:40]!r}")

    if a.dry_run:
        print("\nDry run. Nothing written. Re run without --dry-run when the")
        print("unmatched list is empty or understood.")
        return 0

    out = a.out or os.path.splitext(a.pdf)[0] + " - MARKED UP.pdf"
    d.save(out, garbage=3, deflate=True)
    d.close()

    chk = pymupdf.open(out)
    n_annots = sum(len(list(p.annots())) for p in chk)
    print(f"\nwrote {out}")
    print(f"  pages {len(chk)} (input {pages_in}){'  PAGE COUNT CHANGED' if len(chk) != pages_in else ''}")
    print(f"  annotations present in the output file: {n_annots}")
    chk.close()
    if n_annots < sum(applied.values()):
        print("  FEWER ANNOTATIONS THAN MARKS APPLIED. Do not send this file.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
