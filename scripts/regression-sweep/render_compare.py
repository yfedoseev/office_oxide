"""Structural comparison of office_oxide's markdown against pandoc's GFM for
every .docx in a sweep: headings, list items, table rows, emphasis, links.

  python3 render_compare.py SWEEP_DIR CORPUS_ROOT OUT.jsonl [--jobs N] [--top N]

SWEEP_DIR is a sweep_tree.py output directory (its out/<rel>.markdown.gz
files are read; nothing is re-run on our side). pandoc is run here with
`-t gfm --wrap=none` and cached under OUT.jsonl's directory.

Word-level text agreement is the quorum's job; this asks whether the
*structure* survived: a document pandoc renders with 12 headings and 3
tables should not come back from us as flat paragraphs. Each counter is
compared as a ratio and a file is flagged when a structure pandoc finds
several of is absent or off by more than half on our side (or vice versa).
Both readers have known blind spots — pandoc drops headers/footers and text
boxes, we hide tracked deletions — so a flag is a lead, not a verdict.
"""
import gzip, json, os, re, subprocess, sys
from concurrent.futures import ThreadPoolExecutor

HEADING = re.compile(r"^#{1,6} \S", re.M)
LIST_ITEM = re.compile(r"^\s*(?:[-*+]|\d+[.)]) \S", re.M)
TABLE_ROW = re.compile(r"^\|.*\|\s*$", re.M)
TABLE_SEP = re.compile(r"^\|?\s*:?-{2,}:?\s*(?:\|\s*:?-{2,}:?\s*)*\|?\s*$", re.M)
BOLD = re.compile(r"\*\*\S")
ITALIC = re.compile(r"(?<![*\w])\*(?!\*)\S|(?<!\w)_\S")
LINK = re.compile(r"\]\((?!#)")
IMAGE = re.compile(r"!\[")
CODE = re.compile(r"^```", re.M)

def counts(md):
    return {
        "headings": len(HEADING.findall(md)),
        "list_items": len(LIST_ITEM.findall(md)),
        "table_rows": len(TABLE_ROW.findall(md)) - len(TABLE_SEP.findall(md)),
        "bold": len(BOLD.findall(md)),
        "italic": len(ITALIC.findall(md)),
        "links": len(LINK.findall(md)),
        "images": len(IMAGE.findall(md)),
        "words": len(re.findall(r"\w{2,}", md)),
    }

def pandoc_gfm(path, cache):
    if os.path.exists(cache):
        return open(cache, encoding="utf-8", errors="replace").read()
    try:
        p = subprocess.run(["pandoc", "-t", "gfm", "--wrap=none", path], capture_output=True, timeout=120)
    except subprocess.TimeoutExpired:
        return None
    if p.returncode != 0:
        return None
    md = p.stdout.decode("utf-8", "replace")
    os.makedirs(os.path.dirname(cache), exist_ok=True)
    open(cache, "w", encoding="utf-8").write(md)
    return md

def flags(ours, theirs):
    out = []
    for k in ("headings", "list_items", "table_rows", "links", "bold", "italic"):
        a, b = ours[k], theirs[k]
        if b >= 3 and a < b * 0.5:
            out.append(f"{k}: ours {a} vs pandoc {b}")
        elif a >= 3 and b < a * 0.5 and k in ("headings", "table_rows"):
            out.append(f"{k}: ours {a} vs pandoc {b} (more than pandoc)")
    return out

def main():
    sweep, root, outfile = sys.argv[1:4]
    jobs = 4; top = 40
    a = sys.argv[4:]
    while a:
        k = a.pop(0)
        if k == "--jobs": jobs = int(a.pop(0))
        elif k == "--top": top = int(a.pop(0))
    cache_root = os.path.join(os.path.dirname(os.path.abspath(outfile)), "pandoc_gfm")
    files = []
    for line in open(os.path.join(sweep, "sweep.jsonl")):
        r = json.loads(line)
        if r["surface"] == "markdown" and r["status"] == "ok" and r["path"].startswith("docx/"):
            files.append(r["path"])
    files = sorted(set(files))

    def one(rel):
        ours_path = os.path.join(sweep, "out", rel + ".markdown.gz")
        try:
            ours = gzip.open(ours_path, "rt", encoding="utf-8", errors="replace").read()
        except Exception:
            return None
        theirs = pandoc_gfm(os.path.join(root, rel), os.path.join(cache_root, rel + ".md"))
        if theirs is None:
            return {"path": rel, "pandoc": "err"}
        co, ct = counts(ours), counts(theirs)
        return {"path": rel, "ours": co, "pandoc": ct, "flags": flags(co, ct)}

    out = open(outfile, "w")
    flagged = []
    with ThreadPoolExecutor(jobs) as ex:
        for i, rec in enumerate(ex.map(one, files)):
            if rec is None: continue
            out.write(json.dumps(rec) + "\n")
            if rec.get("flags"): flagged.append(rec)
            if (i + 1) % 500 == 0: print(f"  {i+1}/{len(files)}", file=sys.stderr, flush=True)
    out.close()
    print(f"{len(files)} docx compared, {len(flagged)} flagged")
    from collections import Counter
    kinds = Counter(f.split(":")[0] + (" (more)" if "more than" in f else "") for r in flagged for f in r["flags"])
    for k, v in kinds.most_common(): print(f"  {v:5d}  {k}")
    print("\nlargest structural gaps (by pandoc count of the flagged structure):")
    def weight(r):
        return max(r["pandoc"][f.split(":")[0]] for f in r["flags"])
    for r in sorted(flagged, key=lambda r: -weight(r))[:top]:
        print(f"  {r['path']}\n      " + "; ".join(r["flags"]))

if __name__ == "__main__":
    main()
