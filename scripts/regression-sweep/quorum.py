"""Quorum analysis: for every corpus document, compare office_oxide 0.1.12
(and 0.1.11) plain text against the independent reference panel.

  python3 quorum.py PREV_DIR NEXT_DIR TIKA_OUT PANEL_OUT MANIFEST REPORT.json

Per file: word multisets from each source; for each reference R,
recall(X, R) = sum(min(X[w], R[w])) / sum(R[w]) — the share of R's words that
X also emits. A file is flagged when office_oxide is the odd one out:
at least two references agree with each other (recall >= 0.85 both ways)
while office_oxide's recall against them is below 0.7 — or when office_oxide
emits nothing and the references emit real text.
"""
import collections, gzip, json, os, re, sys

WORD = re.compile(r"[^\W\d_]{3,}", re.UNICODE)
FMTS = {"doc", "docx", "ppt", "pptx", "xls", "xlsx"}

def words(s): return collections.Counter(WORD.findall(s))
def recall(x, r):
    tot = sum(r.values())
    if tot == 0: return None
    return sum(min(x[w], r[w]) for w in r) / tot

def load_sweep(d):
    recs = {}
    for line in open(os.path.join(d, "sweep.jsonl")):
        r = json.loads(line)
        if r["surface"] == "text": recs[r["path"]] = r
    return recs
def sweep_text(d, rel):
    p = os.path.join(d, "out", rel + ".text.gz")
    if not os.path.exists(p): return None
    with gzip.open(p, "rb") as f: return f.read().decode("utf-8", "replace")
def tika_text(tika_out, rel):
    p = os.path.join(tika_out, rel + ".txt")
    if not os.path.exists(p): return None
    return open(p, encoding="utf-8", errors="replace").read()

def main():
    prev_d, next_d, tika_out, panel_out, manifest, report = sys.argv[1:7]
    P, N = load_sweep(prev_d), load_sweep(next_d)
    docs = [d["path"] for d in json.load(open(manifest))["documents"] if d["path"].split("/",1)[0] in FMTS]
    panel = collections.defaultdict(dict)
    if os.path.exists(os.path.join(panel_out, "panel.jsonl")):
        for line in open(os.path.join(panel_out, "panel.jsonl")):
            r = json.loads(line)
            if "path" in r: panel[r["path"]][r["tool"]] = r
    rows = []
    for rel in docs:
        nt, pt = sweep_text(next_d, rel), sweep_text(prev_d, rel)
        refs = {}
        t = tika_text(tika_out, rel)
        if t is not None: refs["tika"] = words(t)
        for tool, r in panel.get(rel, {}).items():
            if r["status"] == "ok":
                p = os.path.join(panel_out, tool, rel + ".txt")
                if os.path.exists(p): refs[tool] = words(open(p, encoding="utf-8", errors="replace").read())
        nw = words(nt) if nt is not None else None
        pw = words(pt) if pt is not None else None
        row = {"path": rel, "fmt": rel.split("/",1)[0],
               "next_status": N.get(rel, {}).get("status", "missing"), "prev_status": P.get(rel, {}).get("status", "missing"),
               "next_words": sum(nw.values()) if nw else 0, "prev_words": sum(pw.values()) if pw else 0,
               "refs": {k: sum(v.values()) for k, v in refs.items()},
               "next_recall": {}, "prev_recall": {}, "flags": []}
        for k, rw in refs.items():
            if nw is not None: row["next_recall"][k] = recall(nw, rw)
            if pw is not None: row["prev_recall"][k] = recall(pw, rw)
        # reference agreement: do any two refs agree with each other?
        ks = list(refs)
        agreeing = set()
        for i in range(len(ks)):
            for j in range(i+1, len(ks)):
                a, b = refs[ks[i]], refs[ks[j]]
                ra, rb = recall(a, b), recall(b, a)
                if ra is not None and rb is not None and ra >= 0.85 and rb >= 0.85 and sum(a.values()) >= 20:
                    agreeing |= {ks[i], ks[j]}
        row["agreeing_refs"] = sorted(agreeing)
        if agreeing:
            low = [k for k in agreeing if row["next_recall"].get(k) is not None and row["next_recall"][k] < 0.7]
            if nw is not None and len(low) >= 2:
                row["flags"].append("odd_one_out")
            if nw is not None and sum(nw.values()) == 0:
                row["flags"].append("empty_vs_refs")
            if nw is None and N.get(rel, {}).get("status") != "ok":
                row["flags"].append("error_vs_refs")
        # regression vs prev on the same refs
        for k in refs:
            a, b = row["prev_recall"].get(k), row["next_recall"].get(k)
            if a is not None and b is not None and a - b >= 0.1 and sum(refs[k].values()) >= 20:
                row["flags"].append(f"regressed_vs_{k}")
        # over-extraction: office_oxide emits far more than every reference
        if nw is not None and refs and sum(nw.values()) >= 50:
            if all(recall(rw, nw) is not None and sum(nw.values()) > 3 * sum(rw.values()) for rw in refs.values()) and len(refs) >= 2:
                row["flags"].append("3x_more_than_all_refs")
        rows.append(row)
    json.dump(rows, open(report, "w"), indent=1)
    flagged = [r for r in rows if r["flags"]]
    c = collections.Counter((r["fmt"], f) for r in flagged for f in r["flags"])
    print(f"documents: {len(rows)}, with >=1 reference: {sum(1 for r in rows if r['refs'])}, flagged: {len(flagged)}")
    for (fmt, f), n in sorted(c.items()): print(f"  {fmt:5} {f:28} {n}")

if __name__ == "__main__":
    main()
