"""Compare two sweep2 output dirs (prev, next).

  python3 compare2.py PREV_DIR NEXT_DIR REPORT_DIR

Axis 1 - status transitions per surface, triaged: an error that replaced
         real bytes is a regression; one that replaced 0 bytes is not.
Axis 2 - content change on files ok in both arms: word-multiset diff with
         split/join awareness (see vanished.py); "lost" = absent from next and
         not explained by re-tokenisation. Also "gained" the same way.
Axis 3 - timing: next >= 3x prev and >= 2 s.
Writes REPORT_DIR/{transitions,content,timing}.json and prints a summary.
"""
import collections, gzip, json, os, re, sys
from concurrent.futures import ProcessPoolExecutor

TAG = re.compile(r"<[^>]+>")
# Inline tags are removed, not spaced: a run wrapped in its own <span> is
# still one word, and spacing it turned `xxxx` into `x x x x` — 500 false
# "lost words" on the html surface in one sweep.
INLINE = re.compile(r"</?(span|b|i|u|a|em|strong|sup|sub|s|del|ins|code|font|mark)\b[^>]*>")
WORD = re.compile(r"[^\W\d_]{3,}", re.UNICODE)

def load(d):
    recs = {}
    with open(os.path.join(d, "sweep.jsonl")) as f:
        for line in f:
            r = json.loads(line); recs[(r["path"], r["surface"])] = r
    return recs

def read_out(d, path, surface):
    p = os.path.join(d, "out", path + "." + surface + ".gz")
    with gzip.open(p, "rb") as f:
        return f.read().decode("utf-8", "replace")

def words(s, surface):
    if surface == "html":
        s = INLINE.sub("", s)
    if surface in ("html", "ir"):
        s = TAG.sub(" ", s)
    return collections.Counter(WORD.findall(s))

def unexplained(gone, other):
    blob = " ".join(other)
    toks = [t for t in other if len(t) >= 3]
    real = []
    for w in gone:
        if w in blob: continue                          # join
        if any(w.startswith(t) or w.endswith(t) for t in toks): continue  # split
        real.append(w)
    return real

def content_one(args):
    prev_d, next_d, path, surface = args
    try:
        a, b = read_out(prev_d, path, surface), read_out(next_d, path, surface)
    except Exception as e:
        return {"path": path, "surface": surface, "error": str(e)}
    wa, wb = words(a, surface), words(b, surface)
    lost = unexplained([w for w in wa if wb[w] == 0], wb)
    gained = unexplained([w for w in wb if wa[w] == 0], wa)
    # count-level drift (same vocabulary, different multiplicity) — dedupe or duplication
    fewer = sum(max(0, wa[w] - wb[w]) for w in wa if wb[w] > 0)
    more = sum(max(0, wb[w] - wa[w]) for w in wb if wa[w] > 0)
    return {"path": path, "surface": surface, "prev_bytes": len(a), "next_bytes": len(b),
            "prev_words": sum(wa.values()), "next_words": sum(wb.values()),
            "lost": len(lost), "gained": len(gained), "fewer": fewer, "more": more,
            "lost_sample": sorted(lost)[:15], "gained_sample": sorted(gained)[:15]}

def main():
    prev_d, next_d, rep = sys.argv[1:4]
    os.makedirs(rep, exist_ok=True)
    P, N = load(prev_d), load(next_d)
    keys = sorted(set(P) | set(N))
    # ---- axis 1
    trans = collections.Counter(); rows = []
    for k in keys:
        a, b = P.get(k), N.get(k)
        sa, sb = (a or {}).get("status", "missing"), (b or {}).get("status", "missing")
        if sa != sb:
            trans[(k[1], sa, sb)] += 1
            rows.append({"path": k[0], "surface": k[1], "prev": sa, "next": sb,
                         "prev_bytes": (a or {}).get("bytes", 0), "next_bytes": (b or {}).get("bytes", 0),
                         "prev_err": (a or {}).get("err", "")[:200], "next_err": (b or {}).get("err", "")[:200]})
    json.dump(rows, open(os.path.join(rep, "transitions.json"), "w"), indent=1)
    print("== status transitions (surface, prev -> next) ==")
    for (s, sa, sb), n in sorted(trans.items(), key=lambda x: -x[1]):
        print(f"  {n:5}  {s:9} {sa:8} -> {sb}")
    reg = [r for r in rows if r["prev"] == "ok" and r["next"] != "ok" and r["prev_bytes"] > 0]
    print(f"  ok->not-ok replacing real bytes: {len(reg)}")
    # ---- axis 2
    changed = [(prev_d, next_d, k[0], k[1]) for k in keys
               if k in P and k in N and P[k]["status"] == "ok" and N[k]["status"] == "ok"
               and P[k]["sha"] != N[k]["sha"]]
    print(f"== content: {len(changed)} (file,surface) pairs differ by hash; diffing words ==")
    with ProcessPoolExecutor(7) as ex:
        res = list(ex.map(content_one, changed, chunksize=32))
    json.dump(res, open(os.path.join(rep, "content.json"), "w"), indent=1)
    by_surface = collections.Counter(r["surface"] for r in res if r.get("lost"))
    print(f"  pairs with unexplained lost words: {sum(by_surface.values())} {dict(by_surface)}")
    print(f"  pairs with only gains: {sum(1 for r in res if r.get('gained') and not r.get('lost'))}")
    print(f"  pairs with count drift only: {sum(1 for r in res if not r.get('lost') and not r.get('gained') and (r.get('fewer') or r.get('more')))}")
    # ---- axis 3
    slow = []
    for k in keys:
        a, b = P.get(k), N.get(k)
        if a and b and "ms" in a and "ms" in b and b["ms"] >= 3 * a["ms"] and b["ms"] >= 2000:
            slow.append({"path": k[0], "surface": k[1], "prev_ms": a["ms"], "next_ms": b["ms"]})
    json.dump(slow, open(os.path.join(rep, "timing.json"), "w"), indent=1)
    print(f"== timing: {len(slow)} pairs >=3x slower and >=2s ==")
    for r in sorted(slow, key=lambda r: -r["next_ms"])[:15]:
        print(f"  {r['next_ms']:9.0f}ms (was {r['prev_ms']:7.0f})  {r['surface']:9} {r['path']}")
    # totals
    for d, R in (("prev", P), ("next", N)):
        c = collections.Counter((r["surface"], r["status"]) for r in R.values())
        print(f"== {d} totals ==", dict(sorted(c.items())))

if __name__ == "__main__":
    main()
