"""For every ok->not-ok transition, report what v0.1.9 actually produced.

An error replacing an *empty* success is an improvement (the file was
unreadable either way, and now says so). An error replacing *real extracted
content* is a genuine regression: the user lost text they used to get.
"""
import collections, json, os

SC = os.path.dirname(os.path.abspath(__file__))

def load(p):
    d = {}
    with open(os.path.join(SC, p)) as f:
        for line in f:
            r = json.loads(line)
            d[(r["path"], r["surface"])] = r
    return d

prev, nxt = load("prev.jsonl"), load("next.jsonl")

# `text` is the surface that decides whether content was lost.
rows = []
for (path, surface), a in prev.items():
    if surface != "text":
        continue
    b = nxt.get((path, surface))
    if b is None or a["status"] != "ok" or b["status"] == "ok":
        continue
    rows.append((a.get("bytes", 0), path, b["status"], b.get("err", "").split("\n")[0][:90]))

rows.sort(reverse=True)
lost = [r for r in rows if r[0] > 0]
empty = [r for r in rows if r[0] == 0]

print(f"ok->not-ok on the `text` surface: {len(rows)} files")
print(f"  v0.1.9 produced NOTHING (0 bytes): {len(empty)}  <- error is an improvement")
print(f"  v0.1.9 produced CONTENT:          {len(lost)}  <- REAL REGRESSION\n")

if lost:
    print("== files that lost real extracted content ==")
    for n, path, st, err in lost:
        print(f"  {n:9} bytes  {path[:58]:58}  {st}  {err}")

print("\n== error classes over files that produced nothing ==")
c = collections.Counter(err for _, _, _, err in empty)
for err, n in c.most_common():
    print(f"  {n:4}  {err}")
