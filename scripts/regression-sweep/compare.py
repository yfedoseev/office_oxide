"""Compare two arms and emit a work list, not a verdict.

Axis 1 (decides shippability): status transitions. 0.1.10 converts several
silent-Ok paths into Err, and a file that used to extract partial text and
now returns Err is a regression from a user's perspective even when the
error is technically correct.

Axis 2: content change on files that are Ok in both arms.
"""
import collections, json, os, sys

def load(p):
    d = {}
    with open(p) as f:
        for line in f:
            r = json.loads(line)
            d[(r["path"], r["surface"])] = r
    return d

SC = os.path.dirname(os.path.abspath(__file__))
prev = load(os.path.join(SC, "prev.jsonl"))
nxt = load(os.path.join(SC, "next.jsonl"))

keys = sorted(set(prev) | set(nxt))
trans = collections.Counter()
regressions = []       # ok -> not-ok
recoveries = []        # not-ok -> ok
content_changed = []
byte_delta = collections.Counter()

for k in keys:
    a, b = prev.get(k), nxt.get(k)
    if a is None or b is None:
        continue
    sa, sb = a["status"], b["status"]
    trans[(sa, sb)] += 1
    if sa == "ok" and sb != "ok":
        regressions.append((k, b.get("err", "")[:200], sb))
    elif sa != "ok" and sb == "ok":
        recoveries.append((k, sa))
    elif sa == "ok" and sb == "ok" and a["sha"] != b["sha"]:
        content_changed.append((k, a["bytes"], b["bytes"]))
        byte_delta[k[1]] += b["bytes"] - a["bytes"]

files = len({k[0] for k in keys})
print(f"corpus: {files} files x 4 surfaces = {len(keys)} observations\n")

print("== status transitions (prev -> next) ==")
for (sa, sb), n in sorted(trans.items(), key=lambda x: -x[1]):
    mark = "  <-- REGRESSION" if (sa == "ok" and sb != "ok") else (
           "  <-- recovery" if (sa != "ok" and sb == "ok") else "")
    print(f"  {sa:8} -> {sb:8} {n:6}{mark}")

print(f"\n== AXIS 1: {len(regressions)} ok->not-ok observations "
      f"({len({r[0][0] for r in regressions})} distinct files) ==")
by_err = collections.Counter()
for (path, surface), err, st in regressions:
    # First line of the message is the error class.
    cls = err.split("\n")[0][:110] if err else f"({st})"
    by_err[cls] += 1
for cls, n in by_err.most_common(30):
    print(f"  {n:5}  {cls}")

print(f"\n== recoveries: {len(recoveries)} observations "
      f"({len({r[0][0] for r in recoveries})} distinct files) ==")

print(f"\n== AXIS 2: {len(content_changed)} ok/ok observations with changed output ==")
per_surface = collections.Counter(k[1] for k, _, _ in content_changed)
for s, n in per_surface.most_common():
    print(f"  {s:9} {n:5} changed, net {byte_delta[s]:+d} bytes")

with open(os.path.join(SC, "regressions.txt"), "w") as f:
    for (path, surface), err, st in sorted(regressions):
        f.write(f"{path}\t{surface}\t{st}\t{err}\n")
with open(os.path.join(SC, "content_changed.txt"), "w") as f:
    for (path, surface), ba, bb in sorted(content_changed):
        f.write(f"{path}\t{surface}\t{ba}\t{bb}\t{bb-ba:+d}\n")
print("\nwrote regressions.txt and content_changed.txt")
