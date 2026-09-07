"""Classify content changes on the `text` surface into intended vs unexplained.

Per the method doc: emit a work list, not a verdict. Every filter lives here
so a later script cannot forget one.
"""
import collections, json, os, re, subprocess, sys

SC = os.path.dirname(os.path.abspath(__file__))
PREV = os.path.join(SC, "target-v019/release/office-oxide")
NEXT = os.path.join(os.environ["REPO"], "target/release/office-oxide")
CORPUS = os.path.join(SC, "corpus")
SURFACE = sys.argv[1] if len(sys.argv) > 1 else "text"

def run(binary, path):
    p = subprocess.run([binary, SURFACE, path], capture_output=True, timeout=120)
    return p.stdout.decode("utf-8", "replace") if p.returncode == 0 else None

def tokens(s):
    """Word multiset. A *set* cannot see a doubled table row, and
    deduplication is often exactly what a release changes."""
    return collections.Counter(re.findall(r"\w+", s))

changed = []
for line in open(os.path.join(SC, "content_changed.txt")):
    path, surface, ba, bb, delta = line.rstrip("\n").split("\t")
    if surface == SURFACE:
        changed.append(path)

print(f"{len(changed)} files changed on `{SURFACE}`\n")

CLASSES = collections.Counter()
work = []
for name in sorted(changed):
    full = os.path.join(CORPUS, name)
    a, b = run(PREV, full), run(NEXT, full)
    if a is None or b is None:
        CLASSES["status changed (handled on axis 1)"] += 1
        continue

    ta, tb = tokens(a), tokens(b)
    lost = ta - tb          # tokens present fewer times in next
    gained = tb - ta

    # --- Intended-change filters, applied before anything is called a defect.
    reasons = []
    # #166: plain-text break marker changed from markdown's `---` to U+000C.
    if "\x0c" in b and "---" in a and "\x0c" not in a:
        reasons.append("break-marker ---/FF (#166)")
        lost.pop("", None)
    # #156: entities resolved, so `ATT` becomes `AT`+`T` under \w tokenising.
    if any(c in b for c in "&<>") and not any(c in a for c in "&<>"):
        reasons.append("entities resolved (#156)")
    # #170: PPT master boilerplate removed.
    if "Click to edit Master" in a and "Click to edit Master" not in b:
        reasons.append("PPT master boilerplate removed (#170)")
    # #137: section title no longer duplicated.
    # #166: headers/footers now included in text.
    if not lost and gained:
        reasons.append("content added only")
    if lost and not gained:
        reasons.append("content removed only")

    net = sum(gained.values()) - sum(lost.values())
    key = "; ".join(reasons) if reasons else "UNEXPLAINED"
    CLASSES[key] += 1
    if not reasons or (lost and sum(lost.values()) > 0):
        work.append((sum(lost.values()), name, net,
                     list(lost.items())[:6], list(gained.items())[:6]))

print("== classification ==")
for k, n in CLASSES.most_common():
    print(f"  {n:5}  {k}")

work.sort(reverse=True)
print(f"\n== work list: {len(work)} files where tokens were LOST ==")
for nlost, name, net, lost, gained in work[:40]:
    print(f"\n  {name}  (-{nlost} tokens, net {net:+d})")
    print(f"    lost:   {lost}")
    print(f"    gained: {gained}")

with open(os.path.join(SC, f"worklist_{SURFACE}.txt"), "w") as f:
    for nlost, name, net, lost, gained in work:
        f.write(f"{name}\t{nlost}\t{net}\t{lost}\t{gained}\n")
