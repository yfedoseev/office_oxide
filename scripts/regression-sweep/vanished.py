"""Words that truly vanished, as opposed to being re-tokenised.

A word only counts as lost when it is absent from the new arm *and* does
not survive inside any new token (a join) and is not itself built from new
tokens (a split). Both of those are re-tokenisations, which the release
does deliberately: restoring `&` splits `Ramp` back into `R` and `amp`,
and separating a footnote mark splits `1Department`.
"""
import collections, json, os, re, subprocess, sys

SC = os.path.dirname(os.path.abspath(__file__))
PREV = os.environ.get("PREV_BIN") or os.path.join(SC, "target-v019/release/office-oxide")
NEXT = os.environ.get("NEXT_BIN") or os.path.join(os.environ["REPO"], "target/release/office-oxide")
CORPUS = os.path.join(SC, "corpus")
for b in (PREV, NEXT):
    if not os.path.isfile(b):
        sys.exit(f"missing binary: {b}")
skipped = collections.Counter()
TAG = re.compile(r"<[^>]+>")
WORD = re.compile(r"[^\W\d_]{3,}", re.UNICODE)   # alphabetic, 3+ chars

def run(binary, surface, path):
    try:
        p = subprocess.run([binary, surface, path], capture_output=True, timeout=120)
        return p.stdout.decode("utf-8", "replace") if p.returncode == 0 else None
    except Exception:
        return None

def words(s, surface):
    if surface in ("html", "ir"):
        s = TAG.sub(" ", s)
    return collections.Counter(WORD.findall(s))

files = sorted(os.listdir(CORPUS))
findings = []
for i, name in enumerate(files):
    if i % 200 == 0:
        print(f"  {i}/{len(files)}", file=sys.stderr, flush=True)
    path = os.path.join(CORPUS, name)
    for surface in ("text", "markdown", "html"):
        a, b = run(PREV, surface, path), run(NEXT, surface, path)
        if a is None or b is None:
            skipped[surface] += 1
            continue
        wa, wb = words(a, surface), words(b, surface)
        gone = [w for w in wa if wb[w] == 0]
        if not gone:
            continue
        newblob = " ".join(wb)
        real = []
        for w in gone:
            # A join: the word now lives inside a longer token.
            if w in newblob:
                continue
            # A split: the word was a concatenation of tokens the new arm
            # emits separately. Accept any prefix/suffix that is a new token.
            if any(w.startswith(t) or w.endswith(t) for t in wb if len(t) >= 3):
                continue
            real.append(w)
        if real:
            findings.append((len(real), name, surface, sorted(real)[:12]))

findings.sort(reverse=True)
with open(os.path.join(SC, "vanished.json"), "w") as fh:
    json.dump([{"n": n, "file": f_, "surface": s_, "words": w}
               for n, f_, s_, w in findings], fh, indent=1)
print(f"skipped (extraction failed on one arm): {dict(skipped)}")
print(f"\n=== files with genuinely vanished words: {len(findings)} ===")
for n, name, surface, sample in findings:
    print(f"  {n:5d}  {surface:9s} {name[:56]:56s} {sample}")
