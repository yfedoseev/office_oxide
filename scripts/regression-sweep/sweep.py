"""Run one arm's CLI over the corpus and record status + output hashes.

Emits one JSON line per (file, surface):
  {path, surface, status, code, bytes, sha, err}

`status` is ok | err | timeout | crash — the axis that matters most for
0.1.10, which converts several silent-Ok paths into Err.
"""
import hashlib, json, os, subprocess, sys

BIN = sys.argv[1]
CORPUS = sys.argv[2]
OUTFILE = sys.argv[3]
SURFACES = ["text", "markdown", "html", "ir"]
TIMEOUT = 120

files = sorted(
    os.path.join(CORPUS, f) for f in os.listdir(CORPUS)
    if os.path.isfile(os.path.join(CORPUS, f))
)

with open(OUTFILE, "w") as out:
    for i, path in enumerate(files):
        for surface in SURFACES:
            rec = {"path": os.path.basename(path), "surface": surface}
            try:
                p = subprocess.run(
                    [BIN, surface, path],
                    capture_output=True, timeout=TIMEOUT,
                )
                rec["code"] = p.returncode
                if p.returncode == 0:
                    rec["status"] = "ok"
                    rec["bytes"] = len(p.stdout)
                    rec["sha"] = hashlib.sha256(p.stdout).hexdigest()[:16]
                else:
                    # A signal shows as a negative return code.
                    rec["status"] = "crash" if p.returncode < 0 else "err"
                    rec["err"] = p.stderr.decode("utf-8", "replace").strip()[:400]
            except subprocess.TimeoutExpired:
                rec["status"] = "timeout"
            except Exception as e:
                rec["status"] = "crash"
                rec["err"] = str(e)[:400]
            out.write(json.dumps(rec) + "\n")
        if (i + 1) % 50 == 0:
            print(f"  {i+1}/{len(files)}", file=sys.stderr, flush=True)
print(f"swept {len(files)} files x {len(SURFACES)} surfaces -> {OUTFILE}", file=sys.stderr)
