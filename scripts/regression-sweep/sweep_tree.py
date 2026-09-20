"""Parallel two-arm sweep: run one CLI over the whole corpus tree, store every
surface's stdout (gzipped) and one JSON line per (file, surface).

  python3 sweep2.py BIN CORPUS_ROOT OUTDIR [--jobs N] [--list FILE]

Record: {path, fmt, surface, status(ok|err|timeout|crash), code, bytes, sha, ms, err}
Outputs land in OUTDIR/out/<relpath>.<surface>.gz so compare/diff scripts can
read them without re-running either arm.
"""
import gzip, hashlib, json, os, subprocess, sys, time
from concurrent.futures import ProcessPoolExecutor, as_completed

SURFACES = ["text", "markdown", "html", "ir"]
TIMEOUT = 120
SKIP_DIRS = {"metadata", "scripts", "test_outputs", "__pycache__", ".git", ".venv"}

def run_one(args):
    bin_, root, out_root, rel = args
    path = os.path.join(root, rel)
    fmt = rel.split("/", 1)[0]
    recs = []
    for surface in SURFACES:
        rec = {"path": rel, "fmt": fmt, "surface": surface}
        t0 = time.perf_counter()
        try:
            p = subprocess.run([bin_, surface, path], capture_output=True, timeout=TIMEOUT)
            rec["ms"] = round((time.perf_counter() - t0) * 1000, 1)
            rec["code"] = p.returncode
            if p.returncode == 0:
                rec["status"] = "ok"
                rec["bytes"] = len(p.stdout)
                rec["sha"] = hashlib.sha256(p.stdout).hexdigest()[:16]
                op = os.path.join(out_root, rel + "." + surface + ".gz")
                os.makedirs(os.path.dirname(op), exist_ok=True)
                with gzip.open(op, "wb", compresslevel=1) as f:
                    f.write(p.stdout)
            else:
                rec["status"] = "crash" if p.returncode < 0 else "err"
                rec["err"] = p.stderr.decode("utf-8", "replace").strip()[:400]
        except subprocess.TimeoutExpired:
            rec["ms"] = TIMEOUT * 1000
            rec["status"] = "timeout"
        except Exception as e:
            rec["status"] = "crash"; rec["err"] = str(e)[:400]
        recs.append(rec)
    return recs

def main():
    bin_, root, outdir = sys.argv[1:4]
    jobs = 7
    listfile = None
    a = sys.argv[4:]
    while a:
        k = a.pop(0)
        if k == "--jobs": jobs = int(a.pop(0))
        elif k == "--list": listfile = a.pop(0)
    if not os.access(bin_, os.X_OK):
        sys.exit(f"binary not executable: {bin_}")
    if listfile:
        files = [l.strip() for l in open(listfile) if l.strip()]
    else:
        files = []
        for d, dirs, fs in os.walk(root):
            dirs[:] = sorted(x for x in dirs if x not in SKIP_DIRS)
            for f in sorted(fs):
                rel = os.path.relpath(os.path.join(d, f), root)
                if rel.split("/", 1)[0] in SKIP_DIRS: continue
                files.append(rel)
    out_root = os.path.join(outdir, "out")
    os.makedirs(out_root, exist_ok=True)
    jl = open(os.path.join(outdir, "sweep.jsonl"), "w")
    done = 0
    with ProcessPoolExecutor(jobs) as ex:
        futs = [ex.submit(run_one, (bin_, root, out_root, rel)) for rel in files]
        for fut in as_completed(futs):
            for rec in fut.result():
                jl.write(json.dumps(rec) + "\n")
            done += 1
            if done % 250 == 0:
                jl.flush(); print(f"  {done}/{len(files)}", file=sys.stderr, flush=True)
    jl.close()
    print(f"swept {len(files)} files x {len(SURFACES)} surfaces -> {outdir}", file=sys.stderr)

if __name__ == "__main__":
    main()
