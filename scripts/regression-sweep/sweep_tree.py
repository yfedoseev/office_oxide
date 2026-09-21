"""Parallel two-arm sweep: run one CLI over the whole corpus tree, store every
surface's stdout (gzipped) and one JSON line per (file, surface).

  python3 sweep_tree.py BIN CORPUS_ROOT OUTDIR [--jobs N] [--list FILE] [--surfaces text,markdown,html,ir]

Record: {path, fmt, surface, status(ok|err|timeout|crash), code, bytes, sha, ms, err}
Outputs land in OUTDIR/out/<relpath>.<surface>.gz so compare/diff scripts can
read them without re-running either arm.
"""
import collections, gzip, hashlib, json, os, select, subprocess, sys, time
from concurrent.futures import ProcessPoolExecutor, as_completed

SURFACES = ["text", "markdown", "html", "ir"]
TIMEOUT = 120
SKIP_DIRS = {"metadata", "scripts", "test_outputs", "__pycache__", ".git", ".venv"}

def run_one(args):
    bin_, root, out_root, rel, surfaces = args
    path = os.path.join(root, rel)
    fmt = rel.split("/", 1)[0]
    recs = []
    for surface in surfaces:
        rec = {"path": rel, "fmt": fmt, "surface": surface}
        t0 = time.perf_counter()
        op = os.path.join(out_root, rel + "." + surface + ".gz")
        os.makedirs(os.path.dirname(op), exist_ok=True)
        try:
            # stdout is streamed to the gzip file in chunks: an `ir` dump of a
            # large spreadsheet is hundreds of MB, and holding one per worker
            # in memory got the sweep itself OOM-killed.
            with subprocess.Popen([bin_, surface, path], stdout=subprocess.PIPE, stderr=subprocess.PIPE) as p, \
                 gzip.open(op, "wb", compresslevel=1) as f:
                h = hashlib.sha256(); n = 0
                fd = p.stdout.fileno()
                try:
                    while True:
                        # A child that hangs before writing anything must
                        # time out too, so wait on the pipe with a deadline.
                        remaining = TIMEOUT - (time.perf_counter() - t0)
                        if remaining <= 0 or not select.select([fd], [], [], remaining)[0]:
                            raise subprocess.TimeoutExpired(bin_, TIMEOUT)
                        chunk = os.read(fd, 1 << 20)
                        if not chunk: break
                        h.update(chunk); f.write(chunk); n += len(chunk)
                    err = p.stderr.read()
                    p.wait(timeout=max(1, TIMEOUT - (time.perf_counter() - t0)))
                except subprocess.TimeoutExpired:
                    p.kill(); p.wait(); raise
            rec["ms"] = round((time.perf_counter() - t0) * 1000, 1)
            rec["code"] = p.returncode
            if p.returncode == 0:
                rec["status"] = "ok"; rec["bytes"] = n; rec["sha"] = h.hexdigest()[:16]
            else:
                rec["status"] = "crash" if p.returncode < 0 else "err"
                rec["err"] = err.decode("utf-8", "replace").strip()[:400]
                os.remove(op)
        except subprocess.TimeoutExpired:
            rec["ms"] = TIMEOUT * 1000; rec["status"] = "timeout"
            if os.path.exists(op): os.remove(op)
        except Exception as e:
            rec["status"] = "crash"; rec["err"] = str(e)[:400]
        recs.append(rec)
    return recs

def main():
    bin_, root, outdir = sys.argv[1:4]
    jobs = 7
    listfile = None
    surfaces = SURFACES
    a = sys.argv[4:]
    while a:
        k = a.pop(0)
        if k == "--jobs": jobs = int(a.pop(0))
        elif k == "--list": listfile = a.pop(0)
        elif k == "--surfaces": surfaces = a.pop(0).split(",")
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
    # Resumable: a file with all its surfaces already recorded is skipped.
    jl_path = os.path.join(outdir, "sweep.jsonl")
    done_files = collections.Counter()
    if os.path.exists(jl_path):
        with open(jl_path) as f:
            for line in f:
                try: done_files[json.loads(line)["path"]] += 1
                except Exception: pass
    files = [f for f in files if done_files.get(f, 0) < len(surfaces)]
    jl = open(jl_path, "a")
    done = 0
    with ProcessPoolExecutor(jobs) as ex:
        futs = [ex.submit(run_one, (bin_, root, out_root, rel, surfaces)) for rel in files]
        for fut in as_completed(futs):
            for rec in fut.result():
                jl.write(json.dumps(rec) + "\n")
            done += 1
            if done % 250 == 0:
                jl.flush(); print(f"  {done}/{len(files)}", file=sys.stderr, flush=True)
    jl.close()
    print(f"swept {len(files)} files x {len(surfaces)} surfaces -> {outdir}", file=sys.stderr)

if __name__ == "__main__":
    main()
