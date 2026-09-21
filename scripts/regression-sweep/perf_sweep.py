"""Per-file performance sweep: wall, CPU (user+sys) and peak RSS of one CLI
surface call per (file, surface), measured for that child alone via os.wait4,
so workers running in parallel do not pollute each other's numbers.

  python3 perf_sweep.py BIN CORPUS_ROOT OUT.jsonl [--jobs N] [--list FILE] [--surfaces text,markdown,html,ir]

Record: {path, fmt, surface, status, code, file_bytes, out_bytes, wall_ms, cpu_ms, rss_mb, err}

CPU time is the number to rank by — wall time under a parallel sweep is load;
peak RSS is per child (ru_maxrss of exactly that pid). Re-measure any
candidate serially with measure.py before filing it.
"""
import json, os, select, subprocess, sys, time
from concurrent.futures import ProcessPoolExecutor, as_completed

SKIP_DIRS = {"metadata", "scripts", "test_outputs", "__pycache__", ".git", ".venv"}
TIMEOUT = 180

def run_one(args):
    bin_, root, rel, surfaces = args
    path = os.path.join(root, rel)
    fmt = rel.split("/", 1)[0]
    fbytes = os.path.getsize(path)
    recs = []
    for surface in surfaces:
        rec = {"path": rel, "fmt": fmt, "surface": surface, "file_bytes": fbytes}
        t0 = time.perf_counter()
        p = subprocess.Popen([bin_, surface, path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        n = 0; err = b""; timed_out = False
        try:
            # Drain stdout so a large `ir` dump cannot block the child, and
            # wait with a deadline so a child that hangs *before* writing
            # anything (v0.1.11's unbounded date loop) is killed too.
            fd = p.stdout.fileno()
            while True:
                remaining = TIMEOUT - (time.perf_counter() - t0)
                if remaining <= 0 or not select.select([fd], [], [], remaining)[0]:
                    timed_out = True; break
                chunk = os.read(fd, 1 << 20)
                if not chunk: break
                n += len(chunk)
            if not timed_out:
                err = p.stderr.read()
        except Exception:
            pass
        if timed_out:
            p.kill()
        _, status, ru = os.wait4(p.pid, 0)
        p.returncode = status  # keep Popen from waiting again
        wall = (time.perf_counter() - t0) * 1000
        rec["wall_ms"] = round(wall, 1)
        rec["cpu_ms"] = round((ru.ru_utime + ru.ru_stime) * 1000, 1)
        rec["rss_mb"] = round(ru.ru_maxrss / 1024, 1)
        rec["out_bytes"] = n
        if timed_out:
            rec["status"] = "timeout"; rec["code"] = -1
        elif os.WIFSIGNALED(status):
            rec["status"] = "crash"; rec["code"] = -os.WTERMSIG(status)
            rec["err"] = err.decode("utf-8", "replace").strip()[:300]
        elif os.WEXITSTATUS(status) != 0:
            rec["status"] = "err"; rec["code"] = os.WEXITSTATUS(status)
            rec["err"] = err.decode("utf-8", "replace").strip()[:300]
        else:
            rec["status"] = "ok"; rec["code"] = 0
        recs.append(rec)
    return recs

def main():
    bin_, root, outfile = sys.argv[1:4]
    jobs = 4; listfile = None; surfaces = ["text", "markdown", "html", "ir"]
    a = sys.argv[4:]
    while a:
        k = a.pop(0)
        if k == "--jobs": jobs = int(a.pop(0))
        elif k == "--list": listfile = a.pop(0)
        elif k == "--surfaces": surfaces = a.pop(0).split(",")
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
    done_paths = set()
    if os.path.exists(outfile):
        with open(outfile) as f:
            for line in f:
                try: done_paths.add(json.loads(line)["path"])
                except Exception: pass
    files = [f for f in files if f not in done_paths]
    out = open(outfile, "a")
    done = 0
    with ProcessPoolExecutor(jobs) as ex:
        futs = [ex.submit(run_one, (bin_, root, rel, surfaces)) for rel in files]
        for fut in as_completed(futs):
            for rec in fut.result():
                out.write(json.dumps(rec) + "\n")
            done += 1
            if done % 250 == 0:
                out.flush(); print(f"  {done}/{len(files)}", file=sys.stderr, flush=True)
    out.close()
    print(f"profiled {len(files)} files x {surfaces} -> {outfile}", file=sys.stderr)

if __name__ == "__main__":
    main()
