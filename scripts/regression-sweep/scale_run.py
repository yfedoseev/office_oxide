"""Scaling probe: generate each gen_scaling.py dimension at doubling sizes,
time every surface, and report the log-log slope between successive sizes.

  python3 scale_run.py BIN OUTDIR [--dims docx.para,xlsx.rows] [--max 32000] [--surfaces text,markdown,html,ir]

Slope 1.0 is linear, 2.0 quadratic. Small N is startup-dominated, so the
verdict column is the slope over the two largest sizes that completed.
"""
import json, os, subprocess, sys, time, math
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gen_scaling

def measure(bin_, surface, path, timeout=180):
    t0 = time.perf_counter()
    p = subprocess.Popen([bin_, surface, path], stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
    n = 0; killed = False
    while True:
        chunk = p.stdout.read(1 << 20)
        if not chunk: break
        n += len(chunk)
        if time.perf_counter() - t0 > timeout:
            p.kill(); killed = True; break
    _, status, ru = os.wait4(p.pid, 0)
    p.returncode = status
    return {"wall_ms": (time.perf_counter() - t0) * 1000, "cpu_ms": (ru.ru_utime + ru.ru_stime) * 1000,
            "rss_mb": ru.ru_maxrss / 1024, "out_bytes": n, "ok": (not killed) and os.WIFEXITED(status) and os.WEXITSTATUS(status) == 0,
            "timeout": killed}

def main():
    bin_, outdir = sys.argv[1:3]
    dims = None; nmax = 32000; surfaces = ["text", "markdown", "html", "ir"]
    a = sys.argv[3:]
    while a:
        k = a.pop(0)
        if k == "--dims": dims = a.pop(0).split(",")
        elif k == "--max": nmax = int(a.pop(0))
        elif k == "--surfaces": surfaces = a.pop(0).split(",")
    if dims is None:
        dims = [f"{f}.{d}" for f, ds in gen_scaling.DIMS.items() for d in ds]
    # Dimensions that create one zip part per element are generation-bound; cap them.
    part_heavy = {"docx.images", "docx.sections", "xlsx.sheets", "pptx.slides", "pptx.notes", "docx.hyperlinks", "xlsx.hyperlinks"}
    os.makedirs(outdir, exist_ok=True)
    jl = open(os.path.join(outdir, "scale.jsonl"), "a")
    print(f"{'dim':18} {'surface':9} {'N':>7} {'cpu_ms':>9} {'rss_mb':>8} {'out_kb':>8} {'slope':>6}")
    for dim in dims:
        fmt, d = dim.split(".")
        cap = min(nmax, 8000) if dim in part_heavy else nmax
        prev = {}
        n = 1000
        stop = set()
        while n <= cap:
            path = gen_scaling.generate(os.path.join(outdir, "files"), fmt, d, n)
            fsize = os.path.getsize(path)
            for s in surfaces:
                if s in stop: continue
                m = measure(bin_, s, path)
                slope = None
                if s in prev and prev[s]["cpu_ms"] > 20:
                    slope = math.log(max(m["cpu_ms"], 1) / prev[s]["cpu_ms"]) / math.log(2)
                rec = {"dim": dim, "surface": s, "n": n, "file_bytes": fsize, **m, "slope": slope}
                jl.write(json.dumps(rec) + "\n"); jl.flush()
                flag = "" if m["ok"] else ("  TIMEOUT" if m["timeout"] else "  ERR")
                print(f"{dim:18} {s:9} {n:7d} {m['cpu_ms']:9.0f} {m['rss_mb']:8.0f} {m['out_bytes']/1024:8.0f} {('%.2f' % slope) if slope is not None else '':>6}{flag}", flush=True)
                prev[s] = m
                if not m["ok"] or m["cpu_ms"] > 60000:
                    stop.add(s)
            os.remove(path)
            if len(stop) == len(surfaces): break
            n *= 2
    jl.close()

if __name__ == "__main__":
    main()
