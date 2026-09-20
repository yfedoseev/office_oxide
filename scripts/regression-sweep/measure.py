"""Wall time and peak RSS of one CLI call, measured from the parent via RUSAGE_CHILDREN.

  python3 measure.py BIN SURFACE FILE

Use this, serially and unloaded, before calling a crash or timeout seen under a
parallel sweep a regression: OOM kills there depend on what else was running.
"""
import resource, subprocess, sys, time
bin_, surf, path = sys.argv[1:4]
t0=time.perf_counter()
p = subprocess.run([bin_, surf, path], stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=600)
dt=time.perf_counter()-t0
ru=resource.getrusage(resource.RUSAGE_CHILDREN)
print(f"{dt:7.1f}s {ru.ru_maxrss/1024:7.0f}MB rc={p.returncode} {len(p.stdout):>10} bytes")
