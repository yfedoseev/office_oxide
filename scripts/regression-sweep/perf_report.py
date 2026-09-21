"""Rank perf_sweep.py records: slowest, worst throughput, worst memory amplification.

  python3 perf_report.py NEXT.jsonl [PREV.jsonl] [--top N]

Throughput is cpu_ms per MB of *input*; memory amplification is peak RSS
over file size. Both are only meaningful above a size floor (tiny files are
process startup), so the floor is 64 KB. With PREV, also lists files whose
cpu or rss regressed >= 2x and >= 200 ms / 50 MB.
"""
import json, sys, collections, statistics

def load(p):
    d = {}
    for line in open(p):
        try: r = json.loads(line)
        except Exception: continue
        d[(r["path"], r["surface"])] = r
    return d

def main():
    nxt = load(sys.argv[1]); prev = load(sys.argv[2]) if len(sys.argv) > 2 and not sys.argv[2].startswith("--") else {}
    top = 15
    if "--top" in sys.argv: top = int(sys.argv[sys.argv.index("--top") + 1])
    recs = list(nxt.values())
    print(f"{len(recs)} records, {len({r['path'] for r in recs})} files")
    st = collections.Counter((r["fmt"], r["surface"], r["status"]) for r in recs)
    bad = [(k, v) for k, v in st.items() if k[2] != "ok"]
    print("\n== non-ok by fmt/surface ==")
    for k, v in sorted(bad): print(f"  {k[0]:5} {k[1]:9} {k[2]:8} {v}")
    print("\n== per fmt/surface: n, median cpu, p95 cpu, max cpu, median MB/s, p5 MB/s ==")
    by = collections.defaultdict(list)
    for r in recs:
        if r["status"] == "ok": by[(r["fmt"], r["surface"])].append(r)
    for k in sorted(by):
        rs = by[k]; cpu = sorted(r["cpu_ms"] for r in rs)
        big = [r for r in rs if r["file_bytes"] >= 65536]
        thr = sorted((r["file_bytes"] / 1e6) / max(r["cpu_ms"], 1) * 1000 for r in big) if big else [0]
        print(f"  {k[0]:5} {k[1]:9} n={len(rs):5d} cpu med={cpu[len(cpu)//2]:7.1f} p95={cpu[int(len(cpu)*0.95)]:8.1f} max={cpu[-1]:9.1f}  MB/s med={thr[len(thr)//2]:6.1f} p5={thr[int(len(thr)*0.05)]:6.1f} (n>=64K: {len(big)})")
    for surface in ["text", "ir"]:
        print(f"\n== slowest {top} ({surface}) ==")
        for r in sorted((r for r in recs if r["surface"] == surface and r["status"] == "ok"), key=lambda r: -r["cpu_ms"])[:top]:
            print(f"  {r['cpu_ms']:9.0f} ms {r['rss_mb']:7.0f} MB {r['file_bytes']/1e6:7.2f} MB in {r['out_bytes']/1e6:7.2f} MB out  {r['path']}")
        print(f"\n== worst throughput {top} ({surface}, files >= 64 KB), ms per MB of input ==")
        for r in sorted((r for r in recs if r["surface"] == surface and r["status"] == "ok" and r["file_bytes"] >= 65536), key=lambda r: -r["cpu_ms"] / r["file_bytes"])[:top]:
            print(f"  {r['cpu_ms']/(r['file_bytes']/1e6):9.0f} ms/MB {r['cpu_ms']:8.0f} ms {r['file_bytes']/1e6:7.2f} MB in {r['out_bytes']/1e6:7.2f} MB out  {r['path']}")
        print(f"\n== worst memory amplification {top} ({surface}, files >= 64 KB), RSS / file size ==")
        for r in sorted((r for r in recs if r["surface"] == surface and r["status"] == "ok" and r["file_bytes"] >= 65536), key=lambda r: -r["rss_mb"] * 1e6 / r["file_bytes"])[:top]:
            print(f"  {r['rss_mb']*1e6/r['file_bytes']:8.0f}x {r['rss_mb']:7.0f} MB {r['file_bytes']/1e6:7.2f} MB in {r['out_bytes']/1e6:7.2f} MB out  {r['path']}")
    print(f"\n== timeouts / crashes / errors (next) ==")
    for r in recs:
        if r["status"] in ("timeout", "crash"): print(f"  {r['status']:8} {r['surface']:9} {r['path']}  {r.get('err','')[:100]}")
    if prev:
        print(f"\n== regressions vs prev (cpu >= 2x and >= 200 ms, or rss >= 2x and >= 50 MB) ==")
        rows = []
        for k, r in nxt.items():
            p = prev.get(k)
            if not p or r["status"] != "ok" or p["status"] != "ok": continue
            if (r["cpu_ms"] >= 2 * p["cpu_ms"] and r["cpu_ms"] >= 200) or (r["rss_mb"] >= 2 * p["rss_mb"] and r["rss_mb"] >= 50):
                rows.append((r["cpu_ms"] / max(p["cpu_ms"], 1), p, r))
        for ratio, p, r in sorted(rows, key=lambda x: -x[0])[:top * 2]:
            print(f"  cpu {p['cpu_ms']:8.0f} -> {r['cpu_ms']:8.0f} ms  rss {p['rss_mb']:6.0f} -> {r['rss_mb']:6.0f} MB  {r['surface']:9} {r['path']}")
        print(f"\n== improvements vs prev (cpu <= 0.5x and prev >= 500 ms) ==")
        rows = []
        for k, r in nxt.items():
            p = prev.get(k)
            if not p or r["status"] != "ok" or p["status"] != "ok": continue
            if r["cpu_ms"] <= 0.5 * p["cpu_ms"] and p["cpu_ms"] >= 500: rows.append((p, r))
        for p, r in sorted(rows, key=lambda x: -x[0]["cpu_ms"])[:top]:
            print(f"  cpu {p['cpu_ms']:8.0f} -> {r['cpu_ms']:8.0f} ms  rss {p['rss_mb']:6.0f} -> {r['rss_mb']:6.0f} MB  {r['surface']:9} {r['path']}")
        # status transitions
        print(f"\n== status transitions prev -> next ==")
        tr = collections.Counter()
        for k, r in nxt.items():
            p = prev.get(k)
            if p and p["status"] != r["status"]: tr[(p["status"], r["status"], r["surface"])] += 1
        for k, v in sorted(tr.items()): print(f"  {k[0]:8} -> {k[1]:8} {k[2]:9} {v}")

if __name__ == "__main__":
    main()
