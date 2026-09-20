"""Join perf_sweep.py (text surface) with refpanel.py timings: per format and
tool, the median and p90 of office_oxide cpu / peer cpu on files both handled,
plus the files where office_oxide is slowest relative to the peer.

  python3 perf_vs_peers.py NEXT.jsonl PANEL.jsonl [--min-ms 50] [--top 10]

Python peers pay import cost once per worker, not per file, so cpu_ms is
comparable; C tools (catdoc, antiword, xls2csv, catppt) include exec.
"""
import json, sys, collections, statistics

def main():
    nxt, panel = sys.argv[1:3]
    min_ms = 50; top = 10
    if "--min-ms" in sys.argv: min_ms = float(sys.argv[sys.argv.index("--min-ms") + 1])
    if "--top" in sys.argv: top = int(sys.argv[sys.argv.index("--top") + 1])
    oo = {}
    for line in open(nxt):
        r = json.loads(line)
        if r["surface"] == "text" and r["status"] == "ok": oo[r["path"]] = r
    peers = collections.defaultdict(dict)
    for line in open(panel):
        r = json.loads(line)
        if r.get("status") == "ok" and "cpu_ms" in r: peers[r["tool"]][r["path"]] = r
    for tool in sorted(peers):
        rows = []
        for path, pr in peers[tool].items():
            r = oo.get(path)
            if not r: continue
            rows.append((r["cpu_ms"], pr["cpu_ms"], path, r["file_bytes"]))
        if not rows: continue
        big = [x for x in rows if max(x[0], x[1]) >= min_ms]
        ratios = sorted(a / max(b, 0.1) for a, b, _, _ in big) or [0]
        tot_oo = sum(a for a, _, _, _ in rows); tot_pr = sum(b for _, b, _, _ in rows)
        print(f"\n== {tool}: {len(rows)} files, {len(big)} with max cpu >= {min_ms} ms ==")
        print(f"  total cpu office_oxide {tot_oo/1000:8.1f} s  vs {tool} {tot_pr/1000:8.1f} s  -> ratio {tot_oo/max(tot_pr,1):.2f}x")
        print(f"  oo/peer ratio on big files: p10 {ratios[int(len(ratios)*0.1)]:.2f}  median {ratios[len(ratios)//2]:.2f}  p90 {ratios[int(len(ratios)*0.9)]:.2f}  max {ratios[-1]:.2f}")
        print(f"  files where office_oxide is slowest relative to {tool}:")
        for a, b, path, fb in sorted(big, key=lambda x: -(x[0] / max(x[1], 0.1)))[:top]:
            print(f"    {a/max(b,0.1):6.1f}x  oo {a:8.0f} ms  {tool} {b:8.0f} ms  {fb/1e6:6.2f} MB  {path}")

if __name__ == "__main__":
    main()
