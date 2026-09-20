"""Files office_oxide refuses (status err/crash/timeout on `text`) that at
least one reference reader opens and extracts real text from.

  python3 unparsed_vs_peers.py SWEEP_DIR PANEL_DIR [--min-chars 40]

Groups by our error message so one root cause shows as one line, and
lists, per group, the peers that succeeded and how much they got. An
encrypted file every peer also fails on is not listed; a file only pandoc
reads is (pandoc is lenient about broken packages, which is a lead in
itself).
"""
import collections, json, os, sys

def main():
    sweep, panel = sys.argv[1:3]
    min_chars = 40
    if "--min-chars" in sys.argv: min_chars = int(sys.argv[sys.argv.index("--min-chars") + 1])
    ours = {}
    for line in open(os.path.join(sweep, "sweep.jsonl")):
        r = json.loads(line)
        if r["surface"] == "text": ours[r["path"]] = r
    peers = collections.defaultdict(dict)
    for line in open(os.path.join(panel, "panel.jsonl")):
        r = json.loads(line)
        if r.get("status") == "ok" and r.get("chars", 0) >= min_chars:
            peers[r["path"]][r["tool"]] = r["chars"]
    groups = collections.defaultdict(list)
    for path, r in ours.items():
        if r["status"] == "ok": continue
        ok_peers = peers.get(path)
        if not ok_peers: continue
        key = (r["status"], (r.get("err") or "").split("\n")[0][:90])
        groups[key].append((path, ok_peers))
    total = sum(len(v) for v in groups.values())
    print(f"{total} files we do not read that a reference reads (>= {min_chars} chars)\n")
    for key, items in sorted(groups.items(), key=lambda kv: -len(kv[1])):
        status, err = key
        print(f"== {len(items):4d}  {status}: {err}")
        tools = collections.Counter(t for _, p in items for t in p)
        print("      peers: " + ", ".join(f"{t} {n}" for t, n in tools.most_common()))
        for path, p in sorted(items, key=lambda x: -max(x[1].values()))[:6]:
            print(f"      {path}  " + " ".join(f"{t}={n}" for t, n in sorted(p.items())))
        print()

if __name__ == "__main__":
    main()
