#!/usr/bin/env python3
"""Benchmark the command-line extractors (catdoc, antiword, xls2csv, catppt)
against the test corpus.

Usage: python3 bench_native.py ~/projects/office_oxide_tests/ [--json OUT]

Each file is one subprocess run, so the times include process start-up, as
they do for anyone who shells out to these tools. A run passes when the tool
exits 0 within the timeout.
"""

import json
import os
import subprocess
import sys
import time

TIMEOUT = 30  # seconds per file

TOOLS = {
    ".doc": [("catdoc", ["catdoc"]), ("antiword", ["antiword"])],
    ".xls": [("xls2csv", ["xls2csv"])],
    ".ppt": [("catppt", ["catppt"])],
}


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: bench_native.py DIR [--json OUT]", file=sys.stderr)
        sys.exit(1)
    root = sys.argv[1]
    json_out = sys.argv[sys.argv.index("--json") + 1] if "--json" in sys.argv else None
    stats: dict[str, dict] = {}
    for dirpath, _, names in os.walk(root):
        for name in sorted(names):
            ext = os.path.splitext(name)[1].lower()
            path = os.path.join(dirpath, name)
            for tool, argv in TOOLS.get(ext, []):
                s = stats.setdefault(f"{tool} {ext}", {"ok": 0, "fail": 0, "total_ms": 0.0})
                t0 = time.monotonic()
                try:
                    r = subprocess.run(
                        argv + [path], capture_output=True, timeout=TIMEOUT, check=False
                    )
                    good = r.returncode == 0
                except subprocess.TimeoutExpired:
                    good = False
                s["total_ms"] += (time.monotonic() - t0) * 1000
                s["ok" if good else "fail"] += 1
    for key, s in sorted(stats.items()):
        total = s["ok"] + s["fail"]
        print(f"{key}: total {total} ok {s['ok']} fail {s['fail']} "
              f"rate {s['ok'] / total * 100:.1f}% mean {s['total_ms'] / total:.1f}ms")
    if json_out:
        with open(json_out, "w") as f:
            json.dump({"libs": stats}, f, indent=2, sort_keys=True)


if __name__ == "__main__":
    main()
