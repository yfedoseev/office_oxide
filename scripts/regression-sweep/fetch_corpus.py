"""Fetch a real Office corpus from public test-data repos.

Sources are the ones README.md already names, all permissively licensed.
Files are downloaded flat into corpus/ with a source prefix so provenance
survives.
"""
import json, os, sys, urllib.request, urllib.error, time

SC = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(SC, "corpus")
os.makedirs(OUT, exist_ok=True)

EXTS = (".docx", ".xlsx", ".pptx", ".doc", ".xls", ".ppt",
        ".docm", ".xlsm", ".pptm", ".dotx", ".xltx", ".potx")

# (label, api url for a directory listing)
SOURCES = [
    ("poi-doc",   "https://api.github.com/repos/apache/poi/contents/test-data/document"),
    ("poi-sheet", "https://api.github.com/repos/apache/poi/contents/test-data/spreadsheet"),
    ("poi-slide", "https://api.github.com/repos/apache/poi/contents/test-data/slideshow"),
    ("poi-int",   "https://api.github.com/repos/apache/poi/contents/test-data/integration"),
    ("pandoc",    "https://api.github.com/repos/jgm/pandoc/contents/test/docx"),
]

def get_json(url, tries=3):
    tok = os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN")
    for i in range(tries):
        req = urllib.request.Request(url, headers={
            "User-Agent": "office-oxide-regression",
            "Accept": "application/vnd.github+json",
            **({"Authorization": f"Bearer {tok}"} if tok else {}),
        })
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            if e.code in (403, 429) and i + 1 < tries:
                time.sleep(5 * (i + 1)); continue
            print(f"  ! {url}: HTTP {e.code}", file=sys.stderr); return []
        except Exception as e:
            print(f"  ! {url}: {e}", file=sys.stderr); return []
    return []

def fetch(url, dest):
    tok = os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN")
    req = urllib.request.Request(url, headers={
        "User-Agent": "office-oxide-regression",
        **({"Authorization": f"Bearer {tok}"} if tok else {}),
    })
    with urllib.request.urlopen(req, timeout=120) as r, open(dest, "wb") as f:
        f.write(r.read())

total = 0
for label, api in SOURCES:
    entries = get_json(api)
    n = 0
    for e in entries:
        if e.get("type") != "file":
            continue
        name = e["name"]
        if not name.lower().endswith(EXTS):
            continue
        # Skip enormous files; the sweep wants breadth, not size.
        if e.get("size", 0) > 12 * 1024 * 1024:
            continue
        dest = os.path.join(OUT, f"{label}__{name}")
        if os.path.exists(dest):
            n += 1; continue
        try:
            fetch(e["download_url"], dest)
            n += 1
        except Exception as ex:
            print(f"  ! {name}: {ex}", file=sys.stderr)
    print(f"{label}: {n} files")
    total += n
print(f"TOTAL {total}")
