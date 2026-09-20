"""Reference panel: extract plain text from every corpus document with
independent libraries, one output file per (tool, document).

  python3 refpanel.py CORPUS_ROOT OUT_ROOT [--jobs N] [--list FILE]

Tools by format (each writes OUT_ROOT/<tool>/<relpath>.txt, or .err):
  docx: python-docx (body paragraphs + tables incl. nested, header/footer, deduped merged cells), pandoc
  xlsx: openpyxl (every cell value, formulas as text), python-calamine
  pptx: python-pptx (every text frame, table cell, notes)
  doc:  catdoc, antiword
  xls:  xlrd, python-calamine, xls2csv
  ppt:  catppt
"""
import os, sys, json, subprocess, signal, traceback
from concurrent.futures import ProcessPoolExecutor, as_completed

TIMEOUT = 120
SKIP_DIRS = {"metadata", "scripts", "test_outputs", "__pycache__", ".git", ".venv"}
FMT_DIRS = {"doc", "docx", "ppt", "pptx", "xls", "xlsx"}

class TO(Exception): pass
def _alarm(*_): raise TO()

# ---------- python-docx
def t_python_docx(path):
    import docx
    d = docx.Document(path)
    out = []
    def tables(tbls):
        for t in tbls:
            for row in t.rows:
                seen = set()
                for cell in row.cells:
                    if id(cell) in seen: continue
                    seen.add(id(cell))
                    out.append(cell.text)
                    tables(cell.tables)
    for p in d.paragraphs: out.append(p.text)
    tables(d.tables)
    for s in d.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer, s.even_page_header, s.even_page_footer):
            try:
                if part.is_linked_to_previous: continue
                for p in part.paragraphs: out.append(p.text)
                tables(part.tables)
            except Exception: pass
    return "\n".join(out)

def t_pandoc(path):
    p = subprocess.run(["pandoc", "-t", "plain", "--wrap=none", path], capture_output=True, timeout=TIMEOUT)
    if p.returncode != 0: raise RuntimeError(p.stderr.decode("utf-8","replace")[:300])
    return p.stdout.decode("utf-8", "replace")

# ---------- openpyxl
def t_openpyxl(path):
    import openpyxl, warnings
    warnings.simplefilter("ignore")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    out = []
    for ws in wb.worksheets:
        out.append(f"## {ws.title}")
        for row in ws.iter_rows(values_only=True):
            vals = [str(v) for v in row if v is not None and str(v) != ""]
            if vals: out.append("\t".join(vals))
    return "\n".join(out)

def t_calamine(path):
    from python_calamine import CalamineWorkbook
    wb = CalamineWorkbook.from_path(path)
    out = []
    for name in wb.sheet_names:
        out.append(f"## {name}")
        for row in wb.get_sheet_by_name(name).to_python(skip_empty_area=True):
            vals = [str(v) for v in row if v is not None and str(v) != ""]
            if vals: out.append("\t".join(vals))
    return "\n".join(out)

# ---------- python-pptx
def t_python_pptx(path):
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    prs = Presentation(path)
    out = []
    def shapes(shs):
        for sh in shs:
            if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                shapes(sh.shapes); continue
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    out.append("".join(r.text for r in p.runs))
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells: out.append(c.text)
    for i, slide in enumerate(prs.slides, 1):
        out.append(f"## slide {i}")
        shapes(slide.shapes)
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
            out.append("[notes] " + slide.notes_slide.notes_text_frame.text)
    return "\n".join(out)

# ---------- xlrd
def t_xlrd(path):
    import xlrd
    wb = xlrd.open_workbook(path, on_demand=True)
    out = []
    for si in range(wb.nsheets):
        ws = wb.sheet_by_index(si)
        out.append(f"## {ws.name}")
        for r in range(ws.nrows):
            vals = [str(v) for v in ws.row_values(r) if v not in ("", None)]
            if vals: out.append("\t".join(vals))
    return "\n".join(out)

def cli(cmd):
    def run(path):
        p = subprocess.run(cmd + [path], capture_output=True, timeout=TIMEOUT)
        if p.returncode != 0 and not p.stdout: raise RuntimeError(p.stderr.decode("utf-8","replace")[:300])
        return p.stdout.decode("utf-8", "replace")
    return run

TOOLS = {
    "docx": [("python_docx", t_python_docx), ("pandoc", t_pandoc)],
    "xlsx": [("openpyxl", t_openpyxl), ("calamine", t_calamine)],
    "pptx": [("python_pptx", t_python_pptx)],
    "doc":  [("catdoc", cli(["catdoc", "-w"])), ("antiword", cli(["antiword", "-w", "0"]))],
    "xls":  [("xlrd", t_xlrd), ("calamine", t_calamine), ("xls2csv", cli(["xls2csv"]))],
    "ppt":  [("catppt", cli(["catppt"]))],
}

def run_one(args):
    root, out_root, rel = args
    fmt = rel.split("/", 1)[0]
    path = os.path.join(root, rel)
    res = []
    for tool, fn in TOOLS.get(fmt, []):
        op = os.path.join(out_root, tool, rel)
        os.makedirs(os.path.dirname(op), exist_ok=True)
        signal.signal(signal.SIGALRM, _alarm); signal.alarm(TIMEOUT)
        try:
            text = fn(path)
            signal.alarm(0)
            open(op + ".txt", "w", encoding="utf-8", errors="replace").write(text)
            res.append({"path": rel, "tool": tool, "status": "ok", "chars": len(text)})
        except TO:
            res.append({"path": rel, "tool": tool, "status": "timeout"})
        except BaseException as e:
            signal.alarm(0)
            msg = f"{type(e).__name__}: {str(e)[:200]}"
            open(op + ".err", "w").write(msg)
            res.append({"path": rel, "tool": tool, "status": "err", "err": msg})
    return res

def main():
    root, out_root = sys.argv[1:3]
    jobs = 6; listfile = None; a = sys.argv[3:]
    while a:
        k = a.pop(0)
        if k == "--jobs": jobs = int(a.pop(0))
        elif k == "--list": listfile = a.pop(0)
    if listfile:
        files = [l.strip() for l in open(listfile) if l.strip()]
    else:
        files = []
        for d, dirs, fs in os.walk(root):
            dirs[:] = sorted(x for x in dirs if x not in SKIP_DIRS)
            for f in sorted(fs):
                rel = os.path.relpath(os.path.join(d, f), root)
                if rel.split("/", 1)[0] in FMT_DIRS: files.append(rel)
    os.makedirs(out_root, exist_ok=True)
    jl = open(os.path.join(out_root, "panel.jsonl"), "w")
    done = 0
    with ProcessPoolExecutor(jobs) as ex:
        futs = [ex.submit(run_one, (root, out_root, rel)) for rel in files]
        for fut in as_completed(futs):
            try:
                for r in fut.result(): jl.write(json.dumps(r) + "\n")
            except Exception as e:
                jl.write(json.dumps({"status": "crash", "err": str(e)[:200]}) + "\n")
            done += 1
            if done % 250 == 0: jl.flush(); print(f"  {done}/{len(files)}", file=sys.stderr, flush=True)
    jl.close()
    print(f"panel over {len(files)} files -> {out_root}", file=sys.stderr)

if __name__ == "__main__":
    main()
