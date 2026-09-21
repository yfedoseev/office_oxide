"""Release regression harness: two arms of office_oxide against a reference
panel over a corpus, with one verdict per file.

  python3 harness.py run    --prev-bin BIN --next-bin BIN --corpus DIR --work DIR
                            [--panel DIR] [--tika DIR] [--validator BIN] [--jobs N] [--reuse]
  python3 harness.py report --work DIR [--top N]

`run` produces (or reuses, with --reuse) every raw input under WORK:
  prev/, next/      sweep_tree.py output for each arm: text, markdown, html,
                    ir per file, with status and timing
  panel/, tika/     the external reference panel (refpanel.py, Tika); these
                    are per-corpus, not per-release, so they are normally
                    passed in and reused
  rt.jsonl          corpus_validate round-trip records for the next arm
  render/           render_compare.py (our markdown structure vs pandoc)
`report` reads those and writes WORK/findings.jsonl + WORK/findings.md.

The decision model — what "expected" means for a file nobody hand-checked:

* Text surface. Every reference's word multiset is compared with every
  other's. References that agree with each other (mutual recall >= 0.85,
  >= 20 words) form the CONSENSUS; the expected output is the consensus
  member closest to the median size. A single reference is a WEAK oracle
  (one opinion; catdoc/antiword/Tika each have their own blind spots), so
  it can only corroborate a two-arm difference, never convict on its own.
* A file is judged on four independent signals, each producing a class:
    REGRESSION        next lost words prev had, unexplained by re-tokenisation
                      (split/join), AND the consensus (or, without one, two
                      references) has them — unless next agrees with the
                      consensus at least as well as prev did.
    PEER_SPLIT        the two arms differ and each sides with a different
                      reference (next ≈ Tika, prev ≈ catdoc, say): a policy
                      difference such as field codes, tracked deletions or
                      hidden text, not a regression. Informational.
    REGRESSION_STATUS prev ok with output, next err/timeout.
    SHARED_DEFECT     both arms < 0.7 recall of the consensus (>= 2 refs).
    PARSE_GAP         we error; >= 1 reference extracts >= 20 words.
    EMPTY_VS_REFS     we succeed with nothing; the consensus has text.
    OVER_EXTRACTION   we emit > 3x every reference's words (>= 2 refs).
    SURFACE_SPLIT     our own text/markdown/html disagree (>= 5 words and
                      >= 2 %), after removing markdown/html syntax.
    ROUNDTRIP_LOSS    IR -> write -> reread loses >= 5 % of words (>= 10).
    STRUCTURE_GAP     pandoc renders headings/lists/tables we flatten
                      (docx only; render_compare.py verdict).
    PERF_REGRESSION   next >= 2x prev wall time and >= 500 ms, same sweep,
                      unless next extracted 1.5x more words (more work).
    IMPROVEMENT       next recall vs consensus >= prev + 0.1 (informational).
* Every finding carries a word sample so triage is a read, not a re-run,
  and an `explained` note when the harness can tell why on its own:
  DOCX words that sit in `w:vanish` runs (hidden text, which Word does not
  show and this crate hides by policy), field instruction vocabulary
  (HYPERLINK, PAGEREF, MERGEFORMAT …). Known deliberate differences it
  cannot see: tracked-change deletions (we hide them, catdoc/antiword
  show them), PowerPoint master placeholder prompts, number formats
  (`General` vs stored value), headers/footers (pandoc drops them).
* Surface comparison ignores what the surfaces legitimately add: URLs
  (markdown link targets), markdown's inline HTML (`<sub>`), image/chart
  placeholders, `## Slide N`/`## Sheet` headings, HTML attribute text.

Findings are grouped by (format, class) with the most severe first;
`findings.md` is the document to triage from.
"""
import argparse, collections, gzip, json, os, re, statistics, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
FMTS = ("doc", "docx", "ppt", "pptx", "xls", "xlsx")
WORD = re.compile(r"[^\W\d_]{3,}", re.UNICODE)
TAG = re.compile(r"<[^>]+>")
INLINE = re.compile(r"</?(span|b|i|u|a|em|strong|sup|sub|s|del|ins|code|font|mark)\b[^>]*>")
URL = re.compile(r"(?:https?|ftp|mailto|file):[^\s)\]>\"']+|www\.[^\s)\]>\"']+", re.I)
# Markdown's inline HTML is a real tag (`<sub>`), never a bare `<` in
# prose (`< 0.5 mg`): matching `<[^>]+>` ate half a pharmaceutical
# document between one `<` and the next `>`.
# A picture's alt text is markdown's own addition (the text surface does
# not print it), so `![alt](target)` goes entirely. An escaped `\<` is
# document text, not a tag.
MD_NOISE = re.compile(r"\[image[^\]]*\]|\[chart[^\]]*\]|!\[[^\]]*\]\([^)]*\)|\(#[^)]*\)|data:[^\s)]+"
                      r"|^#{1,6} (?:Slide|Sheet|Page|Notes|Comments|Charts?)\b.*$|(?<!\\)</?[A-Za-z][\w-]*(?:\s[^<>]*)?>", re.M)

HTML_NOISE = re.compile(r'\b(?:src|href|style|alt|title|class|id|name|data-[\w-]+)="[^"]*"|<!--.*?-->'
                        r'|<h[1-6][^>]*>\s*(?:Slide|Sheet|Page|Notes|Comments|Charts?)\b[^<]*</h[1-6]>'
                        r'|<figcaption>.*?</figcaption>', re.S)
FIELD_VOCAB = {"hyperlink", "pageref", "toc", "mergeformat", "mergefield", "formtext", "formcheckbox",
               "formdropdown", "seq", "docproperty", "includepicture", "ref", "fillin", "autotextlist",
               "styleref", "lastsavedby", "savedate", "createdate", "filename", "numpages", "page",
               "embed", "macrobutton", "arabic", "charformat", "date", "time", "author", "title"}
HIDDEN_RUN = re.compile(r"<w:r\b(?:(?!</w:r>).)*?<w:vanish/>(?:(?!</w:r>).)*?</w:r>", re.S)
WT = re.compile(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>")

# ---------------------------------------------------------------- helpers

def words(s):
    return collections.Counter(w.lower() for w in WORD.findall(s))

def recall(x, r):
    tot = sum(r.values())
    if tot == 0:
        return None
    return sum(min(x[w], r[w]) for w in r) / tot

def unexplained(gone, other):
    """Words in `gone` that `other` does not contain as a join or split.
    `gone` and `other` are word sets; a word joined into a longer token or
    split into shorter ones is re-tokenisation, not loss."""
    if not gone:
        return []
    blob = " ".join(other)
    prefixes = collections.defaultdict(list)
    for t in other:
        prefixes[t[:3]].append(t)
    real = []
    for w in gone:
        if w in blob:
            continue
        cands = prefixes.get(w[:3], ())
        if any(w.startswith(t) for t in cands) or any(w.endswith(t) for t in other if len(t) <= len(w) and w.endswith(t)):
            continue
        real.append(w)
    return real

def top_missing(exp, have, k=12):
    return [w for w, _ in (exp - have).most_common(k)]

def surface_text(raw, surface):
    if surface == "html":
        s = HTML_NOISE.sub(" ", raw)
        s = INLINE.sub("", s)
        s = TAG.sub(" ", s)
        s = (s.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
              .replace("&quot;", '"').replace("&#39;", "'"))
    elif surface == "markdown":
        s = MD_NOISE.sub(" ", raw)
    else:
        # A bracketed line is a placeholder for a picture or object
        # (`[Company logo]`), the text surface's form of markdown's
        # `![alt]()` and html's `<figcaption>`.
        s = re.sub(r"^\[[^\]\n]*\]\s*$", " ", raw, flags=re.M)
    return URL.sub(" ", s)

def docx_hidden_words(corpus, rel):
    """Words inside `w:vanish` runs of a .docx (body, headers, footers,
    notes) — text Word does not display."""
    import zipfile
    out = collections.Counter()
    try:
        with zipfile.ZipFile(os.path.join(corpus, rel)) as z:
            for name in z.namelist():
                if not (name.startswith("word/") and name.endswith(".xml")):
                    continue
                if not any(k in name for k in ("document", "header", "footer", "footnotes", "endnotes")):
                    continue
                xml = z.read(name).decode("utf-8", "replace")
                for run in HIDDEN_RUN.findall(xml):
                    for t in WT.findall(run):
                        out.update(words(t))
    except Exception:
        pass
    return out

def explain_loss(corpus, rel, lost):
    """Why the words in `lost` may be absent on purpose, or None."""
    if not lost:
        return None
    n = sum(lost.values())
    field = sum(c for w, c in lost.items() if w in FIELD_VOCAB or re.fullmatch(r"_?toc\d+|_?ref\d+", w))
    if field / n >= 0.5:
        return "field instruction text (hidden by policy, as Tika/antiword do)"
    if rel.endswith((".docx", ".docm", ".dotx")):
        hidden = docx_hidden_words(corpus, rel)
        if hidden:
            inside = sum(min(c, hidden[w]) for w, c in lost.items())
            if inside / n >= 0.6:
                return "hidden text (w:vanish; Word does not show it, hidden by policy)"
    return None

def read_gz(d, rel, surface):
    p = os.path.join(d, "out", rel + "." + surface + ".gz")
    if not os.path.exists(p):
        return None
    with gzip.open(p, "rb") as f:
        return f.read().decode("utf-8", "replace")

def load_sweep(d):
    recs = {}
    p = os.path.join(d, "sweep.jsonl")
    if not os.path.exists(p):
        return recs
    for line in open(p):
        r = json.loads(line)
        recs[(r["path"], r["surface"])] = r
    return recs

def load_panel(panel_dir):
    panel = collections.defaultdict(dict)
    p = os.path.join(panel_dir, "panel.jsonl") if panel_dir else None
    if p and os.path.exists(p):
        for line in open(p):
            r = json.loads(line)
            if "path" in r:
                panel[r["path"]][r["tool"]] = r
    return panel

def ref_texts(rel, panel, panel_dir, tika_dir):
    refs = {}
    if tika_dir:
        p = os.path.join(tika_dir, rel + ".txt")
        if os.path.exists(p):
            refs["tika"] = open(p, encoding="utf-8", errors="replace").read()
    for tool, r in panel.get(rel, {}).items():
        if r.get("status") == "ok":
            p = os.path.join(panel_dir, tool, rel + ".txt")
            if os.path.exists(p):
                refs[tool] = open(p, encoding="utf-8", errors="replace").read()
    return refs

def consensus(refs):
    """(members, expected) — references that agree with each other, and the
    member closest to the median size as the expected output."""
    ks = [k for k in refs if sum(refs[k].values()) >= 20]
    agree = set()
    for i in range(len(ks)):
        for j in range(i + 1, len(ks)):
            a, b = refs[ks[i]], refs[ks[j]]
            ra, rb = recall(a, b), recall(b, a)
            if ra is not None and rb is not None and ra >= 0.85 and rb >= 0.85:
                agree |= {ks[i], ks[j]}
    if not agree:
        return [], None
    sizes = {k: sum(refs[k].values()) for k in agree}
    med = statistics.median(sizes.values())
    return sorted(agree), min(agree, key=lambda k: abs(sizes[k] - med))

# ---------------------------------------------------------------- run

def sh(cmd, **kw):
    print("+", " ".join(cmd), flush=True)
    subprocess.run(cmd, check=True, **kw)

def cmd_run(a):
    os.makedirs(a.work, exist_ok=True)
    jobs = str(a.jobs)
    for arm, binary in (("prev", a.prev_bin), ("next", a.next_bin)):
        d = os.path.join(a.work, arm)
        if a.reuse and os.path.exists(os.path.join(d, "sweep.jsonl")):
            print(f"reusing {d}")
            continue
        sh([sys.executable, os.path.join(HERE, "sweep_tree.py"), binary, a.corpus, d, "--jobs", jobs])
    if a.validator:
        rt = os.path.join(a.work, "rt.jsonl")
        if not (a.reuse and os.path.exists(rt)):
            lst = os.path.join(a.work, "ooxml.list")
            with open(lst, "w") as f:
                for root, _, files in os.walk(a.corpus):
                    for fn in files:
                        if fn.rsplit(".", 1)[-1].lower() in ("docx", "xlsx", "pptx", "docm", "xlsm", "pptm"):
                            f.write(os.path.relpath(os.path.join(root, fn), a.corpus) + "\n")
            with open(rt, "w") as out:
                subprocess.run([a.validator, "--list", lst], cwd=a.corpus, stdout=out, stderr=subprocess.DEVNULL)
    render = os.path.join(a.work, "render", "render.jsonl")
    if not (a.reuse and os.path.exists(render)):
        os.makedirs(os.path.dirname(render), exist_ok=True)
        sh([sys.executable, os.path.join(HERE, "render_compare.py"), os.path.join(a.work, "next"), a.corpus, render, "--jobs", jobs])
    for name, src in (("panel", a.panel), ("tika", a.tika)):
        dst = os.path.join(a.work, name)
        if src and not os.path.exists(dst):
            os.symlink(os.path.abspath(src), dst)
    cmd_report(a)

# ---------------------------------------------------------------- report

def cmd_report(a):
    W = a.work
    P, N = load_sweep(os.path.join(W, "prev")), load_sweep(os.path.join(W, "next"))
    panel_dir = os.path.join(W, "panel") if os.path.exists(os.path.join(W, "panel")) else None
    tika_dir = os.path.join(W, "tika") if os.path.exists(os.path.join(W, "tika")) else None
    panel = load_panel(panel_dir)
    rt = {}
    rtp = os.path.join(W, "rt.jsonl")
    if os.path.exists(rtp):
        for line in open(rtp):
            try:
                r = json.loads(line)
            except ValueError:
                continue
            if r.get("step") == "rt_words":
                rt[r["path"]] = r
    render = {}
    rp = os.path.join(W, "render", "render.jsonl")
    if os.path.exists(rp):
        for line in open(rp):
            r = json.loads(line)
            if r.get("flags"):
                render[r["path"]] = r
    docs = sorted({p for (p, s) in N if p.split("/", 1)[0] in FMTS})
    ctx = (W, P, N, panel, panel_dir, tika_dir, rt, render)
    findings = []
    stats = collections.Counter()
    from concurrent.futures import ProcessPoolExecutor
    # Python 3.14 starts workers with forkserver, so module globals set
    # here are not inherited: hand the context over via the initializer.
    with ProcessPoolExecutor(a.jobs, initializer=_init, initargs=(ctx, a.corpus)) as ex:
        for fs, st in ex.map(analyse_one, docs, chunksize=16):
            findings.extend(fs)
            stats.update(st)
    write_report(a, findings, stats)

_CTX = None
CORPUS = None

def _init(ctx, corpus):
    global _CTX, CORPUS
    _CTX = ctx
    CORPUS = corpus

def analyse_one(rel):
    W, P, N, panel, panel_dir, tika_dir, rt, render = _CTX
    findings = []
    stats = collections.Counter()
    if True:
        fmt = rel.split("/", 1)[0]
        n, p = N.get((rel, "text"), {}), P.get((rel, "text"), {})
        nt = read_gz(os.path.join(W, "next"), rel, "text") if n.get("status") == "ok" else None
        pt = read_gz(os.path.join(W, "prev"), rel, "text") if p.get("status") == "ok" else None
        nw = words(nt) if nt is not None else None
        pw = words(pt) if pt is not None else None
        refs = {k: words(v) for k, v in ref_texts(rel, panel, panel_dir, tika_dir).items()}
        members, expected = consensus(refs)
        exp = refs[expected] if expected else None
        stats["docs"] += 1
        if refs:
            stats["with_refs"] += 1
        if members:
            stats["with_consensus"] += 1

        def add(cls, sev, detail, sample=None):
            findings.append({"path": rel, "fmt": fmt, "class": cls, "severity": sev,
                             "detail": detail, "sample": sample or [],
                             "next_status": n.get("status"), "prev_status": p.get("status"),
                             "next_words": sum(nw.values()) if nw else 0,
                             "prev_words": sum(pw.values()) if pw else 0,
                             "refs": {k: sum(v.values()) for k, v in refs.items()},
                             "consensus": members})

        # -- status transitions
        if p.get("status") == "ok" and n.get("status") != "ok" and pw and sum(pw.values()) > 0:
            add("REGRESSION_STATUS", 1, f"prev ok ({sum(pw.values())} words) -> next {n.get('status')}: {n.get('err', '')[:160]}")
        # -- two-arm content diff, corroborated by references
        if nw is not None and pw is not None:
            lost = pw - nw
            gone = unexplained(list(lost), list(nw))
            if exp is not None:
                gone = [w for w in gone if w in exp]
            elif len(refs) >= 2:
                gone = [w for w in gone if sum(w in r for r in refs.values()) >= 2]
            elif refs:
                gone = [w for w in gone if any(w in r for r in refs.values())]
            n_gone = sum(lost[w] for w in gone)
            if n_gone >= 5 and n_gone / max(sum(pw.values()), 1) >= 0.02:
                lost_c = collections.Counter({w: lost[w] for w in gone})
                # Each arm agreeing with a different reference is a policy
                # split, not a regression.
                sides = {}
                for k, r in refs.items():
                    if sum(r.values()) < 20:
                        continue
                    rn, rp = recall(nw, r), recall(pw, r)
                    if rn is not None and rn >= 0.97 and recall(r, nw) >= 0.97:
                        sides[k] = "next"
                    elif rp is not None and rp >= 0.97 and recall(r, pw) >= 0.97:
                        sides[k] = "prev"
                why = explain_loss(CORPUS, rel, lost_c)
                if exp is not None and (recall(nw, exp) or 0) >= max(0.95, (recall(pw, exp) or 0)):
                    add("PEER_SPLIT", 5, f"{n_gone} words prev had and next lost, but next matches the consensus {sorted(members)} ({recall(nw, exp):.2f}) at least as well as prev did ({recall(pw, exp):.2f})" + (f"; {why}" if why else ""),
                        [w for w, _ in lost_c.most_common(12)])
                elif "next" in sides.values():
                    # The new arm reproduces a reference's output; what it
                    # lost relative to the old arm, that reference does not
                    # show either (field instructions, tracked deletions).
                    add("PEER_SPLIT", 5, f"{n_gone} words differ; next matches {[k for k, v in sides.items() if v == 'next']} (>= 0.97 both ways)" + (f", prev matches {[k for k, v in sides.items() if v == 'prev']}" if "prev" in sides.values() else "") + (f"; {why}" if why else ""),
                        [w for w, _ in lost_c.most_common(12)])
                elif why:
                    add("PEER_SPLIT", 5, f"{n_gone} words prev had and next lost: {why}", [w for w, _ in lost_c.most_common(12)])
                else:
                    add("REGRESSION", 1, f"{n_gone} words prev had and next lost" + (" (corroborated by the references)" if refs else " (no reference to corroborate)"),
                        [w for w, _ in lost_c.most_common(12)])
        # -- against the consensus
        if exp is not None:
            if nw is not None and pw is not None:
                rn, rpv = recall(nw, exp), recall(pw, exp)
                if rn < 0.7 and rpv < 0.7:
                    add("SHARED_DEFECT", 2, f"recall vs consensus {sorted(members)}: next {rn:.2f}, prev {rpv:.2f}",
                        top_missing(exp, nw))
                elif rn >= rpv + 0.1:
                    add("IMPROVEMENT", 9, f"recall vs consensus: prev {rpv:.2f} -> next {rn:.2f}")
                elif rpv >= rn + 0.1:
                    miss = exp - nw
                    why = explain_loss(CORPUS, rel, miss)
                    if why:
                        add("PEER_SPLIT", 5, f"recall vs consensus fell: prev {rpv:.2f} -> next {rn:.2f}: {why}", top_missing(exp, nw))
                    else:
                        add("REGRESSION", 1, f"recall vs consensus fell: prev {rpv:.2f} -> next {rn:.2f}",
                            top_missing(exp, nw))
            if nw is not None and sum(nw.values()) == 0:
                add("EMPTY_VS_REFS", 2, f"we emit nothing; consensus {sorted(members)} has {sum(exp.values())} words",
                    [w for w, _ in exp.most_common(8)])
            if nw is not None and len(refs) >= 2 and sum(nw.values()) >= 50 and \
               all(sum(nw.values()) > 3 * sum(r.values()) for r in refs.values()):
                add("OVER_EXTRACTION", 3, f"we emit {sum(nw.values())} words, every reference <= {max(sum(r.values()) for r in refs.values())}",
                    [w for w, _ in (nw - exp).most_common(8)])
        if n.get("status") != "ok":
            big = {k: sum(v.values()) for k, v in refs.items() if sum(v.values()) >= 20}
            if big:
                add("PARSE_GAP", 2, f"we {n.get('status')}: {n.get('err', '')[:160]}; references extract {big}")
        # -- our own surfaces against each other
        if all(N.get((rel, s), {}).get("status") == "ok" for s in ("text", "markdown", "html")) and nt is not None:
            mt = read_gz(os.path.join(W, "next"), rel, "markdown")
            ht = read_gz(os.path.join(W, "next"), rel, "html")
            if mt is not None and ht is not None:
                wm, wh = words(surface_text(mt, "markdown")), words(surface_text(ht, "html"))
                base = max(sum(nw.values()), sum(wm.values()), sum(wh.values()), 1)
                if base >= 20:
                    pairs = {"text_not_md": nw - wm, "md_not_text": wm - nw, "text_not_html": nw - wh,
                             "html_not_text": wh - nw, "md_not_html": wm - wh, "html_not_md": wh - wm}
                    bad = {k: v for k, v in pairs.items() if sum(v.values()) >= 5 and sum(v.values()) / base >= 0.02}
                    if bad:
                        k = max(bad, key=lambda k: sum(bad[k].values()))
                        add("SURFACE_SPLIT", 2, "; ".join(f"{k}={sum(v.values())}" for k, v in bad.items()),
                            [w for w, _ in bad[k].most_common(8)])
        # -- round trip
        r = rt.get(rel)
        if r and r.get("total", 0) >= 10 and r["lost"] / r["total"] >= 0.05:
            add("ROUNDTRIP_LOSS", 2, f"IR->write->reread lost {r['lost']}/{r['total']} words, gained {r['gained']}")
        # -- rendering structure
        r = render.get(rel)
        if r:
            add("STRUCTURE_GAP", 3, "vs pandoc: " + ", ".join(r["flags"])[:200])
        # -- timing (same sweep, same load; a lead, re-measure serially).
        # More output is more work: an arm that extracts far more text
        # than the other (v0.1.11 stopped at a record cap and returned
        # nothing from a 29 MB workbook) is not slower for taking longer.
        grew = nw is not None and pw is not None and sum(nw.values()) > 1.5 * max(sum(pw.values()), 20)
        for s in ("text", "ir"):
            ns, ps = N.get((rel, s), {}), P.get((rel, s), {})
            if grew:
                break
            if ns.get("status") == "ok" and ps.get("status") == "ok" and ns.get("ms", 0) >= 500 and ns["ms"] >= 2 * ps.get("ms", 0):
                add("PERF_REGRESSION", 3, f"{s}: prev {ps['ms']:.0f} ms -> next {ns['ms']:.0f} ms (re-measure serially before filing)")
                break
    return findings, stats

def write_report(a, findings, stats):
    W = a.work
    findings.sort(key=lambda f: (f["severity"], f["fmt"], f["class"], f["path"]))
    with open(os.path.join(W, "findings.jsonl"), "w") as f:
        for r in findings:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    by = collections.Counter((f["fmt"], f["class"]) for f in findings)
    lines = ["# Regression harness findings", "",
             f"documents: {stats['docs']}; with >=1 reference: {stats['with_refs']}; with a consensus (>=2 agreeing references): {stats['with_consensus']}", "",
             "| format | class | files |", "|---|---|---|"]
    for (fmt, cls), c in sorted(by.items(), key=lambda kv: (min(f['severity'] for f in findings if f['class'] == kv[0][1]), kv[0])):
        lines.append(f"| {fmt} | {cls} | {c} |")
    lines.append("")
    for cls in sorted({f["class"] for f in findings}, key=lambda c: min(f["severity"] for f in findings if f["class"] == c)):
        lines.append(f"## {cls}")
        lines.append("")
        for f in [f for f in findings if f["class"] == cls][: a.top]:
            lines.append(f"- `{f['path']}` — {f['detail']}" + (f" — e.g. {', '.join(f['sample'][:8])}" if f["sample"] else ""))
        lines.append("")
    open(os.path.join(W, "findings.md"), "w").write("\n".join(lines))
    print("\n".join(lines[:len(by) + 6]))
    print(f"\n{len(findings)} findings -> {os.path.join(W, 'findings.md')}")

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("run")
    r.add_argument("--prev-bin", required=True)
    r.add_argument("--next-bin", required=True)
    r.add_argument("--corpus", required=True)
    r.add_argument("--work", required=True)
    r.add_argument("--panel")
    r.add_argument("--tika")
    r.add_argument("--validator", help="release build of examples/corpus_validate for the round-trip axis")
    r.add_argument("--jobs", type=int, default=4)
    r.add_argument("--reuse", action="store_true")
    r.add_argument("--top", type=int, default=40)
    r.set_defaults(fn=cmd_run)
    p = sub.add_parser("report")
    p.add_argument("--work", required=True)
    p.add_argument("--corpus", required=True, help="the corpus root, for explanations that read the file")
    p.add_argument("--top", type=int, default=40)
    p.add_argument("--jobs", type=int, default=4)
    p.set_defaults(fn=cmd_report)
    a = ap.parse_args()
    a.fn(a)

if __name__ == "__main__":
    main()
