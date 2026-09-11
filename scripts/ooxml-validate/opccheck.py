"""OPC-level checks the XSDs cannot express.

* every part in the package has a content type (a Default or an Override)
* every Override names a part that exists
* every non-external relationship Target resolves
* the package has a root relationships part

Exit code is non-zero on any finding, so this is usable as a CI gate and
needs no downloaded schemas.
"""

import os
import posixpath
import sys
import zipfile

from lxml import etree

CT_NS="http://schemas.openxmlformats.org/package/2006/content-types"
R_NS="http://schemas.openxmlformats.org/package/2006/relationships"

def check(path):
    f=[]
    with zipfile.ZipFile(path) as z:
        names=set(z.namelist())
        if "[Content_Types].xml" not in names:
            return [("FATAL","", "no [Content_Types].xml")]
        ct=etree.fromstring(z.read("[Content_Types].xml"), etree.XMLParser(huge_tree=True))
        defaults={e.get("Extension").lower():e.get("ContentType") for e in ct.findall(f"{{{CT_NS}}}Default")}
        overrides={e.get("PartName").lstrip("/"):e.get("ContentType") for e in ct.findall(f"{{{CT_NS}}}Override")}

        # 1. every part has a content type
        for n in sorted(names):
            if n=="[Content_Types].xml": continue
            if n.endswith("/"): continue
            ext=n.rsplit(".",1)[-1].lower() if "." in n else ""
            if n in overrides: continue
            if ext in defaults: continue
            f.append(("NO-CONTENT-TYPE",n,f"ext '{ext}' has no Default and no Override"))

        # 2. overrides that point at nothing
        for p in overrides:
            if p not in names:
                f.append(("DANGLING-OVERRIDE",p,"Override PartName does not exist in package"))

        # 3. relationship targets resolve
        for n in sorted(names):
            if not n.endswith(".rels"): continue
            base=posixpath.dirname(posixpath.dirname(n))  # strip _rels
            try: r=etree.fromstring(z.read(n), etree.XMLParser(huge_tree=True))
            except Exception as e:
                f.append(("RELS-PARSE",n,str(e))); continue
            for rel in r.findall(f"{{{R_NS}}}Relationship"):
                if rel.get("TargetMode")=="External": continue
                t=rel.get("Target")
                if t.startswith("/"): res=t.lstrip("/")
                else: res=posixpath.normpath(posixpath.join(base,t)) if base else posixpath.normpath(t)
                if res not in names:
                    f.append(("DANGLING-REL",n,f"Id={rel.get('Id')} Target={t} -> {res} missing"))
        # 4. package-level required rel
        if "_rels/.rels" not in names:
            f.append(("MISSING","_rels/.rels","package has no root relationships part"))
    return f

tot=0
for p in sys.argv[1:]:
    r=check(p); b=os.path.basename(p)
    if not r: print(f"PASS  {b}"); continue
    tot+=len(r); print(f"FAIL  {b} ({len(r)})")
    for k,n,m in r: print(f"        [{k}] {n}\n              {m}")
print(f"\n=== {tot} findings ===")
sys.exit(1 if tot else 0)
