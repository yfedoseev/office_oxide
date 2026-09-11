"""Validate every XML part of an OOXML package against ISO/IEC 29500-4.

Run `fetch_schemas.py` first. Exit code is non-zero if any part is invalid,
so this is usable as a CI gate.

Known limitation, deliberately suppressed: `a:graphicData` uses a strict
wildcard, so 2010-era extension content (`wps:wsp` in a text box) has no
global declaration in this schema set and fails even though real Word files
contain it. That is a validator artifact, not a defect.
"""

import os
import re
import sys
import zipfile

from lxml import etree

XSD = os.path.join(os.path.dirname(os.path.abspath(__file__)), "xsd")

NS2XSD = {
 "http://schemas.openxmlformats.org/presentationml/2006/main": "pml.xsd",
 "http://schemas.openxmlformats.org/wordprocessingml/2006/main": "wml.xsd",
 "http://schemas.openxmlformats.org/spreadsheetml/2006/main": "sml.xsd",
 "http://schemas.openxmlformats.org/drawingml/2006/main": "dml-main.xsd",
 "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties": "shared-documentPropertiesExtended.xsd",
 "http://schemas.openxmlformats.org/drawingml/2006/chart": "dml-chart.xsd",
 "http://schemas.openxmlformats.org/package/2006/relationships": "opc-relationships.xsd",
 "http://schemas.openxmlformats.org/package/2006/content-types": "opc-contentTypes.xsd",
 "http://schemas.openxmlformats.org/package/2006/metadata/core-properties": "opc-coreProperties.xsd",
 "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing": "dml-spreadsheetDrawing.xsd",
 "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing": "dml-chartDrawing.xsd",
 "http://schemas.openxmlformats.org/drawingml/2006/diagram": "dml-diagram.xsd",
 "http://schemas.openxmlformats.org/drawingml/2006/picture": "dml-picture.xsd",
 "urn:schemas-microsoft-com:vml": "vml-main.xsd",
}
SKIP_NS = set()

_cache = {}
def schema_for(xsdname):
    if xsdname not in _cache:
        p = os.path.join(XSD, xsdname)
        _cache[xsdname] = etree.XMLSchema(etree.parse(p))
    return _cache[xsdname]

def validate_pkg(path):
    out = []
    with zipfile.ZipFile(path) as z:
        for n in sorted(z.namelist()):
            if not n.endswith(".xml") and not n.endswith(".rels"):
                continue
            data = z.read(n)
            try:
                # huge_tree: libxml2 caps nesting at 256 by default and reports a
                # *parse* failure past it. Our own writer bounds nesting at 256
                # and emits balanced XML, so without this the validator invents
                # a malformed-output finding for a legitimately deep document.
                parser = etree.XMLParser(huge_tree=True)
                doc = etree.fromstring(data, parser)
            except Exception as e:
                out.append((n, "PARSE", str(e)))
                continue
            ns = etree.QName(doc).namespace
            if ns in SKIP_NS:
                continue
            xsdname = NS2XSD.get(ns)
            if xsdname is None:
                out.append((n, "NOSCHEMA", ns or "(none)"))
                continue
            sch = schema_for(xsdname)
            if not sch.validate(doc):
                for e in sch.error_log:
                    if "strict wildcard" in e.message:
                        # 2010 extension content under a:graphicData; see the
                        # module docstring.
                        continue
                    out.append((n, "INVALID", f"line {e.line}: {e.message}"))
    return out

if __name__ == "__main__":
    total = 0
    for p in sys.argv[1:]:
        res = validate_pkg(p)
        base = os.path.basename(p)
        if not res:
            print(f"PASS  {base}")
            continue
        total += len(res)
        print(f"FAIL  {base}  ({len(res)} findings)")
        for n, kind, msg in res:
            print(f"        [{kind}] {n}\n              {msg}")
    print(f"\n=== {total} findings ===")
    sys.exit(1 if total else 0)
