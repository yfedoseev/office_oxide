#!/usr/bin/env python3
"""Fetch the ISO/IEC 29500-4 (transitional) schema set into ./xsd.

The schemas are third-party files, so they are downloaded rather than
committed (CONTRIBUTING #4). Transitional is the right target: it is what
Office actually writes.

Two wrinkles this handles, both of which otherwise look like library bugs:

* `wml.xsd` imports the markup-compatibility namespace from a path that is
  not in the distributed set. A small stub declaring `mc:Ignorable` and
  friends is written so the schema will load at all.
* `a:graphicData` uses a strict wildcard, so 2010-era extension content
  (`wps:wsp` inside a text box) has no global declaration here and fails
  validation even though real Word files contain it. `validate.py` treats
  content under `graphicData` as lax for that reason.
"""

import os
import sys
import urllib.request

BASE = (
    "https://raw.githubusercontent.com/dolanmiu/docx/master/"
    "ooxml-schemas/ISO-IEC29500-4_2016"
)
OPC_BASE = "https://raw.githubusercontent.com/randym/axlsx/master/lib/schema"

MAIN = [
    "wml.xsd", "pml.xsd", "sml.xsd", "xml.xsd",
    "dml-main.xsd", "dml-chart.xsd", "dml-chartDrawing.xsd", "dml-diagram.xsd",
    "dml-lockedCanvas.xsd", "dml-picture.xsd", "dml-spreadsheetDrawing.xsd",
    "dml-wordprocessingDrawing.xsd",
    "shared-additionalCharacteristics.xsd", "shared-bibliography.xsd",
    "shared-commonSimpleTypes.xsd", "shared-customXmlDataProperties.xsd",
    "shared-customXmlSchemaProperties.xsd", "shared-documentPropertiesCustom.xsd",
    "shared-documentPropertiesExtended.xsd",
    "shared-documentPropertiesVariantTypes.xsd", "shared-math.xsd",
    "shared-relationshipReference.xsd",
    "vml-main.xsd", "vml-officeDrawing.xsd", "vml-presentationDrawing.xsd",
    "vml-spreadsheetDrawing.xsd", "vml-wordprocessingDrawing.xsd",
]
OPC = [
    "opc-coreProperties.xsd", "opc-contentTypes.xsd", "opc-relationships.xsd",
    "dc.xsd", "dcterms.xsd", "dcmitype.xsd",
]

MC_STUB = """<?xml version="1.0" encoding="utf-8"?>
<xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema"
  targetNamespace="http://schemas.openxmlformats.org/markup-compatibility/2006"
  elementFormDefault="qualified" attributeFormDefault="qualified">
  <xsd:attribute name="Ignorable" type="xsd:string"/>
  <xsd:attribute name="ProcessContent" type="xsd:string"/>
  <xsd:attribute name="PreserveElements" type="xsd:string"/>
  <xsd:attribute name="PreserveAttributes" type="xsd:string"/>
  <xsd:attribute name="MustUnderstand" type="xsd:string"/>
</xsd:schema>
"""


def main() -> int:
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "xsd")
    os.makedirs(out, exist_ok=True)
    for name, base in [(n, BASE) for n in MAIN] + [(n, OPC_BASE) for n in OPC]:
        dest = os.path.join(out, name)
        if os.path.exists(dest):
            continue
        url = f"{base}/{name}"
        try:
            with urllib.request.urlopen(url, timeout=60) as r:
                data = r.read()
        except Exception as exc:  # noqa: BLE001 - report and fail loudly
            print(f"FAILED {name}: {exc}", file=sys.stderr)
            return 1
        with open(dest, "wb") as f:
            f.write(data)

    with open(os.path.join(out, "mc.xsd"), "w") as f:
        f.write(MC_STUB)
    wml = os.path.join(out, "wml.xsd")
    src = open(wml).read()
    if '../mce/mc.xsd' in src:
        open(wml, "w").write(src.replace('schemaLocation="../mce/mc.xsd"',
                                         'schemaLocation="mc.xsd"'))
    print(f"schemas ready in {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
