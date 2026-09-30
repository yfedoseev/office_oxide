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

The sources are third-party repositories, so each is pinned to a commit and
every file is checked against a SHA-256 recorded here. Fetching from a
mutable branch meant a force-push or an upstream edit could silently change
what this CI gate enforces. To move a pin, update the commit and re-record
the hashes (the script prints the new hash of any file that mismatches).

Provenance: the ISO/IEC 29500-4:2016 transitional schemas as redistributed
by the dolanmiu/docx project (MIT), and the OPC schemas (ECMA-376 Part 2)
plus Dublin Core schemas as redistributed by randym/axlsx (MIT).
"""

import hashlib
import os
import sys
import urllib.request

DOCX_COMMIT = "9737ae7f915a17df496e0294147ae0249acadf4b"
AXLSX_COMMIT = "8e7b4b3b7259103452c191f73ce0bf64f66033fe"

BASE = (
    f"https://raw.githubusercontent.com/dolanmiu/docx/{DOCX_COMMIT}/"
    "ooxml-schemas/ISO-IEC29500-4_2016"
)
OPC_BASE = f"https://raw.githubusercontent.com/randym/axlsx/{AXLSX_COMMIT}/lib/schema"

SHA256 = {
    "wml.xsd": "cf90407251dff19633e9e397ea77a2457764132cf6b736301b48042ea1111354",
    "pml.xsd": "39c6396b7b1f096885c3ceba5dab30e3ee14291c45a1d6fa6038ec4db69d88d0",
    "sml.xsd": "495debc8fa967b77ed37799747b049832f2c95b2ecdb9f19ddcd68f8c9f96ab9",
    "xml.xsd": "70a3c67959f9ad2333016328716101bf8b476e682f29830cc9ee8da63dca7cc2",
    "dml-main.xsd": "6978ba7e889070b0c3cb5b546b23e5a6c3516134afc53b87a21f482ca33f3858",
    "dml-chart.xsd": "4de4390a26e7e44e682ed98b39872cce2620f807f5b0de421567a43f74526388",
    "dml-chartDrawing.xsd": "3fd0586f2637b98bb9886f0e0b67d89e1cc987c2d158cc7deb5f5b9890ced412",
    "dml-diagram.xsd": "809f77f658f71e10f5cfbf99e25c895a23a4e62c81210b59a2fe8c7fb24849d2",
    "dml-lockedCanvas.xsd": "5cb76dabd8b97d1e9308a1700b90c20139be4d50792d21a7f09789f5cccd6026",
    "dml-picture.xsd": "5d389d42befbebd91945d620242347caecd3367f9a3a7cf8d97949507ae1f53c",
    "dml-spreadsheetDrawing.xsd": "b4532b6d258832953fbb3ee4c711f4fe25d3faf46a10644b2505f17010d01e88",
    "dml-wordprocessingDrawing.xsd": "2dbfd4719505bf39b568005994a433997cbeac59cebc9d4867d564c94e9ba03a",
    "shared-additionalCharacteristics.xsd": "3c6709101c6aaa82888df5d8795c33f9e857196790eb320d9194e64be2b6bdd8",
    "shared-bibliography.xsd": "0b364451dc36a48dd6dae0f3b6ada05fd9b71e5208211f8ee5537d7e51a587e2",
    "shared-commonSimpleTypes.xsd": "48675c4f82f6b097434d4b7e313635b790257cca58fa90dd0e8f19c9affa18ed",
    "shared-customXmlDataProperties.xsd": "0ef4bb354ff44b923564c4ddbdda5987919d220225129ec94614a618ceafc281",
    "shared-customXmlSchemaProperties.xsd": "0d103b99a4a8652f8871552a69d42d2a3760ac6a5e3ef02d979c4273257ff6a4",
    "shared-documentPropertiesCustom.xsd": "9c085407751b9061c1f996f6c39ce58451be22a8d334f09175f0e89e42736285",
    "shared-documentPropertiesExtended.xsd": "bc92e36ccd233722d4c5869bec71ddc7b12e2df56059942cce5a39065cc9c368",
    "shared-documentPropertiesVariantTypes.xsd": "7b5b7413e2c895b1e148e82e292a117d53c7ec65b0696c992edca57b61b4a74b",
    "shared-math.xsd": "f812dffca0bb66db6e2ec2e01dd19258636b2a5a3a0bb949a2e7295665e0ce61",
    "shared-relationshipReference.xsd": "12264f3c03d738311cd9237d212f1c07479e70f0cbe1ae725d29b36539aef637",
    "vml-main.xsd": "335e9ecd40bb594ec4b916ac8c7b9eead4e16fe1b79f9d4105fdb1c43b66bb1c",
    "vml-officeDrawing.xsd": "585bedc1313b40888dcc544cb74cd939a105ee674f3b1d3aa1cc6d34f70ff155",
    "vml-presentationDrawing.xsd": "133c9f64a5c5d573b78d0a474122b22506d8eadb5e063f67cdbbb8fa2f161d0e",
    "vml-spreadsheetDrawing.xsd": "6bdeb169c3717eb01108853bd9fc5a3750fb1fa5b82abbdd854d49855a40f519",
    "vml-wordprocessingDrawing.xsd": "475dcae1e7d1ea46232db6f8481040c15e53a52a3c256831d3df204212b0e831",
    "opc-coreProperties.xsd": "bd86c472a1680c708c6217d63efcd0b0052d204c3abe7d56e428a848b6527a2f",
    "opc-contentTypes.xsd": "49ec1c03440a9b5590e69acdf92b200fe133697de0b05f43a129fb94d1885026",
    "opc-relationships.xsd": "607ffb6879cf5d5d765bfdfb2548f3f427ea964ddee6d08a15f0dc781749866d",
    "dc.xsd": "3a3e2858492bb727c1a7732ff8bf918d3bf60cb61cdc348b462f22f5b699625b",
    "dcterms.xsd": "35cf4b3adb1f525bcfcd0a29e6959109acd77d71680211f0850e9d26b30018ed",
    "dcmitype.xsd": "8fe4dddb1c5f45267627af08d3e26cf1bdd2b48f8655e97dd5115b4cb2eb0b7a",
}

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
    assert set(SHA256) == set(MAIN) | set(OPC), "every schema needs a pinned hash"
    for name, base in [(n, BASE) for n in MAIN] + [(n, OPC_BASE) for n in OPC]:
        dest = os.path.join(out, name)
        if os.path.exists(dest):
            # A cached copy is only reused if it is the pinned file (wml.xsd
            # is rewritten below, so its hash is checked at download time).
            if name == "wml.xsd":
                continue
            with open(dest, "rb") as f:
                if hashlib.sha256(f.read()).hexdigest() == SHA256[name]:
                    continue
        url = f"{base}/{name}"
        try:
            with urllib.request.urlopen(url, timeout=60) as r:
                data = r.read()
        except Exception as exc:  # noqa: BLE001 - report and fail loudly
            print(f"FAILED {name}: {exc}", file=sys.stderr)
            return 1
        digest = hashlib.sha256(data).hexdigest()
        if digest != SHA256[name]:
            print(
                f"FAILED {name}: sha256 {digest} does not match the pinned "
                f"{SHA256[name]}",
                file=sys.stderr,
            )
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
