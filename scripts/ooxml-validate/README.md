# OOXML validation gate

Nothing in the test suite used to check that the documents this library
*writes* are valid OOXML. Generated files were only round-tripped through our
own parser, which is deliberately lenient about element order and missing
elements — so a document Word refuses to open passed every test we had.

Two independent bug reports' worth of defects were live in `main` at v0.1.10
with a green suite: [#199](../../issues/199) (every PPTX missing the required
`p:clrMap`) and [#200](../../issues/200) (six DOCX violations, one firing on
every markdown table and one on every document with a list).

## Running it

```bash
python3 -m pip install lxml
python3 scripts/ooxml-validate/fetch_schemas.py         # writes ./xsd, gitignored
cargo run --example gen_validation_corpus -- /tmp/corpus
python3 scripts/ooxml-validate/validate.py  /tmp/corpus/*
python3 scripts/ooxml-validate/opccheck.py  /tmp/corpus/*
```

Both validators exit non-zero on any finding. CI runs exactly this.

## What each part covers

| | catches |
|---|---|
| `validate.py` | missing required elements, wrong element order, out-of-range attribute values — the class behind #199, #200, #202, #204 |
| `opccheck.py` | every part has a content type; every non-external relationship target resolves. Needs no schemas. |
| `gen_validation_corpus.rs` | the inputs. Exercises the builder APIs, not only `create_from_markdown`. |

## Things that bit us, so they are written down

**Schema validation halts at the first content-model error per part.** One
missing element hides everything after it, so a finding count is a lower
bound. Re-run after each fix rather than assuming the count is the defect
count.

**Vary combinations, not one property at a time.** All four schema-invalid
element orders in 0.1.11 needed *two* properties on the same element — an
indent and spacing on one paragraph, borders and shading on one cell. A matrix
that sets one property per document cannot find an ordering bug, and the first
version of this gate missed three of the four for exactly that reason; a
6,000-file real corpus found them instead.

`maximal_properties_corpus` in the generator now sets every child of every
order-sensitive group at once. In WML that surface is closed and enumerable:
`CT_PPrBase`, `CT_TblPrBase`, `CT_TcPrBase` and `CT_SectPr` are sequences;
`EG_RPrBase` and `CT_TrPrBase` are *choices* and therefore order-free. Check the
compositor in the XSD before assuming an ordering bug — I chased a fourth one
in `w:rPr` that does not exist.

**Cover conversion, not just creation.** The `w:pgMar` missing-`w:gutter`
defect was only reachable through `Document::save_as`, because it needs a
section that carries a page setup.

**Include out-of-range values.** `font_size(f64::NAN)`, `"#FF0000"` as a
colour, and a slide under an inch wide are all reachable through the public
API and all produced invalid files.

**`a:graphicData` uses a strict wildcard.** 2010-era extension content
(`wps:wsp` inside a text box) has no global declaration in the 29500-4 set and
fails validation even though real Word files contain it. `validate.py`
suppresses that specific message; it is a validator artifact, not a defect.

## What this does not cover

Schema validity is necessary, not sufficient — Word and PowerPoint accept some
invalid files and reject some valid ones. It also cannot see dangling
cross-part references ([#208](../../issues/208)) or ordering rules the schema
does not encode, such as ascending cell order within a row
([#205](../../issues/205)). Those have their own unit tests.
