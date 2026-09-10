# Reconstructing tables (incl. merged cells) and lists from Word 97‑2003 `.doc`

Independent design research for `office_oxide`. Everything below is grounded in
(a) the primary [MS-DOC] specification, (b) the actual source of Apache POI HWPF and
LibreOffice's `sw/source/filter/ww8` import filter, and (c) byte-level dumps of the three
sample files in `./samples-doc/` that I parsed from scratch during this investigation.

Where I write "verified on `<sample>`" I mean I decoded that structure out of the real file
and printed it; the numbers quoted are literal bytes from the sample, not illustrations.

---

## 0. Executive summary

The current reader (`src/doc/piece_table.rs::extract_text` → `src/convert_doc.rs::doc_to_ir`)
throws away the only information that makes tables and lists recoverable: the **paragraph
property stream**. Word does not store tables or lists as objects. It stores:

* a flat character stream, in which `U+0007` is a **cell mark** and `U+000D` is a paragraph mark;
* a per-paragraph property blob (**PAPX**) reachable through `PlcBtePapx` → `PapxFkp`;
* inside the PAPX of the *last* paragraph of each row (the **TTP mark**), a complete row
  definition (`sprmTDefTable`) giving the column count, the column boundaries in twips, and a
  20-byte `TC80` per cell carrying the merge bits;
* for lists, `sprmPIlfo` / `sprmPIlvl` on each paragraph, resolved through
  `PlfLfo` → `LFO` → `LSTF` (matched by `lsid`) → `LVL[ilvl]`, where `LVLF.nfc` is the
  number-format code (`0x17` = bullet, `0xFF` = none, everything else = a numbering scheme).

So the work splits cleanly into three new pieces:

1. **`src/doc/papx.rs`** — `PlcBtePapx` / `PapxFkp` / `PapxInFkp` / `GrpPrlAndIstd` reader plus a
   generic `Sprm` iterator. This is ~250 lines and is the load-bearing new capability.
2. **`src/doc/table.rs`** — group paragraphs into cells/rows/tables by `fInTable`/`fTtp`/`itap`,
   parse `TDefTableOperand`, build a **union column grid** over all rows of the table, and derive
   `col_span` geometrically and `row_span` from the vertical-merge bits.
3. **`src/doc/list.rs`** — `PlfLst`/`LSTF`/`LVL`/`PlfLfo`/`LFO` reader plus the ilfo→LVL walk.

`DocDocument` then stops being "a `String`" and becomes "a `Vec<DocParagraph>`", and
`convert_doc.rs` becomes a straight structural walk instead of a line heuristic.

**The single pitfall I would bet money an implementation gets wrong**: computing `col_span` from
the *cell index* instead of from `rgdxaCenter` **twip geometry unioned across all rows of the
table**. In Word 97 a horizontal merge is not a flag — Word physically rewrites the row's column
boundaries so the merged cell is one wide cell. `table-merges.doc` proves it: every `TC80.horzMerge`
in that file is `0`, yet row 1 must render as `colspan=3 | colspan=2` and row 4 as `colspan=5`.
POI does exactly this geometric union. LibreOffice reaches the same visual result by a different
route (see §3.1) because Writer tables are hierarchical and need no colspan at all — but its
merge matching is likewise pure twip geometry. Cell-index arithmetic works on neither.

Runner-up bet: **`sprmTDefTable`'s length prefix is a `u16`, not a `u8`** (pitfall P-G1). Getting it
wrong reads `NumberOfColumns` out of the high byte of `cb`, i.e. **zero columns**, while the sprm
walk still resynchronises perfectly — so the failure looks like "tables are empty" rather than
"parsing crashed", and it survives casual testing.

---

## 1. The chain from the FIB to per-paragraph properties

### 1.1 FIB offsets (nFib ≥ 0x00C1)

`FibRgFcLcb97` begins at absolute offset **0x009A** in the `WordDocument` stream and is a flat
array of `(fc: u32, lcb: u32)` pairs. Field *n* (1-based, counting `fc` and `lcb` separately) is at
`0x9A + (n-1)*4`. I extracted the full ordered field list from the spec page and verified the
arithmetic against the constants the existing `fib.rs` already uses:

| field | ordinal | absolute offset | note |
|---|---|---|---|
| `fcPlcfBtePapx` | 27 | **0x0102** | offset of `PlcBtePapx` in the Table stream |
| `lcbPlcfBtePapx` | 28 | **0x0106** | |
| `fcClx` | 67 | **0x01A2** | already used by `fib.rs` ✔ |
| `lcbClx` | 68 | **0x01A6** | |
| `fcPlfLst` | 147 | **0x02E2** | list definitions |
| `lcbPlfLst` | 148 | **0x02E6** | |
| `fcPlfLfo` | 149 | **0x02EA** | list format overrides |
| `lcbPlfLfo` | 150 | **0x02EE** | |
| `fcStshf` / `lcbStshf` | 3 / 4 | 0x00A2 / 0x00A6 | style sheet (phase 2, see §4.4) |

Cross-check: MS-DOC's own worked "Example of a List" prints `0x000002E2 fcPlfLst`,
`0x000002EA fcPlfLfo` — exactly matching. Verified on all three samples: `simple-list.doc`
gives `fcPlfLst=0x160 lcbPlfLst=30 fcPlfLfo=0x1B0 lcbPlfLfo=24`.

Table stream selection is `FibBase.flags` bit 9 (`fWhichTblStm`) at offset 0x0A — already correct
in `fib.rs`. (`table.doc` uses `0Table`; the other two use `1Table`.)

### 1.2 The chain

```
WordDocument[0x0102] ─ fcPlcfBtePapx ──► Table stream
                                          │
                                          ▼
                                    PlcBtePapx  { aFC: [u32; n+1], aPnBtePapx: [PnFkpPapx; n] }
                                          │  aPnBtePapx[j].pn  (low 22 bits)
                                          ▼
                        WordDocument[ pn * 512 ] ──► PapxFkp (exactly 512 bytes)
                                          │
                        ┌─────────────────┴─────────────────┐
                        ▼                                   ▼
        rgfc: [u32; cpara+1]                      rgbx: [BxPap; cpara]   (13 bytes each)
        (fc range of each paragraph)              BxPap.bOffset (1 byte)
                                                          │  ×2
                                                          ▼
                                        PapxInFkp { cb: u8, grpprlInPapx }
                                                          │
                                                          ▼
                                        GrpPrlAndIstd { istd: u16, grpprl: [Prl] }
                                                          │
                                                          ▼
                                        Prl { sprm: u16, operand: [u8] }
```

Exact rules (all verified against the samples):

* **`PlcBtePapx`** ([MS-DOC] 2.8.6). `n = (lcb - 4) / 8`. `aFC` has `n+1` `u32`s;
  `aPnBtePapx` has `n` `u32`s starting at `(n+1)*4`. `PnFkpPapx.pn` is the **low 22 bits**;
  the top 10 bits are undefined and *must be masked* (`pn & 0x003F_FFFF`).
  Note the PLC maps **stream offsets**, not CPs — unusual among PLCs.
* **`PapxFkp`** ([MS-DOC] 2.9.186) is always **512 bytes**, page-aligned at `pn*512` in the
  *WordDocument* stream (not the Table stream — easy to get wrong).
  `cpara` is the **last byte** (offset 511), `1 ≤ cpara ≤ 0x1D`.
  `rgfc` is `cpara+1` `u32`s at offset 0; `rgbx` starts at `(cpara+1)*4`.
* **`BxPap`** ([MS-DOC] 2.9.23) is **13 bytes**: `bOffset: u8` + 12 reserved (legacy `PHE`).
  The PapxInFkp is at byte offset `bOffset * 2` inside the same 512-byte page.
  **`bOffset == 0` means "no PAPX; default properties"** — you must handle it, not skip the entry.
  (`table.doc` has one such paragraph.)
* **`PapxInFkp`** ([MS-DOC] 2.9.187): `cb: u8`. If `cb != 0`, the `GrpPrlAndIstd` is
  `2*cb - 1` bytes starting right after `cb`. If `cb == 0`, the next byte `cb'` (which MUST be ≥ 1)
  gives the size and the `GrpPrlAndIstd` is `2*cb'` bytes after it.
  **Both samples use exclusively the `cb == 0` form** — an implementation that only handles
  `cb != 0` parses nothing at all on real files.
* **`GrpPrlAndIstd`**: `istd: u16` then the `grpprl` byte array. The `istd` is the paragraph
  style index — free heading detection, see §4.4.

### 1.3 Decoding a `Prl` / `Sprm` ([MS-DOC] 2.6.1, 2.2.x)

`sprm: u16` little-endian, bit-packed **LSB-first**:

```
bits 0..8   ispmd   (9 bits)
bit  9      fSpec
bits 10..12 sgc     1=paragraph 2=character 3=picture 4=section 5=table
bits 13..15 spra    operand size code
```

Operand length by `spra`: `0→1, 1→1, 2→2, 3→4, 4→2, 5→2, 7→3`, and `6 → variable`.

**The variable case has an exception that will silently corrupt your whole grpprl walk if you
miss it.** For `spra == 6` the operand normally begins with a 1-byte length. But for
`sprmTDefTable (0xD608)` and `sprmPChgTabs (0xC615)` the length prefix is a **`u16`**. MS-DOC states
this in the `spra` table ("except in the cases of sprmTDefTable and sprmPChgTabs"), and POI encodes
it in `SprmOperation.initSize`:

```java
case 6:
    int offset = _gOffset;
    if ( sprm == SPRM_LONG_TABLE || sprm == SPRM_LONG_PARAGRAPH )
    {
        int retVal = ( 0x0000ffff & LittleEndian.getShort( _grpprl, offset ) ) + 3;
        _gOffset += 2;   // operand body starts after the 2-byte cb
        return retVal;
    }
    return ( 0x000000ff & _grpprl[_gOffset++] ) + 3;
```

For `sprmTDefTable`, total bytes consumed = `2 (sprm) + 2 (cb) + (cb - 1)` = `cb + 3`, because
`TDefTableOperand.cb` is defined as *"the number of bytes used by the remainder of this structure,
incremented by 1"*.

Verified on `table.doc`: the row PAPX contains
`D6 08 | 46 00 | 03 | 94FF 860B 7817 6A23 | <3 × 20-byte TC80>`
→ `cb = 0x0046 = 70`, remainder = 69 = `1 + 2*(3+1) + 3*20`, total sprm entry = 73 bytes.
Reading `cb` as a single byte silently yields `NumberOfColumns = 0` here — see pitfall **P-G1**.

LibreOffice encodes the same exception as a distinct sprm length class. `ww8scan.hxx`:
`enum SprmType {L_FIX=0, L_VAR=1, L_VAR2=2};` — and the enum value doubles as the offset of the
length field. `ww8scan.cxx`:
```cpp
case L_VAR2:
{
    // Variable 2-Byte Length
    // For sprmTDefTable and sprmTDefTable10, the length of the
    // parameter plus 1 is recorded in the two bytes beginning
    // at offset (WW7-) 1 or (WW8+) 2
    sal_uInt8 nIndex = 1 + mnDelta;
    nCount = SVBT16ToUInt16(&pSprm[nIndex]);
    if (nCount) --nCount;
    nL = static_cast<sal_uInt16>(nCount + aSprm.nLen);
}
```
(`mnDelta` is 1 for WW8's 2-byte sprm ids, 0 for WW6/7's 1-byte ids.)

Two more indirections you must at least detect:

* **`sprmPHugePapx` (0x6646)** — `u32` offset into the **Data stream** to a `PrcData`; if present it
  MUST be the first (and only) Prl and the real grpprl lives there. Common on large tables.
* **`sprmPTableProps` (0x646B)** — same mechanism for table properties.
  Both may chain, but the chain must terminate. If you don't follow them you lose whole tables.

---

## 2. How table structure is encoded

### 2.1 The three questions, answered

Everything comes from **direct** paragraph properties on the paragraph mark / cell mark
([MS-DOC] 2.4.3 *Overview of Tables*, 2.4.6.1 *Direct Paragraph Formatting*).

| question | answer |
|---|---|
| is this paragraph in a table? | `sprmPFInTable (0x2416) != 0`, and/or `sprmPItap (0x6649) > 0` |
| how deep? | `itap` = `sprmPItap (0x6649)` ± `sprmPDtap (0x664A)`; 0 = not in a table |
| where does a cell end? | depth 1: the paragraph's text ends with `U+0007`. depth > 1: `U+000D` with `sprmPFInnerTableCell (0x244B) = 1` |
| where does a row end? | depth 1: `U+0007` with `sprmPFTtp (0x2417) = 1`. depth > 1: `U+000D` with `sprmPFInnerTtp (0x244C) = 1` |
| how many columns? | `TDefTableOperand.NumberOfColumns` on the **row-end (TTP) paragraph** |

Spec, verbatim:

> If the table depth is 1, the cell mark MUST be character Unicode 0x0007. If the table depth is
> greater than 1, the cell mark MUST be a paragraph mark (Unicode 0x000D) with sprmPFInnerTableCell
> applied with a value of 1.
> A table row has between 1 and 63 table cells … followed by a Table Terminating Paragraph mark
> (TTP mark, also called a row mark) …
> The properties of each row mark MUST define the cells for that table row.

A crucial simplification the spec grants you: `sprmPIstd` explicitly **preserves** `fTtp`,
`fInTable`, `itap`, `fInnerTableCell` and the table style when a style is applied. So *table
membership is always direct PAPX* — you never need the style sheet to find tables. (Lists are the
opposite; see §4.4.)

### 2.2 Where a table ends

Two adjacent rows of the same depth belong to the same table unless they differ in `sprmTIpgp`,
`sprmTIstd`, `sprmTFBidi`/`sprmTFBidi90`, or any of the position/wrapping sprms
(`sprmTPc, sprmTFNoAllowOverlap, sprmTDxaAbs, sprmTDyaAbs, sprmTDxaFromText, sprmTDyaFromText,
sprmTDxaFromTextRight, sprmTDyaFromTextBottom`). Otherwise the table ends at the first paragraph
with a smaller `itap`.

POI's practical form (`Range.getTable`) is just "extend while `isInTable() && itap >= level`" and it
ignores the split rule entirely — which is why two consecutive distinct tables get glued together
(this is a real, currently-open complaint on the Microsoft Q&A open-specs forum). For v1, matching
POI is acceptable; for correctness, compare at least `sprmTIstd` and the position sprms between
consecutive row marks.

### 2.3 `TDefTableOperand` byte layout ([MS-DOC] 2.9.315)

```
offset  size            field
 0      2               cb                (remainder length + 1)
 2      1               NumberOfColumns   (0..63)
 3      2*(ncols+1)     rgdxaCenter[]     i16 twips, non-decreasing
 3+2n+2 20*ncols        rgTc80[]          TC80
```

`rgdxaCenter[0]` is the logical-left edge of the table (relative to the left page margin — it can
be **negative**; `table.doc` has `-108`, which is Word's marker for "no indent"). `rgdxaCenter[i+1]`
is the right edge of cell *i*. Fewer `TC80`s than columns is legal → the rest get defaults.

**`TC80`** ([MS-DOC] 2.9.310) is exactly **20 bytes**:

```
 0  u16  tcgrf   (TCGRF)
 2  u16  wWidth
 4  4    brcTop     (Brc80MayBeNil)
 8  4    brcLeft
12  4    brcBottom
16  4    brcRight
```

**`TCGRF`** ([MS-DOC] 2.9.311), bit-packed LSB-first in that `u16`:

| bits | field | values |
|---|---|---|
| 0–1 | `horzMerge` | 0 = not merged; 1 = continuation of a horizontal merge (contents not rendered); 2 or 3 = **first** cell of a horizontal merge |
| 2–4 | `textFlow` | rotation |
| 5–6 | `vertMerge` (`VerticalMergeFlag`) | `0x00 fvmClear`; `0x01 fvmMerge` (continuation, must be empty); `0x03 fvmRestart` (first of the merged set) |
| 7–8 | `vertAlign` | |
| 9–11 | `ftsWidth` | unit for `wWidth` |
| 12 | `fFitText` | |
| 13 | `fNoWrap` | |
| 14 | `fHideMark` | |
| 15 | unused | |

This layout is confirmed independently by both reference implementations, which is worth knowing
because MS-DOC only groups the bits into `horzMerge`/`vertMerge` and does not name the individual
flags:

LibreOffice `sw/source/filter/ww8/ww8par2.cxx` (`WW8TabBandDesc::ReadNewShd`/TC read path):
```cpp
sal_uInt16 aBits1 = SVBT16ToUInt16( pTc->aBits1Ver8 );
pCurrentTC->bFirstMerged    = sal_uInt8( ( aBits1 & 0x0001 ) != 0 );
pCurrentTC->bMerged         = sal_uInt8( ( aBits1 & 0x0002 ) != 0 );
pCurrentTC->bVertical       = sal_uInt8( ( aBits1 & 0x0004 ) != 0 );
pCurrentTC->bBackward       = sal_uInt8( ( aBits1 & 0x0008 ) != 0 );
pCurrentTC->bRotateFont     = sal_uInt8( ( aBits1 & 0x0010 ) != 0 );
pCurrentTC->bVertMerge      = sal_uInt8( ( aBits1 & 0x0020 ) != 0 );
pCurrentTC->bVertRestart    = sal_uInt8( ( aBits1 & 0x0040 ) != 0 );
pCurrentTC->nVertAlign      = ( ( aBits1 & 0x0180 ) >> 7 );
```

POI `org/apache/poi/hwpf/model/types/TCAbstractType.java`:
```java
private static final BitField fFirstMerged = new BitField(0x0001);
private static final BitField fMerged      = new BitField(0x0002);
private static final BitField fVertical    = new BitField(0x0004);
private static final BitField fBackward    = new BitField(0x0008);
private static final BitField fRotateFont  = new BitField(0x0010);
private static final BitField fVertMerge   = new BitField(0x0020);
private static final BitField fVertRestart = new BitField(0x0040);
private static final BitField vertAlign    = new BitField(0x0180);
private static final BitField ftsWidth     = new BitField(0x0E00);
```

So `MS-DOC horzMerge` == `{fFirstMerged, fMerged}` and `MS-DOC vertMerge` ==
`{fVertMerge, fVertRestart}`; `fvmMerge (1)` = `fVertMerge && !fVertRestart`,
`fvmRestart (3)` = `fVertMerge && fVertRestart`.

### 2.4 Delta sprms that also change cell definitions

`sprmTDefTable` is the *initial* layout. A conformant reader must also apply, in document order,
the sprms that mutate it:

| sprm | operand | effect |
|---|---|---|
| `sprmTInsert (0x7621)` | `TInsertOperand{itcFirst:u8, ctc:u8, dxaCol:u16}` | insert cells |
| `sprmTDelete (0x5622)` | `ItcFirstLim{itcFirst:u8, itcLim:u8}` | delete cells |
| `sprmTDxaCol (0x7623)` | `TDxaColOperand` | change widths of a cell range |
| `sprmTMerge (0x5624)` | `ItcFirstLim` | **horizontally** merge a cell range |
| `sprmTSplit (0x5625)` | `ItcFirstLim` | undo a merge |
| `sprmTVertMerge (0xD62B)` | `VertMergeOperand{cb=2, itc:u8, vertMergeFlags:u8}` | set one cell's `VerticalMergeFlag` |

`VertMergeOperand.itc` counts **all** cells in the row, including merged continuations.

There is also a **Word 6/95 variant, `sprmTDefTable10` = `0xD606`** (LibreOffice
`sprmids.hxx`: `const sal_uInt16 LN_TDefTable10 = 0xd606;`). Its per-cell `TC` is the older
10-byte `WW8_TCellVer6`, in which only two merge bits exist and they are in a *single byte*:

```cpp
sal_uInt8 aBits1 = pTc->aBits1Ver6;
pCurrentTC->bFirstMerged = sal_uInt8( ( aBits1 & 0x01 ) != 0 );
pCurrentTC->bMerged      = sal_uInt8( ( aBits1 & 0x02 ) != 0 );
```
i.e. **no vertical merge at all in WW6**. If you ever extend the reader to `nFib < 0x00C1`, do not
reuse the 20-byte `TC80` parser.

The spec explicitly says producers *should* write both `sprmTDefTable` (for readers that ignore
`sprmPTableProps`) **and** `sprmTInsert` (for readers that honour it) — so for most files
`sprmTDefTable` alone is sufficient. **POI ignores `sprmTMerge` and `sprmTVertMerge` entirely**
(`TableSprmUncompressor`: `/**@todo handle table sprms from complex files*/ case 0x24: … case 0x2b: break;`),
which is a known blind spot. **LibreOffice ignores them too** — in `ww8scan.cxx` `sprmTMerge`
(0x5624), `sprmTSplit` (0x5625), `sprmTVertMerge` (0xD62B) and `sprmTVertAlign` (0xD62C) appear
only in the sprm *length* table (`InfoRow<NS_sprm::TVertMerge>()`) and are never dispatched by
`WW8TabBandDesc::ProcessSprmT*`. Both reference implementations therefore derive merges
**exclusively** from the `TC80` bits inside `sprmTDefTable`. Handling `0xD62B`/`0x5624` is cheap
and puts you ahead of both — but do not *rely* on them being present, because the files those two
readers were tuned against clearly do not need them.

---

## 3. Merged cells → `col_span` / `row_span`

### 3.1 Horizontal: it is geometry, not a flag

This is the finding that matters most, and my sample dumps prove it.

`table-merges.doc` (paragraph/row structure recovered exactly as described in §2.1):

```
row 0  TDefTable ncols=2  rgdxaCenter = [0,       6872,             9302]        cells: "A" "B"
row 1  TDefTable ncols=4  rgdxaCenter = [0, 1062, 5738,       8148, 9302]        cells: "C" "D" "E" "F"
row 2  TDefTable ncols=4  rgdxaCenter = [0, 1062, 5738,       8148, 9302]        cells: ""  "G" "H" "I\nJ"
row 3  TDefTable ncols=1  rgdxaCenter = [0,                         9302]        cells: "K"
```

Every `TC80.horzMerge` in this file is **0**. The horizontal merges are encoded purely as *wider
cells*. The union of all boundaries across all rows of the table is

```
{0, 1062, 5738, 6872, 8148, 9302}   →  5 grid columns
```

and each cell's `col_span` is the number of grid columns its `[left, right)` twip range covers.
Running that algorithm on the sample produces exactly the intended table:

```
row 0: 'A'(cs=3) | 'B'(cs=2)
row 1: 'C'(cs=1) | 'D'(cs=1) | 'E'(cs=2) | 'F'(cs=1)
row 2: '' (cs=1) | 'G'(cs=1) | 'H'(cs=2) | 'I\nJ'(cs=1)
row 3: 'K'(cs=5)
```

Note row 1's `"E"` spans 2 grid columns even though nothing was "merged" there — it is a genuine
consequence of row 0 introducing the `6872` boundary. That is correct HTML/IR semantics.

Both references do exactly this. POI, `converter/AbstractWordUtils.java`:

```java
static int[] buildTableCellEdgesArray( Table table ) {
    Set<Integer> edges = new TreeSet<>();
    for ( int r = 0; r < table.numRows(); r++ ) {
        TableRow tableRow = table.getRow( r );
        for ( int c = 0; c < tableRow.numCells(); c++ ) {
            TableCell tableCell = tableRow.getCell( c );
            edges.add(tableCell.getLeftEdge());
            edges.add(tableCell.getLeftEdge() + tableCell.getWidth());
        }
    }
    ...
}
```

and `converter/AbstractWordConverter.java`:

```java
protected int getNumberColumnsSpanned(int[] tableCellEdges,
    int currentEdgeIndex, TableCell tableCell) {
    int nextEdgeIndex = currentEdgeIndex;
    int colSpan = 0;
    int cellRightEdge = tableCell.getLeftEdge() + tableCell.getWidth();
    while (tableCellEdges[nextEdgeIndex] < cellRightEdge) {
        colSpan++;
        nextEdgeIndex++;
    }
    return colSpan;
}
```

**LibreOffice takes a different route to the same information, and it is worth knowing why you
should copy POI here and not LO.** Writer's table model is hierarchical (lines own boxes), so LO
never needs a colspan: `WW8TabDesc::CalcDefaults` computes `m_nMinLeft`/`m_nMaxRight` across all
bands, creates the table with `m_nDefaultSwCols = min(nSwCols)` columns ("because inserting cells is
cheaper than merging"), and then `AdjustNewBand()` calls `InsertCells(nSwCols - nDefaultSwCols)` to
top each row up — so rows genuinely keep different box counts and different widths. A pure
horizontal merge of three cells ends up as `[w = total, rowSpan = 1] [w = 0, rowSpan = -1]
[w = 0, rowSpan = -1]`, i.e. zero-width sibling boxes rather than a span. LO's *merge matching* is
still geometric (`FindMergeGroup` on `(nCenter[i], nWidth[i])` twips), but there is **no union grid
and no colspan anywhere in `ww8par2.cxx`**. Since `ir::TableCell` has `col_span`, POI's union-edge
model is the right one to port.

Two LO details are still directly reusable:

* **A cell "does not exist" iff `nCenter[i] == nCenter[i+1]`** (`CalcDefaults`, `bExist[]`).
  Word writes collapsed horizontal merges as degenerate zero-width boundaries; those must be
  dropped, and `nTransCell[]` maps the surviving WW column index to the output column (mapping
  *forward* to the next existing cell, except for a trailing invalid run which maps *backwards*).
* **Ragged left/right edges get up to two synthetic filler boxes per row** (`bLEmptyCol` /
  `bREmptyCol`, triggered when the gap to `m_nMinLeft`/`m_nMaxRight` is ≥ `MINLAY`) so every row
  spans the full table width. With a union grid you get this for free, but if you ever emit a
  strictly rectangular grid you will need the equivalent.

**Recommended `col_span` algorithm**

```
edges = sorted unique set of every rgdxaCenter value over every row of the table
for each row, for each cell i:
    l = rgdxaCenter[i], r = rgdxaCenter[i+1]
    col_span = count of k where edges[k] >= l and edges[k+1] <= r     // grid cols inside [l,r]
    if col_span == 0 { skip the cell entirely }                        // zero-width cell
```

Snap boundaries that differ by ≤ 3 twips onto one grid line before building `edges`. MS-DOC's own
`fVertMerge` description states the tolerance:

> Cells can only be merged vertically if their left and right boundaries are (nearly) identical
> (i.e. if corresponding entries in rgdxaCenter of the table rows differ by at most 3).

POI does **not** apply this tolerance when building the edge array (it uses exact `TreeSet<Integer>`
equality) and consequently invents spurious 1-twip columns on some real documents. Do better.

`TC80.horzMerge` should still be honoured as a *secondary* signal: if `horzMerge == 1` the cell is a
continuation whose contents are not rendered — drop it and fold its grid columns into the preceding
`horzMerge >= 2` cell. Files that set the flags without collapsing the geometry exist (that is why
the flags are there), but they are rare; the geometric path must be the primary one.

### 3.2 Vertical: a downward run of `fvmMerge` under an `fvmRestart`

Spec model ([MS-DOC] 2.4.3, `VerticalMergeFlag`):

> Cells can be vertically merged to create the appearance of a single cell spanning multiple rows.
> The cell mark characters for the merged cells MUST still appear in the file. The second and
> subsequent cells in the merged group MUST NOT contain any content other than their cell marks.

**Recommended `row_span` algorithm** (spec-faithful, matches both references):

```
for each row r, for each cell c with vertMerge == fvmRestart:
    row_span = 1
    for r2 in r+1 .. last_row:
        find the cell c2 in r2 whose [left,right] matches c's within ±3 twips
        if c2 exists and c2.vertMerge == fvmMerge:
            row_span += 1; mark c2 as consumed (emit nothing for it)
        else:
            break
```

Match by **twip geometry**, not by cell index. POI matches by cell ordinal
(`nextRow.getCell(currentColumnIndex)`) and therefore mis-spans whenever the rows have different
cell counts — which is precisely the situation in which vertical merges occur. LibreOffice matches
geometrically via `FindMergeGroup(nCenter[nCol], nWidth[nCol], /*bExact=*/true)`; its merge-group
open condition is:

```cpp
bool bMerge = false;
if ( rCell.bVertRestart && !rCell.bMerged )
    bMerge = true;
else if (rCell.bFirstMerged && m_pActBand->bExist[i])
{ ... }
...
const sal_uInt16 nRowSpan = groupIt->rowsCount();
```

Cells with `vertMerge == fvmMerge` must **not** be emitted at all (they are guaranteed empty).
LibreOffice's `SwWW8ImplReader::IsInvalidOrToBeMergedTabCell()` states the predicate exactly:

```cpp
return !IsValidCell(GetCurrentCol())
    || ( pCell && !pCell->bFirstMerged
                && ( pCell->bMerged || (pCell->bVertMerge && !pCell->bVertRestart) ) );
```

### 3.3 ⚠ `samples-doc/table-merges.doc` does NOT encode its vertical merge correctly

I decoded both row-2 and row-3 TTP PAPX byte-for-byte. They are **identical**, and both set
`TC80[0].tcgrf = 0x0060` — i.e. `fVertMerge | fVertRestart` = **`fvmRestart`**:

```
FKP pn=7 [0] fc=2074..2076   ...08d6 5c00 04 0000 2604 6a16 d41f 5624  6000 0000 ...
FKP pn=8 [0] fc=2094..2096   ...08d6 5c00 04 0000 2604 6a16 d41f 5624  6000 0000 ...
                                                                       ^^^^ tcgrf
```

A conformant file would put `0x0020` (`fvmMerge`) on the row-3 cell. As written, **a spec-correct
implementation — and POI, and LibreOffice — all produce `row_span = 1` plus a stray empty cell**:

* POI `getNumberRowsSpanned` walks down, sees `nextCell.isFirstVerticallyMerged()` and `break`s.
* LibreOffice `MergeCells` opens a *second* group at the same X, which locks the first
  (`p->m_bGroupLocked = true`), leaving both groups of size 1, and
  `if ((1 < groupIt->size()) …)` never fires.

So: **do not tune your algorithm until this sample renders `rowspan=2`.** You would be fitting to a
malformed file and would break correct ones. Either (a) accept the spec answer and regenerate the
sample with real Word, or (b) add an explicitly-labelled *lenient repair* pass, off by default in
the strict path:

> If a cell has `vertMerge != fvmClear`, is **completely empty**, and its `[left,right]` matches
> (±3 twips) a cell in the row above that is itself part of an open vertical-merge run, treat it as
> `fvmMerge` regardless of its `fVertRestart` bit.

That rule fixes `table-merges.doc` and is fairly safe (a genuine second merge-start directly under
another merge would have to be empty to be misclassified, and an empty merge-start renders
identically either way). Gate it behind a `lenient` flag and document why it exists.

---

## 4. Lists

### 4.1 The chain ([MS-DOC] 2.4.6.3 *Determining List Formatting of a Paragraph*)

```
paragraph PAPX
  ├─ sprmPIlfo (0x460B, i16)   →  iLfo
  └─ sprmPIlvl (0x260A, u8)    →  iLvl
                                    │
      Table[fcPlfLfo] ─ PlfLfo { lfoMac:u32, rgLfo:[LFO;lfoMac], rgLfoData:[LFOData;lfoMac] }
                                    │  rgLfo[iLfo-1].lsid
                                    ▼
      Table[fcPlfLst] ─ PlfLst { cLst:u16, rgLstf:[LSTF;cLst] }     ← find lstf.lsid == lfo.lsid
                                    │
                                    │  i = Σ over preceding LSTFs of (fSimpleList ? 1 : 9)
                                    │  i += iLvl
                                    ▼
      Table[fcPlfLst + lcbPlfLst] ─ rgLvl[i]  →  LVL { lvlf: LVLF(28B), grpprlPapx, grpprlChpx, xst }
```

`sprmPIlfo` operand ranges (verbatim from MS-DOC 2.6.2):

| value | meaning |
|---|---|
| `0x0000` | not in a list, list formatting removed |
| `0x0001 – 0x07FE` | 1-based index into `PlfLfo.rgLfo` |
| `0xF801` | **not in a list** |
| `0xF802 – 0xFFFF` | negation of a 1-based index; the paragraph's left/first-line indents must be preserved |

`sprmPIlvl`: `0x0–0x8` = zero-based level; **`0xC` = the list skips this paragraph** (no number).

### 4.2 Exact byte layouts

**`LSTF` — 28 bytes** ([MS-DOC] 2.9.131):

```
 0  i32  lsid            (unique; never 0xFFFFFFFF)
 4  i32  tplc
 8  18   rgistdPara[9]   i16 each; 0x0FFF = no linked style
26  u8   flags:  bit0 fSimpleList, bit1 unused1, bit2 fAutoNum,
                 bit3 unused2, bit4 fHybrid, bits5-7 reserved
27  u8   grfhic
```

⚠ **The flags byte is at offset 26, not 22.** `4 + 4 + 18 = 26`. I got this wrong on my first pass
and it silently produced `fSimpleList = 0` for `simple-list.doc`, which made me expect 9 `LVL`s
where there is 1, which desynchronised the whole `LVL` walk. This is a *very* easy off-by-four.

Verified on `simple-list.doc`:
`744609700f000904 ff0f 0000000000000000000000000000 01 00`
→ `lsid=0x70094674`, `rgistdPara[0]=0x0FFF`, `flags=0x01` → `fSimpleList=1`, `fHybrid=0`.
(Note `rgistdPara[1..8]` here are `0x0000`, not the spec-mandated `0x0FFF` — see pitfall P-L5.)

**`LVL`** ([MS-DOC] 2.9.129) = `LVLF` (28 B) ‖ `grpprlPapx` (`cbGrpprlPapx` B) ‖ `grpprlChpx`
(`cbGrpprlChpx` B) ‖ `Xst` (`cch:u16` then `cch` UTF-16 units). **Papx comes before Chpx** even
though `LVLF` lists `cbGrpprlChpx` first.

**`LVLF` — 28 bytes** ([MS-DOC] 2.9.130):

```
 0  i32  iStartAt
 4  u8   nfc            ← MSONFC
 5  u8   bits: 0-1 jc, 2 fLegal, 3 fNoRestart, 4 fIndentSav,
                5 fConverted, 6 unused1, 7 fTentative
 6  9    rgbxchNums[9]  1-based char offsets of placeholders in xst; zero-terminated
15  u8   ixchFollow     0 = tab, 1 = space, 2 = nothing
16  i32  dxaIndentSav
20  u32  unused2
24  u8   cbGrpprlChpx
25  u8   cbGrpprlPapx
26  u8   ilvlRestartLim
27  u8   grfhic
```

Verified on `simple-list.doc`:
`01000000 00 00 010000000000000000 00 00000000 00000000 00 10 00 00`
→ `iStartAt=1, nfc=0x00 (msonfcArabic), jc=0, ixchFollow=0 (tab), cbGrpprlPapx=16, cbGrpprlChpx=0`,
`rgbxchNums=[1,0,…]`, `Xst = cch:2, {U+0000, U+002E}` → number text = `"«level-0 number»."` → `"1."`.

**`LFO` — 16 bytes** ([MS-DOC] 2.9.126): `lsid:i32, unused1:i32, unused2:i32, clfolvl:u8,
ibstFltAutoNum:u8, grfhic:u8, unused3:u8`.
**`PlfLfo`**: `lfoMac:u32`, then `lfoMac` × `LFO`, then `lfoMac` × `LFOData`.
**`LFOData`**: `cp:i32` then `clfolvl` × `LFOLVL`; each `LFOLVL` is an `LFOLVLBase` optionally
followed by a **full inline `LVL`** when `fFormatting` is set — which *overrides* the `LVL` from
`PlfLst` for that level. Verified on `simple-list.doc`: `lfoMac=1`, `LFO[0].lsid=0x70094674`
(matching `LSTF[0]`), `clfolvl=0`.

### 4.3 Bulleted vs numbered: read `LVLF.nfc`

`nfc` is an `MSONFC` ([MS-OSHARED] 2.2.1.3):

| nfc | name | meaning |
|---|---|---|
| `0x00` | msonfcArabic | `1, 2, 3` → `ListStyle::Decimal`, `ordered = true` |
| `0x01` | msonfcUCRoman | `I, II, III` → `UpperRoman` |
| `0x02` | msonfcLCRoman | `i, ii, iii` → `LowerRoman` |
| `0x03` | msonfcUCLetter | `A, B, C` → `UpperAlpha` |
| `0x04` | msonfcLCLetter | `a, b, c` → `LowerAlpha` |
| `0x05` | msonfcOrdinal | `1st, 2nd` |
| `0x06/0x07` | Cardtext/Ordtext | `One`, `First` |
| `0x16` | msonfcArabicLZ | `01, 02, 03` |
| **`0x17`** | **msonfcBullet** | **bullet → `ordered = false`** |
| `0xFF` | msonfcNone | no numbering at all |

MS-DOC restates it in `LVLF.nfc`: *"If this is equal to 0xFF or 0x17, this level does not have a
number sequence… If this is equal to 0x17, the level uses bullets."*

LibreOffice `ww8par3.cxx::GetSvxNumTypeFromMSONFC` maps `case 23: nType = SVX_NUM_CHAR_SPECIAL;`
and `case 255: nType = SVX_NUM_NUMBER_NONE;`; POI's `HWPFList.getNumberFormat` carries the same
quote in its javadoc.

**Recommended IR mapping**

```rust
let ordered = !matches!(nfc, 0x17 | 0xFF);
let style = match nfc {
    0x00 | 0x16 => ListStyle::Decimal,
    0x01 => ListStyle::UpperRoman,  0x02 => ListStyle::LowerRoman,
    0x03 => ListStyle::UpperAlpha,  0x04 => ListStyle::LowerAlpha,
    0x17 => bullet_style_from_xst(xst, grpprl_chpx),  // Bullet / Square / Circle / Dash
    _    => ListStyle::Decimal,
};
let start_number = (nfc != 0x17 && nfc != 0xFF).then(|| lvlf.iStartAt as u32);
```

For bullets, `Xst.cch` MUST be 1 and that single character is the bullet glyph. Two gotchas:

* MS-DOC 2.4.6.3 Part 2 step 3: *"Let xchBullet be the 16-bit character at xstNumberText.rgtchar[0].
  If xchBullet & 0xF000 is nonzero, let xstNumberText.rgtchar[0] equal xchBullet & 0x0FFF."*
  Bullet glyphs live in the `U+F0xx` private-use area because they are Symbol/Wingdings codepoints.
* The **font** is in `LVL.grpprlChpx` (`sprmCRgFtc0` etc.). `U+F0B7` in Symbol is `•`;
  `U+F0A7` in Wingdings is `▪`. Map the common ones (`0xB7→•`, `0xA7→▪`, `0x6F→○`, `0xFC→✓`) and
  fall back to `•`. Do **not** emit a raw private-use codepoint into the IR.

### 4.4 ⚠ `ilfo` is frequently *not* in the direct PAPX

`sprmPIstd` preserves table-ness but **not** `ilfo`/`ilvl`. Real-world Word documents very often
put list paragraphs into the built-in `List Bullet` / `List Number` / `List Paragraph` styles, and
the `sprmPIlfo` then lives in the **style's** `grpprlPapx` in the `STSH`, not in the paragraph.
POI models this explicitly (`Paragraph.newParagraph` applies style properties, then the LVL's
`grpprlPapx`, then the direct PAPX, re-deriving the PAP twice); LibreOffice has a dedicated
`SwWW8ImplReader::StyleUsingLFO()` lookup.

`simple-list.doc` happens to carry `sprmPIlfo` directly (`460B:0100` on each of the three items,
`istd = 0`), so a direct-PAPX-only v1 passes the shipped corpus — but it will silently miss lists in
a large fraction of real files. Plan for `src/doc/stsh.rs` as phase 2, and note that the same
work gives you `istd → sti → Heading N`, which lets you delete the ALL-CAPS heading heuristic in
`convert_doc.rs` entirely.

### 4.5 Grouping paragraphs into IR `List` elements

Paragraphs sharing the same `iLfo` are the same list ([MS-DOC] 2.4.6.3): *"Paragraphs that share
the same iLfo property, and exist in a range of text that constitutes a Valid Selection, are
considered to be part of the same list. Paragraphs in a list do not need to be consecutive, and a
list can overlap with other lists."*

Neither reference builds a list *tree* — POI emits flat `<p>` with a literal `"1."` prefix and has
no `<ul>/<ol>/<li>` at all. So the IR mapping is yours to define. The pragmatic rule that matches
`ir::build_nested_list`'s contract:

```
accumulate a run of consecutive block-level paragraphs with the same non-zero iLfo
  → collect (ilvl, inline_content) pairs
  → ordered/style/start_number from LVL[run's minimum ilvl]
  → Element::List(ir::build_nested_list(ordered, &items, 0))
break the run on: iLfo change, iLfo == 0, entering/leaving a table cell, or a section break
```

Skip paragraphs with `ilvl == 0xC` (list skips them) — or rather, emit them as plain paragraphs
inside the item.

Verified end-to-end on `simple-list.doc`, my prototype produces:

```
PARA  ilfo=0 'This is a simple word document created using Word 97 – SR2. …'
PARA  ilfo=1 ilvl=0 'First item in list'
PARA  ilfo=1 ilvl=0 'Second item in list'
PARA  ilfo=1 ilvl=0 'Third item in list'
PARA  ilfo=0 'This is the last paragraph.'
```

→ `Paragraph`, `List{ordered: true, style: Decimal, start_number: 1, items: [3]}`, `Paragraph`.

**⚠ The task brief describes `simple-list.doc` as a *bulleted* list. It is not.** `LVLF.nfc = 0x00`
(`msonfcArabic`) and `Xst = "«0»."`, i.e. a **numbered** `1. 2. 3.` list. Any test asserting
`ordered == false` for this file is asserting the wrong thing — which is itself the strongest
possible argument for reading `nfc` instead of guessing from the rendered text.

---

## 5. Recommended implementation shape for `office_oxide`

### 5.1 New/changed modules

```
src/doc/
  fib.rs         + fc/lcb for PlcfBtePapx (0x0102/0x0106), PlfLst (0x2E2/0x2E6),
                   PlfLfo (0x2EA/0x2EE), Stshf (0x00A2/0x00A6)
  sprm.rs        NEW  Sprm { code: u16 } with ispmd/fSpec/sgc/spra accessors,
                        SprmIter<'a> honouring the 2-byte-cb exception for 0xD608/0xC615
  papx.rs        NEW  PlcBtePapx, PapxFkp, BxPap, PapxInFkp, GrpPrlAndIstd
                      -> pub fn paragraph_properties(word_doc, table, fib) -> Vec<ParaPapx>
                         ParaPapx { fc_start, fc_end, istd, sprms: Vec<(u16, Vec<u8>)> }
  para.rs        NEW  ParaProps { in_table, ttp, itap, inner_cell, inner_ttp,
                                   ilfo, ilvl, istd, jc, ... }  (fold sprms -> props)
  table.rs       NEW  TDefTable { ncols, centers: Vec<i16>, tcs: Vec<Tc80> }
                      RowDef, grid building, span computation
  list.rs        NEW  PlfLst/LSTF/LVL/LVLF/PlfLfo/LFO readers + resolve(ilfo, ilvl) -> &Lvl
  document.rs    ~    DocDocument gains `paragraphs: Vec<DocParagraph>` and
                      `blocks(): Vec<DocBlock>`; plain_text() built from it (back-compat)
src/convert_doc.rs   rewritten: DocBlock -> Element, drop the line heuristic
```

Suggested intermediate model (mirrors what `convert_docx.rs` already consumes):

```rust
pub enum DocBlock {
    Paragraph(DocParagraph),
    Table(DocTable),
}
pub struct DocTable { pub rows: Vec<DocRow>, pub column_edges: Vec<i32> }
pub struct DocRow   { pub cells: Vec<DocCell>, pub is_header: bool, pub cant_split: bool,
                      pub height_twips: Option<i32> }
pub struct DocCell  { pub blocks: Vec<DocBlock>,   // recursion handles nested tables
                      pub col_span: u32, pub row_span: u32,
                      pub left_twips: i32, pub right_twips: i32,
                      pub vert_merge: VertMerge, pub horz_merge: HorzMerge,
                      pub background: Option<[u8;3]>, pub border: Option<TableBorder> }
```

### 5.2 The pipeline, in order

1. Parse FIB, CLX/piece table (already done).
2. Build a **CP ↔ FC** map from the piece table (both directions). `PapxFkp.rgfc` is in FCs; you
   need CPs to slice text. For a compressed piece the FC of CP *c* is `fc/2 + (c - cpStart)`;
   for an uncompressed piece it is `fc + 2*(c - cpStart)`.
3. Walk `PlcBtePapx` → for each `PapxFkp` page → for each `k` emit
   `ParaPapx { fc: rgfc[k]..rgfc[k+1], istd, sprms }`. Sort by `fc`.
4. Map each to a CP range; **drop anything with `cp_start >= ccpText`** (see P-G3).
5. Fold sprms into `ParaProps`.
6. Group into blocks with a small state machine (§5.3).
7. Convert to IR.

### 5.3 Table grouping state machine (depth-aware)

```
cell_paras = []; row_cells = []; rows = []
for p in paragraphs:
    if p.itap == 0:
        flush any open table; emit Paragraph
    else if p.itap > current_depth:
        recurse (collect the whole itap>depth run, build the nested table,
                 attach it to the currently-open cell)
    else if p.is_row_end(depth):            // fTtp @1, or fInnerTtp @>1
        rowdef = parse TDefTable/TInsert/... from p's sgc==5 sprms
        rows.push(Row { cells: row_cells, def: rowdef }); row_cells = []; cell_paras = []
    else:
        cell_paras.push(p)
        if p.is_cell_end(depth):            // text ends U+0007 @1, or fInnerTableCell @>1
            row_cells.push(Cell { blocks: cell_paras }); cell_paras = []
close table when the next paragraph's itap < depth OR the row-split rule of §2.2 fires
```

Then per table: build `edges` (§3.1), assign `col_span`, run the vertical-merge pass (§3.2), drop
consumed continuation cells and `col_span == 0` cells, and emit
`ir::Element::Table` with `TableCell { content, col_span, row_span, .. }` and
`Table { column_widths_twips: edges.windows(2).map(|w| (w[1]-w[0]) as u32).collect(), .. }`.

Header rows: `sprmTTableHeader (0x3404) == 1` → `TableRow::repeat_as_header = true` (and
`is_header = true` for the first such run). `sprmTFCantSplit (0x3466)` → `allow_break = false`.
`sprmTDyaRowHeight (0x9407)` → `height_twips` (negative = exact height, use `abs`).

### 5.4 What to test

| test | expectation |
|---|---|
| `table.doc` | 1 table, 3×3, all `col_span=1 row_span=1`, cells `1 2 4 / 6 9 8 / 7 5 3`, `column_widths_twips = [3058,3058,3058]` |
| `table-merges.doc` | 1 table, 4 rows, 5 grid columns; spans `3,2 / 1,1,2,1 / 1,1,2,1 / 5`; cell `(2,3)` content `"I\nJ"` (two paragraphs in one cell) |
| `simple-list.doc` | `[Paragraph, List{ordered:true, style:Decimal, start_number:1, 3 items}, Paragraph]` |
| sprm walk | `sprmTDefTable` with `cb=0x0046` consumes 73 bytes; `sprmPChgTabs` uses the same `cb+3` rule |
| `PapxInFkp` | both `cb != 0` and `cb == 0` (`cb'`) forms |
| `BxPap.bOffset == 0` | yields default properties, not a parse error |
| fuzz/robustness | truncated `TDefTable` (fewer TC80s than columns), `ncols == 0`, `cpara > 0x1D`, `pn*512` past EOF, cyclic `sprmPHugePapx` |

---

## 6. Pitfalls

### Generic / plumbing

**P-G1 — `TDefTableOperand.cb` is a `u16`, unlike every other variable sprm.**
This one is nasty because it fails *quietly*. In `table.doc` the operand starts
`46 00 03 …` (`cb = 0x0046`, `NumberOfColumns = 3`). A reader that assumes a 1-byte length reads
`cb = 0x46 = 70` and then takes the **second byte of `cb`** as `NumberOfColumns` → **0 columns**.
The total sprm size still comes out right (`cb_low + 3 == cb_u16 + 3` for any `cb < 256`), so the
grpprl walk does *not* desynchronise and nothing looks broken — you just get an empty table. Once
the table has ≥ 12 columns (`cb = 22*ncols + 4 ≥ 256`) the size is wrong too and the walk *does*
desynchronise. Same 2-byte rule for `sprmPChgTabs (0xC615)`.

**P-G2 — `PapxFkp` lives in the WordDocument stream, `PlcBtePapx` in the Table stream.**
Mixing them up gives you random 512-byte pages that still parse (because `cpara` is one byte and
`rgfc` is monotonic-ish), producing plausible-looking nonsense.

**P-G3 — the paragraph stream covers all sub-documents, not just the body.**
`PlcBtePapx` describes main text, footnotes, headers, comments, endnotes and textboxes. CP ranges
are `[0, ccpText)`, then `ccpFtn`, `ccpHdd`, `ccpAtn`, `ccpEdn`, `ccpTxbx`, `ccpHdrTxbx` in that
order. `table.doc` has `ccpText = 23` but its `PlcBtePapx` describes 21 paragraphs running to
CP 57 — most of them header/footer material. Emit only `cp < ccpText` or you will inject header
junk (and phantom tables) into the body.

**P-G4 — mask `FcCompressed` with `0xC000_0000`, not `0x4000_0000`.**
Bit 30 is `fCompressed`, bit 31 is `r1` (reserved, must be ignored). The current
`piece_table.rs` uses `fc & !0x40000000`, which leaves bit 31 in the offset if a producer sets it.

**P-G5 — the compressed-piece byte→Unicode table is *not* CP1252.**
MS-DOC lists exactly 24 special bytes (`0x82–0x8C, 0x91–0x9C, 0x9F`). `0x80`, `0x8E`, `0x9E` are
**not** in it — the current `cp1252_to_char` maps them to `€ Ž ž`, which deviates from the spec.
Minor, but it is a real difference and worth a comment either way.

**P-G6 — `BxPap.bOffset == 0` is legal** and means "default paragraph properties". Skipping such
entries shifts your `rgfc[k] ↔ rgbx[k]` pairing and mis-attributes every later paragraph's props.

**P-G7 — `PapxInFkp` with `cb == 0` is the *common* case**, not an edge case. All PAPXs in all
three samples use it. Implement it first.

**P-G8 — `PnFkpPapx.pn` is 22 bits.** Mask with `0x003F_FFFF`.

**P-G9 — the last `rgfc` entry is a limit, not a paragraph.** `rgfc` has `cpara+1` entries.

**P-G10 — `sprmPHugePapx (0x6646)` / `sprmPTableProps (0x646B)` redirect into the Data stream.**
If present, `sprmPHugePapx` must be the first Prl and you must stop processing the rest of the
array and read the `PrcData` instead. Documents with large/complex tables use this routinely; a
reader that ignores it will see a table with no `sprmTDefTable` and give up.

### Tables

**P-T1 — `col_span` from cell index is wrong.** See §3.1. In `table-merges.doc` *every*
`horzMerge` bit is zero and the merges exist only as wider `rgdxaCenter` ranges. This is the one I
would bet an implementation gets wrong.

**P-T2 — the column grid must be unioned over *all* rows of the table, not per row.**
Row 1's `"E"` in `table-merges.doc` legitimately gets `col_span = 2` only because row 0 contributes
the `6872` boundary.

**P-T3 — snap boundaries within ±3 twips (LibreOffice uses ±4).** MS-DOC's own vertical-merge rule
states the ±3 tolerance; `WW8TabDesc::FindMergeGroup` hard-codes `const short nTolerance = 4;` and
widens each merge group's x-range by it on both sides before testing containment.
POI does not apply any tolerance when building its edge array (plain `TreeSet<Integer>` equality)
and consequently invents 1-twip phantom columns on real files.

**P-T4 — `rgdxaCenter[0]` can be negative.** `table.doc` starts at `-108` (Word's "no indent"
sentinel; LibreOffice comments on exactly this value in `CalcDefaults`). Use `i16`/`i32`, never
`u16`, and normalise the grid by subtracting `min(edges)` before computing widths.

**P-T5 — the TTP paragraph's own `U+0007` is not a cell.** The row-end mark is itself a `0x0007`.
If you treat "text ends with 0x0007" as a cell terminator without first checking `fTtp`, every row
gains a phantom trailing empty cell. POI patches this after the fact:
```java
// sometimes there are "fake" cells which we need to exclude
if ( !cells.isEmpty() && cells.size() != expectedCellsCount ) { … cells.remove(cells.size()-1); }
```
Check `fTtp`/`fInnerTtp` **before** the cell-mark test.

**P-T6 — a cell can contain many paragraphs.** `table-merges.doc` row 2 cell 3 is
`"I\r"` + `"J\x07"` — two paragraphs, one cell. Only the paragraph ending in `0x0007` (or carrying
`fInnerTableCell`) closes the cell. Treating every `\r` as a cell boundary produces 5 cells in a
4-column row.

**P-T7 — cells can be empty.** `table-merges.doc` row 2 cell 0 is a bare `0x0007`. Do not filter
empty cells; the column alignment depends on them.

**P-T8 — nested tables use `0x000D` + `sprmPFInnerTableCell` / `sprmPFInnerTtp`, not `0x0007`.**
A reader that only looks for `0x0007` merges nested rows into the parent row. Guard every
cell/row test on `itap == current_depth`.

**P-T9 — rows in one table need not have the same cell count.** MS-DOC: *"There is no requirement
that each row of a table have the same number of cells."* `table-merges.doc` has 2/4/4/1.
Any code path that assumes a rectangular `rows × cols` array is wrong.

**P-T10 — `itap` may be absent while `fInTable` is set.** Word 97 files predate `sprmPItap`.
Default to `itap = 1` when `fInTable` is true and no `sprmPItap`/`sprmPDtap` appeared.

**P-T11 — `sprmPDtap` is a *delta*.** Apply `sprmPItap` and `sprmPDtap` in document order within
the grpprl; the result must be non-negative.

**P-T12 — vertical-merge continuation cells must be dropped, not emitted empty.**
Emitting them adds a spurious column to the row.

**P-T13 — match vertical merges geometrically, not by cell ordinal.** POI's ordinal matching
(`nextRow.getCell(currentColumnIndex)`) misfires exactly when rows differ in cell count, which is
the normal case for merged tables.

**P-T14 — `col_span == 0` happens** (zero-width cells produced by `sprmTDelete`/degenerate rows).
Drop those cells entirely; both POI and LibreOffice do.

**P-T15 — two consecutive tables can look like one.** Without the §2.2 split rule (compare
`sprmTIstd`, `sprmTIpgp`, `sprmTFBidi`, and the eight position/wrapping sprms between adjacent row
marks) two back-to-back tables merge. POI has this bug today.

**P-T16 — fewer `TC80`s than `NumberOfColumns` is legal**, and truncated operands occur in the
wild (POI: *"Sometimes, the grpprl does not contain data at every offset. I have no idea why this
happens."*). Bounds-check every `TC80` read and substitute defaults.

**P-T17 — `NumberOfColumns` is capped at 63** and a row has 1..63 cells. Reject/clamp larger values
before allocating.

**P-T18 — trust the text, not `NumberOfColumns`, when they disagree.** POI overwrites `itcMac` with
the text-derived cell count and logs a warning. Do the same: the `0x0007` marks are the ground truth
for *content*; `rgdxaCenter` is the ground truth for *geometry*. Clamp the geometry lookup with
`get(i)` rather than indexing.

**P-T19 — a `U+0007` is only a cell mark if it sits at the very end of the previous PAPX range.**
LibreOffice checks this explicitly before calling `TabCellEnd()`:
```cpp
//The last paragraph of each cell is terminated by a special
//paragraph mark called a cell mark. ...
//So the 0x7 should be right at the end of the previous
//range to be a real cell-end.
if (pPap->nOrigStartPos == nPosCp+1 || pPap->nOrigStartPos == WW8_CP_MAX)
    TabCellEnd();
else
    bParaMark = true;
```
A stray `U+0007` in body text is a literal character, not a cell boundary.

**P-T20 — `sprmPFInnerTtp` sometimes arrives without `sprmPFInnerTableCell`.**
LibreOffice's `tdf#106799` workaround: *"We expect TTP marks to be also cell marks, but sometimes
sprmPFInnerTtp comes without sprmPFInnerTableCell"* — so a nested row mark must close the open
cell as well as the row. Do not require both flags.

**P-T21 — the PAPX chain can be cyclic; guard your row-end search.**
`SwWW8ImplReader::SearchRowEnd` keeps a seen-set of `(nStartPos, nEndPos)` bounds and bails with
*"SearchRowEnd, loop in paragraph property chain"*. If your row scan walks forward from a cell to
find its terminating TTP, bound the walk. Related: when no row end is found at the expected level,
LO leaves the nesting level unchanged rather than opening a table (`#i19667#`, *"Bad Table, remain
unchanged in level"*).

**P-T22 — `table-merges.doc` is itself non-conformant for vertical merges.** See §3.3. Do not
"fix" your algorithm against it.

### Lists

**P-L1 — `LSTF.flags` is at byte offset 26.** `4 (lsid) + 4 (tplc) + 18 (rgistdPara) = 26`.
Off-by-four here silently flips `fSimpleList`, which changes the expected `LVL` count from 1 to 9
and desynchronises the entire `LVL` array walk. I hit this myself while writing the prototype.

**P-L2 — the `LVL` array is *outside* `lcbPlfLst`.** It begins at `fcPlfLst + lcbPlfLst` and its
size is only discoverable by walking every `LVL` in order (each is variable-length). You cannot seek
to list *k*'s levels; you must walk `Σ (fSimpleList ? 1 : 9)` levels from the start.
Verified: `simple-list.doc` has `fcPlfLst = 0x160, lcbPlfLst = 30`, so the single `LVL` occupies
`0x17E..0x1B0` — and `0x1B0` is exactly `fcPlfLfo`.

**P-L3 — inside `LVL`, `grpprlPapx` precedes `grpprlChpx`** even though `LVLF` declares
`cbGrpprlChpx` before `cbGrpprlPapx`. Swapping them corrupts the `Xst` offset and therefore the
bullet character.

**P-L4 — `ilfo` is signed and the `0xF801`/`0xF802–0xFFFF` range is real.**
POI stores it via `getOperandShortSigned()` and then compares against `0xF801`, so those branches
*never fire* — `0xF801` arrives as `-2047`, `isInList()` returns `true`, and `HWPFList` throws.
Store `ilfo` as `u16`, treat `0` and `0xF801` as "not in a list", and decode
`0xF802..=0xFFFF` as `!ilfo` (one's complement → 1-based index) with "preserve indents" semantics.

**P-L5 — `rgistdPara` entries for unused levels are not always `0x0FFF`.**
`simple-list.doc` has `0x0FFF` for level 0 and `0x0000` for levels 1–8, where `0x0000` is a *valid*
istd (Normal). Only consult `rgistdPara[ilvl]` for levels the list actually defines
(`ilvl < (fSimpleList ? 1 : 9)`), and treat `0x0FFF` as "none".

**P-L6 — since Word 2000, "simple" lists are still written with 9 levels.**
LibreOffice's `#i1869#` comment:
> In word 2000 microsoft got rid of creating new "simple lists" with only 1 level, all new lists are
> created with 9 levels. … create a simple list in 2000 and open it in 97 and 97 will claim
> (correctly) that it is an outline list.
So `fSimpleList` tells you the `LVL` **count** (needed for the walk) but says nothing about how many
levels are used. Do not infer "flat list" from it.

**P-L7 — `LFOLVL` can carry an inline full `LVL` that overrides the `PlfLst` one.**
When `LFOLVLBase.fFormatting` is set, an entire `LVL` follows and takes precedence for that level
([MS-DOC] 2.4.6.3 Part 1 steps 6–7). Ignoring it gives the wrong bullet/number for
per-instance-customised lists. It also carries `iStartAt` overrides (restart numbering).

**P-L8 — bullet glyphs are Symbol/Wingdings private-use codepoints.** Apply
`if (ch & 0xF000) != 0 { ch &= 0x0FFF }` per the spec, then map through the `grpprlChpx` font.
Emitting `U+F0B7` into the IR produces a tofu box downstream.

**P-L9 — `nfc == 0x17` requires `Xst.cch == 1`, but files violate it.** POI warns and continues.
Do the same rather than rejecting the level.

**P-L10 — `ixchFollow == 2` means "nothing follows the number text".** POI's HTML converter
unconditionally strips one trailing character from the label, so with `ixchFollow == 2` it chops a
real character off the bullet text. If you ever render the label, branch on `ixchFollow`.

**P-L11 — the placeholder's format comes from the placeholder's *own* level, not the paragraph's.**
[MS-DOC] `LVL`: *"The level number that replaces a placeholder is formatted according to the
lvlf.nfc of the LVL structure that corresponds to the level that the placeholder specifies."*
POI uses the paragraph's level for every placeholder, so `1.a.iv`-style multilevel labels come out
wrong. Only relevant if you render number text; if you emit structured `List`/`ListItem` you dodge it.

**P-L12 — placeholders in `Xst` are code units `< 9`.** Anything `>= 9` is a literal. `rgbxchNums`
gives their 1-based indices and is zero-terminated. This is how you tell `"1."` from a literal `"1."`.

**P-L13 — `sprmPIlvl == 0x0C` means "the list skips this paragraph"**, not "level 12".

**P-L14 — a huge `iStartAt` with Roman numerals is a DoS vector.** POI added
`IOUtils.safelyAllocateCheck` in `getBulletText` because *"number-format can be roman-numbers, where
very large numbers would have very many 'M' and thus may cause memory to overload."*
`iStartAt` is spec-bounded to `0..=0x7FFF`; enforce it.

**P-L15 — list membership is by `iLfo`, and lists need not be contiguous.** Paragraphs of one list
can be interleaved with other content and with other lists. A naive "consecutive run" grouper is a
reasonable v1 but must break the run on `iLfo` change, not just on "not in a list".

**P-L16 — `simple-list.doc` is numbered, not bulleted.** Any expected-output fixture that says
otherwise is wrong (§4.5).

**P-L17 — `ilfo == 2047` is a trapdoor into the Word 6/95 numbering system.**
LibreOffice's `Read_LFOPosition`:
```cpp
if (m_nLFOPosition != 2047-1) // Normal ww8+ list behaviour
    RegisterNumFormat(m_nLFOPosition, m_nListLevel);
else if (m_xPlcxMan && m_xPlcxMan->HasParaSprm(NS_sprm::LN_PAnld).pSprm)
{
    /* #i8114# Horrific backwards compatible ww7- lists in ww8+ docs */
    Read_ANLevelNo(13 /*equiv ww7- sprm no*/, &m_nListLevel, 1);
}
```
That paragraph's numbering lives in `sprmPAnld` (`WW8_ANLD`: a 16-byte `WW8_ANLV` +
`fNumber1/fNumberAcross/fRestartHdn` + 32 bytes of prefix/suffix text) and `sprmPNLvlAnm`, not in
`PlfLst`. The `ANLV.nfc` table is a *different, smaller* enumeration (0 arabic, 1 upper roman,
2 lower roman, 3 upper letter, 4 lower letter, 5–7 all collapse to arabic) and **has no bullet
code at all** — bullets there are detected by the font's charset being `2` (Symbol/Wingdings).
If you do not implement the ANLD path, at minimum detect `ilfo == 2047` and emit plain paragraphs
rather than resolving `rgLfo[2046]` and producing garbage.
Also note LO's neighbouring branch: `ilfo <= 0` means *cancel* numbering, which is distinct from
"never had any" (it must not inherit list formatting from the style).

**P-L18 — `PlfLfo.rgLfoData` is where spec and practice diverge; be tolerant.**
[MS-DOC] says `rgLfoData` has one `LFOData` per LFO, each `cp:i32` followed by `clfolvl`
`LFOLVL`s. LibreOffice instead only enters the LFOData reader for LFOs with `clfolvl > 0`, and
before each override block it skips a 4-byte header **plus any run of `0xFFFFFFFF` dwords**:
```cpp
//2.2.2.0 skip inter-group of override header ?
//See #i25438# for why I moved this here, compare
//that original bugdoc's binary to what it looks like
//when resaved with word, i.e. there is always a
//4 byte header, there might be more than one if
//that header was 0xFFFFFFFF, e.g. #114412# ?
```
Implement the spec form (one `LFOData` per LFO — `simple-list.doc` has exactly that, with
`cp = 0xFFFFFFFF`), but bounds-check every read and abandon override parsing rather than
propagating a desync. Also: an `LFOLVL` whose flags byte is `0xFF` is disabled yet still consumes
its 8 bytes, and `LFOLVL.iStartAt` applies only when `fStartAt` is set — `fFormatting` alone does
*not* import the inline `LVL`'s `iStartAt`.

**P-L19 — `LVLF` bytes 16..23 are named differently by the two authorities.**
[MS-DOC] calls them `dxaIndentSav:i32` then `unused2:u32`. LibreOffice reads them as
`nV6DxaSpace:i32` then `nV6Indent:i32` and uses the second one as the list-tab position when the
level's `fWord6` bit (0x40 of the flags byte) is set. `LVLF` is 28 bytes either way, so the walk is
unaffected — but do not assume byte 20 is dead. Likewise LO reads flag bits `0x10/0x20/0x40` as
`fPrev/fPrevSpace/fWord6` where MS-DOC names `0x10 fIndentSav`, `0x20 fConverted`, `0x80 fTentative`.

**P-L20 — sanity-gate the list tables before parsing.** LibreOffice refuses outright when
`fcPlcfLst == fcPlfLfo || lcbPlcfLst < 2 || lcbPlfLfo < 2`. Cheap, and it kills a whole class of
garbage-in crashes.

**P-L21 — `LVLF.fNoRestart` (0x08) / `ilvlRestartLim` are read and then ignored by LibreOffice.**
If you want restart semantics right you are on your own; neither reference implements them.

---

## 7. Sources

**Primary specification — [MS-DOC], Word (.doc) Binary File Format** (all fetched and read):

* Retrieving Text (2.4.1) — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/01d5d8c4-cf9c-4ef9-80fd-439e763cfe01
* Determining Paragraph Boundaries (2.4.2) — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/30461a5b-e3ad-44cd-a3fe-038f86639b13
* Overview of Tables (2.4.3) — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/5b45f0e7-7760-4fdb-af88-0146de2feb4c
* Direct Paragraph Formatting (2.4.6.1) — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/61b635c3-2c44-4155-bf17-fec281b30c71
* Determining List Formatting of a Paragraph (2.4.6.3) — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/1a0bc623-211f-44f6-828b-51993f16e586
* FibRgFcLcb97 (2.5.5) — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/0c9df81f-98d0-454e-ad84-b612cd05b1a4
* Sprm (2.6.1 header) — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/099eb99c-a927-4caf-a80c-66254ea83d6a
* Paragraph Properties / paragraph sprm table (2.6.2) — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/484822ee-a9d9-4af4-8423-29fda67a6a58
* Table Properties / table sprm table (2.6.3) — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/b39a6648-501c-4361-8366-4f042f579469
* PlcBtePapx — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/76d3b8e1-337b-4812-a3f1-6b417ba6398d
* PnFkpPapx — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/6b3d10c0-0b95-4533-93fe-caef5c09679b
* PapxFkp — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/34aaeaf3-9578-41af-a3f5-c12f6f66bf1b
* BxPap — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/86df4678-ff4d-4877-b61a-6c621906973f
* PapxInFkp — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/580510b8-df7a-467e-a51c-0d71eb15c7cd
* FcCompressed — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/aa2e55a2-f4f2-4795-bab5-6d9d7a0ed249
* TDefTableOperand — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/de06ec41-a0ac-4046-9096-cdfaa0091ad9
* TC80 — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/9dd62a79-8c0b-4b11-99ee-05742ae7cf6d
* TCGRF — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/11bf5b1c-943f-421d-bbf3-39088cd1b8dd
* VerticalMergeFlag — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/1af35534-516e-4b58-986e-f2084bd6d56f
* VertMergeOperand — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/cf7489ac-7eee-404d-8844-8ebbd279b77d
* LSTF — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/6682e74f-365f-45e6-8b06-0e898bbf69a1
* LVL — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/da06b036-d4e8-4dcf-8759-f157b26c5aaf
* LVLF — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/500d1b05-cd15-46af-bc8f-03bc290d1bdd
* Example of a List (3.7) — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/39d71c1c-1ed0-4cbe-8aad-314d716f9649
* Example of a PlcBtePapx — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/0471d9a7-2475-44e9-a60d-70e2a9af1b24
* MSONFC ([MS-OSHARED] 2.2.1.3) — https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/57c8f7d4-1426-44e0-b58d-6fa1a5a81659
* "Detection of the end of the table" (Microsoft Open Specifications Support answer confirming that
  *all* the §2.2 table-split properties must be compared) —
  https://learn.microsoft.com/answers/a/2015508

**LibreOffice `sw/source/filter/ww8`** (sources downloaded and read directly at `master`).

⚠ `https://cgit.freedesktop.org/libreoffice/core/tree/...` is **dead (HTTP 404)**. The working
canonical browse URL is `https://git.libreoffice.org/core/+/refs/heads/master/sw/source/filter/ww8/<file>`;
raw text is `https://raw.githubusercontent.com/LibreOffice/core/master/sw/source/filter/ww8/<file>`.

* `ww8par2.cxx` — `WW8TabBandDesc::ReadDef` L1081, TC bit layout L1165‑1185, `#i25071#` text-flow
  post-pass L1190, `ProcessSprmTSetBRC` L1208, `ProcessSprmTDxaCol` L1314, `ProcessSprmTInsert` L1337,
  `ProcessSprmTDelete` L1504, `GetTableSprm` L1621, deferred-sprm application L1934, `sprmTDxaLeft`
  row shift L1968, `CalcDefaults` L2116‑2311 (`bExist`/`nTransCell` L2251), `CreateSwTable` L2408,
  `MergeCells` L2546‑2672, `FinishSwTable` rowspan application L2753‑2789, `FindMergeGroup` L2793‑2851
  (`nTolerance = 4`), `AdjustNewBand` L3136, `UpdateTableMergeGroup` L3281,
  `IsInvalidOrToBeMergedTabCell` L3523, `SearchRowEnd` L342, `WW8_ANLV`/ANLD nfc table L498 —
  https://git.libreoffice.org/core/+/refs/heads/master/sw/source/filter/ww8/ww8par2.cxx
* `ww8par3.cxx` — `GetSvxNumTypeFromMSONFC` L490‑643, `ReadLVL` L645‑1000 (LVLF fields L659‑696,
  number text L854, bullet/bitmap branch L858‑878, `rgbxchNums` → `%N%` L919‑954), `PlfLst` read
  L1134‑1250 (`#i1869#` simple-list comment L1197), `PlfLfo` read L1283‑1439 (`#i25438#`/`#114412#`
  header skip L1367), `GetNumRuleForActivation` L1520‑1621, `Read_ListLevel` L1895,
  `Read_LFOPosition` L1942‑2050 (`ilfo == 2047` / `#i8114#`) —
  https://git.libreoffice.org/core/+/refs/heads/master/sw/source/filter/ww8/ww8par3.cxx
* `ww8par.cxx` — `ProcessSpecial` / nesting-level detection L2752‑2914, `TabRowSprm` L2725,
  `0x07` cell-mark position check L3532‑3550, nested-table `0x0D` + magic-tables PLCF L3656‑3689 —
  https://git.libreoffice.org/core/+/refs/heads/master/sw/source/filter/ww8/ww8par.cxx
* `ww8par.hxx` — `WW8TabBandDesc` L1054‑1110 (`MAX_COL 64`, `nCenter[]`, `nWidth[]`, `bExist[]`,
  `nTransCell[]`), `enum ListLevel {nMinLevel=1, nMaxLevel=9}` L166 —
  https://git.libreoffice.org/core/+/refs/heads/master/sw/source/filter/ww8/ww8par.hxx
* `ww8par2.hxx` — `WW8SelBoxInfo` L153‑201 (merge group bucketed by owning `SwTableLine`) —
  https://git.libreoffice.org/core/+/refs/heads/master/sw/source/filter/ww8/ww8par2.hxx
* `ww8struc.hxx` — `WW8_TCell` L526‑556, `WW8_TCellVer6` (10 B) / `WW8_TCellVer8` (20 B) L558‑583,
  `WW8_ANLV`/`WW8_ANLD`/`WW8_OLST` L620‑672 —
  https://git.libreoffice.org/core/+/refs/heads/master/sw/source/filter/ww8/ww8struc.hxx
* `ww8scan.cxx` / `ww8scan.hxx` — sprm length table (`sprmTDefTable` registered `L_VAR2` L700‑703;
  `sprmTVertMerge` length-only L721), `enum SprmType {L_FIX, L_VAR, L_VAR2}`, `L_VAR2` decoding
  L8376‑8397, `DistanceToData` L8438 —
  https://git.libreoffice.org/core/+/refs/heads/master/sw/source/filter/ww8/ww8scan.cxx
* `ww8par6.cxx` — broken-WW6-list-indent workaround L4430‑4440 —
  https://git.libreoffice.org/core/+/refs/heads/master/sw/source/filter/ww8/ww8par6.cxx
* `sprmids.hxx` — verified constants: `PIlvl 0x260A`, `PIlfo 0x460B`, `PFInTable 0x2416`,
  `PFTtp 0x2417`, `PItap 0x6649`, `PDtap 0x664A`, `PFInnerTableCell 0x244B`, `PFInnerTtp 0x244C`,
  `TTableHeader 0x3404`, `TDefTable 0xD608`, `LN_TDefTable10 0xD606`, `TInsert 0x7621`,
  `TDelete 0x5622`, `TMerge 0x5624`, `TVertMerge 0xD62B`, `TFCantSplit 0x3466`,
  `TSetBrc80 0xD620` / `TSetBrc 0xD62F` —
  https://git.libreoffice.org/core/+/refs/heads/master/sw/source/filter/ww8/sprmids.hxx

⚠ Naming note: there is **no `sprmTTableDepth`**. `0x6649` is `sprmPItap`, a *paragraph* sprm
(`sgc == 1`), and its sibling is `sprmPDtap` `0x664A`.

**Apache POI HWPF** (`poi-scratchpad/src/main/java/org/apache/poi/hwpf/…`, browsable at
`https://github.com/apache/poi/blob/trunk/...`):

* `sprm/SprmOperation.java` — the `SPRM_LONG_TABLE` / `SPRM_LONG_PARAGRAPH` 2-byte-cb rule
* `sprm/ParagraphSprmUncompressor.java` — `0x16/0x17/0x49/0x4a/0x4b/0x4c/0x0a/0x0b` handling
* `sprm/TableSprmUncompressor.java` — `sprmTDefTable` parse; `case 0x24 … case 0x2b: break;`
  (i.e. `sprmTMerge`/`sprmTVertMerge` ignored)
* `usermodel/Paragraph.java`, `Range.java`, `Table.java`, `TableRow.java`, `TableCell.java`,
  `TableCellDescriptor.java`
* `model/types/TCAbstractType.java` — authoritative TC bit fields
* `model/types/TAPAbstractType.java` — `rgdxaCenter` semantics
* `model/ListTables.java`, `ListData.java`, `ListLevel.java`, `PlfLfo.java`, `LFOData.java`,
  `model/types/LSTFAbstractType.java`, `LVLFAbstractType.java`
* `converter/AbstractWordUtils.java` (`buildTableCellEdgesArray`, `getBulletText`),
  `converter/AbstractWordConverter.java` (`getNumberColumnsSpanned`, `getNumberRowsSpanned`),
  `converter/WordToHtmlConverter.java` (`processTable`), `converter/NumberFormatter.java`

**Local evidence.** All byte-level facts about `samples-doc/table.doc`,
`samples-doc/table-merges.doc` and `samples-doc/simple-list.doc` quoted above were produced by a
throw-away Python CFB + FIB + CLX + PlcBtePapx + PapxFkp + TDefTable + PlfLst/PlfLfo decoder written
during this investigation (kept in the session scratchpad, `/tmp/dcp/`). No project file was
modified.
