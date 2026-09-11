# SPRM / grpprl / PAPX-FKP / sprmTDefTable — spec-level reference

Primary source: **[MS-DOC] Word (.doc) Binary File Format**, Microsoft Open Specifications,
cross-checked against Apache POI HWPF (`poi-scratchpad`, trunk) and LibreOffice `sw/source/filter/ww8`,
and validated byte-for-byte against `samples-doc/table.doc`, `samples-doc/table-merges.doc`.

All multi-byte integers are **little-endian**. In [MS-DOC] bit diagrams the **first-listed field
occupies the low-order bits** (confirmed by `Sprm.ispmd == opcode & 0x01FF`).

---

## 1. Walking a `grpprl`

A `grpprl` is a flat, unaligned concatenation of `Prl` structures ([MS-DOC] 2.2.5.2):

```
Prl := Sprm (2 bytes)  ||  operand (spra-dependent, 0..n bytes)
```

`Sprm` ([MS-DOC] 2.2.5.1) is a 16-bit value:

| bits | mask     | field   | meaning |
|------|----------|---------|---------|
| 0–8  | `0x01FF` | `ispmd` | property index (with `fSpec`) |
| 9    | `0x0200` | `fSpec` | disambiguates `ispmd` |
| 10–12| `0x1C00` | `sgc`   | 1=PAP, 2=CHP, 3=PIC, 4=SEP, 5=TAP |
| 13–15| `0xE000` | `spra`  | operand-size class |

Walking algorithm:

```
o = 0
while o + 2 <= len(grpprl):
    opcode = LE16(grpprl, o)
    if opcode < 0x0800: break          # invalid / padding (LibreOffice wwSprmParser::GetSprmId)
    total = 2 + operand_size(opcode, grpprl, o + 2)
    yield (opcode, grpprl[o+2 .. o+total])
    o += total
```

### 1.1 `spra` → operand size ([MS-DOC] 2.2.5.1, table of `spra` values)

| `spra` | operand | operand bytes | total `Prl` bytes |
|--------|---------|---------------|-------------------|
| 0 | `ToggleOperand` ([MS-DOC] 2.9.331) | **1** | 3 |
| 1 | 1-byte value | **1** | 3 |
| 2 | 2-byte value | **2** | 4 |
| 3 | 4-byte value | **4** | 6 |
| 4 | 2-byte value | **2** | 4 |
| 5 | 2-byte value | **2** | 4 |
| 6 | **variable** — see §2 | `1 + cb` (general rule) | `3 + cb` |
| 7 | 3-byte value | **3** | 5 |

Note `spra` 2/4/5 are all 2 bytes and differ only in the semantics assigned by the property
tables (4 and 5 typically carry signed/`XAS`-style values). POI's `SprmOperation.initSize()`
encodes exactly this table (`case 0,1 -> 3; case 2,4,5 -> 4; case 3 -> 6; case 7 -> 5`).

Fallback for **unknown** opcodes: derive the size from `spra` alone. LibreOffice
(`wwSprmParser::GetSprmInfo`, `ww8scan.cxx`) says *"We can recover perfectly in this case"* and
switches on `nId >> 13`. Verified in `table.doc`: an undocumented `0xD5FF`
(`sgc=5, ispmd=0x1FF, spra=6`) walks correctly under the generic `spra=6` rule.

---

## 2. Opcodes whose operand size does **not** follow the `spra` rule  ← the crux

[MS-DOC] 2.2.5.1, `spra` value 6, verbatim:

> *"Operand is of variable length. The first byte of the operand indicates the size of the rest of
> the operand, **except in the cases of sprmTDefTable and sprmPChgTabs**."*

That sentence names exactly **two** exceptions. There is a third in practice (a legacy opcode that
[MS-DOC] does not document at all but that Word ≤ 95 files and some converters still emit).

| opcode | name | spec section | operand-length encoding | operand bytes | total `Prl` bytes |
|--------|------|--------------|-------------------------|---------------|-------------------|
| **`0xD608`** | `sprmTDefTable` | [MS-DOC] 2.6.3; operand 2.9.321 | **2-byte** LE `cb` at operand offset 0; `cb` = *bytes of the remainder of the structure, **incremented by 1*** | `2 + (cb − 1)` = **`cb + 1`** | **`cb + 3`** |
| **`0xC615`** | `sprmPChgTabs` | [MS-DOC] 2.6.2; operand 2.9.182 | 1-byte `cb`; **`cb == 255` is an escape** — the length is *computed from the payload*, and the sprm MAY be ignored | `cb != 255`: `1 + cb`; `cb == 255`: `1 + 2 + 4·cTabsDel + 3·cTabsAdd` | `cb != 255`: `cb + 3`; `cb == 255`: `3 + 2 + 4·cTabsDel + 3·cTabsAdd` |
| **`0xD606`** | `sprmTDefTable10` (legacy, *not* in [MS-DOC]; `NS_sprm::LN_TDefTable10` in LibreOffice `sprmids.hxx:70`) | — | same 2-byte `cb` rule as `0xD608` | `cb + 1` | `cb + 3` |

Everything else with `spra == 6` uses the plain rule: 1-byte `cb`, operand = `1 + cb`,
`Prl` = `cb + 3`. This includes `sprmPChgTabsPapx` (`0xC60D`), `sprmTDefTableShd*`,
`sprmTSetBrc`, `sprmTVertMerge`, `sprmTCellWidth`, `CSSAOperand` sprms, `BrcOperand` sprms, …

### 2.1 `sprmTDefTable` (`0xD608`) — the 2-byte length, exactly

`TDefTableOperand.cb` ([MS-DOC] 2.9.321): *"An unsigned integer that specifies the number of bytes
that are used by the remainder of this structure, **incremented by 1**."*

Therefore:

```
cb        = LE16(operand, 0)
rest_len  = cb - 1                  # NumberOfColumns + rgdxaCenter + rgTc80
operand_len = 2 + rest_len = cb + 1
sprm_len    = 2 + operand_len = cb + 3
```

Both reference implementations agree:

* POI `SprmOperation.initSize()`:
  ```java
  case 6:
      if ( sprm == SPRM_LONG_TABLE /*0xd608*/ || sprm == SPRM_LONG_PARAGRAPH /*0xc615*/ ) {
          int retVal = ( 0x0000ffff & LittleEndian.getShort( _grpprl, offset ) ) + 3;
          _gOffset += 2;
          return retVal;
      }
      return ( 0x000000ff & _grpprl[_gOffset++] ) + 3;
  ```
* LibreOffice `wwSprmParser::GetSprmTailLen()` (`ww8scan.cxx`), `L_VAR2` branch:
  > *"For sprmTDefTable and sprmTDefTable10, the length of the parameter plus 1 is recorded in the
  > two bytes beginning at offset (WW7-) 1 or (WW8+) 2"* — it reads `nCount = LE16(...)`, then
  > `--nCount`, then `GetSprmSize()` adds `1 + mnDelta + SprmDataOfs` (= 2 opcode bytes + 2 `cb`
  > bytes for WW8).

**Caveat:** POI also routes `0xC615` (`sprmPChgTabs`) through the 2-byte path. That contradicts
[MS-DOC] 2.9.182 (`PChgTabsOperand.cb` is 1 byte). LibreOffice handles `0xC615` with the
1-byte + 255-escape rule. **Follow [MS-DOC]/LibreOffice, not POI, for `0xC615`.**

### 2.2 `sprmPChgTabs` (`0xC615`) — the 255 escape, exactly

`PChgTabsOperand` ([MS-DOC] 2.9.182):

```
cb               : 1 byte   (MUST be >= 2 and <= 255)
PChgTabsDelClose : variable ([MS-DOC] 2.9.181)
PChgTabsAdd      : variable ([MS-DOC] 2.9.179)
```

> *"A value that is less than 255 specifies the size of the operand in bytes, not including `cb`.
> A value of 255 specifies that this instance of sprmPChgTabs MAY be ignored and that the size of
> the remainder of this operand, in bytes, is calculated by using the following formula:
> `cb = 4 × PChgTabsDelClose.cTabs + 3 × PChgTabsAdd.cTabs`"*

The published formula **omits the two `cTabs` count bytes**. The sub-structures are:

```
PChgTabsDelClose : cTabs(1) + rgdxaDel[cTabs]·2 + rgdxaClose[cTabs]·2   =  1 + 4·cTabs
PChgTabsAdd      : cTabs(1) + rgdxaAdd[cTabs]·2 + rgtbdAdd[cTabs]·1     =  1 + 3·cTabs
```

so the true remainder is `2 + 4·cTabsDel + 3·cTabsAdd`. LibreOffice implements the corrected form
(`ww8scan.cxx`, `GetSprmTailLen`, `case 23: case 0xC615:`):

```cpp
if( pSprm[1 + mnDelta] != 255 )
    nL = pSprm[1 + mnDelta] + aSprm.nLen;      // aSprm.nLen == 0 for L_VAR
else
{
    sal_uInt8 nDel = pSprm[2 + mnDelta];              // PChgTabsDelClose.cTabs
    sal_uInt8 nIns = pSprm[3 + mnDelta + 4 * nDel];   // PChgTabsAdd.cTabs
    nL = 2 + 4 * nDel + 3 * nIns;
}
```
(`mnDelta` = 1 for WW8, 0 for WW6/7 one-byte sprm ids.) Total `Prl` = `nL + 3` for WW8.

Reading order for the escape case: `cb` at +0 (== 255), `cTabsDel` at +1,
`rgdxaDel`/`rgdxaClose` occupy +2 .. +1+4·cTabsDel, `cTabsAdd` at +2+4·cTabsDel.

There is **no** escape for `sprmPChgTabsPapx` (`0xC60D`): [MS-DOC] 2.9.183 says `cb` is simply
*"the size of the operand in bytes, not including `cb`"*, `2 ≤ cb ≤ 255`, with 255 meaning literally 255.

### 2.3 Maximum representable operand length

| case | max operand |
|------|-------------|
| generic `spra=6` (1-byte `cb`) | `cb` = 255 → **256 operand bytes**, 258-byte `Prl` |
| `sprmPChgTabs` non-escape | `cb` = 254 → 255 operand bytes (255 is the escape) |
| `sprmPChgTabs` escape | `2 + 4·64 + 3·64 = 450` remainder → 451 operand bytes (`cTabs ≤ 64`, 2.9.179/2.9.181) |
| `sprmTDefTable` (2-byte `cb`) | theoretical `cb` = 0xFFFF → 65536 operand bytes; **practically bounded by 63 columns**: `1 + 2·64 + 20·63 = 1389` → `cb = 1390`, operand 1391, `Prl` 1393 |

A `grpprl` inside a `PapxInFkp` cannot exceed ~510 bytes (§6), so wide tables force the row
properties into the Data stream via `sprmPTableProps` (`0x646B`) or `sprmPHugePapx` (`0x6646`),
each a 4-byte offset to a `PrcData` ([MS-DOC] 2.9.210). See [MS-DOC] 3.6 for a worked example
where the whole row `grpprl` lives behind `sprmPTableProps`.

---

## 3. `sprmTDefTable` (`0xD608`) — full operand layout

`sprmTDefTable`: `ispmd = 0x08`, `fSpec = 1`, `sgc = 5`, `spra = 6` → `0xD608` ([MS-DOC] 2.6.3).
Operand = `TDefTableOperand` ([MS-DOC] 2.9.321).

| offset (within operand) | size | field | meaning |
|---|---|---|---|
| `0x00` | 2 | `cb` | bytes of the remainder **+ 1**. `operand_len = cb + 1`, `sprm_len = cb + 3` |
| `0x02` | 1 | `NumberOfColumns` (`itcMac`) | number of cells in this row. `0 ≤ n ≤ 63` |
| `0x03` | `2·(n+1)` | `rgdxaCenter[n+1]` | array of `XAS` (signed 16-bit twips, [MS-DOC] 2.9.349). **Cell boundaries**, non-decreasing. `[0]` = logical-left edge of the table relative to the left page margin (commonly `-108` = half of the default 0.08" gap); `[i+1]` = logical-right edge of cell `i`. Interior edges are the midpoint of inter-cell spacing. |
| `0x03 + 2·(n+1)` | `20·m` | `rgTc80[m]` | array of `TC80` ([MS-DOC] 2.9.313), **20 bytes each**. `m` may be **less than** `n` — trailing columns then get default formatting; `m > n` → excess ignored. Compute `m = (cb − 1 − 1 − 2·(n+1)) / 20`. |

Cell width of cell `i` = `rgdxaCenter[i+1] − rgdxaCenter[i]` (twips).

### 3.1 `TC80` (20 bytes) — [MS-DOC] 2.9.313

| offset | size | field |
|---|---|---|
| `+0` | 2 | `tcgrf` — `TCGRF` ([MS-DOC] 2.9.317) |
| `+2` | 2 | `wWidth` — preferred width, unit given by `tcgrf.ftsWidth` |
| `+4` | 4 | `brcTop` — `Brc80MayBeNil` ([MS-DOC] 2.9.18) |
| `+8` | 4 | `brcLeft` |
| `+12`| 4 | `brcBottom` |
| `+16`| 4 | `brcRight` |

(Word 6/95 `TC` is **10** bytes — LibreOffice `ReadDef()` branches on `bVer67` and computes
`nFileCols = nLen / (bVer67 ? 10 : 20)`.)

### 3.2 `TCGRF` bit layout — [MS-DOC] 2.9.317 (low bits first)

| mask | bits | field | values |
|---|---|---|---|
| `0x0003` | 0–1 | **`horzMerge`** | `0` not merged; `1` continuation cell (contributes layout region, contents not rendered); `2`/`3` **first** cell of a horizontally merged set |
| `0x001C` | 2–4 | `textFlow` | `TextFlow` enum |
| `0x0060` | 5–6 | **`vertMerge`** | `VerticalMergeFlag` ([MS-DOC] 2.9.342): `0` = `fvmClear`, `1` = `fvmMerge` (continuation, MUST be empty), `3` = `fvmRestart` (top cell of the merge) |
| `0x0180` | 7–8 | `vertAlign` | `VerticalAlign` enum |
| `0x0E00` | 9–11 | `ftsWidth` | `Fts` ([MS-DOC] 2.9.101); `0` = `ftsNil`, `3` = `ftsDxa`, … |
| `0x1000` | 12 | `fFitText` | |
| `0x2000` | 13 | `fNoWrap` | |
| `0x4000` | 14 | `fHideMark` | |
| `0x8000` | 15 | `fUnused` | MUST be ignored |

Both reference implementations decode the same bits with the older Word-97 field names
(LibreOffice `ww8par2.cxx` `WW8TabBandDesc::ReadDef`, POI `TCAbstractType`):

| bit mask | POI / LibreOffice name | [MS-DOC] equivalent |
|---|---|---|
| `0x0001` | `fFirstMerged` / `bFirstMerged` | `horzMerge` bit 0 |
| `0x0002` | `fMerged` / `bMerged` | `horzMerge` bit 1 |
| `0x0004` | `fVertical` | `textFlow` bit 0 |
| `0x0008` | `fBackward` | `textFlow` bit 1 |
| `0x0010` | `fRotateFont` | `textFlow` bit 2 |
| `0x0020` | `fVertMerge` / `bVertMerge` | `vertMerge` bit 0 |
| `0x0040` | `fVertRestart` / `bVertRestart` | `vertMerge` bit 1 |
| `0x0180` | `vertAlign` | `vertAlign` |
| `0x0E00` | `ftsWidth` | `ftsWidth` |

⚠ The two naming schemes disagree on `horzMerge`. The legacy Word-97 documentation that POI and
LibreOffice inherited says bit0 = *"first cell of a merged range"* and bit1 = *"merged with the
preceding cell"*; [MS-DOC] 2.9.317 says value `1` = continuation and `2`/`3` = first. Note that
`vertMerge` is unambiguously `1` = continuation / `3` = restart, and `horzMerge` is described in
[MS-DOC] with exactly parallel wording — so treat `horzMerge ∈ {2,3}` as "start" and `horzMerge == 1`
as "continuation", and be tolerant of producers that set both bits on the start cell.

### 3.3 Verified against `samples-doc/`

`table-merges.doc` (nFib `0x0101`), row 0, sprm bytes
`08 D6 | 30 00 | 02 | 00 00 D8 1A 56 24 | <TC80 ×2>`:

```
cb              = 0x0030 = 48        -> operand_len = 49, sprm_len = 51   (measured: 51 ✔)
NumberOfColumns = 2
rgdxaCenter     = [0, 6872, 9302]    (3 XAS = 6 bytes)
rgTc80          = 2 × 20 = 40 bytes  (1 + 6 + 40 = 47 = cb - 1 ✔)
  TC80[0] tcgrf=0x0000 wWidth=0
  TC80[1] tcgrf=0x0000 wWidth=0
```

All four `sprmTDefTable` instances in `table-merges.doc` (`cb` = 0x30, 0x5C, 0x5C, 0x1A) and all
three in `table.doc` (`cb` = 0x46) consume **exactly** `cb + 1` operand bytes with zero slack.
`table.doc` rows: `NumberOfColumns = 3`, `rgdxaCenter = [-108, 2950, 6008, 9066]`.

Neither sample uses `sprmTMerge`/`sprmTVertMerge`; horizontal merging is expressed purely through
differing `rgdxaCenter` arrays, and vertical merging through `TC80.tcgrf.vertMerge` (`0x0060`
observed on cell 0 of the two 4-column rows).

---

## 4. Other table sprms that change the cell array ([MS-DOC] 2.6.3)

| sprm | operand | effect |
|---|---|---|
| `sprmTInsert` `0x7621` (spra 3, **4 bytes**) | `TInsertOperand` 2.9.324: `itcFirst`(1) `ctc`(1) `dxaCol`(2) | insert `ctc` default cells of width `dxaCol` at index `itcFirst`; extends `rgdxaCenter`/`rgTc80`. A row MUST have at least one `sprmTInsert` **or** `sprmTDefTable`. |
| `sprmTDelete` `0x5622` (spra 2) | `ItcFirstLim` 2.9.123: `itcFirst`(1) `itcLim`(1) | delete cells `[itcFirst, itcLim)`; shifts `rgdxaCenter`/`rgTc80` down |
| `sprmTDxaCol` `0x7623` (spra 3) | `TDxaColOperand` 2.9.322 | change width of a cell range (shifts subsequent edges) |
| `sprmTMerge` `0x5624` (spra 2) | `ItcFirstLim` | mark `[itcFirst, itcLim)` horizontally merged (sets `horzMerge` bits) |
| `sprmTSplit` `0x5625` (spra 2) | `ItcFirstLim` | undo `sprmTMerge` |
| `sprmTVertMerge` `0xD62B` (spra 6) | `VertMergeOperand` 2.9.343: `cb`(1, MUST be 2), `itc`(1), `vertMergeFlags`(1) | set `VerticalMergeFlag` on one cell |

**Order matters**: `Prl` elements are applied left-to-right ([MS-DOC] 2.4.6 *Applying Properties*),
so `sprmTDefTable` must be evaluated before any later `sprmTInsert`/`sprmTDelete`/`sprmTMerge`
in the same `grpprl`.

---

## 5. Deriving horizontal and vertical spans

### 5.1 Horizontal span — the unified column-edge array

Word does **not** normally record a "colspan". A horizontally merged cell is simply a single cell
whose `rgdxaCenter` interval covers several of the columns used by other rows. The canonical
routine is Apache POI's `AbstractWordUtils.buildTableCellEdgesArray(Table)`
(`poi-scratchpad/.../hwpf/converter/AbstractWordUtils.java`):

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
    Integer[] sorted = edges.toArray(new Integer[0]);
    int[] result = new int[sorted.length];
    for ( int i = 0; i < sorted.length; i++ ) result[i] = sorted[i];
    return result;
}
```

i.e. **collect every left edge and every right edge of every cell of every row into a sorted,
de-duplicated set** — that set is the table's column grid. (`TableCell.getLeftEdge()` is
`rgdxaCenter[i]` and `getWidth()` is `rgdxaCenter[i+1] − rgdxaCenter[i]`; see
`TableRow.initCells()`.)

Column span, `AbstractWordConverter.getNumberColumnsSpanned`:

```java
protected int getNumberColumnsSpanned(int[] tableCellEdges, int currentEdgeIndex, TableCell tableCell) {
    int nextEdgeIndex = currentEdgeIndex;
    int colSpan = 0;
    int cellRightEdge = tableCell.getLeftEdge() + tableCell.getWidth();
    while (tableCellEdges[nextEdgeIndex] < cellRightEdge) { colSpan++; nextEdgeIndex++; }
    return colSpan;
}
```

The caller keeps a running `currentEdgeIndex`, starting at 0 for each row and advancing by the
returned span. Equivalently: `colSpan = index_of(right_edge) − index_of(left_edge)` in the grid.

Recommended algorithm:

```
grid = sorted(set(e for row in rows for e in row.rgdxaCenter))
for row in rows:
    for i in range(row.n):
        colspan = grid.index(row.rgdxaCenter[i+1]) - grid.index(row.rgdxaCenter[i])
    # a leading gap (row.rgdxaCenter[0] > grid[0]) is a "gridBefore"/w:gridBefore;
    # a trailing gap is "gridAfter"
```

**Tolerance.** Real files jitter edges by a twip or three between rows. POI's `TCAbstractType`
documents *"Cells can only be merged vertically if their left and right boundaries are (nearly)
identical (i.e. if corresponding entries in `rgdxaCenter` of the table rows differ by at most 3)"*,
and LibreOffice's `WW8TabDesc::FindMergeGroup` uses `const short nTolerance = 4;`. POI's raw
`TreeSet` does **not** snap, which is a known source of spurious 1-twip columns. **Snap edges that
are within ~3–4 twips of an existing grid entry before inserting.**
Also drop degenerate cells: LibreOffice marks cell `i` non-existent when
`nCenter[i] >= nCenter[i+1]` (`bExist[i] = false`, `ww8par2.cxx`).

Validated on `samples-doc/table-merges.doc`:

```
row rgdxaCenter arrays:
  row0 (2 cells): [0, 6872, 9302]
  row1 (4 cells): [0, 1062, 5738, 8148, 9302]
  row2 (4 cells): [0, 1062, 5738, 8148, 9302]
  row3 (1 cell) : [0, 9302]
unified grid   : [0, 1062, 5738, 6872, 8148, 9302]      (5 columns)
colspans       : row0=[3,2]  row1=[1,1,2,1]  row2=[1,1,2,1]  row3=[5]
```

Document text confirms the layout: `A␇B␇␇  C␇D␇E␇F␇␇  ␇G␇H␇I¶J␇␇  K␇␇¶`
(4 rows of 2/4/4/1 cells, matching `NumberOfColumns`).

If `horzMerge` bits *are* present (produced by `sprmTMerge`), also collapse runs:
a cell with `horzMerge ∈ {2,3}` starts a run, following cells with `horzMerge == 1` join it.
LibreOffice does exactly this (`ww8par2.cxx` ~line 2640: walk forward while
`pTCs[i2].bMerged && !pTCs[i2].bFirstMerged`, summing `nWidth[i2]`).

### 5.2 Vertical span

`vertMerge` (`TCGRF` bits 5–6) or `sprmTVertMerge` (`0xD62B`) per cell:

* `fvmRestart` (3) — top cell of the vertical merge; holds the content.
* `fvmMerge` (1) — continuation; MUST be empty and is not rendered.
* `fvmClear` (0) — not merged.

POI's `AbstractWordConverter.getNumberRowsSpanned` is the reference walk:

```java
if (!tableCell.isFirstVerticallyMerged()) return 1;      // fVertRestart (0x40)
int count = 1;
for (int r1 = currentRowIndex + 1; r1 < numRows; r1++) {
    TableRow nextRow = table.getRow(r1);
    if (currentColumnIndex >= nextRow.numCells()) break;
    // skip rows that have no real cells at this grid position
    ... (uses getNumberColumnsSpanned to advance a per-row edge cursor) ...
    TableCell nextCell = nextRow.getCell(currentColumnIndex);
    if (!nextCell.isVerticallyMerged() || nextCell.isFirstVerticallyMerged()) break;
    count++;
}
return count;
```

Recommended algorithm (grid-index based rather than cell-index based, which is more robust when
rows have different cell counts):

```
for each row r, for each cell c with vertMerge == fvmRestart:
    left = grid_index(cell.left); right = grid_index(cell.right)
    rowspan = 1
    for r2 in r+1 .. last_row:
        find the cell in r2 whose [left,right) grid range matches (within tolerance)
        if none, or its vertMerge != fvmMerge: break
        rowspan += 1; mark that cell as "covered"
cells marked covered are emitted as nothing (or as a continuation marker)
```

**Producer robustness note.** In `samples-doc/table-merges.doc` **both** rows of the vertical merge
carry `vertMerge == 3` (`tcgrf = 0x0060`), i.e. the producer marked the continuation cell as
`fvmRestart` too, even though its cell is empty. A tolerant implementation should therefore also
accept "cell is empty **and** `vertMerge != fvmClear` **and** the geometrically identical cell in
the row above is part of a merge" as a continuation. LibreOffice guards similarly
(`if ( rCell.bVertRestart && !rCell.bMerged ) bMerge = true;`).

---

## 6. PAPX FKP page layout

### 6.1 Locating the pages

`FibRgFcLcb97.fcPlcfBtePapx` / `lcbPlcfBtePapx` ([MS-DOC] 2.5.6) — pair index **13** (0-based),
i.e. byte offset **104 (0x68)** into `FibRgFcLcb97`; for `nFib = 0x00C1` that is FIB offset `0x102`.
It points into the Table stream (`0Table` or `1Table` per `FibBase.fWhichTblStm`, bit 9 of the
flags word at FIB offset 10).

`PlcfBtePapx` ([MS-DOC] 2.8.6): `n+1` 4-byte `FC`s followed by `n` × `PnFkpPapx`
([MS-DOC] 2.9.207) — a 4-byte value whose **low 22 bits** are `pn`; the FKP page lives at
**WordDocument stream offset `pn × 512`**.

### 6.2 `PapxFkp` — [MS-DOC] 2.9.174 (exactly 512 bytes)

```
offset 0                          : rgfc[cpara + 1]   — 4-byte FCs, ascending, no duplicates
offset 4*(cpara+1)                : rgbx[cpara]       — BxPap, 13 bytes each
... free space, then the PapxInFkp blobs (2-byte aligned) ...
offset 511                        : cpara             — 1 byte
```

* `rgfc[k]` = start FC of paragraph/row-mark `k`; `rgfc[cpara]` = end of the last one.
* `cpara` MUST be `1 ≤ cpara ≤ 0x1D` (29) — larger would overflow 512 bytes.
  (`4·30 + 13·29 + 1 = 498 ≤ 512`.)

### 6.3 `BxPap` — [MS-DOC] 2.9.23 (13 bytes)

```
+0  : bOffset (1 byte)
+1  : reserved (12 bytes)  — version-specific paragraph-height info; SHOULD be 0, SHOULD be ignored
```

`bOffset == 0` ⇒ **no** `PapxInFkp`; the paragraph has default properties ([MS-DOC] 2.6.2).
Otherwise the `PapxInFkp` starts at **`bOffset × 2`** bytes from the start of the FKP page
([MS-DOC] 2.4.6.1 step 2: `of + 2 × BxPap.bOffset`).

### 6.4 `PapxInFkp` — [MS-DOC] 2.9.175 — **the length-byte quirk**

```
+0 : cb (1 byte)
+1 : grpprlInPapx (variable)
```

> *"`cb` … specifies the size of the `grpprlInPapx`. If this value is **not 0**, the
> `grpprlInPapx` is **`2×cb − 1`** bytes long. If this value is **0**, the size is specified by the
> first byte of `grpprlInPapx`."*
> *"If `cb` is 0, the first byte of `grpprlInPapx` (call it `cb'`) … `cb'` MUST be at least 1.
> After `cb'`, there are **`2 × cb'`** more bytes in `grpprlInPapx`. The bytes after `cb'` form a
> `GrpPrlAndIstd`."*

So:

```
cb = page[po]
if cb != 0:  grpprlAndIstd = page[po+1 .. po+1 + (2*cb - 1)]     # length 2*cb - 1
else:        cbp = page[po+1]
             grpprlAndIstd = page[po+2 .. po+2 + 2*cbp]          # length 2*cbp
```

Note the asymmetry: the non-zero form yields an **odd** length (`2·cb − 1`), the escape form an
**even** length (`2·cb'`). Getting this wrong shifts the whole `grpprl` by one byte and every
subsequent sprm walk fails.

`GrpPrlAndIstd` ([MS-DOC] 2.9.114):

```
+0 : istd  (2 bytes)  — style index
+2 : grpprl (variable) — array of Prl, MUST be a whole number of Prl
```

Row properties (`sgc == 5`) live in the `PapxInFkp` of the **table terminating paragraph mark**
(the `\x07` whose PAPX carries `sprmPFTtp` `0x2417` = 1); per-cell text is delimited by the other
`\x07` cell marks. `sprmPFInTable` `0x2416` and `sprmPItap` `0x6649` give in-table flag and depth
([MS-DOC] 2.4.3 *Overview of Tables*).

### 6.5 Verified against `samples-doc/`

```
table.doc        nFib=0x00C1  table stream=0Table  PlcfBtePapx fc=2090 lcb=28  pn=[8,9,10]
                   FKP pn=8  cpara=8 : entry[1] bOffset=247 -> po=494 cb=6  -> size 2*6-1 = 11
                   FKP pn=8  entry[4] bOffset=97  -> po=194 cb=0  -> cb'=137 -> size 274
table-merges.doc nFib=0x0101  table stream=1Table  PlcfBtePapx fc=786  lcb=28  pn=[6,7,8]
                   FKP pn=6  cpara=7 : entry[2] bOffset=158 -> po=316 cb=0 -> size 160 (row PAPX)
```

Both the `2·cb − 1` and the `cb == 0 → 2·cb'` forms occur in these two files, and every
resulting `grpprl` walks to its exact end with the rules in §1–§2.

---

## 7. Quick-reference decision table for an implementer

```rust
fn sprm_total_len(grpprl: &[u8], o: usize) -> Option<usize> {
    let op = u16::from_le_bytes([*grpprl.get(o)?, *grpprl.get(o + 1)?]);
    if op < 0x0800 { return None; }                 // invalid / padding
    let spra = (op >> 13) & 7;
    Some(match spra {
        0 | 1 => 3,
        2 | 4 | 5 => 4,
        3 => 6,
        7 => 5,
        6 => match op {
            // EXCEPTION 1 & 3: two-byte cb, "remainder + 1"
            0xD608 | 0xD606 => {
                let cb = u16::from_le_bytes([*grpprl.get(o + 2)?, *grpprl.get(o + 3)?]) as usize;
                cb + 3
            }
            // EXCEPTION 2: one-byte cb with a 255 escape
            0xC615 => {
                let cb = *grpprl.get(o + 2)? as usize;
                if cb != 255 { cb + 3 } else {
                    let n_del = *grpprl.get(o + 3)? as usize;
                    let n_ins = *grpprl.get(o + 4 + 4 * n_del)? as usize;
                    3 + 2 + 4 * n_del + 3 * n_ins
                }
            }
            _ => *grpprl.get(o + 2)? as usize + 3,
        },
        _ => unreachable!(),
    })
}
```

Defensive rules worth keeping: clamp every computed size to the remaining buffer (LibreOffice:
`nSize = std::min(nSize, nLen)` and warns *"sprm longer than remaining bytes, doc or parser is
wrong"*); stop on `total == 0`; and require `MinSprmLen()` (3 bytes for WW8) remaining.

---

## Sources

### [MS-DOC] (Microsoft Open Specifications, `learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/`)

| section | title | URL |
|---|---|---|
| 2.2.5.1 | Sprm (bit layout + `spra` table + the exception sentence) | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/099eb99c-a927-4caf-a80c-66254ea83d6a |
| 2.2.5.2 | Prl | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/4eabffa2-b8b6-444c-9a92-3291ab5035ef |
| 2.4.2 | Determining Paragraph Boundaries | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/30461a5b-e3ad-44cd-a3fe-038f86639b13 |
| 2.4.3 | Overview of Tables | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/5b45f0e7-7760-4fdb-af88-0146de2feb4c |
| 2.4.6 | Applying Properties | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/4e918665-c4da-41d8-aed5-615c2e96c216 |
| 2.4.6.1 | Direct Paragraph Formatting (`of + 2×BxPap.bOffset`) | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/61b635c3-2c44-4155-bf17-fec281b30c71 |
| 2.5.6 | FibRgFcLcb97 | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/0c9df81f-98d0-454e-ad84-b612cd05b1a4 |
| 2.6 | Single Property Modifiers | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/4fae38be-4993-47d2-b82c-8f32e4ab9ff0 |
| 2.6.2 | Paragraph Properties (`sprmPChgTabs` 0xC615, `sprmPChgTabsPapx` 0xC60D, `sprmPHugePapx` 0x6646, `sprmPTableProps` 0x646B) | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/484822ee-a9d9-4af4-8423-29fda67a6a58 |
| 2.6.3 | Table Properties (`sprmTDefTable` 0xD608 and the full TAP sprm table) | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/b39a6648-501c-4361-8366-4f042f579469 |
| 2.8.6 | PlcBtePapx | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/76d3b8e1-337b-4812-a3f1-6b417ba6398d |
| 2.9.18 | Brc80MayBeNil | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/8458edbd-c81c-4ec7-b5ff-c99c50575301 |
| 2.9.23 | BxPap | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/86df4678-ff4d-4877-b61a-6c621906973f |
| 2.9.101 | Fts | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/930b1e39-1132-4f68-9cb0-eb1ebdc39c5e |
| 2.9.114 | GrpPrlAndIstd | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/bd96f2aa-1318-4066-9723-4db035ef412b |
| 2.9.123 | ItcFirstLim | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/61341a2c-b8da-4e52-ba62-b2b0d5efadc4 |
| 2.9.174 | PapxFkp | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/34aaeaf3-9578-41af-a3f5-c12f6f66bf1b |
| 2.9.175 | PapxInFkp (`2×cb − 1` / `cb == 0 → 2×cb'`) | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/580510b8-df7a-467e-a51c-0d71eb15c7cd |
| 2.9.179 | PChgTabsAdd | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/691c20a1-915b-4403-9813-1f0c96ed6f8b |
| 2.9.180 | PChgTabsDel | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/539894c9-8bad-4b8a-9ffb-8c969029d7d4 |
| 2.9.181 | PChgTabsDelClose | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/908cede7-b294-471a-afe8-add61f2078e7 |
| 2.9.182 | PChgTabsOperand (**the 255 escape**) | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/0fa864c2-4660-402f-b726-3e9895748a49 |
| 2.9.183 | PChgTabsPapxOperand (no escape) | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/6ff24fa9-40d8-4b54-93ab-d66a88e9ba5b |
| 2.9.207 | PnFkpPapx | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/6b3d10c0-0b95-4533-93fe-caef5c09679b |
| 2.9.210 | PrcData | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/473fd992-c824-4655-8880-3186bd432f80 |
| 2.9.313 | TC80 | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/9dd62a79-8c0b-4b11-99ee-05742ae7cf6d |
| 2.9.317 | TCGRF | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/11bf5b1c-943f-421d-bbf3-39088cd1b8dd |
| 2.9.321 | TDefTableOperand | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/de06ec41-a0ac-4046-9096-cdfaa0091ad9 |
| 2.9.322 | TDxaColOperand | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/fa2550bf-647e-4cfd-b468-090c6850a213 |
| 2.9.324 | TInsertOperand | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/082d49fd-8f96-4fd9-ac6e-61bd7c00b90d |
| 2.9.342 | VerticalMergeFlag | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/1af35534-516e-4b58-986e-f2084bd6d56f |
| 2.9.343 | VertMergeOperand | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/cf7489ac-7eee-404d-8844-8ebbd279b77d |
| 2.9.349 | XAS | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/7e431354-56f4-4dfb-95df-b8b4409f8039 |
| 3.5 | Example of a PlcBtePapx | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/0471d9a7-2475-44e9-a60d-70e2a9af1b24 |
| 3.6 | Example of Table Row Properties (worked `spra` byte counts) | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/efb41f4c-a920-42d5-860f-850266e0b55b |

### Apache POI HWPF (trunk, `poi-scratchpad/src/main/java/org/apache/poi/hwpf/`)

* `sprm/SprmOperation.java` — `initSize()`, `getOperand()`; `SPRM_LONG_TABLE = 0xd608`,
  `SPRM_LONG_PARAGRAPH = 0xc615`.
  https://raw.githubusercontent.com/apache/poi/trunk/poi-scratchpad/src/main/java/org/apache/poi/hwpf/sprm/SprmOperation.java
* `sprm/TableSprmUncompressor.java` — `case 0x08` (`sprmTDefTable` unpack: `itcMac`,
  `rgdxaCenter[itcMac+1]`, `rgtc[itcMac]`, 20 bytes per TC).
  https://raw.githubusercontent.com/apache/poi/trunk/poi-scratchpad/src/main/java/org/apache/poi/hwpf/sprm/TableSprmUncompressor.java
* `model/types/TCAbstractType.java` — TC bit masks.
  https://raw.githubusercontent.com/apache/poi/trunk/poi-scratchpad/src/main/java/org/apache/poi/hwpf/model/types/TCAbstractType.java
* `usermodel/TableRow.java`, `usermodel/TableCell.java` — `initCells()`, `getLeftEdge()`,
  `isVerticallyMerged()`, `isFirstVerticallyMerged()`.
  https://raw.githubusercontent.com/apache/poi/trunk/poi-scratchpad/src/main/java/org/apache/poi/hwpf/usermodel/TableRow.java
* `converter/AbstractWordUtils.java` — **`buildTableCellEdgesArray()`**.
  https://raw.githubusercontent.com/apache/poi/trunk/poi-scratchpad/src/main/java/org/apache/poi/hwpf/converter/AbstractWordUtils.java
* `converter/AbstractWordConverter.java` — `getNumberColumnsSpanned()`, `getNumberRowsSpanned()`.
  https://raw.githubusercontent.com/apache/poi/trunk/poi-scratchpad/src/main/java/org/apache/poi/hwpf/converter/AbstractWordConverter.java

### LibreOffice ww8 filter (master)

* `sw/source/filter/ww8/ww8scan.cxx` — `wwSprmParser::GetSprmTailLen/GetSprmSize/DistanceToData/GetSprmInfo`,
  `L_FIX / L_VAR / L_VAR2`, the `0xC615` 255 escape, the unknown-sprm `nId >> 13` fallback.
  https://raw.githubusercontent.com/LibreOffice/core/master/sw/source/filter/ww8/ww8scan.cxx
* `sw/source/filter/ww8/ww8par2.cxx` — `WW8TabBandDesc::ReadDef()` (TC bit decode, 10 vs 20 byte TC),
  `bExist` degenerate-cell rule, `WW8TabDesc::FindMergeGroup()` (`nTolerance = 4`), merge-group build.
  https://raw.githubusercontent.com/LibreOffice/core/master/sw/source/filter/ww8/ww8par2.cxx
* `sw/source/filter/ww8/sprmids.hxx` — `LN_TDefTable = 0xd608`, `LN_TDefTable10 = 0xd606`,
  `enum class SPRA` with `spraLen<>` specialisations.
  https://raw.githubusercontent.com/LibreOffice/core/master/sw/source/filter/ww8/sprmids.hxx

### Empirical validation

`samples-doc/table.doc`, `samples-doc/table-merges.doc` — parsed from raw bytes with a standalone
Python CFB + FIB + PlcfBtePapx + PapxFkp + sprm walker (scratchpad only; nothing written into the
repo). Every `sprmTDefTable` operand consumed exactly `cb + 1` bytes; both `PapxInFkp` length forms
exercised; unified-edge colspans reproduced the documents' visible table structure.
