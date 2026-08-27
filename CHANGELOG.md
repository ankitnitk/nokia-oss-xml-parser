# Changelog — Nokia OSS XML Parser

---

## 4G Tool — "Site Name" column in LNCEL Details — August 2026

### Added — Site Name parsed from cellName, flagged when it differs from LNBTS Name
A cell's `cellName` (once renamed to the sector-encoded convention) is `<site>_L<band>_<sector>` — but `<site>` is not always the same text as the LNBTS's own `name`, e.g. LNBTS `NAIROBI_PARKSIDE_TOWERS` hosting cells named `NAIROBI_PARKSIDE_INDOOR_L21_A` etc. The new `Site Name` column (right after `LNBTS Name`) surfaces that parsed site prefix, and is highlighted a neutral light blue whenever it differs from `LNBTS Name` — not red/yellow, since this isn't necessarily a mistake, just an LNBTS worth noticing has two site "parts". Tested against the real 27-Aug dumps (both a fresh raw XML→xlsx re-parse and the live scheduled `4G.270826.xlsb`): 718 of 31,843 cells nationally flagged.

## Version 6.5.2 — August 2026  (exe rebuild, same `oss_xml_to_xlsx_v6.5.py`)

### Rebuilt exe to bundle the updated `hw_tool`
No changes to the core parser script — same pattern as V6.1.2 and V6.5.1 (a same-script exe rebuild for a bundled-subtool change). Picks up the "Combined RMOD" / "Combined BBMOD" columns (previous entry below) into the exe's HW report generation.

## HW Tool — "Combined RMOD" / "Combined BBMOD" columns — August 2026

### Added — per-site combined module summary columns
Both `Site wise (All)` and `Site wise (Working)` sheets now have two trailing columns, `Combined RMOD` and `Combined BBMOD`, that collapse each site's non-zero RMOD / BBMOD unit counts into one summary string, e.g. `1*AHDA+2*ARDA+3*AHEGC`. Descriptive inventoryUnitType labels (e.g. `ABIA AirScale Capacity`, `BB Extension Outdoor Sub-Module FBBA`) are reduced to just the trailing module code (`ABIA`, `FBBA`) for the summary; plain codes and `UNKNOWN` pass through unchanged.

## Version 6.5.1 — August 2026  (exe rebuild, same `oss_xml_to_xlsx_v6.5.py`)

### Rebuilt exe to bundle the updated `4g_tool`
No changes to the core parser script — same pattern as V6.1.2 (a same-script exe rebuild for a bundled-subtool change, that time `3g_tool`). This rebuild picks up everything from the four `4g_tool` entries below (Sector ID / Number of cells in Sector / CA Relation Audit columns, the CAREL Correction sheet, and the duplicate-cellName fix + red highlight) into the standalone exe's interactive "Generate 4G Summary" feature. `DumpWatcher2` already had all of this since it runs `4g_tool` from source — this rebuild is only needed for people running the exe directly.

## 4G Tool — highlight duplicate cellName in LNCEL Details — August 2026

### Added — red highlight on duplicate `LNCEL name` cells
Cells sharing an identical sector-encoded `cellName` with another cell in the same sector (see previous entry) now also get their `LNCEL name` cell highlighted red directly in the LNCEL Details sheet, not just flagged in the `CA Relation Audit` text — makes the underlying naming problem visible without needing to read the audit column.

## 4G Tool — CAREL Correction: fix nonsensical same-band proposals — August 2026

### Fixed — duplicate cellName produced a self-referential "Create" row
Two cells can end up with the identical sector-encoded `cellName` (a real data bug — found `NAIROBI_ARCHIVES_L8_A` assigned to both LNCEL 1 and LNCEL 10, the latter an indoor/IBS cell apparently mis-named after its co-sited outdoor cell). The grouping logic treated both as legitimate sector-mates and proposed a nonsensical CAREL relation between them (band L8_A relating to itself).

CA relations are inherently cross-band — two cells sharing the same band_tag within a "sector" are never a legitimate Missing/Create/Wrong pair between each other, regardless of direction. Fixed in both the LNCEL Details audit and the CAREL Correction sheet: same-band_tag pairs are excluded from Missing/Create, and any *existing* same-band relation is now also caught as Wrong/Delete. Instead, the LNCEL Details `CA Relation Audit` column now reports `Duplicate cellName: -> <cell>` so the actual naming problem stays visible rather than silently disappearing. On the real dump this removed 18 bogus `Create` rows (437 → 419).

## 4G Tool — CAREL Correction sheet — August 2026

### Added — new "CAREL Correction" sheet: concrete delete/create action list
Turns the LNCEL Details "CA Relation Audit" findings into an actionable list of exact CAREL MO changes, one row per relation:

- **Source**: `MRBTS | LNBTS | LNCEL | CAREL | cellName`
- **Target**: `Target lcrId | Target lnBtsId | Target cellName`
- **Remarks**: `Delete` (an existing relation pointing outside its cell's sector) or `Create` (a same-sector band pair missing a relation in one direction).

For `Create` rows, the new CAREL instance ID is the next sequential ID unused by that source cell's *existing* CAREL relations (e.g. IDs 1/2/3 in use → next new relation uses 4) — deleted IDs are never reused for a creation on the same cell. Only covers cells whose `cellName` is already renamed to the sector-encoded convention (see below) — there's no sector to check an un-renamed cell against.

Also fixes a subtlety in the create-side logic: a sector pair only needs ONE creation for whichever direction is actually missing — if e.g. `SITE_L8_A → SITE_L21_A` already exists but the reverse doesn't, only `SITE_L21_A → SITE_L8_A` is proposed, not a duplicate of the already-correct direction.

## 4G Tool — Sector CA-relation audit — August 2026

### Added — "Sector ID", "Number of cells in Sector", "CA Relation Audit" columns in LNCEL Details
Cells are progressively being renamed on the network from the auto-generated sequential `cellName` ("SITE_1", "SITE_2", ...) to a sector-encoded form ("SITE_L8_A", "SITE_L26_C1_B", ...) that spells out the LTE band and sector letter. The LNCEL Details sheet now surfaces this migration:

- **`Sector ID`** — the parsed sector letter (`A`/`B`/`C`/...) if a cell's `cellName` already follows the sector-encoded convention, blank if it's still the old auto-generated form (rename pending).
- **`Number of cells in Sector`** — for renamed cells, how many other cells (across bands) share the same site + sector letter (e.g. `SITE_L8_A` + `SITE_L18_A` are both sector A → count 2). Indoor/outdoor and other differently-prefixed deployments at the "same" site are correctly kept separate since the literal site-name prefix differs (`SITE_TOWERS` vs `SITE_INDOOR`). Blank for cells still pending rename.
- **`CA Relation Audit`** — reads the new `CAREL` MO (carrier-aggregation SCell relations, `lcrId` = target LNCEL ID within the same LNBTS) and validates a strict same-sector mesh: every pair of bands within a sector must have a **mutual** CAREL relation to each other, and *only* to each other. Reports `OK`, `Missing: <bands>` (a same-sector pair with no relation, or only one-directional), `Wrong: -> <cells>` (a relation pointing outside the sector), or both. Blank for cells pending rename, since there's no sector to check against yet.

`4g_tool/network.py` now also loads the `CAREL` sheet (already included in the parser's 4G class filter, no dump re-export needed) and resolves each relation's `lcrId` directly to the target LNCEL's Dist_Name.

**Verified** against a real Kenya 4G dump (31,824 LNCEL cells, 14,028 renamed): spot-checked NDARAGWA, NAIROBI_PARKSIDE (TOWERS vs INDOOR correctly separated), and KISUMU_SAURIMOYO (multi-carrier L26_C1/C2, plus an asymmetric-relation case where one direction of a sector pair already existed) — all matched manually-derived expected results, including the raw CAREL records underlying the audit.

---

## Version 6.5 — August 2026  (`oss_xml_to_xlsx_v6.5.py`)

### Fixed — crash on numeric-looking text values (e.g. hardware serial numbers)
V6.4 crashed with `OverflowError: cannot convert float infinity to integer` while parsing a real Kenya 4G dump. Root cause: `try_numeric()` (and `parse_dist_name()`) call `float(s)` on every parameter value to detect numerics. Some values are alphanumeric **text** that nonetheless happens to be valid float *syntax* — e.g. a hardware `serialNumber` like `"1834E95902"` parses as mantissa `1834` with exponent `95902`. `float()` doesn't raise on that; it silently overflows to `+inf`. The very next step, `int(f)`, then raises `OverflowError` on infinity — an exception type the surrounding `except (ValueError, TypeError)` never caught, so one bad field crashed the entire worker process and aborted the whole parse.

Both `try_numeric()` and `parse_dist_name()` now catch `OverflowError` alongside `ValueError`/`TypeError` and keep the original text in that case, instead of crashing. (An earlier fix pre-checked `math.isinf()`/`isnan()` on every value instead — functionally identical, but added a measurable ~15–25% per-call cost to the hot path per microbenchmark; folding the check into the existing `except` clause keeps V6.4's original zero-cost fast path since Python `try` blocks cost nothing unless an exception actually fires.) Verified against the dump that triggered the bug (`4G DUMP_2108.xml.gz`, 62 MB gz / ~858 MB decompressed): parses clean, 123 MB output.

---

## Version 6.4 — July 2026  (`oss_xml_to_xlsx_v6.4.py`)

### Improved — Streaming Parse (2× faster parse, 3.6× less RAM)
V6.3 still materialised the **entire decompressed XML (~3.4 GB per file) as one string** before any parsing began: full gunzip → full decode → scan → pickle blocks to workers. With two files parsed in parallel threads, peak process-tree RSS reached **~14 GB**.

V6.4 replaces this with a **streaming pipeline** (`parse_xml_stream`):
- The gzip stream is decompressed and scanned in **64 MB chunks**; the full document is never decoded nor held in memory.
- MO blocks are carved from each chunk **as raw bytes** (same open-tag slicing as V6.3, class filter inline) and batched to worker processes **while decompression continues** — gunzip/scan time overlaps worker parse time instead of preceding it.
- Workers decode each kept block individually; in-flight batches are bounded, so peak RAM is O(chunk + batches), not O(document).
- The worker pool is created lazily on the first full batch — small files parse inline with zero process-spawn overhead.
- ZIP containers fall back to in-memory streams (rare and typically small); `.xml.gz` / `.xml` stream straight from disk.

**Measured on a real two-file LTE dump (173 + 178 MB gz → ~7 GB XML), same machine, same session:**

| | V6.3 | V6.4 |
|---|---|---|
| Parse phase | 100 s | **51 s** |
| Total wall | 148 s | **90 s** |
| Peak process-tree RSS | 14.4 GB | **4.0 GB** |
| Output | 165,288,521 B | byte-identical (all data-sheet CRCs equal) |

---

## 4G Tool — HO-trigger "low" is now a cell-level check — July 2026

### Changed — "low" HO-trigger mismatch flagged only when ALL relations are low
Previously a relation was flagged "low" (red) individually whenever `threshold3InterFreq < threshold2InterFreq − 2`. In practice a single low relation isn't a problem — if any other frequency relation is reachable, the cell still has a handover escape route.

Now the "low" flag is **cell-level**: a cell (and all its InterFreq HO Check rows) is flagged red only when **every** relation has `t3 < t2 − 2`. Cells with a mix of low and non-low relations are no longer flagged. The comparison base is `t2` (measurement start, `threshold2InterFreq`), keeping the original 2 dB tolerance. Row reason text reads e.g. `YES (t3 12 dB below start)`; the `HO Thr Issue` column in LNCEL Details shows `All low: <freqs>`. "High" (`t3 > t2`, yellow) is unchanged.

On a live ~31k-cell dump this narrowed the low flag from per-relation noise to **~3,760 genuinely all-low cells (12%)**.

---

## 4G Tool — LNHOIF thresholds in LNCEL Details — July 2026

### Improved — `LNHOIF List` column now shows t3 / t3a per relation
The `LNHOIF List` column in the 4G summary's **LNCEL Details** sheet previously listed only the target EARFCNs (`119, 1351, 40990, 41188`). It now shows **one relation per line** with the trigger thresholds in dBm (raw − 140):

```
119:   t3 -110 / t3a -98
1351:  t3 -102 / t3a -112
40990: t3 -104 / t3a -112
```

`t3` = `LNHOIF.threshold3InterFreq`, `t3a` = `LNHOIF.threshold3aInterFreq` (`?` if absent). The cell uses text-wrap so rows auto-fit; the missing-neighbour red highlight is preserved.

---

## Version 6.3 — June 2026  (`oss_xml_to_xlsx_v6.3.py`)

### Improved — Single-Pass MO Block Slicing (faster parse)
The parse phase previously made **three full passes** over the decompressed XML before any per-record work:
1. `_MO_SELF_CLOSE_RE.sub()` — rewrote self-closing `<managedObject/>` tags to paired open+close, **rebuilding the entire multi-GB string** (~13.7 s on a 3.4 GB dump);
2. `_MO_SPLIT.split()` on `</managedObject>` (~7.4 s);
3. a filter loop over all ~4.75 M split blocks (~1.9 s).

V6.3 replaces all three with a **single `_MO_FIND_RE.finditer()` scan** over managedObject *open* tags. Each block is sliced as `text[this_open : next_open]`, which:
- captures a self-closing `<managedObject .../>` as its own complete block with **no string rewrite** (the old self-close data-loss bug, e.g. MAL-10, stays fixed — `[^>]*>` swallows the trailing `/`);
- **filters by class during the scan**, so only the kept blocks are ever materialised (e.g. 186 k of 4.75 M), instead of allocating every block then discarding.

**Result:** ~2.2× faster block extraction (22 s → 10 s per file on a 3.4 GB dump). End-to-end on a real two-file LTE dump: **parse 76 s → 46 s, total 108 s → 77 s (~29 % faster)**. Output is **byte-identical** to V6.2 (verified: same object counts and same output file size to the byte).

### Includes — 4G InterFreq HO Check
Bundles the `4g_tool` InterFreq measurement-vs-HO-trigger threshold check (see entry below) into the standalone exe.

---

## 4G Tool — InterFreq HO Threshold Check — June 2026

### New — Measurement-vs-HO-trigger threshold validation in 4G Summary
Added inter-frequency threshold consistency checks to the 4G LTE summary report (`4g_tool/reports/lnbts_summary.py`). All four thresholds use the RSRP offset `dBm = raw − 140`.

**Rule A (per frequency relation):** `LNHOIF.threshold3InterFreq` (HO trigger) must sit within `[threshold2InterFreq − 2, threshold2InterFreq]`. Two ways it can fail:
- `threshold3InterFreq < threshold2InterFreq − 2` → trigger too far *below* measurement start (UE measures but HO waits until signal is much worse) — `"trigger N dB low"`.
- `threshold3InterFreq > threshold2InterFreq` → trigger sits *above* measurement start (HO would fire before measurement even begins) — `"trigger N dB high"`.

**Rule B (per cell):** `LNCEL.threshold2a` (measurement stop) must be at least 2 dB better than `threshold2InterFreq` (start), i.e. `threshold2a ≥ threshold2InterFreq + 2`.

**`threshold3aInterFreq`** (target-cell threshold) is captured for reference, no rule.

Presented two ways:
- **LNCEL Details** sheet — 3 new columns (`t2 Start (dBm)`, `t2a Stop (dBm)`, `HO Thr Issue`); the `t2a Stop` cell turns red on Rule B failure, and `HO Thr Issue` lists the offending target EARFCNs (red) when any relation fails Rule A.
- **InterFreq HO Check** sheet (new) — one row per cell × LNHOIF frequency relation, with serving/target EARFCN, all four thresholds in dBm, the signed trigger gap, both rule flags (low/high reason text), and the relation's `LNHOIF Dist_Name` for reference; autofilter on the flag columns for bulk auditing.

**Severity colouring:** the serious `t3 < t2 − 2` ("trigger N dB low") case is **red**; the `t3 > t2` ("trigger N dB high") case — where the serving HO threshold is merely non-binding and is usually an intentional aggressive/neighbour-driven HO design — is downgraded to **yellow/warning** when it is the *only* fault. A "high" row that also fails Rule B (meas stop too low) stays red. Same red/yellow severity applies to the `HO Thr Issue` flag in LNCEL Details (red if the cell has any "low" relation, yellow if only "high").

Picked up automatically by `DumpWatcher2.py`'s scheduled 4G summary (runs `4g_tool` from source). Rebuild the exe to bundle it into the standalone parser.

---

## Version 6.2 — May 2026

### New — 3G Summary integrated into main parser (`oss_xml_to_xlsx_v6.2.py`)
The 3G WCDMA summary tool is now integrated into the main OSS XML parser, following the same pattern as 2G/4G summaries:

- A **Generate 3G Summary** checkbox appears in the post-parse dialog when `RNC`, `WBTS`, and `WCEL` classes are selected.
- Output is saved as `<base>_3G_Summary.xlsx`.
- Uses the pre-read snapshot (no re-read of the output file) — same zero-re-read optimisation as 2G/4G.
- `WNCEL` sheet is included in the snapshot for PMAX lookup even if not explicitly selected, as long as it was parsed.
- 3G summary timing shown in the Grand Total line.

Minimum required classes: **RNC, WBTS, WCEL** (WNCEL included automatically when present).

### Improved — Admin State column in 3G Cell Details
Added **Admin State** column (after WCEL Name) to the 3G summary Cell Details sheet:
- Source: `WCEL` sheet, `AdminCellState` field
- `1` → `Working`, any other value → `Down`

### Improved — Latitude / Longitude columns in HW Report site-wise sheets
Added **Latitude** and **Longitude** columns (after Site Name, before unit-type counts) to both "Site wise (All)" and "Site wise (Working)" sheets:
- Source: `MRBTS` sheet, `latitude` and `longitude` fields (raw internal value ÷ 10,000,000)
- Displayed with 6 decimal places; blank when not present in the dump.
- Freeze panes updated to col 4 to keep all four fixed columns (MRBTS, Site Name, Lat, Lng) visible while scrolling.

### Improved — Pink highlight for duplicate RSI + EARFCN DL in 4G LNCEL Details
If 2 or more LNCELs under the **same MRBTS** share the same combination of **RSI** and **EARFCN DL**, their RSI cells are highlighted in pink.

---

## Version 6.1.2 — May 2026

### New — 3G WCDMA Summary Tool (`3g_tool`)
Added `3g_tool/` package that reads a Nokia 3G WCDMA parameter dump (`.xlsx` / `.xlsb`) and produces a single-sheet `_3G_summary.xlsx`:

**Cell Details** — one row per WCEL, sorted by RNC → WBTS → WCEL ID:

| Column | Sheet | Field |
|---|---|---|
| RNC ID | `RNC` | `RNC` |
| RNC Name | `RNC` | `name` |
| WBTS ID | `WBTS` | `WBTS` |
| WBTS Name | `WBTS` | `name` |
| SBTS | `WBTS` | `SBTSId` |
| WCEL ID | `WCEL` | `WCEL` |
| WCEL Name | `WCEL` | `name` |
| LAC | `WCEL` | `LAC` |
| RAC | `WCEL` | `RAC` |
| PSC | `WCEL` | `PriScrCode` |
| UARFCN | `WCEL` | `UARFCN` |
| Tilt | `WCEL` | `angle` |
| CPICH | `WCEL` | `PtxPrimaryCPICH` ÷ 10 |
| CPICH | `WCEL` | `PtxPrimaryCPICH` ÷ 10 |
| PMAX | `WNCEL` | `maxCarrierPower` ÷ 10 (matched on WNCEL.WNCEL = WCEL.WCEL AND WNCEL.MRBTS = WBTS.SBTSId) |

Minimum required sheets: **RNC, WBTS, WCEL, WNCEL**.
Supports both calamine (fast) and built-in XLSB parser fallback, same as the 2G tool.

---

### Improved — Rotated Column Headers in Parsed Output
Data sheet column-header cells (row 2 of every MO class sheet) now use **Rotate Text Up** (`textRotation="90"`) with horizontal and vertical centre alignment. This makes wide sheets with many columns far more readable — narrow column widths are sufficient since the text reads vertically.

Implementation details:
- A new style `s=3` (blue header + `textRotation="90"`) is added to `STYLES_XML`; `s=1` (blue header, no rotation) is preserved unchanged for the **Info** sheet, which does **not** get rotated headers.
- Header row height is pinned to **95 pt** (~250 px) (`ht="95" customHeight="1"`) so the row never expands beyond that regardless of column-name length.
- Applied in both the in-memory path (`generate_worksheet_xml`) and the streaming path (`_stream_worksheet_xml`).

---

## Version 6.1 — May 2026

### Improved (2G Summary) — NSEI and PSEI columns in Cell Details
Two columns appended to the **Cell Details** sheet:
- **NSEI** — NS Entity Identifier, read from the master BTS record (`nsei` field).
- **PSEI** — Packet Switching Entity Identifier, read from the master BTS record (`psei` field).

Both are written as numbers (added to `_NUMERIC_COLS`).

---

### New — HW Inventory Report (`hw_tool`)
Added `hw_tool/` package that reads `INVUNIT` + `MRBTS`/`LNBTS` sheets from a parsed OSS dump and produces a 3-sheet Excel workbook:
- **Site wise (All)** — one row per MRBTS, one column per `inventoryUnitType`, cell = total count. Row 0 is a colour-coded group banner (RMOD / BBMOD / SMOD / Others); frozen at row 2 / col 2.
- **Site wise (Working)** — same layout filtered to `state=working` units only.
- **Overall** — one row per `inventoryUnitType`: Working | Total | Group (colour-coded).

Column ordering is driven by `vendorUnitFamilyType`: RMOD → BBMOD → SMOD → Others, alphabetical within each group.

Integrated into the main parser (V6.1): a *Generate HW Report* checkbox appears in the post-parse dialog when INVUNIT + MRBTS/LNBTS classes are selected; output is saved as `<base>_HW_Report.xlsx`.

Also integrated into `nokia-kpi-scripts/dump/DumpWatcher2.py` — HW dump jobs now auto-generate the report after conversion.

### New — PyInstaller spec bundles `hw_tool`
`spec/OSS_XML_Parser_V6.1.spec` includes `('../hw_tool', 'hw_tool')` so the compiled exe carries the HW report package.

---

## Version 6.1 — April 2026

### Improved — Sparse Record Flatten (Eliminates Dense Column Scan)
`flatten_records` (called before each worker is submitted) previously iterated every column position for every record, calling `dict.get()` `n_cols` times per row regardless of how many params were actually filled. Nokia MO records are sparse — a class with 200 columns typically has only 25–40 filled params per record. The new `flatten_records_sparse` iterates only the actual keys present in each `(hier, rec)` pair and maps each key to its column index via a single dict lookup. For wide sparse classes this reduces per-record work from O(n\_cols) to O(filled\_keys) — typically 4–8× fewer dict operations.

### Improved — Plain `dict` in Parser (No OrderedDict Overhead)
`parse_mo_block` and `parse_dist_name` now use plain `dict` instead of `OrderedDict`. Python 3.7+ guarantees insertion-order preservation, so column ordering and output are byte-for-byte identical to V6.0. Plain dicts have lower allocation cost and are cheaper to pickle to worker subprocesses than the `OrderedDict` subclass.

---

## Version 6.0 — April 2026

### New — Shared String Table (SST)
String cell values are now stored in `xl/sharedStrings.xml` and referenced by index in each worksheet (`t="s"`), replacing per-cell inline strings (`t="inlineStr"`). SST cells are ~55 % shorter in raw XML — beneficial for large dumps with heavy enum/boolean repetition. A count-based filter ensures only strings that appear more than once enter the SST; unique identifiers such as `Dist_Name` remain as inlineStr, keeping the pickled SST dict small so subprocess IPC stays fast.

### Improved — Column-Order Cache (Eliminates Redundant Record Scan)
The per-class record scan that builds the column order is now folded into the SST pre-scan loop, caching results in `col_orders[cls]`. The write loop reuses the cached order directly — records are iterated once before workers start instead of twice.

### New — Benchmark Script (`benchmark_v5_v6.py`)
Side-by-side harness that runs V5.1 and V6.0 on the same input file(s) and prints a formatted comparison of wall-clock time, peak RAM, average CPU, and output file size. Requires `psutil` for RAM/CPU metrics (`pip install psutil`).

---

## Version 5.1 — April 2026

### Improved — Streaming Worksheet XML (Constant-RAM Sheet Writing)
V5.0 built the entire sheet XML in a `parts=[]` list across all rows, then joined everything into one giant string and encoded it at the end. For a large sheet (e.g. ADCE with 300 K+ rows) this caused a RAM spike in the worker process proportional to the full sheet size.  
V5.1 introduces `_stream_worksheet_xml()` which writes rows to the temp file in batches of 2 000 — peak RAM stays constant regardless of row count. The OS write buffer (1 MB) keeps syscall overhead low.

### Improved — Streaming ZIP Assembly (Low-RAM XLSX Packaging)
V5.0's `assemble_xlsx()` read each temp sheet file fully into RAM (`f.read()`) before writing it into the xlsx zip (`writestr()`). For a 200 MB sheet this meant a second full copy of the data in RAM during assembly.  
V5.1 uses `zipfile.write()` which copies the file into the zip in chunks — large sheets never sit fully in RAM during this step.

### Fix — 4G Summary: LNCEL / LNCEL_FDD cells missing (0 LNCEL records)
The pre-read snapshot built from in-memory parsed data omitted the hierarchy fields (MRBTS, LNBTS, LNCEL etc.) from each record dict. The 4G network module needs these as explicit keys to link cells to their parent LNBTS (`_key_lnbts` / `_key_lncel`). Without them `lncel_fdd_list_by_lnbts_dn` was always empty, producing LNBTS sheets with no associated cells and a "0 LNCEL records" count. Fixed by merging the hierarchy OrderedDict into each record when building `pre_read`: `{**dict(hier), **rec}`. Record parameter values win on any field-name collision.

### Improved — Info Sheet Version Label
The "Created with OSS XML Converter" label in the Info sheet now correctly reads V5.1 (was showing V4 in all prior versions).

---

## Version 5.0 — April 2026

### New — Zero Re-Read Summary (Pre-Read Snapshot)
The 2G/4G summary tools previously re-read the entire output file from disk using calamine after writing it — adding ~52 s on a typical dual-file dump. V5 captures references to all required sheets directly from the in-memory parsed data before the write phase starts. The summary tools receive this snapshot and never touch the output file at all. Re-read time: eliminated.

### New — Parallel XML Files Inside a Single ZIP
When a ZIP archive contains multiple XML files, V5 now parses them in parallel threads (each getting an equal share of CPU cores) instead of sequentially. For a ZIP with two large XMLs this roughly halves parse time for that archive (e.g. ~38 s → ~22 s on a 22-core machine).

### Improved — Dialog Sequencing (Zero Idle Wait Before Write)
The summary options dialog and the Save-As dialog are now both shown during parsing (immediately after class selection). By the time both are answered and the output path is confirmed, parsing is done and write starts with zero additional wait.

Summary dialog logic tightened:
- 2G Summary offered only when BTS, BCF, BSC and TRX are all selected.
- 4G Summary offered only when LNBTS (any variant) AND LNCEL (any variant) are both selected.
- ADCE-dependent sheets (One-Way ADCE, Discrepant ADCE, Co-Site Missing Neighbours) skipped automatically when ADCE was not parsed.

### Improved — XLSB Pre-Warm (Excel Startup Hidden Behind Write Phase)
When `.xlsb` output is chosen, a background thread launches `Excel.Application` via `DispatchEx` immediately — before the write phase starts. Excel's ~10 s startup cost is therefore hidden behind the write phase (~20 s) and costs nothing extra.

COM tweaks for faster conversion: `xlCalculationManual`, `EnableEvents=False`, `ScreenUpdating=False`, `UpdateLinks=0`, `AddToMru=False`, `EnableAutoRecover=False`. `DispatchEx` always spawns a fresh hidden Excel.exe process — never hijacks an already-open window.

### Improved — XLSB + Summary Run in Parallel
After write+assemble completes, XLSB conversion (Excel, background thread) and summary generation (main thread) run simultaneously. On a typical run summary finishes in ~17 s while XLSB takes ~100 s — the summary is ready well before XLSB finishes.

### Improved — Grand Total Shows True Wall-Clock Time
The Grand Total line now reports actual elapsed wall-clock time from tool open to completion, correctly accounting for parallel phases (XLSB and summaries overlap — summing them would over-count).

### Improved — "Press Any Key to Exit" Reliability
Tkinter dialog button-clicks were being buffered into stdin, causing the terminal to close immediately. Fixed by draining the stdin buffer via `msvcrt` before waiting for a real keypress.

---

## Version 4.2 — April 2026

- **Improved** — Larger file parsed first (sorted by size descending).
- **Improved** — Summary file read starts in background thread immediately after write; overlaps the summary dialog display.

---

## Version 4.1 — April 2026

### Fix — Missing MO After Self-Closing Empty Tag
Nokia XML files sometimes contain empty MOs as self-closing tags (e.g. `<managedObject class="SMLC" ... />`). The parser split on `</managedObject>`, so the MO immediately following a self-closing one was merged and silently dropped. Fixed by normalising all self-closing MO tags to paired open+close form before splitting.

### Fix — Leading Zeros Preserved in Parsed Values
Values such as `"03"` were being converted to integer `3`. Fixed: any numeric-looking string starting with `'0'` followed by another digit is kept as text.

### New (2G Summary) — Discrepant ADCE Sheet + Count Column
Lists every ADCE neighbour entry where defined parameters don't match the actual cell. Checked fields: LAC / NCC / BCC / MCC / MNC / BCCH. Mismatched fields highlighted in red.

### New (2G Summary) — Frequency Reuse Sheet
Lists every ARFCN in the network with its occurrence count, split by BCCH and TCH usage.

### Improved (2G Summary) — Hopping Mode & MAL ID Columns in Cell Details
Two columns inserted after BCCH: Hopping Mode (None / BB / RF) and MAL ID (for RF-hopping BTS).

### Improved (2G Summary) — TCH Freq for RF-Hopping Cells
For `hoppingMode = RF`, TCH Freq is sourced from the MAL sheet (frequency list) instead of TRX `initialFrequency`.

### Fix (2G Summary) — Bare `<p>` List Fields Parsed Correctly
List fields where Nokia omits the `name` attribute on `<p>` elements were silently dropped. Affected: `SPC.spcList`, `MAL.frequency`. Fixed with a fallback regex.

---

## Version 4.0 — April 2026

### New — 2G & 4G Summary Report Generation
After the main Excel dump is created, the tool detects whether parsed data contains 2G or 4G objects and offers to generate summary reports:
- **4G Summary** — LNBTS/LNCEL hierarchy, FDD/TDD split, EARFCNs, handover config, network statistics.
- **2G Summary** — Cell details (114 columns), BCF details, one-way ADCE, co-site missing neighbours, network statistics.

Output auto-named `<dump_filename>_4G_Summary.xlsx` / `_2G_Summary.xlsx`. Duplicate filenames get `(1)`, `(2)` suffix automatically.

### New — Nested ZIP Support
Input ZIPs can contain any mix of `.xml`, `.xml.gz`, and nested `.zip` files at any folder depth, all unpacked in-memory.

### New — Multi-Core XML Parsing (bypasses Python GIL)
Parsing distributed across all CPU cores via `ProcessPoolExecutor`. Two modes: one-process-per-core (many files) or intra-file chunking (few large files). Typically 3–6× faster than V3.

### New — Config File (`XML_Parser_AJ.cfg`)
MO class selection saved between runs. Next run pre-ticks the same classes automatically.

### Improved — XLSB Conversion (hidden Excel instance)
Excel instance completely hidden; `DispatchEx` always spawns a fresh process.

### Improved — Parsing Overlaps with Save-As Dialog
XML parsing starts immediately after class selection and runs in background while user browses for an output path.

---

## Version 3.1 — April 2026

- Direct XLSX XML generation (no openpyxl) — significantly faster writes.
- Regex-based XML parser (no lxml dependency).
- Parallel write phase using `ProcessPoolExecutor`.
- Support for `.xml.gz` and `.zip` input files.
- `.xlsb` output option via Excel COM automation.
- Row-splitting for sheets exceeding 1 000 000 rows.
- Info sheet with file metadata and sheet index.
- Freeze panes and styled headers in all sheets.
