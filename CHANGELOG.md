# Changelog — Nokia OSS XML Parser

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
