# Nokia OSS XML Parser

A Windows desktop tool that converts Nokia OSS RAML XML configuration dumps into structured Excel workbooks (`.xlsx` / `.xlsb`), with optional 2G, 3G, 4G, and HW Inventory summary reports.

Built and maintained by **Ankit Jain**.

---

## Features

- **Multi-core parsing** — distributes work across all CPU cores via `ProcessPoolExecutor` (3–6× faster than single-threaded).
- **Shared String Table (SST)** — repeated string values stored once in `xl/sharedStrings.xml`; worksheet XML is 30–60% smaller for Nokia dumps with heavy enum/boolean repetition.
- **Sparse record flatten** — per-record work scales with filled params, not total columns; benefits wide sparse classes (e.g. WCEL, LNREL, ADJW).
- **Streaming write** — rows written to disk in batches; RAM stays flat regardless of output size.
- **Streaming ZIP assembly** — sheets copied into the xlsx zip in chunks; no full in-memory copy.
- **Nested ZIP support** — input can be `.xml`, `.xml.gz`, `.zip`, or ZIPs containing further ZIPs.
- **2G Summary report** — cell details (114 columns), BCF details, one-way ADCE neighbours, discrepant ADCE, co-site missing neighbours, frequency reuse, network statistics.
- **3G Summary report** — WCDMA cell details: RNC/WBTS/WCEL hierarchy, Admin State, LAC/RAC, PSC, UARFCN, Tilt, CPICH, PMAX.
- **4G Summary report** — LNBTS/LNCEL hierarchy, FDD/TDD split, EARFCNs, handover config, network statistics.
- **HW Inventory report** — per-site unit counts (all states + working-only) and network-wide totals; columns grouped and colour-coded by hardware family (RMOD / BBMOD / SMOD / Others).
- **XLSB output** — optional Excel Binary format via hidden COM automation (smaller files, faster to open in Excel).
- **Config persistence** — MO class selection is saved between runs.

---

## Requirements

- Windows 10/11
- Python 3.10+ (to run from source)
- Dependencies: `python-calamine`, `openpyxl`, `xlsxwriter`
- For XLSB output: Microsoft Excel installed

---

## Usage

### Run from source

```
python oss_xml_to_xlsx_v6.4.py
```

A GUI dialog will open to select input files and MO classes. Output is saved as `.xlsx` (or `.xlsb` if selected).

### Run compiled exe

Download the latest `OSS_XML_Parser_vX.X.exe` from [Releases](../../releases) and run it directly — no Python installation needed.

---

## Building the exe

Requires [PyInstaller](https://pyinstaller.org/):

```
pip install pyinstaller python-calamine openpyxl xlsxwriter
pyinstaller spec/OSS_XML_Parser_V6.4.spec --distpath dist_v64 --workpath build_v64
```

The compiled exe will appear in `dist_v64/`.

> **Note:** The spec file references `../2g_tool`, `../3g_tool`, `../4g_tool`, and `../hw_tool` relative to its location in `spec/`. Run PyInstaller from the repo root as shown above.

---

## Benchmark

Compare any two versions side-by-side:

```
python benchmark_v5_v6.py your_dump.xml.gz --runs 2
# or compare specific versions:
python benchmark_v5_v6.py your_dump.xml.gz --v51 oss_xml_to_xlsx_v6.0.py --v60 oss_xml_to_xlsx_v6.1.py
```

Requires `psutil` for RAM/CPU metrics (`pip install psutil`).

---

## Repository Structure

```
nokia-oss-xml-parser/
├── oss_xml_to_xlsx_v6.4.py   ← current release (V6.4)
├── oss_xml_to_xlsx_v6.3.py   ← previous release (V6.3)
├── oss_xml_to_xlsx_v6.2.py   ← previous release (V6.2)
├── oss_xml_to_xlsx_v6.1.py
├── oss_xml_to_xlsx_v6.0.py
├── oss_xml_to_xlsx_v5.1.py   ← stable baseline (V5.1)
├── benchmark_v5_v6.py        ← side-by-side benchmark harness
├── 2g_tool/                  ← 2G summary report package
│   ├── main.py
│   ├── network.py
│   ├── xlsb_reader.py
│   ├── xlsx_reader.py
│   └── reports/
│       └── cell_summary.py
├── 3g_tool/                  ← 3G WCDMA summary report package
│   ├── main.py
│   ├── network.py
│   ├── xlsb_reader.py
│   ├── xlsx_reader.py
│   └── reports/
│       └── cell_summary.py
├── 4g_tool/                  ← 4G summary report package
│   ├── main.py
│   ├── network.py
│   ├── xlsx_reader.py
│   └── reports/
│       └── lnbts_summary.py
├── hw_tool/                  ← HW inventory report package
│   ├── main.py
│   └── report.py
├── spec/                     ← PyInstaller build specs
│   ├── OSS_XML_Parser_V6.4.spec
│   ├── OSS_XML_Parser_V6.3.spec
│   ├── OSS_XML_Parser_V6.2.spec
│   ├── OSS_XML_Parser_V6.1.2.spec
│   ├── OSS_XML_Parser_V6.0.spec
│   ├── OSS_XML_Parser_V5.1.spec
│   ├── version_info_v64.txt
│   ├── version_info_v63.txt
│   ├── version_info_v62.txt
│   ├── version_info_v612.txt
│   ├── version_info_v61.txt
│   ├── version_info_v6.txt
│   └── version_info_v5.txt
├── archive/                  ← historical source versions (V1–V4)
├── CHANGELOG.md
└── README.md
```

---

## Version History (summary)

| Version | Key improvement |
|---------|----------------|
| **V6.4** | Streaming parse — 2× faster parse phase, 3.6× lower peak RAM |
| V6.3 | Single-pass MO block slicing (~29% faster parse); bundles 4G InterFreq HO Check |
| V6.2 | 3G WCDMA summary report (`3g_tool`) integrated into main parser |
| V6.1.2 | Rotated column headers (`textRotation=90`, centred) in all parsed sheets |
| V6.1 | Sparse record flatten; plain `dict` parser; HW Inventory report (`hw_tool`) |
| V6.0 | Shared String Table (SST); column-order cache |
| V5.1 | Streaming worksheet XML; streaming ZIP assembly |
| V5.0 | Pre-read snapshot (zero re-read); parallel ZIP parsing; dialog overlap |
| V4.x | 2G/4G summary reports; nested ZIP support; multi-core parsing |

See [CHANGELOG.md](CHANGELOG.md) for the full version history.

---

## License

Personal project. Not affiliated with Nokia.
