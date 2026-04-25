# Nokia OSS XML Parser

A Windows desktop tool that converts Nokia OSS RAML XML configuration dumps into structured Excel workbooks (`.xlsx` / `.xlsb`), with optional 2G and 4G network summary reports.

Built and maintained by **Ankit Jain**.

---

## Features

- **Multi-core parsing** — distributes work across all CPU cores via `ProcessPoolExecutor` (3–6× faster than single-threaded).
- **Streaming write** — rows are written to disk in batches; RAM usage stays flat regardless of output size.
- **Nested ZIP support** — input can be `.xml`, `.xml.gz`, `.zip`, or ZIPs containing further ZIPs.
- **2G Summary report** — cell details (114 columns), BCF details, one-way ADCE neighbours, discrepant ADCE, co-site missing neighbours, frequency reuse, network statistics.
- **4G Summary report** — LNBTS/LNCEL hierarchy, FDD/TDD split, EARFCNs, handover config, network statistics.
- **XLSB output** — optional Excel Binary format via hidden COM automation (smaller files, faster to open in Excel).
- **Config persistence** — MO class selection is saved between runs.

---

## Requirements

- Windows 10/11
- Python 3.10+ (to run from source)
- Dependencies: `python-calamine`, `openpyxl`
- For XLSB output: Microsoft Excel installed

---

## Usage

### Run from source

```
python oss_xml_to_xlsx_v5.1.py
```

A GUI dialog will open to select input files and MO classes. Output is saved as `.xlsx` (or `.xlsb` if selected).

### Run compiled exe

Download the latest `OSS_XML_Parser_vX.X.exe` from [Releases](../../releases) and run it directly — no Python installation needed.

---

## Building the exe

Requires [PyInstaller](https://pyinstaller.org/):

```
pip install pyinstaller python-calamine openpyxl
cd spec
pyinstaller OSS_XML_Parser_V5.1.spec
```

The compiled exe will appear in `dist/`.

> **Note:** The `2g_tool/` and `4g_tool/` directories must be one level above the spec file when building (the spec references `../2g_tool` and `../4g_tool`). The directory layout in this repo already matches that requirement.

---

## Repository Structure

```
nokia-oss-xml-parser/
├── oss_xml_to_xlsx_v5.1.py   ← current release (V5.1)
├── oss_xml_to_xlsx_v5.py     ← previous stable (V5.0)
├── 2g_tool/                  ← 2G summary report package
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
├── spec/                     ← PyInstaller build specs
│   ├── OSS_XML_Parser_V5.0.spec
│   ├── OSS_XML_Parser_V5.1.spec
│   └── version_info_v5.txt
├── archive/                  ← historical source versions (V1–V4)
├── CHANGELOG.md
└── README.md
```

---

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for the full version history.

---

## License

Personal project. Not affiliated with Nokia.
