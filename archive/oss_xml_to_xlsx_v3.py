#!/usr/bin/env python3
"""
OSS XML to XLSX Converter  — Version 3  (Blazing Fast)
Converts Nokia OSS RAML XML dump files to Excel format.

Key improvements over V2:
  - Regex-based XML parser (no lxml dependency) — 3-5x faster
  - Direct XLSX XML generation (no openpyxl) — raw string assembly
  - Parallel parse (ThreadPoolExecutor) + parallel write (ProcessPoolExecutor)
  - Zero heavy dependencies for core engine (only stdlib)

Usage (interactive GUI):   python oss_xml_to_xlsx_v3.py
Usage (command-line):      python oss_xml_to_xlsx_v3.py file1.xml.gz file2.gz -o out.xlsx
                           python oss_xml_to_xlsx_v3.py file1.zip --classes BTS,BCF -o out.xlsx
"""

import gzip
import io
import math
import os
import re
import sys
import zipfile
import argparse
import threading
from collections import defaultdict, OrderedDict
from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor, as_completed
from datetime import datetime
from multiprocessing import freeze_support

try:
    import win32com.client
    XLSB_SUPPORTED = True
except ImportError:
    XLSB_SUPPORTED = False

AUTHOR = 'Ankit Jain'

# Excel hard limit is 1,048,576 rows.  With 2 header rows that leaves 1,048,574
# data rows.  We use a round 1,000,000 as the split threshold so oversized
# classes are automatically tiled into like LNREL(1), LNREL(2) … sheets.
MAX_ROWS_PER_SHEET = 1_000_000

# ---------------------------------------------------------------------------
# Column letter helpers — handles any column count (A, B, ..., Z, AA, ..., ZZZ)
# ---------------------------------------------------------------------------
_CL = []
for _i in range(26):
    _CL.append(chr(65 + _i))
for _i in range(26):
    for _j in range(26):
        _CL.append(chr(65 + _i) + chr(65 + _j))

def _col_letter(n):
    """0-based column index to Excel column letter. Handles unlimited columns."""
    if n < 702:          # fast path for pre-computed table
        return _CL[n]
    # General algorithm for AAA+ columns
    result = ''
    n += 1              # switch to 1-based for the algorithm
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result

# ---------------------------------------------------------------------------
# Hardcoded styles.xml — 3 styles:
#   s=0: default (General)
#   s=1: blue header with white bold text + thin borders
#   s=2: hyperlink (blue underlined text)
# ---------------------------------------------------------------------------
STYLES_XML = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="3">
  <font><sz val="11"/><name val="Calibri"/></font>
  <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  <font><sz val="11"/><color rgb="FF0563C1"/><u/><name val="Calibri"/></font>
</fonts>
<fills count="3">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>
</fills>
<borders count="2">
  <border><left/><right/><top/><bottom/><diagonal/></border>
  <border>
    <left style="thin"><color auto="1"/></left>
    <right style="thin"><color auto="1"/></right>
    <top style="thin"><color auto="1"/></top>
    <bottom style="thin"><color auto="1"/></bottom>
    <diagonal/>
  </border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
  <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
</styleSheet>'''

# ---------------------------------------------------------------------------
# Thread-safe print
# ---------------------------------------------------------------------------
_print_lock = threading.Lock()
def tprint(*a, **k):
    with _print_lock:
        print(*a, **k)

def fmt_elapsed(s):
    if s < 60: return f'{s:.1f}s'
    m, sec = divmod(int(s), 60)
    return f'{m}m {sec:02d}s'

def ts():
    return datetime.now().strftime('%H:%M:%S')


# ---------------------------------------------------------------------------
# File reading
# ---------------------------------------------------------------------------

def iter_xml_streams(filepath):
    name = os.path.basename(filepath)
    low  = filepath.lower()
    if low.endswith('.zip'):
        with zipfile.ZipFile(filepath, 'r') as zf:
            entries = [e for e in zf.namelist()
                       if e.lower().endswith('.xml') and not e.startswith('__MACOSX')]
            if not entries:
                raise ValueError(f'No .xml files found inside {name}')
            for entry in entries:
                with zf.open(entry) as f:
                    yield f.read(), f'{name}/{os.path.basename(entry)}'
    elif low.endswith('.gz'):
        with gzip.open(filepath, 'rb') as f:
            yield f.read(), name
    else:
        with open(filepath, 'rb') as f:
            yield f.read(), name


# ---------------------------------------------------------------------------
# Quick scan (regex — same as V2)
# ---------------------------------------------------------------------------

_CLASS_RE = re.compile(rb'managedObject\s[^>]*class="([^"]+)"')

def quick_scan_classes(filepath):
    counts = defaultdict(int)
    low = filepath.lower()
    if low.endswith('.zip'):
        with zipfile.ZipFile(filepath, 'r') as zf:
            for entry in [e for e in zf.namelist() if e.lower().endswith('.xml')]:
                with zf.open(entry) as f:
                    for m in _CLASS_RE.finditer(f.read()):
                        counts[m.group(1).decode()] += 1
    elif low.endswith('.gz'):
        with gzip.open(filepath, 'rb') as f:
            data = f.read()
            for m in _CLASS_RE.finditer(data):
                counts[m.group(1).decode()] += 1
    else:
        with open(filepath, 'rb') as f:
            data = f.read()
            for m in _CLASS_RE.finditer(data):
                counts[m.group(1).decode()] += 1
    return counts

def scan_all_files(filepaths):
    merged = defaultdict(int)
    with ThreadPoolExecutor(max_workers=min(len(filepaths), 4)) as ex:
        for counts in ex.map(quick_scan_classes, filepaths):
            for cls, cnt in counts.items():
                merged[cls] += cnt
    return merged


# ---------------------------------------------------------------------------
# Regex-based XML Parser  (NO lxml — operates on raw bytes decoded to str)
# ---------------------------------------------------------------------------

# Pre-compiled regexes for attribute extraction
_MO_ATTR = re.compile(r'<(?:\w+:)?managedObject\s([^>]*)>')
_ATTR_RE = re.compile(r'(\w+)="([^"]*)"')
_LOG_DT  = re.compile(r'<(?:\w+:)?log\s[^>]*dateTime="([^"]*)"')

# For extracting <p name="...">...</p> (non-greedy, handles empty)
_P_RE = re.compile(r'<(?:\w+:)?p\s+name="([^"]*)">(.*?)</(?:\w+:)?p>', re.DOTALL)
_P_EMPTY_RE = re.compile(r'<(?:\w+:)?p\s+name="([^"]*)"\s*/>')

# For extracting <list name="...">...</list>
_LIST_RE = re.compile(r'<(?:\w+:)?list\s+name="([^"]*)">(.*?)</(?:\w+:)?list>', re.DOTALL)

# For extracting <item>...</item> blocks
_ITEM_RE = re.compile(r'<(?:\w+:)?item>(.*?)</(?:\w+:)?item>', re.DOTALL)

# Split on </managedObject> to get individual MO blocks
_MO_SPLIT = re.compile(r'</(?:\w+:)?managedObject>')


def try_numeric(text):
    if text is None:
        return None
    s = text.strip()
    if not s:
        return None
    try:
        f = float(s)
        return int(f) if f == int(f) else f
    except (ValueError, TypeError):
        return s


def parse_dist_name(dist_name):
    h = OrderedDict()
    for part in dist_name.split('/'):
        if '-' not in part:
            continue
        idx = part.index('-')
        cls, oid = part[:idx], part[idx + 1:]
        if cls == 'PLMN':
            continue
        try:
            f = float(oid)
            h[cls] = int(f) if f == int(f) else f
        except ValueError:
            h[cls] = oid
    return h


def parse_mo_block(block, filename):
    """Parse a single managedObject block using regex. Returns (mo_class, hierarchy, record) or None."""
    m = _MO_ATTR.search(block)
    if not m:
        return None

    attrs = dict(_ATTR_RE.findall(m.group(1)))
    mo_class  = attrs.get('class', '')
    dist_name = attrs.get('distName', '')
    obj_id    = attrs.get('id', '')
    version   = attrs.get('version', '')
    operation = attrs.get('operation', '')

    hierarchy = parse_dist_name(dist_name)
    record = OrderedDict()
    if operation:
        record['operation'] = operation
    record['id']         = try_numeric(obj_id)
    record['File_Name']  = filename
    record['Dist_Name']  = dist_name
    record['SW_Version'] = version

    # Extract lists first, then remove them to find top-level <p> tags
    remainder = block
    for lm in _LIST_RE.finditer(block):
        list_name = lm.group(1)
        list_body = lm.group(2)

        # Check if list contains <item> elements
        items = _ITEM_RE.findall(list_body)
        if items:
            record[list_name] = 'List'
            item_fields = OrderedDict()
            for item_body in items:
                for pm in _P_RE.finditer(item_body):
                    item_fields.setdefault(pm.group(1), []).append(pm.group(2) or '')
                for pm in _P_EMPTY_RE.finditer(item_body):
                    item_fields.setdefault(pm.group(1), []).append('')
            for fname, vals in item_fields.items():
                record[f'Item-{list_name}-{fname}'] = ';'.join(vals)
        else:
            vals = []
            for pm in _P_RE.finditer(list_body):
                vals.append(pm.group(2) or '')
            for pm in _P_EMPTY_RE.finditer(list_body):
                vals.append('')
            if vals:
                record[list_name] = 'List;' + ';'.join(vals)

    # Remove all <list>...</list> blocks, then extract top-level <p>
    remainder = _LIST_RE.sub('', block)
    for pm in _P_RE.finditer(remainder):
        record[pm.group(1)] = try_numeric(pm.group(2))
    for pm in _P_EMPTY_RE.finditer(remainder):
        record[pm.group(1)] = None

    return mo_class, hierarchy, record


def parse_xml_bytes_v3(data, display_name, filter_classes=None):
    """Parse XML bytes using regex. Returns (classes_data, file_info)."""
    classes_data = defaultdict(list)
    file_info = {'filename': display_name, 'dateTime': ''}
    count = 0
    t0 = datetime.now()
    tprint(f'  [{ts()}] Parsing {display_name} ...')

    # Decode bytes to string
    text = data.decode('utf-8', errors='replace') if isinstance(data, (bytes, bytearray)) else data

    # Extract dateTime from <log> element
    log_m = _LOG_DT.search(text)
    if log_m:
        file_info['dateTime'] = log_m.group(1)

    # If filter_classes specified, do a fast pre-check per block
    filter_set = filter_classes if filter_classes else None

    # Split on </managedObject> and process each block
    blocks = _MO_SPLIT.split(text)
    for block in blocks:
        if 'managedObject' not in block:
            continue

        # Quick class extraction for filtering
        if filter_set:
            cm = re.search(r'class="([^"]+)"', block)
            if not cm or cm.group(1) not in filter_set:
                continue

        result = parse_mo_block(block, display_name)
        if result is None:
            continue

        mo_class, hierarchy, record = result
        classes_data[mo_class].append((hierarchy, record))
        count += 1

    elapsed = (datetime.now() - t0).total_seconds()
    tprint(f'  [{ts()}] Done    {display_name} -- {count:,} objects in {fmt_elapsed(elapsed)}')
    return classes_data, file_info


def parse_input_file(filepath, filter_classes=None):
    results = []
    for xml_bytes, display_name in iter_xml_streams(filepath):
        results.append(parse_xml_bytes_v3(xml_bytes, display_name, filter_classes))
    return results


# ---------------------------------------------------------------------------
# Column ordering + record flattening (same as V2)
# ---------------------------------------------------------------------------

def build_column_order(records):
    META = ('operation', 'id', 'File_Name', 'Dist_Name', 'SW_Version')
    hier_seen, param_seen = OrderedDict(), OrderedDict()
    has_op = False
    for hierarchy, record in records:
        for k in hierarchy:
            hier_seen[k] = True
        for k in record:
            if k == 'operation':
                has_op = True
            if k not in META:
                param_seen[k] = True
    meta = list(META) if has_op else [m for m in META if m != 'operation']
    return list(hier_seen.keys()), meta, list(param_seen.keys())


def flatten_records(records, all_cols, hier_set):
    """Convert list of (hierarchy, record) to list of plain lists for fast pickling."""
    col_info = [(col in hier_set, col) for col in all_cols]
    return [
        [hierarchy.get(col) if is_h else record.get(col)
         for is_h, col in col_info]
        for hierarchy, record in records
    ]


# ---------------------------------------------------------------------------
# Direct XLSX XML generation  (NO openpyxl — raw string assembly)
# ---------------------------------------------------------------------------

def _xml_escape(s):
    """Fast XML escape — only replaces when needed."""
    if s is None:
        return ''
    s = str(s)
    if '&' in s or '<' in s or '>' in s or '"' in s:
        return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
    return s


def _cell_xml(col_idx, row_idx, value, style=0):
    """Generate XML for a single cell. row_idx is 1-based."""
    ref = _col_letter(col_idx) + str(row_idx)
    if value is None:
        return ''
    if isinstance(value, int):
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    if isinstance(value, float):
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    # String — use inline string
    escaped = _xml_escape(value)
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{escaped}</t></is></c>'


def generate_worksheet_xml(cls_name, flat_records, all_cols, n_hier, return_hyperlinks=True):
    """
    Generate raw worksheet XML for one sheet.
    Returns (sheet_xml_bytes, rels_xml_bytes_or_None).
    """
    n_cols = len(all_cols)
    has_op = 'operation' in all_cols

    # Freeze pane position
    dist_freeze_col = n_hier + (1 if has_op else 0) + 3  # 0-based index of column AFTER Dist_Name
    freeze_col_letter = _col_letter(dist_freeze_col)
    freeze_ref = f'{freeze_col_letter}3'

    # Column widths
    col_widths = []
    ci = 0
    for i in range(n_hier):
        col_widths.append((ci, 12)); ci += 1
    if has_op:
        col_widths.append((ci, 10)); ci += 1
    col_widths.append((ci, 12)); ci += 1   # id
    col_widths.append((ci, 22)); ci += 1   # File_Name
    col_widths.append((ci, 45)); ci += 1   # Dist_Name
    col_widths.append((ci, 10)); ci += 1   # SW_Version
    for i in range(ci, n_cols):
        col_widths.append((i, 18))

    # Build XML using list of strings joined at end
    parts = []
    parts.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    parts.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                 ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')

    # Sheet views with freeze pane
    parts.append('<sheetViews><sheetView tabSelected="0" workbookViewId="0">')
    parts.append(f'<pane xSplit="{dist_freeze_col}" ySplit="2" topLeftCell="{freeze_ref}" activePane="bottomRight" state="frozen"/>')
    parts.append('</sheetView></sheetViews>')

    # Column widths
    parts.append('<cols>')
    for col_idx, width in col_widths:
        c1 = col_idx + 1
        parts.append(f'<col min="{c1}" max="{c1}" width="{width}" customWidth="1"/>')
    parts.append('</cols>')

    # Sheet data
    parts.append('<sheetData>')

    # Row 1: "Info" link cell (style=2 for hyperlink)
    parts.append('<row r="1">')
    parts.append(f'<c r="A1" s="2" t="inlineStr"><is><t>Info</t></is></c>')
    parts.append('</row>')

    # Row 2: headers (style=1 for blue header)
    parts.append('<row r="2">')
    for ci, col_name in enumerate(all_cols):
        parts.append(_cell_xml(ci, 2, col_name, style=1))
    parts.append('</row>')

    # Data rows (rows 3+)
    for ri, row in enumerate(flat_records, start=3):
        parts.append(f'<row r="{ri}">')
        for ci, val in enumerate(row):
            if val is not None:
                parts.append(_cell_xml(ci, ri, val, style=0))
        parts.append('</row>')

    parts.append('</sheetData>')

    # Auto-filter
    last_col = _col_letter(n_cols - 1)
    parts.append(f'<autoFilter ref="A2:{last_col}2"/>')

    # Hyperlink for Info cell
    hyperlinks_xml = ''
    if return_hyperlinks:
        hyperlinks_xml = f"<hyperlinks><hyperlink ref=\"A1\" location=\"'Info'!A1\"/></hyperlinks>"
        parts.append(hyperlinks_xml)

    parts.append('</worksheet>')

    sheet_xml = ''.join(parts).encode('utf-8')

    # No .rels needed for internal hyperlinks (location-based)
    return sheet_xml, None


def generate_info_sheet_xml(files_info, class_names):
    """Generate Info sheet XML directly."""
    parts = []
    parts.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    parts.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                 ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')

    parts.append('<cols>')
    parts.append('<col min="1" max="1" width="3" customWidth="1"/>')
    parts.append('<col min="2" max="2" width="22" customWidth="1"/>')
    parts.append('<col min="3" max="3" width="32" customWidth="1"/>')
    parts.append('<col min="4" max="4" width="38" customWidth="1"/>')
    parts.append('</cols>')

    parts.append('<sheetData>')

    row = 1
    # Row 1: empty
    parts.append(f'<row r="{row}"/>')
    row += 1

    # Row 2: title
    parts.append(f'<row r="{row}">')
    parts.append(_cell_xml(1, row, 'Created with OSS XML Converter  --  V3', style=1))
    parts.append('</row>')
    row += 1

    # Row 3: author
    parts.append(f'<row r="{row}">')
    parts.append(_cell_xml(1, row, f'Created by: {AUTHOR}', style=0))
    parts.append('</row>')
    row += 1

    # Row 4: empty
    parts.append(f'<row r="{row}"/>')
    row += 1

    # Row 5: file headers
    parts.append(f'<row r="{row}">')
    parts.append(_cell_xml(2, row, 'File name', style=1))
    parts.append(_cell_xml(3, row, 'Dump Extracted Time', style=1))
    parts.append('</row>')
    row += 1

    # File info rows
    for fi in files_info:
        parts.append(f'<row r="{row}">')
        parts.append(_cell_xml(1, row, 'Used Netact export:', style=0))
        parts.append(_cell_xml(2, row, fi['filename'], style=0))
        parts.append(_cell_xml(3, row, fi['dateTime'], style=0))
        parts.append('</row>')
        row += 1

    # Empty row
    parts.append(f'<row r="{row}"/>')
    row += 1

    # "Contents" label
    parts.append(f'<row r="{row}">')
    parts.append(_cell_xml(1, row, 'Contents', style=1))
    parts.append('</row>')
    row += 1

    # Class links
    hyperlinks = []
    for cls in class_names:
        parts.append(f'<row r="{row}">')
        parts.append(f'<c r="B{row}" s="2" t="inlineStr"><is><t>{_xml_escape(cls)}</t></is></c>')
        parts.append('</row>')
        hyperlinks.append(f"<hyperlink ref=\"B{row}\" location=\"'{_xml_escape(cls)}'!A1\"/>")
        row += 1

    parts.append('</sheetData>')

    # Hyperlinks
    if hyperlinks:
        parts.append('<hyperlinks>')
        parts.extend(hyperlinks)
        parts.append('</hyperlinks>')

    parts.append('</worksheet>')

    return ''.join(parts).encode('utf-8')


# ---------------------------------------------------------------------------
# Sheet writing worker  (runs in separate process)
# ---------------------------------------------------------------------------

def write_sheet_worker(args):
    """Runs in a subprocess. Generates one MO class sheet XML and writes to a temp file."""
    cls, flat_records, all_cols, n_hier, temp_dir = args
    t_start = datetime.now()

    sheet_xml, rels_xml = generate_worksheet_xml(cls, flat_records, all_cols, n_hier)

    # Write to disk so main process never accumulates all sheets in RAM
    safe_name = cls.replace('/', '_').replace('\\', '_')
    xml_path  = os.path.join(temp_dir, f'{safe_name}.xml')
    rels_path = None
    with open(xml_path, 'wb') as f:
        f.write(sheet_xml)
    if rels_xml:
        rels_path = os.path.join(temp_dir, f'{safe_name}.rels')
        with open(rels_path, 'wb') as f:
            f.write(rels_xml)

    elapsed = (datetime.now() - t_start).total_seconds()
    return cls, xml_path, rels_path, elapsed


# ---------------------------------------------------------------------------
# XLSX assembly  (build the zip directly)
# ---------------------------------------------------------------------------

WS_CT  = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
WB_CT  = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
STY_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
CORE_CT = 'application/vnd.openxmlformats-package.core-properties+xml'
WS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
STY_REL= 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
DOC_REL= 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
CORE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'


def _content_types_xml(n_sheets):
    overrides = ''.join(
        f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="{WS_CT}"/>'
        for i in range(n_sheets)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'<Override PartName="/xl/workbook.xml" ContentType="{WB_CT}"/>'
        f'<Override PartName="/xl/styles.xml" ContentType="{STY_CT}"/>'
        f'<Override PartName="/docProps/core.xml" ContentType="{CORE_CT}"/>'
        + overrides +
        '</Types>'
    ).encode()


def _root_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{DOC_REL}" Target="xl/workbook.xml"/>'
        f'<Relationship Id="rId2" Type="{CORE_REL}" Target="docProps/core.xml"/>'
        '</Relationships>'
    ).encode()


def _workbook_xml(sheet_names):
    sheets = ''.join(
        f'<sheet name="{_xml_escape(name)}" sheetId="{i+1}" r:id="rId{i+1}"/>'
        for i, name in enumerate(sheet_names)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<bookViews><workbookView activeTab="0"/></bookViews>'
        f'<sheets>{sheets}</sheets>'
        '</workbook>'
    ).encode()


def _workbook_rels_xml(n_sheets):
    ws_rels = ''.join(
        f'<Relationship Id="rId{i+1}" Type="{WS_REL}" Target="worksheets/sheet{i+1}.xml"/>'
        for i in range(n_sheets)
    )
    sty_rel = f'<Relationship Id="rId{n_sheets+1}" Type="{STY_REL}" Target="styles.xml"/>'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + ws_rels + sty_rel +
        '</Relationships>'
    ).encode()


def _docprops_core_xml():
    now = datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
        ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
        ' xmlns:dcterms="http://purl.org/dc/terms/"'
        ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>OSS XML Dump</dc:title>'
        f'<dc:creator>{AUTHOR}</dc:creator>'
        f'<cp:lastModifiedBy>{AUTHOR}</cp:lastModifiedBy>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        '</cp:coreProperties>'
    ).encode()


def assemble_xlsx(sheet_order, sheet_paths, output_path):
    """
    Build the final xlsx from worksheet XML files on disk.
    sheet_paths: {name: (xml_path_or_bytes, rels_path_or_None)}
    Sheets missing from sheet_paths (failed workers) are silently skipped.
    """
    # Only include sheets that succeeded
    present = [name for name in sheet_order if name in sheet_paths]
    n = len(present)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=1) as out:
        out.writestr('[Content_Types].xml',        _content_types_xml(n))
        out.writestr('_rels/.rels',                _root_rels_xml())
        out.writestr('xl/workbook.xml',            _workbook_xml(present))
        out.writestr('xl/_rels/workbook.xml.rels', _workbook_rels_xml(n))
        out.writestr('xl/styles.xml',              STYLES_XML)
        out.writestr('docProps/core.xml',          _docprops_core_xml())

        for i, name in enumerate(present):
            xml_src, rels_src = sheet_paths[name]
            if isinstance(xml_src, (bytes, bytearray)):
                out.writestr(f'xl/worksheets/sheet{i+1}.xml', xml_src)
            else:
                with open(xml_src, 'rb') as f:
                    out.writestr(f'xl/worksheets/sheet{i+1}.xml', f.read())
            if rels_src:
                if isinstance(rels_src, (bytes, bytearray)):
                    out.writestr(f'xl/worksheets/_rels/sheet{i+1}.xml.rels', rels_src)
                else:
                    with open(rels_src, 'rb') as f:
                        out.writestr(f'xl/worksheets/_rels/sheet{i+1}.xml.rels', f.read())


# ---------------------------------------------------------------------------
# XLSB conversion
# ---------------------------------------------------------------------------

def convert_to_xlsb(xlsx_path, xlsb_path):
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
        wb.SaveAs(os.path.abspath(xlsb_path), FileFormat=50)
        wb.Close(False)
    finally:
        excel.Quit()


# ---------------------------------------------------------------------------
# Core conversion
# ---------------------------------------------------------------------------

def run_conversion(input_paths, output_path, filter_classes=None):
    t_total = datetime.now()

    want_xlsb = output_path.lower().endswith('.xlsb')
    if want_xlsb and not XLSB_SUPPORTED:
        print('WARNING: .xlsb requires pywin32 (pip install pywin32). Falling back to .xlsx.')
        output_path = output_path[:-5] + '.xlsx'
        want_xlsb   = False

    xlsx_path = output_path if not want_xlsb else output_path[:-5] + '_tmp.xlsx'

    print(f'[{ts()}] Output: {output_path}')
    if filter_classes:
        print(f'[{ts()}] Filtering to classes: {", ".join(sorted(filter_classes))}')
    print(f'[{ts()}] Parsing {len(input_paths)} file(s) in parallel...')

    # --- Parse ---
    all_cd, all_fi = [], []
    t_parse = datetime.now()
    with ThreadPoolExecutor(max_workers=min(len(input_paths), 4)) as ex:
        futs = {ex.submit(parse_input_file, p, filter_classes): p for p in input_paths}
        for fut in as_completed(futs):
            for cd, fi in fut.result():
                all_cd.append(cd)
                all_fi.append(fi)
    t_parse = (datetime.now() - t_parse).total_seconds()

    merged = defaultdict(list)
    for cd in all_cd:
        for cls, recs in cd.items():
            merged[cls].extend(recs)

    # Re-order file info to match input order
    fi_by_name = {fi['filename']: fi for fi in all_fi}
    ordered_fi = []
    for path in input_paths:
        name = os.path.basename(path)
        if name in fi_by_name and fi_by_name[name] not in ordered_fi:
            ordered_fi.append(fi_by_name[name])
    for fi in all_fi:
        if fi not in ordered_fi:
            ordered_fi.append(fi)

    class_names   = sorted(merged.keys())
    total_records = sum(len(v) for v in merged.values())
    print(f'[{ts()}] Parsing done in {fmt_elapsed(t_parse)} -- '
          f'{total_records:,} records, {len(class_names)} classes: {", ".join(class_names)}')

    # --- Parallel write (with temp dir to avoid RAM accumulation) ---
    import tempfile, shutil
    temp_dir = tempfile.mkdtemp(prefix='oss_v3_')
    try:
        print(f'[{ts()}] Writing {len(class_names)} sheets in parallel...')
        t_write = datetime.now()

        sheet_paths = {}   # cls -> (xml_path_or_bytes, rels_path_or_None)

        # Build sheet plan — split oversized classes into LNREL(1), LNREL(2) …
        # sheet_plan: list of (sheet_name, cls, start_row, end_row)
        sheet_plan = []
        for cls in class_names:
            n_rows = len(merged[cls])
            if n_rows > MAX_ROWS_PER_SHEET:
                n_parts = math.ceil(n_rows / MAX_ROWS_PER_SHEET)
                for i in range(n_parts):
                    sheet_plan.append((f'{cls}({i+1})', cls,
                                       i * MAX_ROWS_PER_SHEET,
                                       min((i + 1) * MAX_ROWS_PER_SHEET, n_rows)))
            else:
                sheet_plan.append((cls, cls, 0, n_rows))

        all_sheet_names = [sp[0] for sp in sheet_plan]  # may include split names

        # Info sheet (fast — in main process, keep as bytes since it's tiny)
        print(f'  [{ts()}] Sheet "Info" ...', end='', flush=True)
        t0 = datetime.now()
        info_xml = generate_info_sheet_xml(ordered_fi, all_sheet_names)
        sheet_paths['Info'] = (info_xml, None)
        print(f' done ({fmt_elapsed((datetime.now()-t0).total_seconds())})')

        # Data sheets — submit workers one class at a time to free RAM promptly
        n_workers = min(len(sheet_plan), os.cpu_count() or 4)
        if not sheet_plan:
            print(f'  [{ts()}] No data sheets to write.')
        else:
            split_count = len(sheet_plan) - len(class_names)
            split_note  = f'  ({split_count} extra sheet(s) from row-splitting)' if split_count else ''
            print(f'  [{ts()}] Launching {n_workers} parallel workers'
                  f' for {len(sheet_plan)} sheets{split_note}...')

        with ProcessPoolExecutor(max_workers=max(n_workers, 1)) as ex:
            futs       = {}
            prev_cls   = None
            flat_cur   = None
            cols_cur   = None
            nhier_cur  = 0

            for sheet_name, cls, start, end in sheet_plan:
                if cls != prev_cls:
                    # Free previous class data before loading next
                    if flat_cur is not None:
                        del flat_cur
                        flat_cur = None
                    recs = merged.pop(cls)
                    hier_cols, meta_cols, param_cols = build_column_order(recs)
                    cols_cur  = hier_cols + meta_cols + param_cols
                    nhier_cur = len(hier_cols)
                    flat_cur  = flatten_records(recs, cols_cur, set(hier_cols))
                    del recs
                    prev_cls = cls

                # Slice only when splitting; avoid copy for single-sheet classes
                chunk = flat_cur[start:end] if (end - start) < len(flat_cur) else flat_cur
                args  = (sheet_name, chunk, cols_cur, nhier_cur, temp_dir)
                futs[ex.submit(write_sheet_worker, args)] = sheet_name
                del chunk

            if flat_cur is not None:
                del flat_cur

            for fut in as_completed(futs):
                sheet_name = futs[fut]
                try:
                    name_result, xml_path, rels_path, elapsed = fut.result()
                    sheet_paths[name_result] = (xml_path, rels_path)
                    print(f'  [{ts()}] Sheet "{name_result}" done ({fmt_elapsed(elapsed)})')
                except Exception as e:
                    print(f'  [{ts()}] Sheet "{sheet_name}" FAILED: {e}')

        t_write = (datetime.now() - t_write).total_seconds()
        print(f'[{ts()}] All sheets written in {fmt_elapsed(t_write)}')

        # --- Assemble ---
        failed = [n for n in all_sheet_names if n not in sheet_paths]
        if failed:
            print(f'[{ts()}] WARNING: {len(failed)} sheet(s) failed and will be omitted: {", ".join(failed)}')
        print(f'[{ts()}] Assembling {os.path.basename(xlsx_path)} ...', end='', flush=True)
        t0 = datetime.now()
        sheet_order = ['Info'] + all_sheet_names
        assemble_xlsx(sheet_order, sheet_paths, xlsx_path)
        t_asm = (datetime.now() - t0).total_seconds()
        print(f' done ({fmt_elapsed(t_asm)})')

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    if want_xlsb:
        t0 = datetime.now()
        print(f'[{ts()}] Converting to .xlsb via Excel...')
        convert_to_xlsb(xlsx_path, output_path)
        os.remove(xlsx_path)
        print(f'[{ts()}] Conversion done in {fmt_elapsed((datetime.now()-t0).total_seconds())}')

    t_total = (datetime.now() - t_total).total_seconds()
    size_mb = os.path.getsize(output_path) / 1024 / 1024
    print(f'\n[{ts()}] Done!  {output_path}  ({size_mb:.1f} MB)')
    print(f'         Parse: {fmt_elapsed(t_parse)}  |  '
          f'Write: {fmt_elapsed(t_write)}  |  '
          f'Assemble: {fmt_elapsed(t_asm)}  |  '
          f'Total: {fmt_elapsed(t_total)}')
    return output_path


# ---------------------------------------------------------------------------
# Class selection dialog (same as V2)
# ---------------------------------------------------------------------------

def ask_class_selection(class_counts):
    import tkinter as tk
    selected = {}
    result   = [None]
    n_total  = len(class_counts)
    total_obj = sum(class_counts.values())

    dlg = tk.Toplevel()
    dlg.title('Select MO Classes')
    dlg.resizable(True, True)
    dlg.grab_set()

    # ── top summary ──────────────────────────────────────────────────────────
    tk.Label(dlg, text=f'{n_total} MO classes  |  {total_obj:,} total objects',
             font=('Arial', 10, 'bold'), pady=4).pack()

    # live counter — updated whenever a checkbox changes
    sel_var = tk.StringVar(value=f'{n_total} / {n_total} selected')
    tk.Label(dlg, textvariable=sel_var, font=('Arial', 9), fg='#444').pack()

    def _update_count(*_):
        n = sum(1 for v in selected.values() if v.get())
        sel_var.set(f'{n} / {n_total} selected')

    # ── scrollable checkbox area (vertical + horizontal) ──────────────────
    container = tk.Frame(dlg)
    container.pack(padx=10, pady=(4, 2), fill='both', expand=True)

    canvas  = tk.Canvas(container, width=420, height=300, bd=1,
                        relief='sunken', highlightthickness=0)
    vscroll = tk.Scrollbar(container, orient='vertical',   command=canvas.yview)
    hscroll = tk.Scrollbar(container, orient='horizontal', command=canvas.xview)
    canvas.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)

    vscroll.pack(side='right',  fill='y')
    hscroll.pack(side='bottom', fill='x')
    canvas.pack(side='left', fill='both', expand=True)

    inner = tk.Frame(canvas)
    cw    = canvas.create_window((0, 0), window=inner, anchor='nw')

    def _on_inner_resize(event):
        canvas.configure(scrollregion=canvas.bbox('all'))
    inner.bind('<Configure>', _on_inner_resize)
    canvas.bind('<MouseWheel>',
                lambda e: canvas.yview_scroll(int(-1 * e.delta / 120), 'units'))

    # 3-column layout — no fixed width so long names aren't clipped
    for i, cls in enumerate(sorted(class_counts)):
        var = tk.BooleanVar(value=True)
        selected[cls] = var
        var.trace_add('write', _update_count)
        row, col = divmod(i, 3)
        tk.Checkbutton(inner,
                       text=f'{cls}  ({class_counts[cls]:,})',
                       variable=var, anchor='w',
                       font=('Consolas', 9)).grid(
            row=row, column=col, sticky='w', padx=8, pady=1)

    # ── buttons ──────────────────────────────────────────────────────────────
    btn_frame = tk.Frame(dlg)
    btn_frame.pack(pady=6)

    def _set_all(val):
        for v in selected.values():
            v.set(val)

    def ok():
        result[0] = {c for c, v in selected.items() if v.get()}
        dlg.destroy()

    tk.Button(btn_frame, text='All',    width=8,
              command=lambda: _set_all(True)).grid(row=0, column=0, padx=3)
    tk.Button(btn_frame, text='None',   width=8,
              command=lambda: _set_all(False)).grid(row=0, column=1, padx=3)
    tk.Button(btn_frame, text='OK',     width=8, command=ok,
              bg='#4472C4', fg='white').grid(row=0, column=2, padx=3)
    tk.Button(btn_frame, text='Cancel', width=8,
              command=dlg.destroy).grid(row=0, column=3, padx=3)

    dlg.wait_window()
    return result[0]


# ---------------------------------------------------------------------------
# Interactive GUI
# ---------------------------------------------------------------------------

def run_interactive():
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()

    input_paths = list(filedialog.askopenfilenames(
        title='Select OSS XML dump file(s)',
        filetypes=[('All supported', '*.xml *.gz *.zip'),
                   ('XML files', '*.xml'), ('GZ files', '*.gz'),
                   ('ZIP files', '*.zip'), ('All files', '*.*')],
    ))
    if not input_paths:
        print('No files selected.'); return

    print(f'[{ts()}] Scanning for MO classes...')
    t0 = datetime.now()
    class_counts = scan_all_files(input_paths)
    print(f'[{ts()}] Scan done in {fmt_elapsed((datetime.now()-t0).total_seconds())} '
          f'-- classes: {", ".join(sorted(class_counts))}')

    filter_classes = ask_class_selection(class_counts)
    if not filter_classes:
        print('No classes selected.'); return

    # Always offer both formats; warn if xlsb libs are missing
    out_types = [('Excel Workbook',        '*.xlsx'),
                 ('Excel Binary Workbook', '*.xlsb')]
    base    = os.path.basename(input_paths[0]).split('.')[0]
    defname = f'{base}_{datetime.now().strftime("%Y%m%d_%H%M%S")}_dump.xlsx'
    output_path = filedialog.asksaveasfilename(
        title='Save output as', initialdir=os.path.dirname(input_paths[0]),
        initialfile=defname, defaultextension='.xlsx', filetypes=out_types,
    )
    if not output_path:
        print('No output file selected.'); return

    # Early check: xlsb requires pywin32
    if output_path.lower().endswith('.xlsb') and not XLSB_SUPPORTED:
        from tkinter import messagebox
        messagebox.showwarning(
            'pywin32 not installed',
            '.xlsb format requires the pywin32 package.\n\n'
            '  pip install pywin32\n\n'
            'Output will be saved as .xlsx instead.')
        output_path = output_path[:-5] + '.xlsx'

    root.destroy()
    run_conversion(input_paths, output_path, filter_classes or None)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    os.system("title Ankit's XML Parser  [V3 - Blazing Fast]")

    if len(sys.argv) == 1:
        run_interactive()
        input('\nPress Enter to exit...')
        return

    parser = argparse.ArgumentParser(description='OSS XML -> Excel Converter V3 (Blazing Fast)')
    parser.add_argument('inputs',    nargs='+', metavar='FILE')
    parser.add_argument('-o', '--output',  default='')
    parser.add_argument('--classes', default='')
    args = parser.parse_args()

    for path in args.inputs:
        if not os.path.isfile(path):
            print(f'ERROR: File not found: {path}', file=sys.stderr)
            input('\nPress Enter to exit...')
            sys.exit(1)

    if not args.output:
        base = os.path.basename(args.inputs[0]).split('.')[0]
        args.output = f'{base}_{datetime.now().strftime("%Y%m%d_%H%M%S")}_dump.xlsx'

    filter_classes = {c.strip() for c in args.classes.split(',') if c.strip()} or None
    run_conversion(args.inputs, args.output, filter_classes)
    input('\nPress Enter to exit...')


if __name__ == '__main__':
    freeze_support()   # required for PyInstaller + multiprocessing on Windows
    main()
