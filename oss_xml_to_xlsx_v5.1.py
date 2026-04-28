#!/usr/bin/env python3
"""
OSS XML to XLSX Converter  — Version 5.1  (Streaming Write + Low-RAM Assembly)
Converts Nokia OSS RAML XML dump files to Excel format.

Key improvements over V5.0:
  - Streaming worksheet XML: rows written directly to disk in batches — peak RAM stays
    constant regardless of sheet size (no giant in-memory string accumulation)
  - Streaming ZIP assembly: sheets copied into the xlsx zip in chunks instead of being
    fully loaded into RAM again during assembly

Key improvements over V4.2 (carried from V5.0):
  - Pre-read snapshot: summary tools read from memory, never re-read the xlsx (~52 s saved)
  - Parallel XML files in ZIP: multiple XMLs inside one archive parsed simultaneously
  - Both dialogs (summary + save-as) shown during parsing — write starts with zero wait
  - XLSB pre-warm: Excel.Application launched before write, startup cost is free
  - XLSB + summaries run in parallel; 4G and 2G summaries also run in parallel
  - Grand Total shows true wall-clock time (parallelism accounted for correctly)

Key improvements over V3 (carried from V4):
  - Regex-based XML parser (no lxml dependency) — 3-5x faster
  - Direct XLSX XML generation (no openpyxl) — raw string assembly
  - Parallel parse (ThreadPoolExecutor) + parallel write (ProcessPoolExecutor)
  - Zero heavy dependencies for core engine (only stdlib)

Usage (interactive GUI):   python oss_xml_to_xlsx_v5.1.py
Usage (command-line):      python oss_xml_to_xlsx_v5.1.py file1.xml.gz file2.gz -o out.xlsx
                           python oss_xml_to_xlsx_v5.1.py file1.zip --classes BTS,BCF -o out.xlsx
"""

import gzip
import io
import math
import os
import queue
import re
import sys
import zipfile
import argparse
import threading
from collections import defaultdict, OrderedDict
from concurrent.futures import (ThreadPoolExecutor, ProcessPoolExecutor,
                                as_completed, wait as cf_wait, FIRST_COMPLETED)
from datetime import datetime
from multiprocessing import freeze_support

try:
    import win32com.client
    XLSB_SUPPORTED = True
except ImportError:
    XLSB_SUPPORTED = False

AUTHOR = 'Ankit Jain'

# ---------------------------------------------------------------------------
# Config file — persists the user's last MO class selection across runs.
# Stored next to the script (or exe when frozen by PyInstaller).
# ---------------------------------------------------------------------------
import json as _json

def _cfg_path():
    """Return absolute path to XML_Parser_AJ.cfg beside the script/exe."""
    base = (sys.executable if getattr(sys, 'frozen', False)
            else os.path.abspath(__file__))
    return os.path.join(os.path.dirname(base), 'XML_Parser_AJ.cfg')

def load_saved_classes():
    """Return set of previously saved class names, or None if no cfg exists."""
    path = _cfg_path()
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = _json.load(f)
        classes = data.get('selected_classes')
        if isinstance(classes, list):
            return set(classes)
    except (FileNotFoundError, _json.JSONDecodeError, OSError):
        pass
    return None

def save_selected_classes(class_set):
    """Persist the current selection to the cfg file."""
    path = _cfg_path()
    try:
        with open(path, 'w', encoding='utf-8') as f:
            _json.dump({'selected_classes': sorted(class_set)}, f, indent=2)
    except OSError as e:
        print(f'[WARN] Could not save class config: {e}')

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

def _iter_xml_from_zip(zf, prefix=''):
    """
    Recursively yield (xml_bytes, display_name) for every XML found inside a
    ZipFile object, no matter how deeply nested the folder structure is or
    whether intermediate containers are .zip or .gz files.

    prefix  — human-readable path accumulated from outer containers, e.g.
              "outer.zip/inner.zip/"
    """
    for entry in zf.namelist():
        if entry.startswith('__MACOSX') or entry.endswith('/'):
            continue  # skip Mac metadata and directory entries

        low_entry = entry.lower()
        display   = f'{prefix}{entry}'

        with zf.open(entry) as f:
            raw = f.read()

        if low_entry.endswith('.zip'):
            # Nested zip — recurse in-memory
            try:
                inner_zf = zipfile.ZipFile(io.BytesIO(raw))
                yield from _iter_xml_from_zip(inner_zf, prefix=f'{display}/')
            except zipfile.BadZipFile:
                tprint(f'  [WARN] Skipping bad zip entry: {display}')

        elif low_entry.endswith('.gz'):
            # Compressed XML — decompress and yield
            try:
                xml_bytes = gzip.decompress(raw)
                # Strip .gz from display name so it shows as the xml filename
                clean = os.path.basename(entry)
                if clean.lower().endswith('.gz'):
                    clean = clean[:-3]
                yield xml_bytes, f'{prefix}{clean}'
            except OSError:
                tprint(f'  [WARN] Skipping unreadable gz entry: {display}')

        elif low_entry.endswith('.xml'):
            yield raw, display


def iter_xml_streams(filepath):
    """Yield (xml_bytes, display_name) for every XML reachable from filepath."""
    name = os.path.basename(filepath)
    low  = filepath.lower()
    if low.endswith('.zip'):
        with zipfile.ZipFile(filepath, 'r') as zf:
            found = list(_iter_xml_from_zip(zf, prefix=f'{name}/'))
        if not found:
            raise ValueError(f'No XML files found inside {name}')
        yield from found
    elif low.endswith('.gz'):
        with gzip.open(filepath, 'rb') as f:
            yield f.read(), name
    else:
        with open(filepath, 'rb') as f:
            yield f.read(), name


# ---------------------------------------------------------------------------
# Quick scan (regex — same as V2/V3)
# ---------------------------------------------------------------------------

_CLASS_RE = re.compile(rb'managedObject\s[^>]*class="([^"]+)"')

def quick_scan_classes(filepath):
    counts = defaultdict(int)
    # iter_xml_streams handles all nesting (zip/gz/xml) recursively
    for xml_bytes, _ in iter_xml_streams(filepath):
        for m in _CLASS_RE.finditer(xml_bytes):
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
# Single regex handles both <p name="x">value</p>  and  <p name="x"/>
# Group 1 = name,  Group 2 = value (None when self-closing)
_P_ANY_RE = re.compile(
    r'<(?:\w+:)?p\s+name="([^"]*)"(?:>(.*?)</(?:\w+:)?p>|\s*/>)', re.DOTALL)

# For extracting <list name="...">...</list>
_LIST_RE = re.compile(r'<(?:\w+:)?list\s+name="([^"]*)">(.*?)</(?:\w+:)?list>', re.DOTALL)

# For extracting <item>...</item> blocks
_ITEM_RE = re.compile(r'<(?:\w+:)?item>(.*?)</(?:\w+:)?item>', re.DOTALL)

# For extracting bare <p>value</p> tags that have NO name attribute.
# Used as a fallback inside <list> blocks where Nokia omits the name attr,
# e.g.  <list name="spcList"><p>1284</p></list>
_P_BARE_RE = re.compile(r'<(?:\w+:)?p>(.*?)</(?:\w+:)?p>', re.DOTALL)

# Split on </managedObject> to get individual MO blocks
_MO_SPLIT = re.compile(r'</(?:\w+:)?managedObject>')

# Matches self-closing <managedObject .../> (no closing tag).
# Nokia XML sometimes uses self-closing tags for empty MOs (e.g. SMLC class).
# If left as-is, the MO immediately following gets merged into the same split
# block and is silently lost.  We normalise them to paired open+close before
# splitting so every MO becomes its own block.
_MO_SELF_CLOSE_RE = re.compile(r'(<(?:\w+:)?managedObject\s[^>]*)/>')


def try_numeric(text):
    if text is None:
        return None
    s = text.strip()
    if not s:
        return None
    # Preserve leading zeros — "03" must stay "03", not 3.
    # Any numeric-looking string that starts with '0' and has a second digit
    # (e.g. "03", "007", "0123") is kept as text to avoid silent data loss.
    if len(s) > 1 and s[0] == '0' and s[1].isdigit():
        return s
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

    # ── Pass 1: extract <list> blocks ────────────────────────────────────────
    # Collect spans so we can build `remainder` without a second regex scan.
    list_spans = []
    for lm in _LIST_RE.finditer(block):
        list_name = lm.group(1)
        list_body = lm.group(2)
        list_spans.append((lm.start(), lm.end()))

        items = _ITEM_RE.findall(list_body)
        if items:
            record[list_name] = 'List'
            item_fields = OrderedDict()
            for item_body in items:
                for pm in _P_ANY_RE.finditer(item_body):
                    item_fields.setdefault(pm.group(1), []).append(pm.group(2) or '')
            for fname, vals in item_fields.items():
                record[f'Item-{list_name}-{fname}'] = ';'.join(vals)
        else:
            vals = [pm.group(2) or '' for pm in _P_ANY_RE.finditer(list_body)]
            if not vals:
                # Fallback: bare <p>value</p> with no name attribute
                # e.g. <list name="spcList"><p>1284</p></list>
                vals = [pm.group(1).strip() for pm in _P_BARE_RE.finditer(list_body)]
            if vals:
                record[list_name] = 'List;' + ';'.join(vals)

    # ── Build remainder without a second _LIST_RE scan ────────────────────────
    # Stitch together the parts of `block` that fall outside any <list> span.
    if list_spans:
        parts_r, prev = [], 0
        for s, e in list_spans:
            parts_r.append(block[prev:s])
            prev = e
        parts_r.append(block[prev:])
        remainder = ''.join(parts_r)
    else:
        remainder = block

    # ── Pass 2: extract top-level <p> from remainder (single pass) ───────────
    for pm in _P_ANY_RE.finditer(remainder):
        val = pm.group(2)           # None when self-closing
        record[pm.group(1)] = try_numeric(val) if val is not None else None

    return mo_class, hierarchy, record


# ---------------------------------------------------------------------------
# Subprocess worker: parse one batch of pre-split MO block strings.
# Must be a module-level function so it is picklable by ProcessPoolExecutor.
# ---------------------------------------------------------------------------

def _parse_blocks_worker(args):
    """
    Runs in a worker process.
    args = (blocks: list[str], display_name: str, filter_classes: set|None)
    Returns a plain dict  {mo_class: [(hierarchy, record), ...]}
    """
    blocks, display_name, filter_classes = args
    classes_data = {}
    filter_set = filter_classes if filter_classes else None

    for block in blocks:
        if 'managedObject' not in block:
            continue
        if filter_set:
            cm = re.search(r'class="([^"]+)"', block)
            if not cm or cm.group(1) not in filter_set:
                continue
        result = parse_mo_block(block, display_name)
        if result is None:
            continue
        mo_class, hierarchy, record = result
        if mo_class not in classes_data:
            classes_data[mo_class] = []
        classes_data[mo_class].append((hierarchy, record))

    return classes_data


def parse_xml_bytes_v3(data, display_name, filter_classes=None, n_workers=1):
    """
    Parse XML bytes using regex. Returns (classes_data, file_info).

    n_workers > 1  →  split MO blocks across that many worker processes so
                      all available CPU cores are used (bypasses the GIL).
                      Safe to call from the main process or from a thread.
    """
    file_info = {'filename': display_name, 'dateTime': ''}
    t0 = datetime.now()
    tprint(f'  [{ts()}] Parsing {display_name} ...')

    # Decode bytes to string
    text = (data.decode('utf-8', errors='replace')
            if isinstance(data, (bytes, bytearray)) else data)

    # Normalise self-closing <managedObject .../> → <managedObject ...></managedObject>
    # so _MO_SPLIT correctly isolates every MO as its own block.
    # Without this, a self-closing empty MO (e.g. SMLC) and the MO that follows
    # it end up in the same block; parse_mo_block() then picks up only the first
    # opening tag and silently drops the second MO's data (e.g. MAL-10 was lost).
    text = _MO_SELF_CLOSE_RE.sub(r'\1></managedObject>', text)

    # Extract dateTime from <log> element
    log_m = _LOG_DT.search(text)
    if log_m:
        file_info['dateTime'] = log_m.group(1)

    # Split once into all MO blocks; filter immediately to reduce work
    filter_set = filter_classes if filter_classes else None
    all_blocks = []
    for block in _MO_SPLIT.split(text):
        if 'managedObject' not in block:
            continue
        if filter_set:
            cm = re.search(r'class="([^"]+)"', block)
            if not cm or cm.group(1) not in filter_set:
                continue
        all_blocks.append(block)

    n_blocks = len(all_blocks)
    classes_data = defaultdict(list)

    # ── Single-process path ───────────────────────────────────────────────────
    # Use it when: only 1 worker requested, or too few blocks to justify the
    # ~0.5 s process-spawn overhead on Windows.
    if n_workers <= 1 or n_blocks < 2_000:
        for block in all_blocks:
            result = parse_mo_block(block, display_name)
            if result:
                mo_class, hierarchy, record = result
                classes_data[mo_class].append((hierarchy, record))

    # ── Multi-process path ────────────────────────────────────────────────────
    # Divide blocks into n_workers chunks; each chunk runs in its own process.
    # Threads in the main process spawning subprocesses is safe on Windows.
    else:
        chunk_sz = math.ceil(n_blocks / n_workers)
        chunks   = [all_blocks[i: i + chunk_sz]
                    for i in range(0, n_blocks, chunk_sz)]
        args_list = [(chunk, display_name, filter_classes) for chunk in chunks]

        with ProcessPoolExecutor(max_workers=len(chunks)) as ex:
            for partial in ex.map(_parse_blocks_worker, args_list):
                for cls, recs in partial.items():
                    classes_data[cls].extend(recs)

    count   = sum(len(v) for v in classes_data.values())
    elapsed = (datetime.now() - t0).total_seconds()
    mode    = f'{n_workers}-core' if n_workers > 1 else '1-core'
    tprint(f'  [{ts()}] Done    {display_name} -- {count:,} objects '
           f'in {fmt_elapsed(elapsed)}  [{mode}]')
    return classes_data, file_info


def parse_input_file(filepath, filter_classes=None, n_workers=1):
    # Read all XML streams out of the file first (zip may contain many XMLs).
    streams = list(iter_xml_streams(filepath))

    if len(streams) <= 1:
        # Single stream — give it the full core budget.
        return [parse_xml_bytes_v3(xb, dn, filter_classes, n_workers)
                for xb, dn in streams]

    # Multiple streams (e.g. a zip with 2+ XML files) — split the core budget
    # equally and parse all streams in parallel threads.
    # Threads → inner ProcessPoolExecutor is safe on Windows.
    workers_each = max(1, n_workers // len(streams))
    results = [None] * len(streams)

    def _stream_worker(idx, xb, dn):
        results[idx] = parse_xml_bytes_v3(xb, dn, filter_classes, workers_each)

    threads = [threading.Thread(target=_stream_worker, args=(i, xb, dn), daemon=True)
               for i, (xb, dn) in enumerate(streams)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    return results


# ---------------------------------------------------------------------------
# Column ordering + record flattening (same as V2/V3)
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
    col_widths.append((ci, 10)); ci += 1   # File_Name
    col_widths.append((ci, 10)); ci += 1   # Dist_Name
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
    # ── Hot path optimisations ────────────────────────────────────────────────
    # • col_letters: pre-compute all column letter strings once (n_cols values)
    #   instead of calling _col_letter() for every cell (up to 100 M calls).
    # • ri_str: convert row index to string once per row, not once per cell.
    # • Inline _xml_escape: avoids function-call overhead on every string cell.
    # • Omit s="0" style attribute — 0 is the XLSX default, so it's redundant.
    # • parts.extend(row_parts): one C-level extend per row vs N append calls.
    # ─────────────────────────────────────────────────────────────────────────
    col_letters = [_col_letter(i) for i in range(n_cols)]

    for ri, row in enumerate(flat_records, start=3):
        ri_str   = str(ri)
        row_parts = [f'<row r="{ri_str}">']
        for ci, val in enumerate(row):
            if val is None:
                continue
            ref = col_letters[ci] + ri_str
            if isinstance(val, (int, float)):
                row_parts.append(f'<c r="{ref}"><v>{val}</v></c>')
            else:
                s = str(val)
                if '&' in s or '<' in s or '>' in s or '"' in s:
                    s = (s.replace('&', '&amp;')
                          .replace('<', '&lt;')
                          .replace('>', '&gt;')
                          .replace('"', '&quot;'))
                row_parts.append(f'<c r="{ref}" t="inlineStr"><is><t>{s}</t></is></c>')
        row_parts.append('</row>')
        parts.extend(row_parts)

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
    parts.append(_cell_xml(1, row, 'Created with OSS XML Converter  --  V5.1', style=1))
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

_STREAM_BATCH = 2000   # rows flushed per write call — balances call overhead vs RAM


def _stream_worksheet_xml(xml_path, flat_records, all_cols, n_hier):
    """
    Write worksheet XML directly to xml_path, streaming rows in batches.

    V5.1 improvement over V5.0:
      V5.0 built a parts=[] list across ALL rows then did one giant ''.join() + encode
      at the end — a large RAM spike proportional to sheet size.
      V5.1 writes _STREAM_BATCH rows at a time; peak RAM is O(batch) not O(sheet).
    """
    n_cols = len(all_cols)
    has_op = 'operation' in all_cols
    dist_freeze_col = n_hier + (1 if has_op else 0) + 3
    freeze_col_letter = _col_letter(dist_freeze_col)
    freeze_ref        = f'{freeze_col_letter}3'
    col_letters       = [_col_letter(i) for i in range(n_cols)]

    # newline='' prevents Windows CRLF translation (values may contain \n).
    # buffering=1<<20 gives a 1 MB write buffer — fewer OS-level syscalls.
    with open(xml_path, 'w', encoding='utf-8', newline='', buffering=1 << 20) as f:
        w = f.write

        # ── Header ────────────────────────────────────────────────────────────
        w('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
          ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          '<sheetViews><sheetView tabSelected="0" workbookViewId="0">')
        w(f'<pane xSplit="{dist_freeze_col}" ySplit="2" topLeftCell="{freeze_ref}"'
          f' activePane="bottomRight" state="frozen"/>')
        w('</sheetView></sheetViews>')

        # ── Column widths ─────────────────────────────────────────────────────
        w('<cols>')
        ci = 0
        for _ in range(n_hier):
            w(f'<col min="{ci+1}" max="{ci+1}" width="12" customWidth="1"/>'); ci += 1
        if has_op:
            w(f'<col min="{ci+1}" max="{ci+1}" width="10" customWidth="1"/>'); ci += 1
        w(f'<col min="{ci+1}" max="{ci+1}" width="12" customWidth="1"/>'); ci += 1  # id
        w(f'<col min="{ci+1}" max="{ci+1}" width="10" customWidth="1"/>'); ci += 1  # File_Name
        w(f'<col min="{ci+1}" max="{ci+1}" width="10" customWidth="1"/>'); ci += 1  # Dist_Name
        w(f'<col min="{ci+1}" max="{ci+1}" width="10" customWidth="1"/>'); ci += 1  # SW_Version
        for i in range(ci, n_cols):
            w(f'<col min="{i+1}" max="{i+1}" width="18" customWidth="1"/>')
        w('</cols>')

        # ── Sheet data ────────────────────────────────────────────────────────
        # Row 1: Info hyperlink
        w('<sheetData>'
          '<row r="1"><c r="A1" s="2" t="inlineStr"><is><t>Info</t></is></c></row>')

        # Row 2: header row
        hdr = ['<row r="2">']
        for ci, col_name in enumerate(all_cols):
            esc = _xml_escape(col_name)
            hdr.append(f'<c r="{col_letters[ci]}2" s="1" t="inlineStr"><is><t>{esc}</t></is></c>')
        hdr.append('</row>')
        w(''.join(hdr))

        # ── Data rows — batched ───────────────────────────────────────────────
        batch = []
        for ri, row in enumerate(flat_records, start=3):
            ri_str    = str(ri)
            row_parts = [f'<row r="{ri_str}">']
            for ci, val in enumerate(row):
                if val is None:
                    continue
                ref = col_letters[ci] + ri_str
                if type(val) is int or type(val) is float:
                    row_parts.append(f'<c r="{ref}"><v>{val}</v></c>')
                else:
                    s = str(val)
                    if '&' in s or '<' in s or '>' in s or '"' in s:
                        s = (s.replace('&', '&amp;')
                              .replace('<', '&lt;')
                              .replace('>', '&gt;')
                              .replace('"', '&quot;'))
                    row_parts.append(f'<c r="{ref}" t="inlineStr"><is><t>{s}</t></is></c>')
            row_parts.append('</row>')
            batch.append(''.join(row_parts))
            if len(batch) >= _STREAM_BATCH:
                w(''.join(batch))
                batch.clear()
        if batch:
            w(''.join(batch))

        # ── Footer ────────────────────────────────────────────────────────────
        last_col = col_letters[n_cols - 1]
        w('</sheetData>')
        w(f'<autoFilter ref="A2:{last_col}2"/>')
        w("<hyperlinks><hyperlink ref=\"A1\" location=\"'Info'!A1\"/></hyperlinks>")
        w('</worksheet>')


def write_sheet_worker(args):
    """
    Runs in a subprocess.  Streams one MO class sheet XML directly to a temp file.
    V5.1: uses _stream_worksheet_xml — no giant in-memory string accumulation.
    """
    cls, flat_records, all_cols, n_hier, temp_dir = args
    t_start = datetime.now()

    safe_name = cls.replace('/', '_').replace('\\', '_')
    xml_path  = os.path.join(temp_dir, f'{safe_name}.xml')

    _stream_worksheet_xml(xml_path, flat_records, all_cols, n_hier)

    elapsed = (datetime.now() - t_start).total_seconds()
    return cls, xml_path, None, elapsed   # rels always None (internal hyperlinks)


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
            arc_name = f'xl/worksheets/sheet{i+1}.xml'
            if isinstance(xml_src, (bytes, bytearray)):
                out.writestr(arc_name, xml_src)
            else:
                # V5.1: write() reads the file in chunks — avoids loading the
                # entire sheet XML into RAM again during assembly.
                out.write(xml_src, arcname=arc_name,
                          compress_type=zipfile.ZIP_DEFLATED, compresslevel=1)
            if rels_src:
                rels_arc = f'xl/worksheets/_rels/sheet{i+1}.xml.rels'
                if isinstance(rels_src, (bytes, bytearray)):
                    out.writestr(rels_arc, rels_src)
                else:
                    out.write(rels_src, arcname=rels_arc,
                              compress_type=zipfile.ZIP_DEFLATED, compresslevel=1)


# ---------------------------------------------------------------------------
# XLSB conversion
# ---------------------------------------------------------------------------

def convert_to_xlsb(xlsx_path, xlsb_path, _excel=None):
    """Convert xlsx_path → xlsb_path using Excel COM.

    _excel: optional already-running Excel.Application COM object (pre-warmed).
            When supplied the caller owns the object and must call excel.Quit().
            When None (default) this function creates and quits its own instance.

    COM tweaks applied:
      Calculation = -4135  (xlCalculationManual) — skip full recalc on Open
      EnableEvents = False  — suppress add-in / macro events
      UpdateLinks  = 0      — don't prompt/refresh external links on Open
      AddToMru     = False  — don't pollute the MRU list
      EnableAutoRecover = False — skip periodic auto-save overhead
    NOTE: Interactive=False is intentionally omitted — it can cause Excel to
    freeze during SaveAs when a suppressed dialog can't be auto-dismissed.
    """
    # DispatchEx always creates a FRESH, separate Excel process.
    # Dispatch() would hijack an already-open visible Excel window instead.
    owned = _excel is None
    if owned:
        _excel = win32com.client.DispatchEx('Excel.Application')
        _excel.Visible        = False
        _excel.DisplayAlerts  = False
        _excel.ScreenUpdating = False
        try:
            _excel.Calculation = -4135  # xlCalculationManual — skip recalc on open
        except Exception:
            pass   # non-critical optimisation; some headless sessions reject it
        _excel.EnableEvents   = False   # suppress add-in / macro events
    try:
        wb = _excel.Workbooks.Open(os.path.abspath(xlsx_path),
                                   UpdateLinks=0, AddToMru=False)
        wb.EnableAutoRecover = False    # skip periodic auto-save overhead
        wb.SaveAs(os.path.abspath(xlsb_path), FileFormat=50)
        wb.Close(False)
    finally:
        if owned:
            _excel.Quit()


# ---------------------------------------------------------------------------
# Phase 1: Parse  (can run while the UI "Save as" dialog is open)
# ---------------------------------------------------------------------------

def _parse_phase(input_paths, filter_classes, progress=None):
    """
    Parse all input files and return (merged, ordered_fi, t_parse).
    Designed to be called in a background thread so the user can pick
    the output path concurrently.
    """
    n_files = len(input_paths)
    n_cpu   = os.cpu_count() or 4

    print(f'[{ts()}] Parsing {n_files} file(s)  [{n_cpu} logical CPU(s) available]')
    if filter_classes:
        print(f'[{ts()}] Filtering to classes: {", ".join(sorted(filter_classes))}')

    if progress:
        progress.update(phase=f'Parsing {n_files} file(s)...',
                        mode='determinate', maximum=n_files, value=0,
                        status='')

    # ── CPU budget strategy ───────────────────────────────────────────────────
    #
    # Parsing is CPU-bound; the GIL makes threads useless for this work.
    # We must use processes to get real parallelism.
    #
    # Two modes depending on file count vs CPU count:
    #
    # A) Many files  (n_files >= n_cpu)
    #    → one worker process per CPU, each handles one file sequentially.
    #      No intra-file chunking — avoids nested ProcessPoolExecutors
    #      (subprocess spawning subprocess is risky on Windows).
    #
    # B) Few files  (n_files < n_cpu)
    #    → one thread per file (threads are fine as the outer layer),
    #      each thread runs parse_xml_bytes_v3 with intra-file chunking
    #      (n_workers = n_cpu // n_files).  Threads → processes is safe.
    #
    # Minimum blocks to justify process-spawn overhead (~0.5 s on Windows)
    # is handled inside parse_xml_bytes_v3 (falls back if < 2 000 blocks).
    # ─────────────────────────────────────────────────────────────────────────

    all_cd, all_fi = [], []
    done   = [0]
    t_parse = datetime.now()

    if n_files >= n_cpu:
        # Mode A: ProcessPoolExecutor — one process per CPU, round-robin files
        print(f'[{ts()}] Mode A: {n_cpu} file-worker process(es), 1-core each')
        with ProcessPoolExecutor(max_workers=n_cpu) as ex:
            futs = {ex.submit(parse_input_file, p, filter_classes, 1): p
                    for p in input_paths}
            for fut in as_completed(futs):
                for cd, fi in fut.result():
                    all_cd.append(cd)
                    all_fi.append(fi)
                done[0] += 1
                if progress:
                    fname = os.path.basename(futs[fut])
                    progress.update(value=done[0],
                                    status=f'[{done[0]}/{n_files}]  {fname}')
    else:
        # Mode B: ThreadPoolExecutor (outer) + ProcessPoolExecutor (inner chunks)
        workers_per_file = max(1, n_cpu // n_files)
        print(f'[{ts()}] Mode B: {n_files} file thread(s), '
              f'{workers_per_file}-core intra-file chunking each')
        with ThreadPoolExecutor(max_workers=n_files) as ex:
            futs = {ex.submit(parse_input_file, p, filter_classes,
                              workers_per_file): p
                    for p in input_paths}
            for fut in as_completed(futs):
                for cd, fi in fut.result():
                    all_cd.append(cd)
                    all_fi.append(fi)
                done[0] += 1
                if progress:
                    fname = os.path.basename(futs[fut])
                    progress.update(value=done[0],
                                    status=f'[{done[0]}/{n_files}]  {fname}')

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

    return merged, ordered_fi, t_parse


# ---------------------------------------------------------------------------
# Phase 2: Write + Assemble
# ---------------------------------------------------------------------------

def _write_phase(merged, ordered_fi, output_path, t_parse, progress=None,
                 defer_xlsb=False):
    """
    Write parsed data to an xlsx (or xlsb) file.

    defer_xlsb=True  — skip the COM conversion step even if .xlsb was requested.
                       Returns the intermediate .xlsx temp path so the caller can
                       run conversion in a background thread alongside other work.
                       The caller is responsible for converting and deleting the
                       temp file.  The "Done!" summary line is also suppressed so
                       the caller can print it at the right moment.
    """
    import tempfile, shutil

    want_xlsb = output_path.lower().endswith('.xlsb')
    if want_xlsb and not XLSB_SUPPORTED:
        print('WARNING: .xlsb requires pywin32 (pip install pywin32). Falling back to .xlsx.')
        output_path = output_path[:-5] + '.xlsx'
        want_xlsb   = False

    # For xlsb: write intermediate .xlsx to system temp dir (not next to output),
    # so the output folder stays clean and temp is on the fastest available drive.
    if want_xlsb:
        tmp_fd, xlsx_path = tempfile.mkstemp(prefix='oss_v5_', suffix='.xlsx')
        os.close(tmp_fd)
    else:
        xlsx_path = output_path

    class_names = sorted(merged.keys())
    temp_dir = tempfile.mkdtemp(prefix='oss_v4_')
    try:
        print(f'[{ts()}] Writing {len(class_names)} sheets in parallel...')
        t_write = datetime.now()

        sheet_paths = {}   # cls -> (xml_path_or_bytes, rels_path_or_None)

        # Build sheet plan — split oversized classes into LNREL(1), LNREL(2) …
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

        all_sheet_names = [sp[0] for sp in sheet_plan]

        # Info sheet (fast — in main process, keep as bytes since it's tiny)
        print(f'  [{ts()}] Sheet "Info" ...', end='', flush=True)
        t0 = datetime.now()
        info_xml = generate_info_sheet_xml(ordered_fi, all_sheet_names)
        sheet_paths['Info'] = (info_xml, None)
        print(f' done ({fmt_elapsed((datetime.now()-t0).total_seconds())})')

        # Data sheets — submit workers one class at a time to free RAM promptly
        n_workers = min(len(sheet_plan), os.cpu_count() or 4)
        n_total_sheets = len(sheet_plan)
        if not sheet_plan:
            print(f'  [{ts()}] No data sheets to write.')
        else:
            split_count = n_total_sheets - len(class_names)
            split_note  = f'  ({split_count} extra sheet(s) from row-splitting)' if split_count else ''
            print(f'  [{ts()}] Launching {n_workers} parallel workers'
                  f' for {n_total_sheets} sheets{split_note}...')

        if progress:
            progress.update(phase=f'Writing {n_total_sheets} sheet(s)...',
                            mode='determinate', maximum=n_total_sheets, value=0,
                            status='')

        with ProcessPoolExecutor(max_workers=max(n_workers, 1)) as ex:
            futs       = {}
            prev_cls   = None
            flat_cur   = None
            cols_cur   = None
            nhier_cur  = 0

            for sheet_name, cls, start, end in sheet_plan:
                if cls != prev_cls:
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

                chunk = flat_cur[start:end] if (end - start) < len(flat_cur) else flat_cur
                args  = (sheet_name, chunk, cols_cur, nhier_cur, temp_dir)
                futs[ex.submit(write_sheet_worker, args)] = sheet_name
                del chunk

            if flat_cur is not None:
                del flat_cur

            # Poll with a short timeout so progress.tick() fires every ~150 ms,
            # keeping the bar animated even when one sheet takes a long time.
            sheets_done = 0
            pending = set(futs.keys())
            while pending:
                if progress:
                    progress.tick()
                done_set, pending = cf_wait(pending, timeout=0.15,
                                            return_when=FIRST_COMPLETED)
                for fut in done_set:
                    sheet_name = futs[fut]
                    try:
                        name_result, xml_path, rels_path, elapsed = fut.result()
                        sheet_paths[name_result] = (xml_path, rels_path)
                        print(f'  [{ts()}] Sheet "{name_result}" done ({fmt_elapsed(elapsed)})')
                    except Exception as e:
                        print(f'  [{ts()}] Sheet "{sheet_name}" FAILED: {e}')
                    sheets_done += 1
                    if progress:
                        progress.update(
                            value=sheets_done,
                            status=f'[{sheets_done}/{n_total_sheets}]  Sheet "{sheet_name}"'
                        )

        t_write = (datetime.now() - t_write).total_seconds()
        print(f'[{ts()}] All sheets written in {fmt_elapsed(t_write)}')

        # --- Assemble ---
        failed = [n for n in all_sheet_names if n not in sheet_paths]
        if failed:
            print(f'[{ts()}] WARNING: {len(failed)} sheet(s) failed and will be omitted: {", ".join(failed)}')
        print(f'[{ts()}] Assembling {os.path.basename(xlsx_path)} ...', end='', flush=True)
        if progress:
            progress.update(phase='Assembling workbook...', mode='indeterminate', status='')
            progress.tick()
        t0 = datetime.now()
        sheet_order = ['Info'] + all_sheet_names
        assemble_xlsx(sheet_order, sheet_paths, xlsx_path)
        t_asm = (datetime.now() - t0).total_seconds()
        print(f' done ({fmt_elapsed(t_asm)})')

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

    if want_xlsb and not defer_xlsb:
        # Synchronous conversion (CLI path, or caller that doesn't parallelise).
        t0 = datetime.now()
        print(f'[{ts()}] Converting to .xlsb via Excel...')
        if progress:
            progress.update(phase='Converting to .xlsb via Excel...', status='')
            progress.tick()
        try:
            convert_to_xlsb(xlsx_path, output_path)
        finally:
            try:
                os.remove(xlsx_path)
            except OSError:
                pass
        t_xlsb = (datetime.now() - t0).total_seconds()
        print(f'[{ts()}] Conversion done in {fmt_elapsed(t_xlsb)}')
    else:
        t_xlsb = 0.0

    if not (want_xlsb and defer_xlsb):
        # Print the Done! line now — caller handles it when deferring xlsb.
        t_total = t_parse + t_write + t_asm + t_xlsb
        size_mb = os.path.getsize(output_path) / 1024 / 1024
        print(f'\n[{ts()}] Done!  {output_path}  ({size_mb:.1f} MB)')
        line = (f'         Parse: {fmt_elapsed(t_parse)}  |  '
                f'Write: {fmt_elapsed(t_write)}  |  '
                f'Assemble: {fmt_elapsed(t_asm)}')
        if t_xlsb:
            line += f'  |  XLSB: {fmt_elapsed(t_xlsb)}'
        line += f'  |  Total: {fmt_elapsed(t_total)}'
        print(line)

    # When defer_xlsb=True, xlsx_path is the intermediate temp file (not the
    # final xlsb path).  Return it so the caller can kick off conversion.
    returned_path = xlsx_path if (want_xlsb and defer_xlsb) else output_path
    return returned_path, t_write, t_asm, t_xlsb


# ---------------------------------------------------------------------------
# 2G / 4G Summary tool integration
# ---------------------------------------------------------------------------

# MO classes that indicate each technology is present in the output
_4G_CLASSES = {'LNBTS', 'LNBTS_FDD', 'LNBTS_TDD', 'LNCEL', 'LNCEL_FDD', 'LNCEL_TDD'}
_2G_CLASSES = {'BSC', 'BCF', 'BTS', 'TRX'}

# Modules that belong to the 2G/4G tool packages — must be reloaded when
# switching between tools so 2G's network.py isn't confused with 4G's.
_TOOL_MODULE_NAMES = {'xlsx_reader', 'xlsb_reader', 'network', 'reports'}


def _unique_path(path):
    """If *path* already exists, append (1), (2), … before the extension until unique."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 1
    while True:
        candidate = f'{base}({i}){ext}'
        if not os.path.exists(candidate):
            return candidate
        i += 1


def _tool_base_dir():
    """
    Return the directory that contains the 2g_tool/ and 4g_tool/ sub-folders.

    Script mode : tools live at  .../Claude/2g_tool/  and  .../Claude/4g_tool/
                  (parent of the script's own directory).
    Frozen exe  : PyInstaller extracts --add-data bundles into sys._MEIPASS,
                  so 2g_tool/ and 4g_tool/ land directly inside _MEIPASS.
    """
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS                              # PyInstaller temp dir
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # ../


def _clean_tool_modules():
    """
    Remove any previously imported tool-specific modules from sys.modules so
    that a subsequent import picks up the correct package (2G vs 4G).
    """
    for key in list(sys.modules.keys()):
        if key in _TOOL_MODULE_NAMES or key.startswith('reports.'):
            del sys.modules[key]


def _ask_summary_dialog(root, has_4g, has_2g):
    """
    Show a small Toplevel asking whether to generate 2G/4G summary reports.
    Returns (want_4g: bool, want_2g: bool).  Both False if user clicks Skip.
    """
    import tkinter as tk

    result = [False, False]   # [want_4g, want_2g]

    dlg = tk.Toplevel(root)
    dlg.title('Generate Summary Reports?')
    dlg.resizable(False, False)
    dlg.grab_set()

    tk.Label(dlg, text='Summary report(s) can be generated from the parsed data:',
             font=('Calibri', 10), pady=8).pack(padx=20, anchor='w')

    var_4g = tk.BooleanVar(value=has_4g)
    var_2g = tk.BooleanVar(value=has_2g)

    if has_4g:
        tk.Checkbutton(dlg,
                       text='  4G Summary   (LNBTS/LNCEL data found)',
                       variable=var_4g,
                       font=('Calibri', 10)).pack(padx=28, anchor='w')
    if has_2g:
        tk.Checkbutton(dlg,
                       text='  2G Summary   (BTS/BCF/BSC/TRX data found)',
                       variable=var_2g,
                       font=('Calibri', 10)).pack(padx=28, anchor='w')

    tk.Label(dlg, text='', height=1).pack()   # spacer

    def _on_generate():
        result[0] = bool(var_4g.get()) if has_4g else False
        result[1] = bool(var_2g.get()) if has_2g else False
        dlg.destroy()

    def _on_skip():
        dlg.destroy()

    btn_frame = tk.Frame(dlg)
    btn_frame.pack(pady=(0, 12))
    tk.Button(btn_frame, text='Generate', width=12,
              bg='#4472C4', fg='white',
              command=_on_generate).pack(side='left', padx=10)
    tk.Button(btn_frame, text='Skip', width=10,
              command=_on_skip).pack(side='left', padx=10)

    # Centre on screen
    dlg.update_idletasks()
    sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
    w, h   = dlg.winfo_reqwidth(), dlg.winfo_reqheight()
    dlg.geometry(f'+{(sw - w) // 2}+{(sh - h) // 2}')

    dlg.wait_window()
    return result[0], result[1]


def _run_4g_summary(input_file, output_path, pre_read=None):
    """
    Load the 4G tool from ../4g_tool/, build the summary, write to *output_path*.
    If *pre_read* is supplied (dict of sheet_name -> rows), skip the xlsx read.
    Returns True on success, False on error.
    """
    tool_dir = os.path.join(_tool_base_dir(), '4g_tool')
    if not os.path.isdir(tool_dir):
        print(f'[{ts()}] 4G: tool directory not found: {tool_dir}')
        return False

    _clean_tool_modules()
    if tool_dir in sys.path:
        sys.path.remove(tool_dir)
    sys.path.insert(0, tool_dir)

    try:
        from network import Network                # 4g_tool/network.py
        from reports.lnbts_summary import build   # 4g_tool/reports/lnbts_summary.py

        if pre_read is not None:
            sheets = pre_read
        else:
            from xlsx_reader import read_xlsx
            NEEDED = ['LNBTS', 'LNBTS_FDD', 'LNBTS_TDD', 'LNCEL', 'LNCEL_FDD',
                      'LNCEL_TDD', 'IRFIM', 'LNHOIF', 'SIB', 'REDRT', 'CAPR']
            print(f'[{ts()}] 4G: Reading {os.path.basename(input_file)} ...')
            sheets = read_xlsx(input_file, sheet_names=NEEDED,
                               progress_fn=lambda m: tprint(f'  [4G] {m}'))

        net   = Network(sheets)
        print(f'[{ts()}] 4G: Building summary -> {os.path.basename(output_path)} ...')
        count = build(net, output_path,
                      progress_fn=lambda m: tprint(f'  [4G] {m}'))
        print(f'[{ts()}] 4G: Done — {count} LNBTS written to '
              f'{os.path.basename(output_path)}')
        return True

    except Exception as exc:
        print(f'[{ts()}] 4G: ERROR — {exc}')
        import traceback; traceback.print_exc()
        return False


def _run_2g_summary(input_file, output_path, pre_read=None):
    """
    Load the 2G tool from ../2g_tool/, build the summary, write to *output_path*.
    If *pre_read* is supplied (dict of sheet_name -> rows), skip the xlsx read.
    Returns True on success, False on error.
    """
    tool_dir = os.path.join(_tool_base_dir(), '2g_tool')
    if not os.path.isdir(tool_dir):
        print(f'[{ts()}] 2G: tool directory not found: {tool_dir}')
        return False

    _clean_tool_modules()
    if tool_dir in sys.path:
        sys.path.remove(tool_dir)
    sys.path.insert(0, tool_dir)

    try:
        from network import Network                # 2g_tool/network.py
        from reports.cell_summary import build    # 2g_tool/reports/cell_summary.py

        if pre_read is not None:
            sheets = pre_read
        else:
            from xlsx_reader import read_xlsx
            NEEDED = ['BSC', 'BCF', 'BTS', 'TRX', 'ADCE', 'ADJW', 'ADJL', 'HOC', 'POC', 'MAL']
            print(f'[{ts()}] 2G: Reading {os.path.basename(input_file)} ...')
            sheets = read_xlsx(input_file, sheet_names=NEEDED,
                               progress_fn=lambda m: tprint(f'  [2G] {m}'))

        net   = Network(sheets)
        print(f'[{ts()}] 2G: Building summary -> {os.path.basename(output_path)} ...')
        # Only generate ADCE-dependent sheets (One-Way ADCE, Discrepant ADCE,
        # Co-Site Missing Neighbours) when ADCE data was actually parsed.
        has_adce = bool(sheets.get('ADCE'))
        count = build(net, output_path,
                      progress_fn=lambda m: tprint(f'  [2G] {m}'),
                      neighbour_checks=has_adce)
        print(f'[{ts()}] 2G: Done — {count} cells written to '
              f'{os.path.basename(output_path)}')
        return True

    except Exception as exc:
        print(f'[{ts()}] 2G: ERROR — {exc}')
        import traceback; traceback.print_exc()
        return False


def _post_process_summaries(output_file, class_names, root):
    """
    After the main xlsx/xlsb is written, offer to generate 2G/4G summaries.

    output_file  — final output path (.xlsx or .xlsb)
    class_names  — iterable of MO class names present in the output
    root         — tkinter root window (for Toplevel dialog parenting)

    The output file is read in a background thread immediately (using the union
    of all 2G + 4G needed sheets) while the summary dialog is shown.  By the
    time the user clicks OK, the read is often already done.
    """
    class_set = set(class_names)
    has_4g = (bool(class_set & {'LNBTS', 'LNBTS_FDD', 'LNBTS_TDD'}) and
              bool(class_set & {'LNCEL', 'LNCEL_FDD', 'LNCEL_TDD'}))
    has_2g = {'BTS', 'BCF', 'BSC', 'TRX'}.issubset(class_set)

    if not has_4g and not has_2g:
        return 0.0, 0.0

    # Load a reader module (2G and 4G tools share the same xlsx_reader interface).
    tool_dir_2g = os.path.join(_tool_base_dir(), '2g_tool')
    _clean_tool_modules()
    if tool_dir_2g in sys.path:
        sys.path.remove(tool_dir_2g)
    sys.path.insert(0, tool_dir_2g)

    try:
        from xlsx_reader import read_xlsx
    except Exception as exc:
        print(f'[{ts()}] Summary: failed to load xlsx_reader — {exc}')
        return 0.0, 0.0

    # Union of all sheets that either summary tool might need.
    NEEDED_ALL = [
        'BSC', 'BCF', 'BTS', 'TRX', 'ADCE', 'ADJW', 'ADJL', 'HOC', 'POC', 'MAL',
        'LNBTS', 'LNBTS_FDD', 'LNBTS_TDD', 'LNCEL', 'LNCEL_FDD', 'LNCEL_TDD',
        'IRFIM', 'LNHOIF', 'SIB', 'REDRT', 'CAPR',
    ]

    # Start reading the xlsx in background — overlaps with user time on dialog.
    _read_result = {}
    _read_error  = [None]

    def _bg_read():
        try:
            print(f'[{ts()}] Summary: Reading {os.path.basename(output_file)} ...')
            sheets = read_xlsx(output_file, sheet_names=NEEDED_ALL,
                               progress_fn=lambda m: tprint(f'  [Read] {m}'))
            _read_result['sheets'] = sheets
        except Exception as exc:
            _read_error[0] = exc

    read_thread = threading.Thread(target=_bg_read, daemon=True, name='sum-read')
    read_thread.start()

    # Show the dialog while reading runs in parallel.
    want_4g, want_2g = _ask_summary_dialog(root, has_4g, has_2g)

    if not want_4g and not want_2g:
        print(f'[{ts()}] Summary generation skipped.')
        return 0.0, 0.0

    # Wait for reading to finish (may already be done by the time user clicks OK).
    read_thread.join()

    if _read_error[0]:
        print(f'[{ts()}] Summary: Read error — {_read_error[0]}')
        import traceback; traceback.print_exc()
        return 0.0, 0.0

    pre_read = _read_result['sheets']
    base = os.path.splitext(output_file)[0]
    t_4g = t_2g = 0.0

    if want_4g:
        out_4g = _unique_path(f'{base}_4G_Summary.xlsx')
        t0 = datetime.now()
        _run_4g_summary(output_file, out_4g, pre_read=pre_read)
        t_4g = (datetime.now() - t0).total_seconds()

    if want_2g:
        out_2g = _unique_path(f'{base}_2G_Summary.xlsx')
        t0 = datetime.now()
        _run_2g_summary(output_file, out_2g, pre_read=pre_read)
        t_2g = (datetime.now() - t0).total_seconds()

    return t_4g, t_2g


# ---------------------------------------------------------------------------
# Core conversion  (sequential — used by CLI path)
# ---------------------------------------------------------------------------

def run_conversion(input_paths, output_path, filter_classes=None):
    t_total = datetime.now()
    print(f'[{ts()}] Output: {output_path}')
    merged, ordered_fi, t_parse = _parse_phase(input_paths, filter_classes)
    output_path, _, _, _ = _write_phase(merged, ordered_fi, output_path, t_parse)
    t_total = (datetime.now() - t_total).total_seconds()
    print(f'         Wall clock total: {fmt_elapsed(t_total)}')
    return output_path


# ---------------------------------------------------------------------------
# Class selection dialog (same as V2/V3)
# ---------------------------------------------------------------------------

def ask_class_selection(class_counts):
    import tkinter as tk
    selected = {}
    result    = [None]
    n_total   = len(class_counts)
    total_obj = sum(class_counts.values())

    # Load last-used class selection from config (None = first run, tick all)
    saved = load_saved_classes()
    if saved is not None:
        n_pre = sum(1 for c in class_counts if c in saved)
        cfg_note = f'  |  {n_pre} pre-selected from last run'
    else:
        cfg_note = ''

    dlg = tk.Toplevel()
    dlg.title('Select MO Classes')
    dlg.resizable(True, True)
    dlg.grab_set()

    # ── top summary ──────────────────────────────────────────────────────────
    tk.Label(dlg, text=f'{n_total} MO classes  |  {total_obj:,} total objects{cfg_note}',
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
        # Pre-tick based on saved config; fall back to True (all) if no config
        pre_checked = (cls in saved) if saved is not None else True
        var = tk.BooleanVar(value=pre_checked)
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
        chosen = {c for c, v in selected.items() if v.get()}
        result[0] = chosen
        save_selected_classes(chosen)   # persist for next run
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
# Progress window  (thread-safe, polls a queue via after())
# ---------------------------------------------------------------------------

class ProgressWindow:
    """
    A small Toplevel window that shows phase / progress bar / status text.
    Thread-safe: update() can be called from any thread.
    tick()  must be called from the main thread during long synchronous loops
            (write phase) to keep the UI responsive.
    """

    def __init__(self, parent):
        import tkinter as tk
        import tkinter.ttk as ttk

        self._q        = queue.Queue()
        self._closed   = False
        self._t0       = datetime.now()
        self._after_id = None   # set by _poll; cancelled by _on_close

        self.win = tk.Toplevel(parent)
        self.win.title("OSS XML Parser  V4")
        self.win.resizable(False, False)
        self.win.protocol('WM_DELETE_WINDOW', lambda: None)  # locked while working

        pad = dict(padx=18, pady=3)

        # Phase label (bold)
        self._phase_var = tk.StringVar(value='Starting...')
        tk.Label(self.win, textvariable=self._phase_var,
                 font=('Calibri', 11, 'bold'), anchor='w').pack(fill='x', **pad)

        # Progress bar
        self._pb      = ttk.Progressbar(self.win, length=460, mode='indeterminate')
        self._pb_mode = 'indeterminate'
        self._pb.pack(fill='x', padx=18, pady=2)
        self._pb.start(12)

        # Status line
        self._status_var = tk.StringVar(value='')
        tk.Label(self.win, textvariable=self._status_var,
                 font=('Consolas', 9), fg='#555', anchor='w',
                 wraplength=460, justify='left').pack(fill='x', **pad)

        # Elapsed time (right-aligned)
        self._elapsed_var = tk.StringVar(value='Elapsed: 0.0s')
        tk.Label(self.win, textvariable=self._elapsed_var,
                 font=('Calibri', 9), fg='#888', anchor='e').pack(fill='x', padx=18)

        # Close button — hidden until done
        self._btn_frame = tk.Frame(self.win)
        tk.Button(self._btn_frame, text='  Close  ', width=10,
                  bg='#4472C4', fg='white',
                  command=self._on_close).pack()

        self._tk = tk   # store for methods that need it after __init__

        # Size & center
        w, h = 500, 155
        self.win.update_idletasks()
        sw = self.win.winfo_screenwidth()
        sh = self.win.winfo_screenheight()
        self.win.geometry(f'{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}')

        self._poll()

    # ── internal ─────────────────────────────────────────────────────────────

    def _on_close(self):
        self._closed = True
        # Cancel the pending after() before destroying — prevents the
        # "invalid command name" TclError that fires when a scheduled
        # callback runs against an already-destroyed widget.
        try:
            self.win.after_cancel(self._after_id)
        except Exception:
            pass
        self.win.destroy()

    def _poll(self):
        if self._closed:
            return
        self._drain()
        elapsed = (datetime.now() - self._t0).total_seconds()
        self._elapsed_var.set(f'Elapsed: {fmt_elapsed(elapsed)}')
        self._after_id = self.win.after(100, self._poll)

    def _drain(self):
        try:
            while True:
                self._apply(self._q.get_nowait())
        except queue.Empty:
            pass

    def _apply(self, msg):
        if 'phase' in msg:
            self._phase_var.set(msg['phase'])
        if 'status' in msg:
            self._status_var.set(msg['status'])
        if 'mode' in msg and msg['mode'] != self._pb_mode:
            self._pb_mode = msg['mode']
            self._pb.config(mode=msg['mode'])
            if msg['mode'] == 'indeterminate':
                self._pb.start(12)
            else:
                self._pb.stop()
        if 'maximum' in msg:
            self._pb.config(maximum=msg['maximum'])
        if 'value' in msg:
            self._pb['value'] = msg['value']
        if msg.get('_done'):
            self._show_done(msg)

    def _show_done(self, msg):
        self._pb.stop()
        self._pb.config(mode='determinate', maximum=100, value=100)
        self._phase_var.set(msg.get('phase', 'Done!'))
        self._status_var.set(msg.get('status', ''))
        elapsed = (datetime.now() - self._t0).total_seconds()
        self._elapsed_var.set(f'Total time: {fmt_elapsed(elapsed)}')
        self.win.protocol('WM_DELETE_WINDOW', self._on_close)
        self._btn_frame.pack(pady=(4, 8))
        self.win.geometry('')   # auto-resize to fit button

    # ── public API ────────────────────────────────────────────────────────────

    def update(self, **kwargs):
        """Thread-safe: put a message on the queue from any thread."""
        self._q.put(kwargs)

    def tick(self):
        """
        Drain the queue and refresh the window.
        Call from the MAIN thread inside long synchronous loops (write phase)
        so the bar animates even though we never return to the event loop.
        Uses win.update() — stronger than update_idletasks — to drive the
        indeterminate spinner's internal timer events.
        """
        self._drain()
        elapsed = (datetime.now() - self._t0).total_seconds()
        self._elapsed_var.set(f'Elapsed: {fmt_elapsed(elapsed)}')
        try:
            self.win.update()   # drives animation timers + redraws
        except Exception:
            pass                # window may be destroyed; ignore

    def done(self, phase='Done!', status=''):
        """Signal completion (thread-safe)."""
        self._q.put({'_done': True, 'phase': phase, 'status': status})

    def wait(self):
        """Block the main thread until the user closes the window."""
        if not self._closed:
            self.win.wait_window()


# ---------------------------------------------------------------------------
# Interactive GUI  — V4 key change: parsing overlaps with "Save as" dialog
# ---------------------------------------------------------------------------

def run_interactive():
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()

    # ── Step 1: Select input files ───────────────────────────────────────────
    input_paths = list(filedialog.askopenfilenames(
        title='Select OSS XML dump file(s)',
        filetypes=[('All supported', '*.xml *.gz *.zip'),
                   ('XML files', '*.xml'), ('GZ files', '*.gz'),
                   ('ZIP files', '*.zip'), ('All files', '*.*')],
    ))
    if not input_paths:
        print('No files selected.'); return

    # ── Step 2: Quick class scan ─────────────────────────────────────────────
    print(f'[{ts()}] Scanning for MO classes...')
    t0 = datetime.now()
    class_counts = scan_all_files(input_paths)
    print(f'[{ts()}] Scan done in {fmt_elapsed((datetime.now()-t0).total_seconds())} '
          f'-- classes: {", ".join(sorted(class_counts))}')

    # ── Step 3: User selects classes ─────────────────────────────────────────
    filter_classes = ask_class_selection(class_counts)
    if not filter_classes:
        print('No classes selected.'); return

    # ── Step 4: Kick off parsing in background thread immediately ────────────
    # Both dialogs below (summary + save-as) run while parsing is in progress.
    # Sort by file size descending so the heaviest file gets a head start.
    input_paths.sort(key=lambda p: os.path.getsize(p), reverse=True)

    _parse_result = {}
    _parse_error  = [None]

    def _bg_parse():
        try:
            merged, ordered_fi, t_parse = _parse_phase(input_paths, filter_classes)
            _parse_result['merged']  = merged
            _parse_result['fi']      = ordered_fi
            _parse_result['t_parse'] = t_parse
        except Exception as exc:
            _parse_error[0] = exc

    parse_thread = threading.Thread(target=_bg_parse, daemon=True, name='bg-parse')
    parse_thread.start()
    _t_start = datetime.now()   # wall-clock from parse start (excludes file-browsing)

    # ── Step 4b: Summary options — asked NOW, while parsing runs ─────────────
    # filter_classes already tells us what data is coming — no parsed data needed.
    # Asking here (before save-as) means BOTH dialogs overlap with parsing,
    # so write starts immediately after the save-as closes with zero extra wait.
    _has_4g = (bool(filter_classes & {'LNBTS', 'LNBTS_FDD', 'LNBTS_TDD'}) and
               bool(filter_classes & {'LNCEL', 'LNCEL_FDD', 'LNCEL_TDD'}))
    _has_2g = {'BTS', 'BCF', 'BSC', 'TRX'}.issubset(filter_classes)
    want_4g = want_2g = False
    if _has_4g or _has_2g:
        want_4g, want_2g = _ask_summary_dialog(root, _has_4g, _has_2g)
        if not want_4g and not want_2g:
            print(f'[{ts()}] Summary generation skipped.')

    # ── Step 5: Ask for output path (also runs while parsing is happening) ────
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

    # ── Step 6: Progress window (currently dormant — set to None to disable) ───
    # To re-enable: replace `progress = None` with `progress = ProgressWindow(root)`
    # The ProgressWindow class is fully implemented above and ready to use.
    progress = None
    # progress = ProgressWindow(root)

    # ── Step 7: Wait for parse thread ────────────────────────────────────────
    if parse_thread.is_alive():
        print(f'[{ts()}] Output path selected — waiting for parse to finish...')
        while parse_thread.is_alive():
            if progress:
                progress.tick()
            parse_thread.join(timeout=0.05)
    else:
        if progress:
            progress.update(phase='Parse complete — preparing to write...',
                            mode='indeterminate', status='')
            progress.tick()

    if _parse_error[0]:
        raise _parse_error[0]

    # Capture class names NOW — _write_phase() pops from merged, leaving it empty
    class_names = sorted(_parse_result['merged'].keys())

    # ── Step 8: Pre-read snapshot (before write pops merged) ────────────────
    # Grab references to the ~20 sheets the summary tools need directly from
    # the already-parsed merged dict — no data copied, just dict references.
    # We snapshot ALL candidate sheets now (cost is negligible) so write can
    # start immediately in the next step without waiting for the dialog answer.
    _SUMMARY_SHEETS = {
        'BSC', 'BCF', 'BTS', 'TRX', 'ADCE', 'ADJW', 'ADJL', 'HOC', 'POC', 'MAL',
        'LNBTS', 'LNBTS_FDD', 'LNBTS_TDD', 'LNCEL', 'LNCEL_FDD', 'LNCEL_TDD',
        'IRFIM', 'LNHOIF', 'SIB', 'REDRT', 'CAPR',
    }
    # Each entry in merged[cls] is a (hierarchy, record) tuple.
    # Hierarchy is an OrderedDict of {ClassName: id} pairs from parse_dist_name()
    # e.g. {'MRBTS': 1, 'LNBTS': 1, 'LNCEL': 1} for an LTE cell.
    # The 4G/2G summary tools (network.py) need these hierarchy fields as explicit
    # keys in the record dict — the same way xlsx_reader supplies them as columns.
    # Merge hierarchy INTO each record so summary tools find MRBTS/LNBTS/LNCEL etc.
    # Record fields win on collision (they carry the actual parameter values).
    pre_read = {k: [{**dict(hier), **rec} for hier, rec in v]
                for k, v in _parse_result['merged'].items()
                if k in _SUMMARY_SHEETS}

    # ── Step 9: XLSB pre-warm + Write + Assemble ─────────────────────────────
    #
    # Both dialogs are already answered (Steps 4b + 5).
    # Write starts immediately — no dialog blocking.
    #
    # ① XLSB pre-warm: Excel.exe launches NOW, warms up during write (~12–30 s)
    # ② Write runs on main thread
    # ③ After write, signal XLSB — it opens the file on already-warm Excel
    # ④ Summaries run on main thread while XLSB converts in background
    #
    want_xlsb    = output_path.lower().endswith('.xlsb')
    final_path   = output_path
    summary_base = os.path.splitext(final_path)[0]

    # ── ① XLSB pre-warm thread ───────────────────────────────────────────────
    # DispatchEx always creates a FRESH separate Excel.exe process — never
    # interferes with any Excel windows already open on the user's desktop.
    # NOTE: Interactive=False intentionally omitted — can freeze SaveAs.
    t_xlsb_box      = [0.0]
    _xlsb_ready     = threading.Event()
    _xlsb_paths_box = [None]
    xlsb_thread     = None

    if want_xlsb and XLSB_SUPPORTED:
        def _bg_xlsb():
            import pythoncom
            pythoncom.CoInitialize()

            # Phase A: launch Excel immediately (warms up during write phase).
            # Only Visible + DisplayAlerts are safe right after DispatchEx.
            try:
                excel = win32com.client.DispatchEx('Excel.Application')
                excel.Visible       = False
                excel.DisplayAlerts = False
            except Exception as exc:
                print(f'[{ts()}] XLSB pre-warm ERROR — {exc}')
                pythoncom.CoUninitialize()
                return

            # Phase B: wait for write to finish, then convert.
            # t0 starts HERE — after the wait — so the reported time is only
            # actual conversion time, not the write-phase wait.
            _xlsb_ready.wait()
            xlsx_path, xlsb_path = _xlsb_paths_box[0]
            t0 = datetime.now()   # ← timer starts now, not at thread start
            try:
                excel.ScreenUpdating = False
                excel.Calculation    = -4135   # xlCalculationManual
                excel.EnableEvents   = False
            except Exception:
                pass   # non-critical

            print(f'[{ts()}] Converting to .xlsb via Excel (pre-warmed)...')
            try:
                convert_to_xlsb(xlsx_path, xlsb_path, _excel=excel)
            except Exception as exc:
                print(f'[{ts()}] XLSB ERROR — {exc}')
            finally:
                try:
                    excel.Quit()
                except Exception:
                    pass
                try:
                    os.remove(xlsx_path)
                except OSError:
                    pass
                pythoncom.CoUninitialize()

            t_xlsb_box[0] = (datetime.now() - t0).total_seconds()
            print(f'[{ts()}] XLSB conversion done in {fmt_elapsed(t_xlsb_box[0])}')

        xlsb_thread = threading.Thread(target=_bg_xlsb, daemon=True, name='xlsb')
        xlsb_thread.start()   # Excel warms up while write runs below

    # ── ② Write + assemble ───────────────────────────────────────────────────
    intermediate_path, t_write, t_asm, _ = _write_phase(
        _parse_result['merged'], _parse_result['fi'],
        output_path, _parse_result['t_parse'], progress=progress,
        defer_xlsb=True)

    # Signal XLSB pre-warm — Excel already warm, file is ready now.
    if xlsb_thread:
        _xlsb_paths_box[0] = (intermediate_path, final_path)
        _xlsb_ready.set()

    # ── Step 10: Summaries on main thread + wait for XLSB ────────────────────
    # Summary build and XLSB conversion run simultaneously.
    t_4g = t_2g = 0.0
    if want_4g or want_2g:
        if want_4g:
            out_4g = _unique_path(f'{summary_base}_4G_Summary.xlsx')
            t0 = datetime.now()
            _run_4g_summary(final_path, out_4g, pre_read=pre_read)
            t_4g = (datetime.now() - t0).total_seconds()
        if want_2g:
            out_2g = _unique_path(f'{summary_base}_2G_Summary.xlsx')
            t0 = datetime.now()
            _run_2g_summary(final_path, out_2g, pre_read=pre_read)
            t_2g = (datetime.now() - t0).total_seconds()

    # Wait for XLSB to finish (may already be done if summary took longer).
    if xlsb_thread:
        xlsb_thread.join()

    t_xlsb = t_xlsb_box[0]

    # Print the "Done!" line now that both outputs are complete.
    # Guard against xlsb conversion failure (file may not exist if Excel errored).
    if os.path.isfile(final_path):
        report_path = final_path
    else:
        report_path = intermediate_path   # fall back to the temp xlsx
        print(f'[{ts()}] WARNING: XLSB output not created — reporting xlsx instead.')
    size_mb = os.path.getsize(report_path) / 1024 / 1024
    print(f'\n[{ts()}] Done!  {report_path}  ({size_mb:.1f} MB)')
    line = (f'         Parse: {fmt_elapsed(_parse_result["t_parse"])}  |  '
            f'Write: {fmt_elapsed(t_write)}  |  '
            f'Assemble: {fmt_elapsed(t_asm)}')
    if t_xlsb:
        line += f'  |  XLSB: {fmt_elapsed(t_xlsb)}'
    line += f'  |  Total: {fmt_elapsed(_parse_result["t_parse"] + t_write + t_asm + t_xlsb)}'
    print(line)

    if progress:
        progress.wait()

    # ── Grand total ───────────────────────────────────────────────────────────
    # Wall-clock total = true elapsed since the tool opened (includes all
    # parallelism correctly — XLSB and summaries overlap, so summing phases
    # would over-count; wall clock never lies).
    t_parse    = _parse_result['t_parse']
    t_wall_end = (datetime.now() - _t_start).total_seconds()
    grand_line = (f'\n[{ts()}] GRAND TOTAL  Wall clock: {fmt_elapsed(t_wall_end)}'
                  f'  (Parse: {fmt_elapsed(t_parse)}'
                  f'  |  Write: {fmt_elapsed(t_write)}'
                  f'  |  Assemble: {fmt_elapsed(t_asm)}')
    if t_xlsb:
        grand_line += f'  |  XLSB: {fmt_elapsed(t_xlsb)}'
    if t_4g:
        grand_line += f'  |  4G Summary: {fmt_elapsed(t_4g)}'
    if t_2g:
        grand_line += f'  |  2G Summary: {fmt_elapsed(t_2g)}'
    grand_line += ')'
    print(grand_line)

    root.destroy()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def _wait_for_keypress():
    """Wait for a real keypress — drains any stdin buffered by Tkinter dialogs first."""
    try:
        import msvcrt
        # Drain any keys that Tkinter button-clicks may have left in the buffer.
        while msvcrt.kbhit():
            msvcrt.getch()
        print('\nPress any key to exit...')
        msvcrt.getch()
    except ImportError:
        # Non-Windows fallback (shouldn't happen for this exe)
        input('\nPress Enter to exit...')


def main():
    os.system("title Ankit's XML Parser  [V5.1 - Streaming Write]")

    if len(sys.argv) == 1:
        try:
            run_interactive()
        except Exception as _e:
            print(f'\n[ERROR] {_e}')
        finally:
            _wait_for_keypress()
        return

    parser = argparse.ArgumentParser(description='OSS XML -> Excel Converter V5.1 (Streaming Write + Low-RAM Assembly)')
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
