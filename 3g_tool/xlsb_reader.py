"""
xlsb_reader.py
Reads Nokia-format .xlsb parameter dump files.
Sheet structure: Row 1 = ignored, Row 2 = headers (trimmed), Row 3+ = data.
Dist_Name is the unique key per sheet (first occurrence kept).
"""

import struct
import zipfile


# ---------------------------------------------------------------------------
# BIFF12 low-level helpers
# ---------------------------------------------------------------------------

def _read_rt(data, pos):
    b = data[pos]
    if b & 0x80:
        return ((data[pos + 1] << 7) | (b & 0x7F)), pos + 2
    return b, pos + 1


def _read_varint(data, pos):
    b = data[pos]
    if b & 0x80 == 0:
        return b, pos + 1
    b2 = data[pos + 1]
    if b2 & 0x80 == 0:
        return ((b & 0x7F) | (b2 << 7)), pos + 2
    b3 = data[pos + 2]
    if b3 & 0x80 == 0:
        return ((b & 0x7F) | ((b2 & 0x7F) << 7) | (b3 << 14)), pos + 3
    b4 = data[pos + 3]
    return ((b & 0x7F) | ((b2 & 0x7F) << 7) | ((b3 & 0x7F) << 14) | (b4 << 21)), pos + 4


def _decode_rk(rk):
    """Decode BIFF12 RK compressed number."""
    f_x100 = rk & 1
    f_int  = (rk >> 1) & 1
    vb     = rk & 0xFFFFFFFC
    val    = (vb >> 2) if f_int else struct.unpack('<d', struct.pack('<Q', vb << 32))[0]
    return val / 100 if f_x100 else val


def _parse_sst_item(rec_data):
    """Parse a BrtSSTItem record into a Python string."""
    if len(rec_data) < 5:
        return ''
    str_len = struct.unpack_from('<I', rec_data, 1)[0]
    if 0 < str_len < 32768:
        try:
            return rec_data[5:5 + str_len * 2].decode('utf-16-le')
        except Exception:
            pass
    return ''


# ---------------------------------------------------------------------------
# Shared string table
# ---------------------------------------------------------------------------

def _load_shared_strings(xlsb_path):
    strings = []
    with open(xlsb_path, "rb") as _f, zipfile.ZipFile(_f) as z:
        with z.open('xl/sharedStrings.bin') as f:
            buf = b''
            while True:
                chunk = f.read(5 * 1024 * 1024)
                if not chunk:
                    break
                buf += chunk
                p = 0
                while p < len(buf) - 20:
                    try:
                        rt, np = _read_rt(buf, p)
                        size, np = _read_varint(buf, np)
                        if np + size > len(buf):
                            break
                        if rt == 0x0013:
                            strings.append(_parse_sst_item(buf[np:np + size]))
                        p = np + size
                    except Exception:
                        p += 1
                buf = buf[p:]
    return strings


# ---------------------------------------------------------------------------
# Sheet bin reader
# ---------------------------------------------------------------------------

def _read_sheet_bin(xlsb_path, bin_name, ss):
    """
    Read a worksheet .bin file from inside the xlsb zip.
    Returns dict: {row_index: {col_index: value}}
    All cell records have layout: col(4) + iStyleRef(4) + value
    """
    rows = {}
    current_row = -1

    with open(xlsb_path, "rb") as _f, zipfile.ZipFile(_f) as z:
        with z.open(f'xl/worksheets/{bin_name}') as f:
            buf = b''
            while True:
                chunk = f.read(10 * 1024 * 1024)
                if not chunk:
                    break
                buf += chunk
                p = 0
                while p < len(buf) - 4:
                    try:
                        rt, np = _read_rt(buf, p)
                        size, np = _read_varint(buf, np)
                        if np + size > len(buf):
                            break
                        rec = buf[np:np + size]
                        p = np + size
                    except Exception:
                        p += 1
                        continue

                    if rt == 0:  # BrtRowHdr
                        current_row = struct.unpack_from('<I', rec, 0)[0]
                        if current_row not in rows:
                            rows[current_row] = {}

                    elif current_row >= 0 and len(rec) >= 4:
                        col = struct.unpack_from('<I', rec, 0)[0]

                        if rt == 7 and len(rec) >= 12:    # BrtCellIsst
                            idx = struct.unpack_from('<I', rec, 8)[0]
                            rows[current_row][col] = ss[idx] if idx < len(ss) else ''

                        elif rt == 2 and len(rec) >= 12:  # BrtCellRk
                            rows[current_row][col] = _decode_rk(
                                struct.unpack_from('<I', rec, 8)[0])

                        elif rt == 5 and len(rec) >= 16:  # BrtCellReal
                            rows[current_row][col] = struct.unpack_from('<d', rec, 8)[0]

                        elif rt == 4 and len(rec) >= 9:   # BrtCellBool
                            rows[current_row][col] = int(rec[8])

                        elif rt == 6 and len(rec) >= 12:  # BrtCellSt (inline string)
                            cch = struct.unpack_from('<I', rec, 8)[0]
                            try:
                                rows[current_row][col] = rec[12:12 + cch * 2].decode('utf-16-le')
                            except Exception:
                                pass

                        elif rt == 8 and len(rec) >= 12:  # BrtFmlaString
                            cch = struct.unpack_from('<I', rec, 8)[0]
                            try:
                                rows[current_row][col] = rec[12:12 + cch * 2].decode('utf-16-le')
                            except Exception:
                                pass

                buf = buf[p:]
    return rows


# ---------------------------------------------------------------------------
# Sheet name → bin file mapping (dynamic, from workbook.bin)
# ---------------------------------------------------------------------------

def _detect_sheet_map(xlsb_path):
    import re

    def _read_xl_wstring(data, pos):
        cch = struct.unpack_from('<I', data, pos)[0]
        s = data[pos + 4:pos + 4 + cch * 2].decode('utf-16-le')
        return s, pos + 4 + cch * 2

    with open(xlsb_path, "rb") as _f, zipfile.ZipFile(_f) as z:
        with z.open('xl/_rels/workbook.bin.rels') as f:
            rels_text = f.read().decode('utf-8')
        with z.open('xl/workbook.bin') as f:
            wb_data = f.read()

    rid_to_bin = {}
    for m in re.finditer(r'Id="([^"]+)"[^>]+Target="worksheets/([^"]+)"', rels_text):
        rid_to_bin[m.group(1)] = m.group(2)

    sheet_map = {}
    pos = 0
    while pos < len(wb_data) - 4:
        try:
            rt, np = _read_rt(wb_data, pos)
            size, np = _read_varint(wb_data, np)
            rec = wb_data[np:np + size]
            if rt == 0x009C and size > 8:  # BrtBundleSh
                rel_id, p2 = _read_xl_wstring(rec, 8)
                name, _   = _read_xl_wstring(rec, p2)
                if rel_id in rid_to_bin:
                    sheet_map[name] = rid_to_bin[rel_id]
            pos = np + size
        except Exception:
            pos += 1

    return sheet_map


# ---------------------------------------------------------------------------
# Convert raw rows → list of record dicts  (dedup by Dist_Name)
# ---------------------------------------------------------------------------

def _rows_to_records(rows, header_row=1):
    """
    Convert raw {row_index: {col_index: value}} dict to list of record dicts.

    header_row : 0-based index of the row containing column headers.
                 Rows before it are ignored; rows after it are data.
                 Default 1 = Excel row 2 (standard Nokia format).
                 Pass 0 for files whose headers start on Excel row 1.
    """
    hdr_data = rows.get(header_row, {})
    if not hdr_data:
        return []

    col_to_name = {col: str(val).strip() for col, val in hdr_data.items()}

    records = []

    for rn in sorted(rows.keys()):
        if rn <= header_row:
            continue
        row = rows[rn]
        rec = {}
        for col, name in col_to_name.items():
            val = row.get(col, '')
            if isinstance(val, float) and val == int(val):
                val = int(val)
            rec[name] = val

        records.append(rec)

    return records


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def read_xlsb(path, sheet_names=None, progress_fn=None, header_row=1):
    """
    Read an .xlsb file and return a dict of {sheet_name: [records]}.

    sheet_names : list of sheet names to load; None = load all.
    progress_fn : optional callable(msg: str) for progress reporting.
    """
    def log(msg):
        if progress_fn:
            progress_fn(msg)

    log('Loading shared strings...')
    ss = _load_shared_strings(path)
    log(f'  {len(ss):,} shared strings loaded')

    log('Detecting sheet layout...')
    sheet_map = _detect_sheet_map(path)
    log(f'  Sheets found: {list(sheet_map.keys())}')

    result = {}
    for name, bin_file in sheet_map.items():
        if sheet_names and name not in sheet_names:
            continue
        log(f'  Reading sheet: {name}...')
        raw = _read_sheet_bin(path, bin_file, ss)
        records = _rows_to_records(raw, header_row=header_row)
        result[name] = records
        log(f'    {len(records):,} records')

    return result
