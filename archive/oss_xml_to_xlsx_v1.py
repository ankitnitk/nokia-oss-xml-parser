#!/usr/bin/env python3
"""
OSS XML to XLSX/XLSB Converter
Converts Nokia OSS RAML XML dump files to Excel format.

Usage (interactive GUI):   python oss_xml_to_xlsx.py
Usage (command-line):      python oss_xml_to_xlsx.py file1.xml.gz file2.gz -o out.xlsx
                           python oss_xml_to_xlsx.py file1.zip -o out.xlsx --classes BTS,BCF,TRX

Supported input formats:   .xml  .xml.gz  .gz  .zip
Supported output formats:  .xlsx  (+ .xlsb if pywin32 is installed)
"""

import gzip
import io
import os
import re
import sys
import zipfile
import argparse
import threading
from collections import defaultdict, OrderedDict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

from lxml import etree
import xlsxwriter

try:
    import win32com.client
    XLSB_SUPPORTED = True
except ImportError:
    XLSB_SUPPORTED = False

AUTHOR = 'Ankit Jain'

# Match by local name only — handles raml20.xsd, raml21.xsd, and any future versions
def _local(tag):
    return tag.split('}', 1)[-1] if '}' in tag else tag

TAG_MO   = 'managedObject'
TAG_P    = 'p'
TAG_LIST = 'list'
TAG_ITEM = 'item'
TAG_LOG  = 'log'

_print_lock = threading.Lock()

def tprint(*args, **kwargs):
    with _print_lock:
        print(*args, **kwargs)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def fmt_elapsed(s):
    if s < 60:
        return f'{s:.1f}s'
    m, sec = divmod(int(s), 60)
    return f'{m}m {sec:02d}s'

def ts():
    return datetime.now().strftime('%H:%M:%S')


# ---------------------------------------------------------------------------
# File reading — .xml / .gz / .zip
# ---------------------------------------------------------------------------

def iter_xml_streams(filepath):
    """Yield (stream_or_bytes, display_name) for every XML inside filepath."""
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
# Quick class scan (fast — regex, no XML parse)
# ---------------------------------------------------------------------------

_CLASS_RE = re.compile(rb'managedObject class="([^"]+)"')

def quick_scan_classes(filepath):
    """Return {class_name: count} by scanning bytes (no full XML parse)."""
    counts = defaultdict(int)
    low = filepath.lower()
    if low.endswith('.zip'):
        with zipfile.ZipFile(filepath, 'r') as zf:
            entries = [e for e in zf.namelist() if e.lower().endswith('.xml')]
            for entry in entries:
                with zf.open(entry) as f:
                    for m in _CLASS_RE.finditer(f.read()):
                        counts[m.group(1).decode()] += 1
    elif low.endswith('.gz'):
        with gzip.open(filepath, 'rb') as f:
            for chunk in iter(lambda: f.read(1 << 20), b''):
                for m in _CLASS_RE.finditer(chunk):
                    counts[m.group(1).decode()] += 1
    else:
        with open(filepath, 'rb') as f:
            for chunk in iter(lambda: f.read(1 << 20), b''):
                for m in _CLASS_RE.finditer(chunk):
                    counts[m.group(1).decode()] += 1
    return counts


def scan_all_files(filepaths):
    """Parallel quick-scan of all input files. Returns merged {class: count}."""
    merged = defaultdict(int)
    with ThreadPoolExecutor(max_workers=min(len(filepaths), 4)) as ex:
        futs = {ex.submit(quick_scan_classes, p): p for p in filepaths}
        for fut in as_completed(futs):
            for cls, cnt in fut.result().items():
                merged[cls] += cnt
    return merged


# ---------------------------------------------------------------------------
# XML Parsing
# ---------------------------------------------------------------------------

def try_numeric(text):
    if text is None:
        return None
    s = text.strip()
    try:
        f = float(s)
        return int(f) if f == int(f) else f
    except (ValueError, TypeError):
        return s or None


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


def parse_managed_object(elem, filename):
    mo_class  = elem.get('class', '')
    dist_name = elem.get('distName', '')
    obj_id    = elem.get('id', '')
    version   = elem.get('version', '')
    operation = elem.get('operation', '')  # present in plan/change XMLs

    hierarchy = parse_dist_name(dist_name)

    record = OrderedDict()
    if operation:
        record['operation']  = operation
    record['id']         = try_numeric(obj_id)
    record['File_Name']  = filename
    record['Dist_Name']  = dist_name
    record['SW_Version'] = version

    for child in elem:
        ltag = _local(child.tag)
        if ltag == TAG_P:
            record[child.get('name')] = try_numeric(child.text)
        elif ltag == TAG_LIST:
            list_name = child.get('name')
            children  = list(child)
            if not children:
                pass
            elif _local(children[0].tag) == TAG_ITEM:
                record[list_name] = 'List'
                item_fields = OrderedDict()
                for item in children:
                    for p in item:
                        if _local(p.tag) == TAG_P:
                            item_fields.setdefault(p.get('name'), []).append(str(p.text or ''))
                for fname, vals in item_fields.items():
                    record[f'Item-{list_name}-{fname}'] = ';'.join(vals)
            else:
                vals = [str(p.text or '') for p in children if _local(p.tag) == TAG_P]
                record[list_name] = 'List;' + ';'.join(vals)

    return mo_class, hierarchy, record


def parse_xml_bytes(data, display_name, filter_classes=None):
    """Parse XML from bytes or file-like stream."""
    classes_data = defaultdict(list)
    file_info    = {'filename': display_name, 'dateTime': ''}
    count        = 0
    t0           = datetime.now()

    tprint(f'  [{ts()}] Parsing {display_name} ...', end='', flush=True)

    stream = io.BytesIO(data) if isinstance(data, (bytes, bytearray)) else data
    for event, elem in etree.iterparse(
        stream, events=('end',),
        load_dtd=False, no_network=True, recover=True, huge_tree=True
    ):
        ltag = _local(elem.tag)
        if ltag == TAG_LOG and not file_info['dateTime']:
            file_info['dateTime'] = elem.get('dateTime', '')
            elem.clear()
        elif ltag == TAG_MO:
            mo_class = elem.get('class', '')
            if filter_classes is None or mo_class in filter_classes:
                _, hierarchy, record = parse_managed_object(elem, display_name)
                classes_data[mo_class].append((hierarchy, record))
                count += 1
            elem.clear()

    elapsed = (datetime.now() - t0).total_seconds()
    tprint(f' done — {count:,} objects in {fmt_elapsed(elapsed)}')
    return classes_data, file_info


def parse_input_file(filepath, filter_classes=None):
    results = []
    for xml_bytes, display_name in iter_xml_streams(filepath):
        results.append(parse_xml_bytes(xml_bytes, display_name, filter_classes))
    return results


# ---------------------------------------------------------------------------
# Column ordering
# ---------------------------------------------------------------------------

def build_column_order(records):
    META = ('operation', 'id', 'File_Name', 'Dist_Name', 'SW_Version')
    hier_seen  = OrderedDict()
    param_seen = OrderedDict()
    has_op = False
    for hierarchy, record in records:
        for k in hierarchy:
            hier_seen[k] = True
        for k in record:
            if k == 'operation':
                has_op = True
            if k not in META:
                param_seen[k] = True
    # Only include 'operation' in meta if at least one record has it
    meta = list(META) if has_op else [m for m in META if m != 'operation']
    return list(hier_seen.keys()), meta, list(param_seen.keys())


# ---------------------------------------------------------------------------
# Excel writing
# ---------------------------------------------------------------------------

def make_formats(wb):
    header_fmt = wb.add_format({
        'bold': True, 'bg_color': '#4472C4', 'font_color': '#FFFFFF',
        'border': 1, 'valign': 'vcenter',
    })
    label_fmt = wb.add_format({'bold': True})
    data_fmt  = wb.add_format({'valign': 'vcenter'})
    return header_fmt, label_fmt, data_fmt


def write_info_sheet(wb, files_info, class_names, label_fmt):
    ws = wb.add_worksheet('Info')
    ws.set_column(0, 0, 3)
    ws.set_column(1, 1, 22)
    ws.set_column(2, 2, 32)
    ws.set_column(3, 3, 38)
    ws.write(1, 1, 'Created with OSS XML Converter', label_fmt)
    ws.write(2, 1, f'Created by: {AUTHOR}')
    ws.write(3, 2, 'File name',           label_fmt)
    ws.write(3, 3, 'Dump Extracted Time', label_fmt)
    row = 3
    for fi in files_info:
        row += 1
        ws.write(row, 1, 'Used Netact export:')
        ws.write(row, 2, fi['filename'])
        ws.write(row, 3, fi['dateTime'])
    row += 2
    ws.write(row, 1, 'Contents', label_fmt)
    link_fmt = wb.add_format({'color': '#0563C1', 'underline': True})
    for cls in class_names:
        row += 1
        ws.write_url(row, 1, f"internal:'{cls}'!A1", link_fmt, cls)


def write_class_sheet(wb, class_name, records, header_fmt, data_fmt):
    ws = wb.add_worksheet(class_name)
    if not records:
        return

    hier_cols, meta_cols, param_cols = build_column_order(records)
    all_cols  = hier_cols + meta_cols + param_cols
    hier_set  = set(hier_cols)

    # Row 0: "Info" hyperlinked back to Info sheet; Row 1: headers
    link_fmt = wb.add_format({'color': '#0563C1', 'underline': True})
    ws.write_url(0, 0, "internal:'Info'!A1", link_fmt, 'Info')
    for ci, col in enumerate(all_cols):
        ws.write(1, ci, col, header_fmt)
    # Freeze after Dist_Name — offset shifts by 1 if 'operation' column present
    has_op   = 'operation' in meta_cols
    dist_idx = len(hier_cols) + (1 if has_op else 0) + 3  # hier + [op] + id + File_Name + Dist_Name
    ws.freeze_panes(2, dist_idx)
    ws.autofilter(1, 0, 1, len(all_cols) - 1)

    # Pre-compute extraction info once per sheet
    col_info = [(col in hier_set, col) for col in all_cols]

    # Write rows — pre-build each row as a list, then write_row in one call
    for ri, (hierarchy, record) in enumerate(records):
        row = [
            hierarchy.get(col) if is_h else record.get(col)
            for is_h, col in col_info
        ]
        ws.write_row(ri + 2, 0, row, data_fmt)

    # Column widths
    n  = len(hier_cols)
    ci = n  # current column index
    if has_op:
        ws.set_column(ci, ci, 10); ci += 1   # operation
    ws.set_column(0, max(n - 1, 0), 12)      # hierarchy
    ws.set_column(ci, ci, 12);     ci += 1   # id
    ws.set_column(ci, ci, 22);     ci += 1   # File_Name
    ws.set_column(ci, ci, 45);     ci += 1   # Dist_Name
    ws.set_column(ci, ci, 10);     ci += 1   # SW_Version
    if ci <= len(all_cols) - 1:
        ws.set_column(ci, len(all_cols) - 1, 18)


# ---------------------------------------------------------------------------
# XLSB via Excel COM
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

    all_cd, all_fi = [], []
    t_parse = datetime.now()

    # Parse all files in parallel (lxml releases GIL — threads are effective)
    with ThreadPoolExecutor(max_workers=min(len(input_paths), 4)) as ex:
        futs = {ex.submit(parse_input_file, p, filter_classes): p for p in input_paths}
        for fut in as_completed(futs):
            for cd, fi in fut.result():
                all_cd.append(cd)
                all_fi.append(fi)

    t_parse = (datetime.now() - t_parse).total_seconds()

    # Merge
    merged = defaultdict(list)
    for cd in all_cd:
        for cls, recs in cd.items():
            merged[cls].extend(recs)

    # Re-order file_info to match input order (parallel may reorder)
    fi_map = {fi['filename']: fi for fi in all_fi}
    ordered_fi = []
    for path in input_paths:
        name = os.path.basename(path)
        if name in fi_map:
            ordered_fi.append(fi_map[name])
        else:
            # zip may have sub-names
            for fi in all_fi:
                if fi not in ordered_fi:
                    ordered_fi.append(fi)

    class_names   = sorted(merged.keys())
    total_records = sum(len(v) for v in merged.values())
    print(f'[{ts()}] Parsing done in {fmt_elapsed(t_parse)} — '
          f'{total_records:,} records, {len(class_names)} classes: {", ".join(class_names)}')

    print(f'[{ts()}] Writing Excel...')
    t_write = datetime.now()

    wb = xlsxwriter.Workbook(xlsx_path, {'constant_memory': False})
    wb.set_properties({
        'title':   'OSS XML Dump',
        'author':  AUTHOR,
        'company': 'Nokia',
        'created': datetime.now(),
    })
    header_fmt, label_fmt, data_fmt = make_formats(wb)
    write_info_sheet(wb, ordered_fi, class_names, label_fmt)

    for cls in class_names:
        recs = merged[cls]
        t0   = datetime.now()
        print(f'  [{ts()}] Sheet "{cls}": {len(recs):,} rows ... ', end='', flush=True)
        write_class_sheet(wb, cls, recs, header_fmt, data_fmt)
        print(f'done ({fmt_elapsed((datetime.now() - t0).total_seconds())})')

    wb.close()
    t_write = (datetime.now() - t_write).total_seconds()
    print(f'[{ts()}] Writing done in {fmt_elapsed(t_write)}')

    if want_xlsb:
        t0 = datetime.now()
        print(f'[{ts()}] Converting to .xlsb via Excel...')
        convert_to_xlsb(xlsx_path, output_path)
        os.remove(xlsx_path)
        print(f'[{ts()}] Conversion done in {fmt_elapsed((datetime.now() - t0).total_seconds())}')

    t_total = (datetime.now() - t_total).total_seconds()
    size_mb = os.path.getsize(output_path) / 1024 / 1024
    print(f'\n[{ts()}] Done!  {output_path}  ({size_mb:.1f} MB)')
    print(f'         Parse: {fmt_elapsed(t_parse)}  |  Write: {fmt_elapsed(t_write)}  |  Total: {fmt_elapsed(t_total)}')
    return output_path


# ---------------------------------------------------------------------------
# Class selection dialog (tkinter)
# ---------------------------------------------------------------------------

def ask_class_selection(class_counts):
    """Show a compact scrollable checkbox dialog. Returns set of selected class names, or None to cancel."""
    import tkinter as tk
    from tkinter import ttk

    selected = {}
    result   = [None]

    dlg = tk.Toplevel()
    dlg.title('Select MO Classes')
    dlg.resizable(False, False)
    dlg.grab_set()

    total = sum(class_counts.values())
    tk.Label(dlg, text=f'{len(class_counts)} MO classes  |  {total:,} total objects',
             font=('Arial', 10, 'bold'), pady=6).pack()

    # --- Scrollable canvas area ---
    container = tk.Frame(dlg)
    container.pack(padx=12, pady=(0, 4), fill='both')

    canvas = tk.Canvas(container, width=340, height=280, bd=1, relief='sunken', highlightthickness=0)
    scrollbar = tk.Scrollbar(container, orient='vertical', command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side='right', fill='y')
    canvas.pack(side='left', fill='both')

    inner = tk.Frame(canvas)
    canvas_window = canvas.create_window((0, 0), window=inner, anchor='nw')

    def on_resize(event):
        canvas.configure(scrollregion=canvas.bbox('all'))
        canvas.itemconfig(canvas_window, width=event.width)

    inner.bind('<Configure>', on_resize)
    canvas.bind('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1 * e.delta / 120), 'units'))

    # Checkboxes — 2 columns
    classes = sorted(class_counts)
    for i, cls in enumerate(classes):
        var = tk.BooleanVar(value=True)
        selected[cls] = var
        cnt = class_counts[cls]
        row, col = divmod(i, 2)
        tk.Checkbutton(inner, text=f'{cls}  ({cnt:,})',
                       variable=var, anchor='w',
                       font=('Consolas', 9), width=18).grid(
            row=row, column=col, sticky='w', padx=6, pady=1)

    # --- Buttons ---
    btn_frame = tk.Frame(dlg)
    btn_frame.pack(pady=6)

    def select_all():
        for v in selected.values(): v.set(True)
    def select_none():
        for v in selected.values(): v.set(False)
    def ok():
        result[0] = {cls for cls, v in selected.items() if v.get()}
        dlg.destroy()
    def cancel():
        dlg.destroy()

    tk.Button(btn_frame, text='All',    width=8,  command=select_all).grid(row=0, column=0, padx=3)
    tk.Button(btn_frame, text='None',   width=8,  command=select_none).grid(row=0, column=1, padx=3)
    tk.Button(btn_frame, text='OK',     width=8,  command=ok, bg='#4472C4', fg='white').grid(row=0, column=2, padx=3)
    tk.Button(btn_frame, text='Cancel', width=8,  command=cancel).grid(row=0, column=3, padx=3)

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

    # Step 1 — pick input files
    input_paths = filedialog.askopenfilenames(
        title='Select OSS XML dump file(s)',
        filetypes=[
            ('All supported', '*.xml *.gz *.zip'),
            ('XML files',     '*.xml'),
            ('GZ files',      '*.gz'),
            ('ZIP files',     '*.zip'),
            ('All files',     '*.*'),
        ],
    )
    if not input_paths:
        print('No files selected. Exiting.')
        return
    input_paths = list(input_paths)

    # Step 2 — quick scan + class selection
    print(f'[{ts()}] Scanning files for MO classes...')
    t0 = datetime.now()
    class_counts = scan_all_files(input_paths)
    print(f'[{ts()}] Scan done in {fmt_elapsed((datetime.now()-t0).total_seconds())} '
          f'— found classes: {", ".join(sorted(class_counts))}')

    filter_classes = ask_class_selection(class_counts)
    if filter_classes is None:
        print('Cancelled.')
        return
    if not filter_classes:
        print('No classes selected. Exiting.')
        return

    # Step 3 — pick output file
    out_types = [('Excel Binary Workbook', '*.xlsb'), ('Excel Workbook', '*.xlsx')] \
                if XLSB_SUPPORTED else [('Excel Workbook', '*.xlsx')]
    base    = os.path.basename(input_paths[0]).split('.')[0]
    defname = f'{base}_{datetime.now().strftime("%Y%m%d_%H%M%S")}_dump.xlsx'
    defdir  = os.path.dirname(input_paths[0])

    output_path = filedialog.asksaveasfilename(
        title='Save output as',
        initialdir=defdir,
        initialfile=defname,
        defaultextension='.xlsx',
        filetypes=out_types,
    )
    if not output_path:
        print('No output file selected. Exiting.')
        return

    root.destroy()

    run_conversion(input_paths, output_path, filter_classes or None)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    os.system("title Ankit's XML Parser  [V1]")

    if len(sys.argv) == 1:
        run_interactive()
        input('\nPress Enter to exit...')
        return

    parser = argparse.ArgumentParser(description='Convert Nokia OSS RAML XML dumps to Excel')
    parser.add_argument('inputs',    nargs='+', metavar='FILE',
                        help='Input file(s): .xml  .xml.gz  .gz  .zip')
    parser.add_argument('-o', '--output',  default='',
                        help='Output file (.xlsx or .xlsb)')
    parser.add_argument('--classes', default='',
                        help='Comma-separated MO classes to include, e.g. BTS,BCF,TRX  (default: all)')
    args = parser.parse_args()

    for path in args.inputs:
        if not os.path.isfile(path):
            print(f'ERROR: File not found: {path}', file=sys.stderr)
            input('\nPress Enter to exit...')
            sys.exit(1)

    if not args.output:
        base        = os.path.basename(args.inputs[0]).split('.')[0]
        args.output = f'{base}_{datetime.now().strftime("%Y%m%d_%H%M%S")}_dump.xlsx'

    filter_classes = {c.strip() for c in args.classes.split(',') if c.strip()} or None

    run_conversion(args.inputs, args.output, filter_classes)
    input('\nPress Enter to exit...')


if __name__ == '__main__':
    main()
