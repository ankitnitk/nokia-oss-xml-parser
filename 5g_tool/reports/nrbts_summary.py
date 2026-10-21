"""
nrbts_summary.py
Builds the 5G NR summary report Excel file.

Sheets:
  1. NRBTS Details   — one row per NRBTS (gNB)
  2. NRCELL Details  — one row per NR cell
  3. Network Stats   — Working / Other / Total summary
"""

import re
from collections import defaultdict

import xlsxwriter
from network import get, to_num, _nrbts_dn


# dlMimoMode enum (5G NR -- differs from the LTE map in 4g_tool).
_MIMO_MODE = {
    12: '1x1 Single Antenna Transmission',
    30: '2x2 Closed Loop Spatial Multiplexing',
    40: '4x4 or 4x2 Closed Loop Spatial Multiplexing',
    50: '2x2 Open Loop Spatial Multiplexing',
    60: '4x4 or 4x2 Open Loop Spatial Multiplexing',
    70: '8x4 Closed Loop Spatial Multiplexing',
}

# NRCELL administrativeState -- note this is NOT the LNCEL enum (1/3) that
# 4g_tool decodes; NRCELL uses 2 = unlocked / in service, 0 = locked.
_ADMIN_STATE = {'0': 'Locked', '2': 'Working'}

# cellName convention is <SITE>_NR_<sector>, e.g. NAIROBI_EASTLEIGH_NR_3.
_SECTOR_RE = re.compile(r'_(\d+)$')


def _admin_state(val):
    return _ADMIN_STATE.get(val, val)


def _mimo_mode(val):
    n = to_num(val, default=None)
    if n is None:
        return ''
    return _MIMO_MODE.get(int(n), str(int(n)))


def _sector_id(cell_name):
    """Trailing number of the cellName = sector. '' if it doesn't follow the
    convention."""
    if not cell_name:
        return ''
    m = _SECTOR_RE.search(cell_name)
    return m.group(1) if m else ''


# ---------------------------------------------------------------------------
# Shared format factory
# ---------------------------------------------------------------------------

def _make_formats(wb):
    hdr = wb.add_format({
        'bold': True, 'bg_color': '#1F4E79', 'font_color': '#FFFFFF',
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
    })
    cell = wb.add_format({'border': 1, 'valign': 'vcenter'})
    num  = wb.add_format({'border': 1, 'valign': 'vcenter', 'num_format': '0'})
    dec1 = wb.add_format({'border': 1, 'valign': 'vcenter', 'num_format': '0.0'})
    red  = wb.add_format({'border': 1, 'valign': 'vcenter',
                          'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    return dict(hdr=hdr, cell=cell, num=num, dec1=dec1, red=red)


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def build(network, output_path, progress_fn=None):
    """Build the 5G summary workbook. Returns number of NRBTS rows written."""
    def log(msg):
        if progress_fn:
            progress_fn(msg)

    wb = xlsxwriter.Workbook(output_path, {'strings_to_numbers': False})
    fmt = _make_formats(wb)

    nrbts_rows = list(_iter_nrbts_rows(network))
    _build_nrbts_details(wb, fmt, log, nrbts_rows)
    nrcell_rows = list(_iter_nrcell_rows(network))
    _build_nrcell_details(wb, fmt, log, nrcell_rows)
    _build_network_stats(wb, fmt, log, nrbts_rows, nrcell_rows)

    wb.close()
    log(f'Workbook saved: {output_path}')
    return len(nrbts_rows)


# ---------------------------------------------------------------------------
# Sheet 1 — NRBTS Details
# ---------------------------------------------------------------------------

_NRBTS_COLS = [
    'MRBTS ID', 'NRBTS ID', 'NRBTS Name', 'SW Version',
    'Cell Count', 'Band Count', 'Band List',
]
_NRBTS_NUM = {'MRBTS ID', 'NRBTS ID', 'Cell Count', 'Band Count'}
_NRBTS_WIDTHS = {
    'MRBTS ID': 12, 'NRBTS ID': 12, 'NRBTS Name': 28, 'SW Version': 24,
    'Cell Count': 11, 'Band Count': 11, 'Band List': 34,
}


def _iter_nrbts_rows(network):
    rows = []
    for nrbts_dn, rec in network.nrbts_by_dn.items():
        arfcns = network.nrarfcns_for_nrbts(nrbts_dn)
        rows.append({
            'MRBTS ID':   get(rec, 'MRBTS'),
            'NRBTS ID':   get(rec, 'NRBTS'),
            'NRBTS Name': get(rec, 'name'),
            'SW Version': get(rec, 'SW_Version'),
            'Cell Count': len(network.nrcell_for_nrbts(nrbts_dn)),
            'Band Count': len(arfcns),
            'Band List':  ', '.join(str(a) for a in arfcns),
        })
    rows.sort(key=lambda r: (to_num(r['MRBTS ID']), to_num(r['NRBTS ID'])))
    return rows


def _build_nrbts_details(wb, fmt, log, rows):
    log('Building NRBTS Details sheet...')
    ws = wb.add_worksheet('NRBTS Details')
    ws.freeze_panes(1, 0)

    for ci, col in enumerate(_NRBTS_COLS):
        ws.write(0, ci, col, fmt['hdr'])
    ws.set_row(0, 30)
    log(f'  {len(rows):,} NRBTS records')

    for ri, rd in enumerate(rows, 1):
        for ci, col in enumerate(_NRBTS_COLS):
            val = rd.get(col, '')
            if col in _NRBTS_NUM:
                n = to_num(val, default=None)
                if n is not None:
                    ws.write_number(ri, ci, n, fmt['num'])
                else:
                    ws.write_blank(ri, ci, fmt['num'])
            else:
                ws.write(ri, ci, val, fmt['cell'])

    for ci, col in enumerate(_NRBTS_COLS):
        ws.set_column(ci, ci, _NRBTS_WIDTHS.get(col, 15))
    if rows:
        ws.autofilter(0, 0, len(rows), len(_NRBTS_COLS) - 1)


# ---------------------------------------------------------------------------
# Sheet 2 — NRCELL Details
# ---------------------------------------------------------------------------

_NRCELL_COLS = [
    'MRBTS ID', 'NRBTS ID', 'NRBTS Name', 'NRCELL ID', 'Name', 'Cell Name',
    'Sector ID', 'Admin State', 'TAC', 'PCI', 'RSI', 'NRARFCN',
    'Bandwidth (MHz)', 'SSPBCH Block Power (dBm)', 'CSIRS Power (dBm)',
    'Mimo Mode', 'Tilt Offset', 'Tilt', 'Left Edge Angle', 'Right Edge Angle',
    'Beam Set in Use',
]
_NRCELL_NUM = {
    'MRBTS ID', 'NRBTS ID', 'NRCELL ID', 'Sector ID', 'TAC', 'PCI', 'RSI',
    'NRARFCN', 'Bandwidth (MHz)', 'SSPBCH Block Power (dBm)',
    'CSIRS Power (dBm)', 'Tilt Offset', 'Tilt', 'Left Edge Angle',
    'Right Edge Angle',
}
_NRCELL_WIDTHS = {
    'MRBTS ID': 12, 'NRBTS ID': 12, 'NRBTS Name': 28, 'NRCELL ID': 10,
    'Name': 28, 'Cell Name': 28, 'Sector ID': 9, 'Admin State': 12,
    'TAC': 9, 'PCI': 7, 'RSI': 7, 'NRARFCN': 11, 'Bandwidth (MHz)': 14,
    'SSPBCH Block Power (dBm)': 15, 'CSIRS Power (dBm)': 14,
    'Mimo Mode': 38, 'Tilt Offset': 11, 'Tilt': 8,
    'Left Edge Angle': 13, 'Right Edge Angle': 13, 'Beam Set in Use': 15,
}


def _iter_nrcell_rows(network):
    rows = []
    for nrcell_dn, rec in network.nrcell_by_dn.items():
        mrbts = get(rec, 'MRBTS')
        nrbts = get(rec, 'NRBTS')
        nrbts_rec = network.nrbts_by_dn.get(_nrbts_dn(mrbts, nrbts), {})
        cell_name = get(rec, 'cellName')
        rows.append({
            'MRBTS ID':    mrbts,
            'NRBTS ID':    nrbts,
            'NRBTS Name':  get(nrbts_rec, 'name'),
            'NRCELL ID':   get(rec, 'NRCELL'),
            'Name':        get(rec, 'name'),
            'Cell Name':   cell_name,
            'Sector ID':   _sector_id(cell_name),
            'Admin State': _admin_state(get(rec, 'administrativeState')),
            'TAC':         get(rec, 'configuredEpsTac'),
            'PCI':         get(rec, 'physCellId'),
            'RSI':         get(rec, 'prachRootSequenceIndex'),
            'NRARFCN':     get(rec, 'nrarfcn'),
            'Bandwidth (MHz)':          get(rec, 'chBw'),
            'SSPBCH Block Power (dBm)': get(rec, 'ssPbchBlockPower'),
            'CSIRS Power (dBm)':        get(rec, 'Item-periodicCsiAcquisition-csirsPower'),
            'Mimo Mode':       _mimo_mode(get(rec, 'dlMimoMode')),
            'Tilt Offset':     get(rec, 'Item-beamSet-tiltOffset'),
            'Tilt':            get(rec, 'angle'),
            'Left Edge Angle': get(rec, 'Item-beamSet-leftEdgeAngle'),
            'Right Edge Angle': get(rec, 'Item-beamSet-rightEdgeAngle'),
            'Beam Set in Use': get(rec, 'Item-beamSet-beamSetInUse'),
            '_admin_raw':      get(rec, 'administrativeState'),
        })
    rows.sort(key=lambda r: (to_num(r['MRBTS ID']), to_num(r['NRBTS ID']),
                             to_num(r['NRCELL ID'])))
    return rows


def _build_nrcell_details(wb, fmt, log, rows):
    log('Building NRCELL Details sheet...')
    ws = wb.add_worksheet('NRCELL Details')
    ws.freeze_panes(1, 0)

    for ci, col in enumerate(_NRCELL_COLS):
        ws.write(0, ci, col, fmt['hdr'])
    ws.set_row(0, 30)
    log(f'  {len(rows):,} NRCELL records')

    # Duplicate (MRBTS, PCI, NRARFCN) within the same site -- same check the
    # 4G tool does for RSI+EARFCN; a PCI clash on one carrier at one site is a
    # real planning fault.
    pci_count = defaultdict(int)
    for rd in rows:
        if rd['PCI'] != '' and rd['NRARFCN'] != '':
            pci_count[(rd['MRBTS ID'], rd['PCI'], rd['NRARFCN'])] += 1
    dup_pci = {k for k, v in pci_count.items() if v > 1}

    for ri, rd in enumerate(rows, 1):
        locked = rd.get('_admin_raw') == '0'
        for ci, col in enumerate(_NRCELL_COLS):
            val = rd.get(col, '')
            if col == 'Admin State':
                ws.write(ri, ci, val, fmt['red'] if locked else fmt['cell'])
            elif col == 'PCI':
                is_dup = (rd['MRBTS ID'], rd['PCI'], rd['NRARFCN']) in dup_pci
                f = fmt['red'] if is_dup else fmt['num']
                n = to_num(val, default=None)
                if n is not None:
                    ws.write_number(ri, ci, n, f)
                else:
                    ws.write_blank(ri, ci, f)
            elif col in _NRCELL_NUM:
                n = to_num(val, default=None)
                if n is not None:
                    ws.write_number(ri, ci, n, fmt['num'])
                else:
                    ws.write_blank(ri, ci, fmt['num'])
            else:
                ws.write(ri, ci, val, fmt['cell'])

    for ci, col in enumerate(_NRCELL_COLS):
        ws.set_column(ci, ci, _NRCELL_WIDTHS.get(col, 15))
    if rows:
        ws.autofilter(0, 0, len(rows), len(_NRCELL_COLS) - 1)
    log(f'  {len(dup_pci):,} duplicate PCI+NRARFCN combos within a site')


# ---------------------------------------------------------------------------
# Sheet 3 — Network Stats
# ---------------------------------------------------------------------------

def _build_network_stats(wb, fmt, log, nrbts_rows, nrcell_rows):
    log('Building Network Stats sheet...')
    ws = wb.add_worksheet('Network Stats')

    title_fmt = wb.add_format({'bold': True, 'font_size': 13, 'valign': 'vcenter'})
    cat_fmt = wb.add_format({'bold': True, 'border': 1, 'valign': 'vcenter',
                             'bg_color': '#D6E4F0'})
    col_hdr_fmt = wb.add_format({'bold': True, 'border': 1, 'align': 'center',
                                 'valign': 'vcenter', 'bg_color': '#1F4E79',
                                 'font_color': '#FFFFFF'})
    working_fmt = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',
                                 'bg_color': '#E2EFDA', 'num_format': '0'})
    other_fmt = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter',
                               'bg_color': '#FFF2CC', 'num_format': '0'})
    total_fmt = wb.add_format({'bold': True, 'border': 1, 'align': 'center',
                               'valign': 'vcenter', 'bg_color': '#DEEAF1',
                               'num_format': '0'})

    cells_total   = len(nrcell_rows)
    cells_working = sum(1 for r in nrcell_rows if r.get('_admin_raw') == '2')
    sites_total   = len(nrbts_rows)
    # A site counts as working if at least one of its cells is unlocked.
    working_sites = {(r['MRBTS ID'], r['NRBTS ID'])
                     for r in nrcell_rows if r.get('_admin_raw') == '2'}
    sites_working = sum(1 for r in nrbts_rows
                        if (r['MRBTS ID'], r['NRBTS ID']) in working_sites)

    stats = {
        'NRBTS (Sites)': (sites_working, sites_total - sites_working, sites_total),
        'NR Cells':      (cells_working, cells_total - cells_working, cells_total),
    }

    ws.write(0, 0, '5G NR Network Statistics', title_fmt)
    ws.set_row(0, 22)
    for ci, h in enumerate(['Category', 'Working', 'Other', 'Total']):
        ws.write(2, ci, h, col_hdr_fmt)
    ws.set_row(2, 20)

    r = 3
    for cat, (w, o, t) in stats.items():
        ws.write(r, 0, cat, cat_fmt)
        ws.write_number(r, 1, w, working_fmt)
        ws.write_number(r, 2, o, other_fmt)
        ws.write_number(r, 3, t, total_fmt)
        r += 1

    # Cell count per NRARFCN (band distribution)
    r += 1
    ws.write(r, 0, 'Cells per NRARFCN', title_fmt)
    r += 2
    for ci, h in enumerate(['NRARFCN', 'Cells', 'Bandwidth (MHz)', '']):
        ws.write(r, ci, h, col_hdr_fmt)
    r += 1
    per_arfcn = defaultdict(int)
    bw_by_arfcn = defaultdict(set)
    for rd in nrcell_rows:
        a = rd.get('NRARFCN', '')
        if a != '':
            per_arfcn[a] += 1
            if rd.get('Bandwidth (MHz)', '') != '':
                bw_by_arfcn[a].add(rd['Bandwidth (MHz)'])
    for a in sorted(per_arfcn, key=lambda x: to_num(x)):
        ws.write_number(r, 0, to_num(a), total_fmt)
        ws.write_number(r, 1, per_arfcn[a], working_fmt)
        ws.write(r, 2, ', '.join(sorted(bw_by_arfcn[a], key=lambda x: to_num(x))),
                 other_fmt)
        ws.write_blank(r, 3, None, fmt['cell'])
        r += 1

    ws.set_column(0, 0, 26)
    ws.set_column(1, 3, 16)
    log(f'  {len(stats)} stat categories, {len(per_arfcn)} NRARFCNs')
