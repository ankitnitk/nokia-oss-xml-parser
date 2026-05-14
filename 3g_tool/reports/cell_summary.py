"""
reports/cell_summary.py
Builds the 3G WCDMA Cell Details Excel report from a Network object.
One row per WCEL, sorted by RNC ID → WBTS ID → WCEL ID.
"""

import xlsxwriter
from network import get, to_num, _rnc_dn, _wbts_dn


# ---------------------------------------------------------------------------
# Column definitions
# ---------------------------------------------------------------------------

COLUMNS = [
    'RNC ID',
    'RNC Name',
    'WBTS ID',
    'WBTS Name',
    'SBTS',
    'WCEL ID',
    'WCEL Name',
    'LAC',
    'RAC',
    'PSC',
    'UARFCN',
    'Tilt',
    'CPICH',
    'PMAX',
]

_NUMERIC_COLS = {
    'RNC ID', 'WBTS ID', 'WCEL ID',
    'LAC', 'RAC', 'PSC', 'UARFCN', 'Tilt', 'CPICH', 'PMAX',
}

_COL_WIDTHS = [
    10, 22, 10, 22, 14,   # RNC ID/Name, WBTS ID/Name, SBTS
    10, 22,               # WCEL ID/Name
    10, 10,               # LAC, RAC
     8, 10,               # PSC, UARFCN
     8, 10, 10,           # Tilt, CPICH, PMAX
]


# ---------------------------------------------------------------------------
# Builder
# ---------------------------------------------------------------------------

def build(network, output_path, progress_fn=None):
    def log(msg):
        if progress_fn:
            progress_fn(msg)

    log(f'Building Cell Details — {len(network.wcel_list):,} cells...')

    wb = xlsxwriter.Workbook(output_path, {'strings_to_numbers': False})

    hdr_fmt = wb.add_format({
        'font_name': 'Arial', 'font_size': 10, 'bold': True,
        'font_color': '#FFFFFF', 'bg_color': '#1F4E79',
        'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        'left': 1, 'left_color': '#FFFFFF',
        'right': 1, 'right_color': '#FFFFFF',
        'bottom': 1, 'bottom_color': '#FFFFFF',
    })
    data_fmt = wb.add_format({
        'font_name': 'Arial', 'font_size': 9, 'valign': 'vcenter',
        'left': 1, 'left_color': '#BDD7EE',
        'right': 1, 'right_color': '#BDD7EE',
        'bottom': 1, 'bottom_color': '#BDD7EE',
    })
    alt_fmt = wb.add_format({
        'font_name': 'Arial', 'font_size': 9, 'valign': 'vcenter',
        'bg_color': '#EBF3FB',
        'left': 1, 'left_color': '#BDD7EE',
        'right': 1, 'right_color': '#BDD7EE',
        'bottom': 1, 'bottom_color': '#BDD7EE',
    })

    ws = wb.add_worksheet('Cell Details')
    ws.freeze_panes(1, 0)
    ws.set_row(0, 30)

    for ci, (col, w) in enumerate(zip(COLUMNS, _COL_WIDTHS)):
        ws.write(0, ci, col, hdr_fmt)
        ws.set_column(ci, ci, w)

    rows = _build_rows(network)
    log(f'  {len(rows):,} cells')

    for ri, rd in enumerate(rows, 1):
        fmt = alt_fmt if ri % 2 == 1 else data_fmt
        for ci, col in enumerate(COLUMNS):
            val = rd.get(col, '')
            if col in _NUMERIC_COLS and val != '':
                try:
                    ws.write_number(ri, ci, float(val), fmt)
                except (ValueError, TypeError):
                    ws.write(ri, ci, val, fmt)
            else:
                ws.write(ri, ci, val, fmt)

    ws.autofilter(0, 0, len(rows), len(COLUMNS) - 1)
    wb.close()
    log(f'Saved: {output_path}')
    return len(rows)


def _build_rows(network):
    rows = []
    for wcel_r in network.wcel_list:
        if not wcel_r.get('Dist_Name', ''):
            continue

        rnc_id  = get(wcel_r, 'RNC')
        wbts_id = get(wcel_r, 'WBTS')

        rnc_r  = network.rnc_by_dn.get(_rnc_dn(rnc_id), {})
        wbts_r = network.wbts_by_dn.get(_wbts_dn(rnc_id, wbts_id), {})

        rows.append({
            'RNC ID':    get(rnc_r, 'RNC') or rnc_id,
            'RNC Name':  get(rnc_r, 'name'),
            'WBTS ID':   get(wbts_r, 'WBTS') or wbts_id,
            'WBTS Name': get(wbts_r, 'name'),
            'SBTS':      get(wbts_r, 'SBTSId'),
            'WCEL ID':   get(wcel_r, 'WCEL'),
            'WCEL Name': get(wcel_r, 'name'),
            'LAC':       get(wcel_r, 'LAC'),
            'RAC':       get(wcel_r, 'RAC'),
            'PSC':       get(wcel_r, 'PriScrCode'),
            'UARFCN':    get(wcel_r, 'UARFCN'),
            'Tilt':      get(wcel_r, 'angle'),
            'CPICH':     get(wcel_r, 'PtxPrimaryCPICH'),
            'PMAX':      get(wcel_r, 'maximumTransmissionPower'),
        })

    rows.sort(key=lambda r: (
        to_num(r['RNC ID']),
        to_num(r['WBTS ID']),
        to_num(r['WCEL ID']),
    ))
    return rows
