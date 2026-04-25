"""
lnbts_summary.py
Builds the 4G LTE summary report Excel file.

Sheets:
  1. LNBTS Details  — one row per LNBTS (≈ BCF Details in 2G)
  2. LNCEL Details  — one row per cell (FDD or TDD)
  3. Network Stats  — Working / Other / Total summary
"""

import math
import xlsxwriter
from collections import defaultdict
from network import get, to_num, _lnbts_dn, _lncel_dn

# LTE channel bandwidth (MHz) → number of resource blocks
_BW_TO_NRB = {1.4: 6, 3.0: 15, 5.0: 25, 10.0: 50, 15.0: 75, 20.0: 100}


# ---------------------------------------------------------------------------
# Shared format factory
# ---------------------------------------------------------------------------

def _make_formats(wb):
    hdr = wb.add_format({
        'bold': True, 'bg_color': '#1F4E79', 'font_color': '#FFFFFF',
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
    })
    cell  = wb.add_format({'border': 1, 'valign': 'vcenter'})
    num   = wb.add_format({'border': 1, 'valign': 'vcenter', 'num_format': '0'})
    dec1  = wb.add_format({'border': 1, 'valign': 'vcenter', 'num_format': '0.0'})
    dec2  = wb.add_format({'border': 1, 'valign': 'vcenter', 'num_format': '0.00'})
    fdd   = wb.add_format({'border': 1, 'valign': 'vcenter', 'bg_color': '#DEEAF1'})
    tdd   = wb.add_format({'border': 1, 'valign': 'vcenter', 'bg_color': '#E2EFDA'})
    mixed = wb.add_format({'border': 1, 'valign': 'vcenter', 'bg_color': '#FFF2CC'})
    red   = wb.add_format({'border': 1, 'valign': 'vcenter', 'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    return dict(hdr=hdr, cell=cell, num=num, dec1=dec1, dec2=dec2,
                fdd=fdd, tdd=tdd, mixed=mixed, red=red)


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def build(network, output_path, progress_fn=None):
    """Build the full summary workbook. Returns number of LNBTS rows written."""
    def log(msg):
        if progress_fn:
            progress_fn(msg)

    wb = xlsxwriter.Workbook(output_path, {'strings_to_numbers': False})
    fmt = _make_formats(wb)

    _build_lnbts_details(wb, fmt, network, log)
    lncel_rows = list(_iter_lncel_rows(network))
    _build_lncel_details(wb, fmt, network, log, lncel_rows)
    _build_network_stats(wb, fmt, network, log, lncel_rows)

    wb.close()
    log(f'Workbook saved: {output_path}')
    return len(network.lnbts_list)


# ---------------------------------------------------------------------------
# Sheet 1 — LNBTS Details
# ---------------------------------------------------------------------------

_LNBTS_COLS = [
    'MRBTS ID', 'LNBTS ID', 'LNBTS Name', 'SW Version',
    'Cell Count', 'LNCEL Count', 'Band Count', 'Band List', 'LTE Mode',
]
_LNBTS_NUM = {'MRBTS ID', 'LNBTS ID', 'Cell Count', 'LNCEL Count', 'Band Count'}
_LNBTS_WIDTHS = {
    'MRBTS ID': 12, 'LNBTS ID': 12, 'LNBTS Name': 22, 'SW Version': 22,
    'Cell Count': 10, 'LNCEL Count': 11, 'Band Count': 10,
    'Band List': 35, 'LTE Mode': 11,
}


def _build_lnbts_details(wb, fmt, network, log):
    log('Building LNBTS Details sheet...')
    ws = wb.add_worksheet('LNBTS Details')
    ws.freeze_panes(1, 0)

    for ci, col in enumerate(_LNBTS_COLS):
        ws.write(0, ci, col, fmt['hdr'])
    ws.set_row(0, 30)

    rows = list(_iter_lnbts_rows(network))
    log(f'  {len(rows):,} LNBTS records')

    for ri, rd in enumerate(rows, 1):
        mode = rd.get('LTE Mode', '')
        mode_fmt = fmt['fdd'] if mode == 'FDD' else fmt['tdd'] if mode == 'TDD' else fmt['mixed'] if mode == 'FDD+TDD' else fmt['cell']
        for ci, col in enumerate(_LNBTS_COLS):
            val = rd.get(col, '')
            if col in _LNBTS_NUM:
                ws.write_number(ri, ci, to_num(val), fmt['num'])
            elif col == 'LTE Mode':
                ws.write(ri, ci, val, mode_fmt)
            else:
                ws.write(ri, ci, val, fmt['cell'])

    for ci, col in enumerate(_LNBTS_COLS):
        ws.set_column(ci, ci, _LNBTS_WIDTHS.get(col, 15))
    ws.autofilter(0, 0, len(rows), len(_LNBTS_COLS) - 1)


def _iter_lnbts_rows(network):
    for lnbts_dn, rec in network.lnbts_by_dn.items():
        fdd_cells  = network.fdd_cells_for_lnbts(lnbts_dn)
        tdd_cells  = network.tdd_cells_for_lnbts(lnbts_dn)
        lncel_list = network.lncel_for_lnbts(lnbts_dn)
        earfcns    = network.earfcns_for_lnbts(lnbts_dn)
        yield {
            'MRBTS ID':    get(rec, 'MRBTS'),
            'LNBTS ID':    get(rec, 'LNBTS'),
            'LNBTS Name':  get(rec, 'name'),
            'SW Version':  get(rec, 'SW_Version'),
            'Cell Count':  len(fdd_cells) + len(tdd_cells),
            'LNCEL Count': len(lncel_list),
            'Band Count':  len(earfcns),
            'Band List':   ', '.join(str(e) for e in earfcns),
            'LTE Mode':    network.lte_mode(lnbts_dn),
        }


# ---------------------------------------------------------------------------
# Sheet 2 — LNCEL Details
# ---------------------------------------------------------------------------

_LNCEL_COLS = [
    'MRBTS ID', 'LNBTS ID', 'LNBTS Name', 'LNCEL Name', 'LNCEL ID',
    'Admin State',
    'MCC', 'MNC', 'PCI', 'RSI', 'EARFCN DL', 'Ch BW (MHz)',
    'PMAX (dBm)', 'dlRsBoost', 'RS Power (dBm)', 'DL MIMO Mode', 'Array Mode', 'TAC', 'Tilt',
    'Cell Type', 'SIB Priority', 'IRFIM {Prio} List', 'LNHOIF List', 'CAPR {Prio} List',
]
_LNCEL_NUM = {
    'MRBTS ID', 'LNBTS ID', 'LNCEL ID', 'PCI', 'RSI', 'EARFCN DL', 'SIB Priority',
}
# TAC is numeric but handled separately for conditional formatting
_LNCEL_DEC1 = {'Ch BW (MHz)', 'PMAX (dBm)', 'RS Power (dBm)', 'Tilt'}
_LNCEL_DEC2 = {'dlRsBoost'}
_LNCEL_WIDTHS = {
    'MRBTS ID': 12, 'LNBTS ID': 12, 'LNBTS Name': 22, 'LNCEL Name': 22,
    'LNCEL ID': 9, 'Admin State': 12, 'MCC': 7, 'MNC': 7,
    'PCI': 7, 'RSI': 7, 'EARFCN DL': 11, 'Ch BW (MHz)': 12,
    'PMAX (dBm)': 11, 'dlRsBoost': 11, 'RS Power (dBm)': 14, 'DL MIMO Mode': 32, 'Array Mode': 28,
    'TAC': 8, 'Tilt': 7, 'Cell Type': 9,
    'SIB Priority': 12, 'IRFIM {Prio} List': 40, 'LNHOIF List': 40, 'CAPR {Prio} List': 40,
}


def _build_lncel_details(wb, fmt, network, log, rows):
    log('Building LNCEL Details sheet...')
    ws = wb.add_worksheet('LNCEL Details')
    ws.freeze_panes(1, 0)

    for ci, col in enumerate(_LNCEL_COLS):
        ws.write(0, ci, col, fmt['hdr'])
    ws.set_row(0, 30)

    log(f'  {len(rows):,} LNCEL records')

    # Pre-compute inconsistent TAC LNBTS set
    tac_by_lnbts = defaultdict(set)
    for rd in rows:
        tac = rd.get('TAC', '')
        if tac:
            tac_by_lnbts[rd['LNBTS ID']].add(tac)
    inconsistent_tac_lnbts = {k for k, v in tac_by_lnbts.items() if len(v) > 1}

    for ri, rd in enumerate(rows, 1):
        cell_type      = rd.get('Cell Type', '')
        row_fmt        = fmt['fdd'] if cell_type == 'FDD' else fmt['tdd'] if cell_type == 'TDD' else fmt['cell']
        irfim_missing  = rd.get('_irfim_missing', False)
        lnhoif_missing = rd.get('_lnhoif_missing', False)
        capr_missing   = rd.get('_capr_missing', False)
        tac_red        = rd.get('LNBTS ID', '') in inconsistent_tac_lnbts

        for ci, col in enumerate(_LNCEL_COLS):
            val = rd.get(col, '')
            if col in _LNCEL_NUM:
                n = to_num(val)
                if n != 0 or val != '':
                    ws.write_number(ri, ci, n, fmt['num'])
                else:
                    ws.write_blank(ri, ci, fmt['num'])
            elif col in _LNCEL_DEC1:
                n = to_num(val, default=None)
                if n is not None:
                    ws.write_number(ri, ci, n, fmt['dec1'])
                else:
                    ws.write_blank(ri, ci, fmt['dec1'])
            elif col in _LNCEL_DEC2:
                n = to_num(val, default=None)
                if n is not None:
                    ws.write_number(ri, ci, n, fmt['dec2'])
                else:
                    ws.write_blank(ri, ci, fmt['dec2'])
            elif col == 'Cell Type':
                ws.write(ri, ci, val, row_fmt)
            elif col == 'TAC':
                f = fmt['red'] if tac_red else fmt['num']
                n = to_num(val, default=None)
                if n is not None:
                    ws.write_number(ri, ci, n, f)
                else:
                    ws.write_blank(ri, ci, f)
            elif col == 'IRFIM {Prio} List':
                ws.write(ri, ci, val, fmt['red'] if irfim_missing else fmt['cell'])
            elif col == 'LNHOIF List':
                ws.write(ri, ci, val, fmt['red'] if lnhoif_missing else fmt['cell'])
            elif col == 'CAPR {Prio} List':
                ws.write(ri, ci, val, fmt['red'] if capr_missing else fmt['cell'])
            else:
                ws.write(ri, ci, val, fmt['cell'])

    for ci, col in enumerate(_LNCEL_COLS):
        ws.set_column(ci, ci, _LNCEL_WIDTHS.get(col, 15))
    ws.autofilter(0, 0, len(rows), len(_LNCEL_COLS) - 1)


_ARRAY_MODE = {
    0: 'Full 64TRX Array (4x8x2)',
    1: 'Left 32TRX Array (4x4x2)',
    2: 'Right 32TRX Array (4x4x2)',
    3: 'Top 32TRX Array (2x8x2)',
    4: 'Bottom 32TRX Array (2x8x2)',
    5: 'Full 32TRX Array (2x8x2)',
}

_MIMO_MODE = {
    0:  'SingleTX',
    10: 'TXDiv',
    11: '4-way TXDiv',
    30: 'Dynamic Open Loop MIMO (2x2)',
    40: 'Closed Loop MIMO (2x2)',
    41: 'Closed Loop MIMO (4x2)',
    42: 'Closed Loop MIMO (8x2)',
    43: 'Closed Loop MIMO (4x4)',
    44: 'Closed Loop MIMO (8x4)',
    50: 'Single Stream Beamforming',
    60: 'Dual Stream Beamforming',
}


def _admin_state(val):
    if val == '1':
        return 'Working'
    if val == '3':
        return 'Down'
    return val


def _iter_lncel_rows(network):
    # Gather all cells from both FDD and TDD, sort by (MRBTS, LNBTS, LNCEL)
    all_cells = []
    for recs in network.lncel_fdd_list_by_lnbts_dn.values():
        for r in recs:
            all_cells.append(('FDD', r))
    for recs in network.lncel_tdd_list_by_lnbts_dn.values():
        for r in recs:
            all_cells.append(('TDD', r))

    all_cells.sort(key=lambda x: (
        to_num(get(x[1], 'MRBTS')),
        to_num(get(x[1], 'LNBTS')),
        to_num(get(x[1], 'LNCEL')),
    ))

    for cell_type, mode_rec in all_cells:
        mrbts    = get(mode_rec, 'MRBTS')
        lnbts    = get(mode_rec, 'LNBTS')
        lncel_id = get(mode_rec, 'LNCEL')

        lnbts_k = _lnbts_dn(mrbts, lnbts)
        lncel_k = _lncel_dn(mrbts, lnbts, lncel_id)

        lnbts_rec  = network.lnbts_by_dn.get(lnbts_k, {})
        lnbts_name = get(lnbts_rec, 'name')

        lncel_rec   = network.lncel_by_dn.get(lncel_k, {})
        lncel_name  = get(lncel_rec, 'cellName') or get(lncel_rec, 'name')
        admin_state = _admin_state(get(lncel_rec, 'administrativeState'))
        mcc         = get(lncel_rec, 'mcc')
        mnc         = get(lncel_rec, 'mnc')
        pci         = get(lncel_rec, 'phyCellId')
        tac         = get(lncel_rec, 'tac')
        tilt        = get(lncel_rec, 'angle')
        pmax_raw    = to_num(get(lncel_rec, 'pMax'), default=None)
        pmax        = round(pmax_raw / 10, 1) if pmax_raw is not None else ''

        if cell_type == 'FDD':
            earfcn_dl = get(mode_rec, 'earfcnDL')
            chbw_raw  = to_num(get(mode_rec, 'dlChBw'), default=None)
        else:
            earfcn_dl = get(mode_rec, 'earfcn')
            chbw_raw  = to_num(get(mode_rec, 'chBw'), default=None)

        chbw              = round(chbw_raw / 10, 1) if chbw_raw is not None else ''
        mimo_raw          = to_num(get(mode_rec, 'dlMimoMode'), default=None)
        dl_mimo_mode      = _MIMO_MODE.get(int(mimo_raw), str(int(mimo_raw))) if mimo_raw is not None else ''
        rsi               = get(mode_rec, 'rootSeqIndex')
        if cell_type == 'TDD':
            arr_raw    = to_num(get(mode_rec, 'mMimoAntArrayMode'), default=None)
            array_mode = _ARRAY_MODE.get(int(arr_raw), str(int(arr_raw))) if arr_raw is not None else ''
        else:
            array_mode = ''

        rs_raw      = to_num(get(mode_rec, 'dlRsBoost'), default=None)
        dl_rs_boost = round((rs_raw - 1000) / 100, 2) if rs_raw is not None else ''

        n_rb = _BW_TO_NRB.get(chbw) if chbw != '' else None
        if (cell_type == 'TDD'
                and mimo_raw is not None and int(mimo_raw) == 60
                and arr_raw is not None
                and pmax != '' and dl_rs_boost != '' and n_rb is not None):
            # TDD Dual Stream Beamforming: PMAX - 10*log10(N_RB*12) + 10*log10(NumTRX/2) + dlRsBoost
            num_trx  = 64 if int(arr_raw) == 0 else 32
            rs_power = round(pmax - 10 * math.log10(n_rb * 12) + 10 * math.log10(num_trx / 2) + dl_rs_boost, 1)
        elif pmax != '' and dl_rs_boost != '' and n_rb is not None:
            # Standard formula: PMAX - 10*log10(N_RB*12) + dlRsBoost
            rs_power = round(pmax + dl_rs_boost - 10 * math.log10(n_rb * 12), 1)
        else:
            rs_power = ''

        own_earfcn  = to_num(earfcn_dl) if earfcn_dl else None
        all_earfcns = set(network.earfcns_for_lnbts(lnbts_k))
        required    = all_earfcns - ({own_earfcn} if own_earfcn is not None else set())

        # Build IRFIM list as "freq {prio}" sorted by descending eutCelResPrio
        irfim_freq_prio = {}   # freq → highest prio seen
        for r in network.irfim_list_by_lncel_dn.get(lncel_k, []):
            freq_raw = get(r, 'dlCarFrqEut')
            if not freq_raw:
                continue
            freq = to_num(freq_raw)
            prio = to_num(get(r, 'eutCelResPrio'), default=0)
            if freq not in irfim_freq_prio or prio > irfim_freq_prio[freq]:
                irfim_freq_prio[freq] = prio
        # Sort by descending priority, then ascending frequency
        irfim_sorted  = sorted(irfim_freq_prio.items(), key=lambda x: (-x[1], x[0]))
        irfim_list    = ', '.join(f'{int(f)} {{{p}}}' for f, p in irfim_sorted)
        irfim_freqs   = {f for f, _ in irfim_sorted}
        irfim_missing = bool(required - irfim_freqs)

        lnhoif_freqs = sorted({
            to_num(get(r, 'eutraCarrierInfo'))
            for r in network.lnhoif_list_by_lncel_dn.get(lncel_k, [])
            if get(r, 'eutraCarrierInfo')
        })
        lnhoif_list    = ', '.join(str(int(f)) for f in lnhoif_freqs)
        lnhoif_missing = bool(required - set(lnhoif_freqs))

        # Build CAPR list as "freq {prio}" sorted by descending sFreqPrio
        capr_freq_prio = {}   # freq → highest prio seen
        for r in network.capr_list_by_lncel_dn.get(lncel_k, []):
            freq_raw = get(r, 'earfcnDL')
            if not freq_raw:
                continue
            freq = to_num(freq_raw)
            prio = to_num(get(r, 'sFreqPrio'), default=0)
            if freq not in capr_freq_prio or prio > capr_freq_prio[freq]:
                capr_freq_prio[freq] = prio
        capr_sorted  = sorted(capr_freq_prio.items(), key=lambda x: (-x[1], x[0]))
        capr_list    = ', '.join(f'{int(f)} {{{p}}}' for f, p in capr_sorted)
        capr_freqs   = {f for f, _ in capr_sorted}
        capr_missing = bool(required - capr_freqs)

        # Cell reselection priority from SIB sheet
        sib_rec         = network.sib_by_lncel_dn.get(lncel_k, {})
        cell_resel_prio = get(sib_rec, 'cellReSelPrio')

        yield {
            'MRBTS ID':        mrbts,
            'LNBTS ID':        lnbts,
            'LNBTS Name':      lnbts_name,
            'LNCEL Name':      lncel_name,
            'LNCEL ID':        lncel_id,
            'Admin State':     admin_state,
            'MCC':             mcc,
            'MNC':             mnc,
            'PCI':             pci,
            'RSI':             rsi,
            'EARFCN DL':       earfcn_dl,
            'Ch BW (MHz)':     chbw,
            'PMAX (dBm)':      pmax,
            'dlRsBoost':       dl_rs_boost,
            'RS Power (dBm)':  rs_power,
            'DL MIMO Mode':    dl_mimo_mode,
            'Array Mode':      array_mode,
            'TAC':             tac,
            'Tilt':            tilt,
            'Cell Type':       cell_type,
            'SIB Priority':     cell_resel_prio,
            'IRFIM {Prio} List': irfim_list,
            'LNHOIF List':     lnhoif_list,
            'CAPR {Prio} List': capr_list,
            '_irfim_missing':  irfim_missing,
            '_lnhoif_missing': lnhoif_missing,
            '_capr_missing':   capr_missing,
        }


# ---------------------------------------------------------------------------
# Sheet 3 — Network Stats
# ---------------------------------------------------------------------------

def _build_network_stats(wb, fmt, network, log, lncel_rows):
    log('Building Network Stats sheet...')
    ws = wb.add_worksheet('Network Stats')

    # Extra formats for this sheet
    title_fmt = wb.add_format({
        'bold': True, 'font_size': 13, 'valign': 'vcenter',
    })
    cat_fmt = wb.add_format({
        'bold': True, 'border': 1, 'valign': 'vcenter',
        'bg_color': '#D6E4F0',
    })
    col_hdr_fmt = wb.add_format({
        'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#1F4E79', 'font_color': '#FFFFFF',
    })
    working_fmt = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#E2EFDA', 'num_format': '0',
    })
    other_fmt = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#FFF2CC', 'num_format': '0',
    })
    total_fmt = wb.add_format({
        'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter',
        'bg_color': '#DEEAF1', 'num_format': '0',
    })

    # ---- Compute stats ----
    stats = _compute_stats(network, lncel_rows)

    # ---- Layout ----
    ws.set_column(0, 0, 22)   # Category
    ws.set_column(1, 3, 12)   # Working / Other / Total

    ws.write(0, 0, '4G LTE Network Statistics', title_fmt)
    ws.set_row(0, 24)

    # Header row
    ws.write(2, 0, 'Category',  col_hdr_fmt)
    ws.write(2, 1, 'Working',   col_hdr_fmt)
    ws.write(2, 2, 'Other',     col_hdr_fmt)
    ws.write(2, 3, 'Total',     col_hdr_fmt)
    ws.set_row(2, 20)

    # Data rows
    for ri, (label, (working, other, total)) in enumerate(stats.items(), start=3):
        ws.write(ri, 0, label,   cat_fmt)
        ws.write_number(ri, 1, working, working_fmt)
        ws.write_number(ri, 2, other,   other_fmt)
        ws.write_number(ri, 3, total,   total_fmt)
        ws.set_row(ri, 18)

    log(f'  {len(stats)} stat categories written')


def _compute_stats(network, lncel_rows):
    """Returns OrderedDict of label → (working, other, total)."""

    # -- LNBTS --
    lnbts_total   = len(network.lnbts_by_dn)
    lnbts_working = sum(
        1 for r in network.lnbts_by_dn.values()
        if get(r, 'blockingState') == '1'
    )

    # -- FDD cells --
    fdd_rows    = [r for r in lncel_rows if r['Cell Type'] == 'FDD']
    fdd_total   = len(fdd_rows)
    fdd_working = sum(1 for r in fdd_rows if r['Admin State'] == 'Working')

    # -- TDD cells --
    tdd_rows    = [r for r in lncel_rows if r['Cell Type'] == 'TDD']
    tdd_total   = len(tdd_rows)
    tdd_working = sum(1 for r in tdd_rows if r['Admin State'] == 'Working')

    # -- FDD+TDD sites --
    fdd_tdd_dns     = [dn for dn in network.lnbts_by_dn if network.lte_mode(dn) == 'FDD+TDD']
    fdd_tdd_total   = len(fdd_tdd_dns)
    fdd_tdd_working = sum(
        1 for dn in fdd_tdd_dns
        if get(network.lnbts_by_dn[dn], 'blockingState') == '1'
    )

    return {
        'LNBTS (Sites)':  (_w(lnbts_working),   lnbts_total   - lnbts_working,   lnbts_total),
        'FDD Cells':      (_w(fdd_working),      fdd_total     - fdd_working,     fdd_total),
        'TDD Cells':      (_w(tdd_working),      tdd_total     - tdd_working,     tdd_total),
        'FDD+TDD Sites':  (_w(fdd_tdd_working),  fdd_tdd_total - fdd_tdd_working, fdd_tdd_total),
    }


def _w(n):
    """Passthrough — just for readability."""
    return n
