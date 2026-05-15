"""
reports/cell_summary.py
Builds the 2G Cell Summary Excel report from a Network object.
One row per unique Segment (2G cell).
"""

import xlsxwriter

from network import get, extract_id


# ---------------------------------------------------------------------------
# Column definitions
# Each entry: (header_label, extractor_fn(seg, network) -> value)
# To add a new column: just append a new _col() entry here.
# ---------------------------------------------------------------------------

def _col(label, fn):
    return {'label': label, 'fn': fn}


def _master(seg):
    return seg.get('master') or {}

def _slaves(seg):
    return seg.get('slaves') or []

def _admin_state(val):
    if val == '1': return 'Working'
    if val == '3': return 'Down'
    return val

def _phys_site_id(seg_name):
    """Left(3) + '0' + Mid(4,3)  e.g. 'NBI4123' -> 'NBI0412'"""
    if len(seg_name) >= 6:
        return seg_name[:3] + '0' + seg_name[3:6]
    return ''


COLUMNS = [
    _col('Physical Site ID',     lambda s, n: _phys_site_id(get(_master(s), 'segmentName'))),
    _col('BSC ID',               lambda s, n: extract_id(s['seg_dn'], 'BSC')),
    _col('BSC Name',             lambda s, n: get(n.get_bsc(
                                     'PLMN-PLMN/BSC-' + extract_id(s['seg_dn'], 'BSC')), 'name')),
    _col('BCF ID',               lambda s, n: extract_id(n.bcf_dn_for_segment(s), 'BCF')),
    _col('BCF Name',             lambda s, n: get(n.get_bcf(n.bcf_dn_for_segment(s)), 'name')),
    _col('Seg ID',               lambda s, n: extract_id(s['seg_dn'], 'SEG')),
    _col('Seg Name',             lambda s, n: get(_master(s), 'segmentName')),
    _col('Cell Name',            lambda s, n: get(_master(s), 'name')),
    _col('Bands',                lambda s, n: n.bands_for_segment(s)),
    _col('Master BTS ID',        lambda s, n: extract_id(get(_master(s), 'Dist_Name'), 'BTS')),
    _col('Slave BTS ID',         lambda s, n: ' & '.join(filter(None, [
                                     extract_id(get(r, 'Dist_Name'), 'BTS')
                                     for r in _slaves(s)]))),
    _col('MCC',                  lambda s, n: get(_master(s), 'locationAreaIdMCC')),
    _col('MNC',                  lambda s, n: get(_master(s), 'locationAreaIdMNC')),
    _col('NCC',                  lambda s, n: get(_master(s), 'bsIdentityCodeNCC')),
    _col('BCC',                  lambda s, n: get(_master(s), 'bsIdentityCodeBCC')),
    _col('LAC',                  lambda s, n: get(_master(s), 'locationAreaIdLAC')),
    _col('RAC',                  lambda s, n: get(_master(s), 'rac')),
    _col('Cell ID',              lambda s, n: get(_master(s), 'cellId')),
    _col('BCCH',                 lambda s, n: n.bcch_freq(_master(s))),
    _col('Hopping Mode',         lambda s, n: n.hopping_mode_and_mal(s)[0]),
    _col('MAL ID',               lambda s, n: n.hopping_mode_and_mal(s)[1]),
    _col('TCH Freq',             lambda s, n: n.tch_freqs(s)),
    _col('Master TRX Count',     lambda s, n: n.trx_count(_master(s))),
    _col('Slave TRX Count',      lambda s, n: sum(n.trx_count(r) for r in _slaves(s))),
    _col('Total TRX Count',      lambda s, n: (n.trx_count(_master(s)) +
                                     sum(n.trx_count(r) for r in _slaves(s)))),
    _col('Master TRX Power (W)', lambda s, n: n.max_trx_power_w(_master(s))),
    _col('Slave TRX Power (W)',  lambda s, n: (
                                     max((p for r in _slaves(s)
                                          for p in [n.max_trx_power_w(r)] if p != ''),
                                         default='')
                                     if _slaves(s) else '')),
    _col('Power Reduction 900',   lambda s, n: get(n.poc_for_bts(_master(s)), 'bsTxPwrMax')),
    _col('Power Reduction 1x00', lambda s, n: get(n.poc_for_bts(_master(s)), 'bsTxPwrMax1x00')),
    _col('Master NBL',           lambda s, n: get(_master(s), 'nonBCCHLayerOffset')),
    _col('Slave NBL',            lambda s, n: ' & '.join(filter(None, [
                                     get(r, 'nonBCCHLayerOffset') for r in _slaves(s)]))),
    _col('LAR',                  lambda s, n: n.lar(_master(s))),
    _col('LER',                  lambda s, n: n.ler(_master(s))),
    _col('Master Tilt',          lambda s, n: get(_master(s), 'angle')),
    _col('Slave Tilt',           lambda s, n: ' & '.join(filter(None, [
                                     get(r, 'angle') for r in _slaves(s)]))),
    _col('BCF Admin State',      lambda s, n: _admin_state(get(n.get_bcf(n.bcf_dn_for_segment(s)), 'adminState'))),
    _col('Master Admin State',   lambda s, n: _admin_state(get(_master(s), 'adminState'))),
    _col('Slave Admin State',    lambda s, n: ' & '.join(filter(None, [
                                     _admin_state(get(r, 'adminState')) for r in _slaves(s)]))),
    _col('Master LSEG',          lambda s, n: get(_master(s), 'btsLoadInSeg')),
    _col('Slave LSEG',           lambda s, n: ' & '.join(filter(None, [
                                     get(r, 'btsLoadInSeg') for r in _slaves(s)]))),
    _col('FRL',                  lambda s, n: get(_master(s), 'btsSpLoadDepTchRateLower')),
    _col('FRU',                  lambda s, n: get(_master(s), 'btsSpLoadDepTchRateUpper')),
    _col('AFRL',                 lambda s, n: get(_master(s), 'amrSegLoadDepTchRateLower')),
    _col('AFRU',                 lambda s, n: get(_master(s), 'amrSegLoadDepTchRateUpper')),
    _col('BLT',                  lambda s, n: get(_master(s), 'btsLoadThreshold')),
    _col('ADCE Count',            lambda s, n: len(n.adce_for_bts(_master(s)))),
    _col('One-Way ADCE Count',   lambda s, n: len(n.oneway_adce_for_bts(_master(s)))),
    _col('Discrepant ADCE Count',lambda s, n: n.discrepant_adce_count(_master(s))),
    _col('ADJW Count',           lambda s, n: len(n.adjw_for_bts(_master(s)))),
    _col('ADJL Count',           lambda s, n: len(n.adjl_for_bts(_master(s)))),
    _col('Diff LAC ADCE Count',  lambda s, n: n.diff_lac_adce_count(
                                     _master(s), get(_master(s), 'locationAreaIdLAC'))),
    _col('SDCCH Count',          lambda s, n: n.channel_type_counts(s)['sdcch']),
    _col('CCCH Count',           lambda s, n: n.channel_type_counts(s)['ccch']),
    _col('Not-Used RTSL Count',  lambda s, n: n.channel_type_counts(s)['not_used']),
    _col('TCH Count',            lambda s, n: n.channel_type_counts(s)['tch']),
    _col('Master GTCH Count',    lambda s, n: n.channel_type_counts(s)['gtch_master']),
    _col('Slave GTCH Count',     lambda s, n: n.channel_type_counts(s)['gtch_slave']),
    _col('Master CDED Count',    lambda s, n: n.channel_type_counts(s)['cded_master']),
    _col('Slave CDED Count',     lambda s, n: n.channel_type_counts(s)['cded_slave']),
    _col('Master CDEF Count',    lambda s, n: n.channel_type_counts(s)['cdef_master']),
    _col('Slave CDEF Count',     lambda s, n: n.channel_type_counts(s)['cdef_slave']),
    _col('Master Cell ID',       lambda s, n: get(_master(s), 'cellId')[1:]
                                     if get(_master(s), 'cellId') else ''),
    _col('NSEI',                 lambda s, n: get(_master(s), 'nsei')),
    _col('PSEI',                 lambda s, n: get(_master(s), 'psei')),
]

# Columns whose values should be written as numbers (not text)
_NUMERIC_COLS = {
    'BSC ID', 'BCF ID', 'Seg ID', 'Master BTS ID', 'Slave BTS ID',
    'NCC', 'BCC', 'LAC', 'RAC', 'Cell ID', 'BCCH',
    'Master NBL', 'Slave NBL', 'Master Tilt', 'Slave Tilt',
    'Power Reduction 900', 'Power Reduction 1x00',
    'Master LSEG', 'Slave LSEG',
    'FRL', 'FRU', 'AFRL', 'AFRU', 'BLT',
    'Master Cell ID', 'NSEI', 'PSEI',
}


def _to_num(val):
    """Convert val to int or float if possible, else return as-is."""
    if val == '' or val is None:
        return val
    s = str(val).strip()
    try:
        return int(s)
    except ValueError:
        pass
    try:
        return float(s)
    except ValueError:
        return val


_COL_WIDTHS = [
    15, 10, 18, 10, 25,                                    # Physical Site ID, BSC ID/Name, BCF ID/Name
    10, 14, 25, 12, 14, 14,                               # Seg ID ... Slave BTS ID
     8,  8,  8,  8,  8,  8, 10,                           # MCC ... Cell ID
     8, 14, 14, 35, 16, 14, 14, 18, 16,                    # BCCH, Hopping Mode, MAL ID, TCH Freq ... Slave TRX Power
    16, 18, 12, 12, 10, 10, 12, 12, 16, 16, 14, 12, 12,   # Pwr 900/1x00, NBL x2, LAR, LER, Tilt x2, Admin States x3, LSEG x2
     8,  8,  8,  8,  8,                                   # FRL, FRU, AFRL, AFRU, BLT
    12, 18, 20, 12, 12, 18,                                # ADCE, One-Way ADCE, Discrepant ADCE, ADJW, ADJL, Diff LAC ADCE
    12, 12, 18, 12, 16, 14, 16, 14, 16, 14, 14,            # SDCCH, CCCH, Not-Used RTSL, TCH, GTCH/CDED/CDEF M+S, Master Cell ID
    10, 10,                                                 # NSEI, PSEI
]


# ---------------------------------------------------------------------------
# Builder
# ---------------------------------------------------------------------------

def build(network, output_path, progress_fn=None, neighbour_checks=False):
    def log(msg):
        if progress_fn:
            progress_fn(msg)

    segments = network.segments
    log(f'Building Cell Details — {len(segments):,} segments...')

    wb = xlsxwriter.Workbook(output_path, {'strings_to_numbers': False})

    # --- Formats ---
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

    _border = {'left': 1, 'left_color': '#BDD7EE',
               'right': 1, 'right_color': '#BDD7EE',
               'bottom': 1, 'bottom_color': '#BDD7EE'}
    _base   = {'font_name': 'Arial', 'font_size': 9, 'valign': 'vcenter', 'bold': True}
    red_fmt = wb.add_format({**_base, **_border, 'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    yel_fmt = wb.add_format({**_base, **_border, 'bg_color': '#FFEB9C', 'font_color': '#9C6500'})
    blu_fmt = wb.add_format({**_base, **_border, 'bg_color': '#9DC3E6', 'font_color': '#1F3864'})

    # --- Cell Summary sheet ---
    ws = wb.add_worksheet('Cell Details')

    for ci, w in enumerate(_COL_WIDTHS):
        ws.set_column(ci, ci, w)

    ws.freeze_panes(1, 0)
    ws.set_row(0, 30)

    for ci, col_def in enumerate(COLUMNS):
        ws.write(0, ci, col_def['label'], hdr_fmt)

    # --- Pre-compute frequency highlights (shared BCF grouping) ---
    from collections import Counter

    bcch_col     = next((ci for ci, c in enumerate(COLUMNS) if c['label'] == 'BCCH'),     None)
    tch_freq_col = next((ci for ci, c in enumerate(COLUMNS) if c['label'] == 'TCH Freq'), None)
    lac_col      = next((ci for ci, c in enumerate(COLUMNS) if c['label'] == 'LAC'),      None)
    rac_col      = next((ci for ci, c in enumerate(COLUMNS) if c['label'] == 'RAC'),      None)
    cell_id_col  = next((ci for ci, c in enumerate(COLUMNS) if c['label'] == 'Cell ID'),  None)
    bcch_flag = {}     # seg_dn -> 'red' | 'yellow'
    tch_flag  = {}     # seg_dn -> 'red' | 'yellow'  (BCF-level check)
    tch_blue  = set()  # seg_dns where adjacent TCH exists within the same sector
    lac_red   = set()  # seg_dns where LAC is inconsistent within BCF
    rac_red   = set()  # seg_dns where RAC is inconsistent within BCF
    cell_id_flag = {}  # seg_dn -> 'red' | 'yellow'  (only populated for duplicates)

    if cell_id_col is not None:
        # Group seg_dns by Cell ID
        cid_groups = {}
        for seg_dn, seg in segments.items():
            cid = get(seg.get('master') or {}, 'cellId')
            if cid:
                cid_groups.setdefault(cid, []).append(seg_dn)

        for cid, dn_list in cid_groups.items():
            if len(dn_list) < 2:
                continue
            # Red only when EVERY duplicate has both BCF Admin State AND Master Admin State = Working
            all_working = all(
                _admin_state(get(segments[dn].get('master') or {}, 'adminState')) == 'Working'
                and _admin_state(get(network.get_bcf(
                    network.bcf_dn_for_segment(segments[dn])), 'adminState')) == 'Working'
                for dn in dn_list
            )
            flag = 'red' if all_working else 'yellow'
            for dn in dn_list:
                cell_id_flag[dn] = flag

    # Group segments by BCF DN (used by both checks)
    bcf_segs = {}
    for seg_dn, seg in segments.items():
        bcf_dn = network.bcf_dn_for_segment(seg)
        if bcf_dn:
            bcf_segs.setdefault(bcf_dn, []).append((seg_dn, seg))

    for bcf_dn, seg_list in bcf_segs.items():

        # -- BCCH check: pool = one BCCH freq per segment --
        if bcch_col is not None:
            bcch_map = {}
            for seg_dn, seg in seg_list:
                try:
                    bcch_map[seg_dn] = int(network.bcch_freq(seg.get('master') or {}))
                except (ValueError, TypeError):
                    pass
            if len(bcch_map) >= 2:
                freq_counts = Counter(bcch_map.values())
                all_freqs   = set(bcch_map.values())
                for seg_dn, freq in bcch_map.items():
                    if freq_counts[freq] > 1:
                        bcch_flag[seg_dn] = 'red'
                    elif (freq - 1) in all_freqs or (freq + 1) in all_freqs:
                        bcch_flag[seg_dn] = 'yellow'

        # -- TCH Freq check: pool = ALL TRX frequencies (BCCH + TCH) in the BCF --
        if tch_freq_col is not None:
            # Use pre-built BCF frequency pool from network model (no re-iteration needed)
            bcf_pool      = network.bcf_freq_pool.get(bcf_dn, Counter())
            all_bcf_freqs = set(bcf_pool.keys())

            if bcf_pool:
                for seg_dn, seg in seg_list:
                    # TCH freqs already cached in network model — parse the stored string
                    seg_tch_freqs = []
                    tch_str = network.tch_freqs(seg)
                    if tch_str:
                        for f in tch_str.split(', '):
                            try:
                                seg_tch_freqs.append(int(f))
                            except (ValueError, TypeError):
                                pass

                    flag = None
                    for freq in seg_tch_freqs:
                        if bcf_pool[freq] > 1:
                            flag = 'red'
                            break
                        elif (freq - 1) in all_bcf_freqs or (freq + 1) in all_bcf_freqs:
                            flag = 'yellow'
                    if flag:
                        tch_flag[seg_dn] = flag

                    # Blue check: adjacent TCH within the same sector (segment).
                    # Tighter than yellow — checks adjacency among this segment's
                    # own freqs only, not the whole BCF pool.
                    if len(seg_tch_freqs) >= 2:
                        seg_freq_set = set(seg_tch_freqs)
                        if any((f - 1) in seg_freq_set or (f + 1) in seg_freq_set
                               for f in seg_tch_freqs):
                            tch_blue.add(seg_dn)

        # -- LAC / RAC consistency check --
        lacs = {seg_dn: get(seg.get('master') or {}, 'locationAreaIdLAC')
                for seg_dn, seg in seg_list}
        racs = {seg_dn: get(seg.get('master') or {}, 'rac')
                for seg_dn, seg in seg_list}
        if len(set(lacs.values())) > 1:
            lac_red.update(lacs.keys())
        if len(set(racs.values())) > 1:
            rac_red.update(racs.keys())

    numeric_cols = {ci for ci, c in enumerate(COLUMNS) if c['label'] in _NUMERIC_COLS}

    for ri, (seg_dn, seg) in enumerate(sorted(segments.items()), 1):
        fmt = alt_fmt if ri % 2 == 1 else data_fmt
        for ci, col_def in enumerate(COLUMNS):
            try:
                val = col_def['fn'](seg, network)
            except Exception:
                val = ''
            if   ci == bcch_col     and seg_dn in bcch_flag:
                cell_fmt = red_fmt if bcch_flag[seg_dn] == 'red' else yel_fmt
            elif ci == tch_freq_col and (seg_dn in tch_flag or seg_dn in tch_blue):
                if tch_flag.get(seg_dn) == 'red':
                    cell_fmt = red_fmt
                elif seg_dn in tch_blue:
                    cell_fmt = blu_fmt
                else:
                    cell_fmt = yel_fmt
            elif ci == lac_col      and seg_dn in lac_red:
                cell_fmt = red_fmt
            elif ci == rac_col      and seg_dn in rac_red:
                cell_fmt = red_fmt
            elif ci == cell_id_col  and seg_dn in cell_id_flag:
                cell_fmt = red_fmt if cell_id_flag[seg_dn] == 'red' else yel_fmt
            else:
                cell_fmt = fmt
            if ci in numeric_cols:
                val = _to_num(val)
            ws.write(ri, ci, val, cell_fmt)

    # --- BCF Details sheet (tab 2) ---
    ws2 = wb.add_worksheet('BCF Details')

    bcf_headers = ['Physical Site ID', 'BSC ID', 'BSC Name', 'BCF ID', 'BCF Name', 'SBTS ID',
                   'M-Plane IP', 'Admin State', 'Cell Count', 'BTS Count', 'TRX Count']
    bcf_widths  = [20, 10, 18, 10, 25, 12, 28, 14, 12, 12, 12]

    for ci, w in enumerate(bcf_widths):
        ws2.set_column(ci, ci, w)
    ws2.freeze_panes(1, 0)
    ws2.set_row(0, 30)
    for ci, label in enumerate(bcf_headers):
        ws2.write(0, ci, label, hdr_fmt)

    # Group segments by BCF DN
    bcf_to_segs = {}
    for seg in segments.values():
        bcf_dn = network.bcf_dn_for_segment(seg)
        if bcf_dn:
            bcf_to_segs.setdefault(bcf_dn, []).append(seg)

    for ri, (bcf_dn, bcf_r) in enumerate(sorted(network.bcf_by_dn.items()), 1):
        fmt = alt_fmt if ri % 2 == 1 else data_fmt
        bsc_dn = 'PLMN-PLMN/BSC-' + extract_id(bcf_dn, 'BSC')
        bsc_r  = network.get_bsc(bsc_dn)

        segs_in_bcf = bcf_to_segs.get(bcf_dn, [])
        cell_count  = len(segs_in_bcf)
        bts_count   = sum(len(seg['all_bts']) for seg in segs_in_bcf)
        trx_count   = sum(
            network.trx_count(bts_r)
            for seg in segs_in_bcf
            for bts_r in seg['all_bts']
        )

        # Physical Site ID: unique IDs from all segments in this BCF, order-preserved
        site_ids = sorted(set(
            _phys_site_id(get(seg.get('master') or {}, 'segmentName'))
            for seg in segs_in_bcf
        ))
        phys_site = ' & '.join(s for s in site_ids if s)

        ip = get(bcf_r, 'btsMPlaneIpAddress')
        row = [
            phys_site,
            extract_id(bcf_dn, 'BSC'),
            get(bsc_r, 'name'),
            extract_id(bcf_dn, 'BCF'),
            get(bcf_r, 'name'),
            get(bcf_r, 'SBTSId'),
            ('https://' + ip) if ip else '',
            _admin_state(get(bcf_r, 'adminState')),
            cell_count,
            bts_count,
            trx_count,
        ]
        for ci, val in enumerate(row):
            ws2.write(ri, ci, val, fmt)

    # --- Optional neighbour-check sheets (skipped unless neighbour_checks=True) ---
    if neighbour_checks:
        log('Generating One-Way ADCE...')
        ws_ow = wb.add_worksheet('One-Way ADCE')
        ow_headers = ['Source Seg Name', 'Source DN', 'Target Cell ID', 'Target LAC',
                      'Target Seg Name', 'Remarks']
        ow_widths  = [20, 55, 16, 12, 20, 35]
        for ci, w in enumerate(ow_widths):
            ws_ow.set_column(ci, ci, w)
        ws_ow.freeze_panes(1, 0)
        ws_ow.set_row(0, 30)
        for ci, label in enumerate(ow_headers):
            ws_ow.write(0, ci, label, hdr_fmt)

        for ow_ri, row in enumerate(network.all_oneway_adce_rows(), 1):
            fmt = alt_fmt if ow_ri % 2 == 1 else data_fmt
            ws_ow.write(ow_ri, 0, row['src_seg_name'], fmt)
            ws_ow.write(ow_ri, 1, row['src_bts_dn'],   fmt)
            ws_ow.write(ow_ri, 2, row['tgt_ci'],        fmt)
            ws_ow.write(ow_ri, 3, row['tgt_lac'],       fmt)
            ws_ow.write(ow_ri, 4, row['tgt_seg_name'],  fmt)
            ws_ow.write(ow_ri, 5, row['remark'],        fmt)

        log('Generating Discrepant ADCE...')
        ws_disc = wb.add_worksheet('Discrepant ADCE')
        disc_headers = [
            'Source BSC', 'Source Seg DN', 'Source Seg Name', 'Source Cell Name', 'Neighbour CI',
            'Neighbour LAC (ADCE)', 'Neighbour LAC (Actual)',
            'NCC (ADCE)', 'NCC (Actual)',
            'BCC (ADCE)', 'BCC (Actual)',
            'MCC (ADCE)', 'MCC (Actual)',
            'MNC (ADCE)', 'MNC (Actual)',
            'BCCH (ADCE)', 'BCCH (Actual)',
            'Remarks',
        ]
        disc_widths = [
            20, 55, 18, 25, 14,      # Source BSC … Neighbour CI
            18, 18,                   # LAC pair
            14, 14,                   # NCC pair
            14, 14,                   # BCC pair
            14, 14,                   # MCC pair
            14, 14,                   # MNC pair
            14, 14,                   # BCCH pair
            40,                       # Remarks
        ]
        for ci, w in enumerate(disc_widths):
            ws_disc.set_column(ci, ci, w)
        ws_disc.freeze_panes(1, 0)
        ws_disc.set_row(0, 30)
        for ci, label in enumerate(disc_headers):
            ws_disc.write(0, ci, label, hdr_fmt)

        # Field order must match disc_headers columns 5-16
        _DISC_FIELDS = ['LAC', 'NCC', 'BCC', 'MCC', 'MNC', 'BCCH']
        _DISC_COL_START = 5  # first ADCE-value column index

        disc_ri = 1
        for row in network.all_discrepant_adce_rows():
            fmt = alt_fmt if disc_ri % 2 == 1 else data_fmt
            ws_disc.write(disc_ri, 0, row['src_bsc_name'],  fmt)
            ws_disc.write(disc_ri, 1, row['src_seg_dn'],    fmt)
            ws_disc.write(disc_ri, 2, row['src_seg_name'],  fmt)
            ws_disc.write(disc_ri, 3, row['src_cell_name'], fmt)
            ws_disc.write(disc_ri, 4, row['tgt_ci'],        fmt)
            for fi, field in enumerate(_DISC_FIELDS):
                adce_ci   = _DISC_COL_START + fi * 2
                actual_ci = adce_ci + 1
                cell_fmt  = red_fmt if field in row['mismatched'] else fmt
                ws_disc.write(disc_ri, adce_ci,   row['adce_vals'][field],   cell_fmt)
                ws_disc.write(disc_ri, actual_ci, row['actual_vals'][field], cell_fmt)
            ws_disc.write(disc_ri, 17, row['remark'], fmt)
            disc_ri += 1

        log('Generating Co-Site Missing Neighbours...')
        ws_co = wb.add_worksheet('Co-Site Missing Neighbours')
        co_headers = ['BCF Name', 'BCF Admin State',
                      'Source Seg Name', 'Source Cell Name', 'Source Admin State',
                      'Target Seg Name', 'Target Cell Name', 'Target Admin State']
        co_widths  = [25, 14, 20, 25, 16, 20, 25, 16]
        for ci, w in enumerate(co_widths):
            ws_co.set_column(ci, ci, w)
        ws_co.freeze_panes(1, 0)
        ws_co.set_row(0, 30)
        for ci, label in enumerate(co_headers):
            ws_co.write(0, ci, label, hdr_fmt)

        co_ri = 1
        for bcf_dn, seg_list in sorted(bcf_segs.items()):
            bcf_r     = network.get_bcf(bcf_dn)
            bcf_name  = get(bcf_r, 'name')
            bcf_admin = _admin_state(get(bcf_r, 'adminState'))

            for seg_dn_x, seg_x in seg_list:
                master_x    = seg_x.get('master') or {}
                master_x_dn = get(master_x, 'Dist_Name')
                x_neighbors = network._adce_neighbors_by_bts.get(master_x_dn, set())

                for seg_dn_y, seg_y in seg_list:
                    if seg_dn_x == seg_dn_y:
                        continue
                    master_y = seg_y.get('master') or {}
                    y_ci     = get(master_y, 'cellId')
                    y_lac    = get(master_y, 'locationAreaIdLAC')
                    if not y_ci or not y_lac:
                        continue
                    if (y_ci, y_lac) in x_neighbors:
                        continue   # neighbour already defined

                    fmt = alt_fmt if co_ri % 2 == 1 else data_fmt
                    ws_co.write(co_ri, 0, bcf_name,                                  fmt)
                    ws_co.write(co_ri, 1, bcf_admin,                                  fmt)
                    ws_co.write(co_ri, 2, get(master_x, 'segmentName'),               fmt)
                    ws_co.write(co_ri, 3, get(master_x, 'name'),                      fmt)
                    ws_co.write(co_ri, 4, _admin_state(get(master_x, 'adminState')),  fmt)
                    ws_co.write(co_ri, 5, get(master_y, 'segmentName'),               fmt)
                    ws_co.write(co_ri, 6, get(master_y, 'name'),                      fmt)
                    ws_co.write(co_ri, 7, _admin_state(get(master_y, 'adminState')),  fmt)
                    co_ri += 1

    # --- Frequency Reuse sheet ---
    log('Generating Frequency Reuse...')
    ws_fr = wb.add_worksheet('Frequency Reuse')
    fr_headers = ['ARFCN', 'BCCH/TCH', 'Number of Occurrences']
    fr_widths  = [12, 12, 24]
    for ci, w in enumerate(fr_widths):
        ws_fr.set_column(ci, ci, w)
    ws_fr.freeze_panes(1, 0)
    ws_fr.set_row(0, 30)
    for ci, label in enumerate(fr_headers):
        ws_fr.write(0, ci, label, hdr_fmt)

    # Build per-type frequency counters across all segments.
    # BCCH: one per segment (master BTS BCCH TRX).
    # TCH : all non-BCCH effective frequencies per segment
    #        (MAL freqs for RF-hopping, TRX initialFrequency otherwise).
    bcch_counter = Counter()
    tch_counter  = Counter()

    for seg in segments.values():
        master = seg.get('master') or {}
        # BCCH
        bcch_val = network.bcch_freq(master)
        if bcch_val:
            try:
                bcch_counter[int(bcch_val)] += 1
            except (ValueError, TypeError):
                bcch_counter[bcch_val] += 1
        # TCH — already computed & cached by tch_freqs()
        tch_str = network.tch_freqs(seg)
        if tch_str:
            for f in tch_str.split(', '):
                f = f.strip()
                if not f:
                    continue
                try:
                    tch_counter[int(f)] += 1
                except (ValueError, TypeError):
                    tch_counter[f] += 1

    # Sort helper: integers first (numerically), then stray strings alphabetically
    def _arfcn_key(f):
        return (isinstance(f, str), f)

    bcch_arfcns = sorted(bcch_counter, key=_arfcn_key)
    tch_arfcns  = sorted(tch_counter,  key=_arfcn_key)

    fr_ri = 1
    # --- BCCH block first ---
    for arfcn in bcch_arfcns:
        fmt = alt_fmt if fr_ri % 2 == 1 else data_fmt
        ws_fr.write(fr_ri, 0, arfcn,               fmt)
        ws_fr.write(fr_ri, 1, 'BCCH',              fmt)
        ws_fr.write(fr_ri, 2, bcch_counter[arfcn], fmt)
        fr_ri += 1
    # --- TCH block second ---
    for arfcn in tch_arfcns:
        fmt = alt_fmt if fr_ri % 2 == 1 else data_fmt
        ws_fr.write(fr_ri, 0, arfcn,              fmt)
        ws_fr.write(fr_ri, 1, 'TCH',              fmt)
        ws_fr.write(fr_ri, 2, tch_counter[arfcn], fmt)
        fr_ri += 1

    # --- Network Stats sheet ---
    ws3 = wb.add_worksheet('Network Stats')
    ws3.set_column(0, 0, 28)
    ws3.set_column(1, 3, 12)

    title_fmt = wb.add_format({'font_name': 'Arial', 'font_size': 12, 'bold': True})
    hdr2_fmt  = wb.add_format({'font_name': 'Arial', 'font_size': 10, 'bold': True, 'align': 'center'})
    lbl_fmt   = wb.add_format({'font_name': 'Arial', 'font_size': 10})
    val_fmt   = wb.add_format({'font_name': 'Arial', 'font_size': 10, 'bold': True, 'align': 'center'})

    ws3.write(0, 0, 'Summary Statistics', title_fmt)
    for ci, label in enumerate(['Working', 'Other', 'Total'], 1):
        ws3.write(1, ci, label, hdr2_fmt)

    all_segs = list(segments.values())
    all_bcf  = list(network.bcf_by_dn.values())
    all_trx  = [trx for trxs in network.trx_by_bts_dn.values() for trx in trxs]

    def _wot(records):
        working = sum(1 for r in records if get(r, 'adminState') == '1')
        total   = len(records)
        return working, total - working, total

    def _seg_wot(filtered):
        return _wot([s['master'] or {} for s in filtered])

    seg_w, seg_o, seg_t = _wot([s['master'] or {} for s in all_segs])
    bcf_w, bcf_o, bcf_t = _wot(all_bcf)
    trx_w, trx_o, trx_t = _wot(all_trx)
    bsc_w = len(set(extract_id(s['seg_dn'], 'BSC') for s in all_segs))
    bsc_t = len(network.bsc_by_dn)

    seg_bands = {s['seg_dn']: network.bands_for_segment(s) for s in all_segs}
    dual_w,  dual_o,  dual_t  = _seg_wot([s for s in all_segs if '&' in seg_bands[s['seg_dn']]])
    g900_w,  g900_o,  g900_t  = _seg_wot([s for s in all_segs if seg_bands[s['seg_dn']] == '900'])
    g1800_w, g1800_o, g1800_t = _seg_wot([s for s in all_segs if seg_bands[s['seg_dn']] == '1800'])
    slv_w,   slv_o,   slv_t   = _seg_wot([s for s in all_segs if s['slaves']])

    stats = [
        ('Total Segments (Cells)',   seg_w,         seg_o,         seg_t),
        ('Dual-band Segments',       dual_w,        dual_o,        dual_t),
        ('900-only Segments',        g900_w,        g900_o,        g900_t),
        ('1800-only Segments',       g1800_w,       g1800_o,       g1800_t),
        ('Segments with Slave BTS',  slv_w,         slv_o,         slv_t),
        ('Unique BSCs',              bsc_w,         bsc_t - bsc_w, bsc_t),
        ('Unique BCFs',              bcf_w,         bcf_o,         bcf_t),
        ('Total TRX in Network',     trx_w,         trx_o,         trx_t),
    ]

    for i, (label, working, other, total) in enumerate(stats, 2):
        ws3.write(i, 0, label,   lbl_fmt)
        ws3.write(i, 1, working, val_fmt)
        ws3.write(i, 2, other,   val_fmt)
        ws3.write(i, 3, total,   val_fmt)

    wb.close()
    log(f'Saved: {output_path}')
    return len(segments)
