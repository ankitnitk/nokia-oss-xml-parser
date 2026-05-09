"""
hw_tool/report.py  —  HW Inventory Report Builder

Builds a 3-sheet Excel workbook from INVUNIT + MRBTS/LNBTS data:

  Sheet 1 "Site wise (All)"     — one row per MRBTS, one column per
                                   inventoryUnitType, cell = total count.
                                   Row 0 = group banner (RMOD / BBMOD / SMOD / Others),
                                   Row 1 = column headers, Row 2+ = data.
  Sheet 2 "Site wise (Working)" — same layout but only state=working units
  Sheet 3 "Overall"             — one row per inventoryUnitType,
                                   columns: Working | Total | Group

Column ordering for inventoryUnitType (determined by vendorUnitFamilyType):
  Group 0 — family contains "RMOD"   (radio modules / RRH)
  Group 1 — family contains "BBMOD"  (baseband modules)
  Group 2 — family contains "SMOD"   (system modules)
  Group 3 — everything else          (cabinets, fans, PSU, …)
  Within each group: alphabetical by inventoryUnitType name.
"""

from collections import defaultdict
from itertools import groupby

import xlsxwriter


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_GROUP_LABELS = {0: 'RMOD', 1: 'BBMOD', 2: 'SMOD', 3: 'Others'}

_GROUP_COLORS = {
    0: '#C6EFCE',   # light green  — RMOD
    1: '#DDEBF7',   # light blue   — BBMOD
    2: '#FCE4D6',   # light orange — SMOD
    3: '#F2F2F2',   # light grey   — Others
}

_GROUP_FONT_COLORS = {
    0: '#276221',   # dark green
    1: '#1F4E79',   # dark blue
    2: '#833C00',   # dark orange
    3: '#595959',   # dark grey
}

_DARK_BLUE = '#1F497D'


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _is_working(state_val):
    """True if the INVUNIT state value means 'working'."""
    if state_val is None:
        return False
    s = str(state_val).strip().lower()
    return s in ('working', '1')


def _group_key(fam_type):
    """
    Return sort-group integer for a vendorUnitFamilyType string.
    0 = RMOD,  1 = BBMOD,  2 = SMOD,  3 = everything else.
    """
    u = (fam_type or '').upper()
    if 'RMOD'  in u: return 0
    if 'BBMOD' in u: return 1
    if 'SMOD'  in u: return 2
    return 3


def _group_spans(all_inv_types, fam_map):
    """
    Return list of (group_int, col_start, col_end) for consecutive groups
    in all_inv_types.  col indices are relative to the inv-type columns only
    (i.e. col 0 here = spreadsheet col 2 in the site-wise sheets).
    """
    spans = []
    for grp, items in groupby(enumerate(all_inv_types),
                              key=lambda x: _group_key(fam_map[x[1]])):
        idxs = [i for i, _ in items]
        spans.append((grp, idxs[0], idxs[-1]))
    return spans


# ---------------------------------------------------------------------------
# Workbook formats
# ---------------------------------------------------------------------------

def _make_formats(wb, group_colors, group_font_colors):
    # Column header (dark blue, row 1 of site-wise sheets)
    col_hdr = wb.add_format({
        'bold': True, 'bg_color': _DARK_BLUE, 'font_color': 'white',
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
    })
    # Group banner formats (row 0 of site-wise sheets), one per group
    grp_fmt = {}
    for g, bg in group_colors.items():
        grp_fmt[g] = wb.add_format({
            'bold': True,
            'bg_color': bg,
            'font_color': group_font_colors[g],
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
        })
    # Group label format for Overall col 4 (same colours, no border on all sides)
    grp_cell_fmt = {}
    for g, bg in group_colors.items():
        grp_cell_fmt[g] = wb.add_format({
            'bold': False,
            'bg_color': bg,
            'font_color': group_font_colors[g],
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
        })

    str_cell  = wb.add_format({'border': 1, 'valign': 'vcenter'})
    num_cell  = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    zero_cell = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_color': '#CCCCCC',
    })
    int_cell  = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0',
    })
    return col_hdr, grp_fmt, grp_cell_fmt, str_cell, num_cell, zero_cell, int_cell


# ---------------------------------------------------------------------------
# Core builder
# ---------------------------------------------------------------------------

def build_hw_report(sheets, output_path):
    """
    Build the 3-sheet HW report and write it to *output_path*.

    sheets       — dict: sheet_name → list of row-dicts.
                   Required key: 'INVUNIT'.
                   Optional keys: 'MRBTS', 'LNBTS' (for site names).

    output_path  — destination .xlsx file path.

    Returns the number of MRBTS sites processed.
    """
    invunit_rows = sheets.get('INVUNIT', [])
    mrbts_rows   = sheets.get('MRBTS',   [])
    lnbts_rows   = sheets.get('LNBTS',   [])

    # ── Site-name lookup: MRBTS id (str) → site name ─────────────────────────
    site_name = {}
    for row in (mrbts_rows or lnbts_rows):
        mid  = row.get('MRBTS')
        name = row.get('name') or row.get('btsName')
        if mid is not None and name:
            site_name[str(mid)] = str(name)

    # ── Aggregate INVUNIT rows ───────────────────────────────────────────────
    all_counts     = defaultdict(lambda: defaultdict(int))
    working_counts = defaultdict(lambda: defaultdict(int))
    fam_counter    = defaultdict(lambda: defaultdict(int))
    seen_mrbts     = set()

    for row in invunit_rows:
        mrbts    = row.get('MRBTS')
        inv_type = row.get('inventoryUnitType')
        fam_type = row.get('vendorUnitFamilyType') or ''
        state    = row.get('state')

        if mrbts is None or not inv_type:
            continue

        inv_type  = str(inv_type).strip()
        fam_type  = str(fam_type).strip()
        mrbts_key = str(mrbts)

        if mrbts_key not in seen_mrbts:
            seen_mrbts.add(mrbts_key)

        all_counts[mrbts_key][inv_type] += 1
        fam_counter[inv_type][fam_type] += 1

        if _is_working(state):
            working_counts[mrbts_key][inv_type] += 1

    # ── Column order for inventoryUnitType ───────────────────────────────────
    # Representative family = most common vendorUnitFamilyType for each inv type
    fam_map = {
        inv_type: max(fam_counts, key=fam_counts.get)
        for inv_type, fam_counts in fam_counter.items()
    }

    all_inv_types = sorted(
        fam_counter.keys(),
        key=lambda t: (_group_key(fam_map[t]), t.upper())
    )

    # Group per inv_type (for quick lookup)
    inv_group = {t: _group_key(fam_map[t]) for t in all_inv_types}

    # Consecutive group spans: list of (group, first_inv_col_idx, last_inv_col_idx)
    # inv_col_idx is 0-based within all_inv_types; sheet col = inv_col_idx + 2
    spans = _group_spans(all_inv_types, fam_map)

    # Sort MRBTS numerically
    all_mrbts = sorted(seen_mrbts, key=lambda x: (int(x) if x.isdigit() else x))

    # ── Write workbook ───────────────────────────────────────────────────────
    wb = xlsxwriter.Workbook(output_path)
    col_hdr_fmt, grp_fmt, grp_cell_fmt, str_fmt, num_fmt, zero_fmt, int_fmt = \
        _make_formats(wb, _GROUP_COLORS, _GROUP_FONT_COLORS)

    def _write_site_sheet(ws_name, count_dict):
        ws = wb.add_worksheet(ws_name)
        # freeze at row 2 (below group banner + header), col 2 (after MRBTS + Site Name)
        ws.freeze_panes(2, 2)
        ws.set_zoom(85)
        ws.set_default_row(15)

        n_inv = len(all_inv_types)
        n_hdr = n_inv + 2   # total columns

        # ── Row 0: group banner ───────────────────────────────────────────────
        ws.set_row(0, 18)
        # First two columns: blank with dark-blue header style
        ws.write(0, 0, '', col_hdr_fmt)
        ws.write(0, 1, '', col_hdr_fmt)

        for grp, i_start, i_end in spans:
            c_start = i_start + 2   # sheet column
            c_end   = i_end   + 2
            label   = _GROUP_LABELS[grp]
            fmt     = grp_fmt[grp]
            if c_start == c_end:
                ws.write(0, c_start, label, fmt)
            else:
                ws.merge_range(0, c_start, 0, c_end, label, fmt)

        # ── Row 1: column headers ─────────────────────────────────────────────
        ws.set_row(1, 45)
        ws.write(1, 0, 'MRBTS',     col_hdr_fmt)
        ws.write(1, 1, 'Site Name', col_hdr_fmt)
        for c, inv_type in enumerate(all_inv_types, start=2):
            ws.write(1, c, inv_type, col_hdr_fmt)

        # ── Column widths ─────────────────────────────────────────────────────
        ws.set_column(0, 0, 10)           # MRBTS
        ws.set_column(1, 1, 32)           # Site Name
        ws.set_column(2, n_hdr - 1, 9)   # count columns

        # ── Data rows (start at row 2) ────────────────────────────────────────
        for r, mrbts in enumerate(all_mrbts, start=2):
            mrbts_val = int(mrbts) if mrbts.isdigit() else mrbts
            ws.write(r, 0, mrbts_val, int_fmt)
            ws.write(r, 1, site_name.get(mrbts, ''), str_fmt)
            for c, inv_type in enumerate(all_inv_types, start=2):
                cnt = count_dict[mrbts].get(inv_type, 0)
                if cnt:
                    ws.write(r, c, cnt, num_fmt)
                else:
                    ws.write_blank(r, c, None, zero_fmt)

    _write_site_sheet('Site wise (All)',     all_counts)
    _write_site_sheet('Site wise (Working)', working_counts)

    # ── Overall sheet ────────────────────────────────────────────────────────
    ws_ov = wb.add_worksheet('Overall')
    ws_ov.freeze_panes(1, 1)
    ws_ov.set_zoom(100)

    for c, h in enumerate(['inventoryUnitType', 'Working', 'Total', 'Group']):
        ws_ov.write(0, c, h, col_hdr_fmt)
    ws_ov.set_row(0, 20)
    ws_ov.set_column(0, 0, 34)
    ws_ov.set_column(1, 2, 14)
    ws_ov.set_column(3, 3, 10)

    for r, inv_type in enumerate(all_inv_types, start=1):
        total_w = sum(working_counts[m].get(inv_type, 0) for m in all_mrbts)
        total_a = sum(all_counts[m].get(inv_type, 0)     for m in all_mrbts)
        grp     = inv_group[inv_type]
        ws_ov.write(r, 0, inv_type,              str_fmt)
        ws_ov.write(r, 1, total_w,               num_fmt)
        ws_ov.write(r, 2, total_a,               num_fmt)
        ws_ov.write(r, 3, _GROUP_LABELS[grp],    grp_cell_fmt[grp])

    wb.close()
    return len(all_mrbts)
