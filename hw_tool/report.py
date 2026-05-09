"""
hw_tool/report.py  —  HW Inventory Report Builder

Builds a 3-sheet Excel workbook from INVUNIT + MRBTS/LNBTS data:

  Sheet 1 "Site wise (All)"     — one row per MRBTS, one column per
                                   inventoryUnitType, cell = total count
  Sheet 2 "Site wise (Working)" — same layout but only state=working units
  Sheet 3 "Overall"             — one row per inventoryUnitType,
                                   columns: Working count | Total count

Column ordering for inventoryUnitType (determined by vendorUnitFamilyType):
  Group 0 — family contains "RMOD"   (radio modules / RRH)
  Group 1 — family contains "BBMOD"  (baseband modules)
  Group 2 — family contains "SMOD"   (system modules)
  Group 3 — everything else          (cabinets, fans, PSU, …)
  Within each group: alphabetical by inventoryUnitType name.
"""

from collections import defaultdict

import xlsxwriter


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


# ---------------------------------------------------------------------------
# Workbook formats
# ---------------------------------------------------------------------------

_DARK_BLUE = '#1F497D'


def _make_formats(wb):
    hdr = wb.add_format({
        'bold': True, 'bg_color': _DARK_BLUE, 'font_color': 'white',
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
    })
    str_cell  = wb.add_format({'border': 1, 'valign': 'vcenter'})
    num_cell  = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    zero_cell = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_color': '#CCCCCC',
    })
    int_cell  = wb.add_format({
        'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '0',
    })
    return hdr, str_cell, num_cell, zero_cell, int_cell


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
    # Prefer MRBTS sheet ('name' column); fall back to LNBTS ('name' column).
    site_name = {}
    for row in (mrbts_rows or lnbts_rows):
        mid  = row.get('MRBTS')
        name = row.get('name') or row.get('btsName')
        if mid is not None and name:
            site_name[str(mid)] = str(name)

    # ── Aggregate INVUNIT rows ───────────────────────────────────────────────
    #
    # all_counts[mrbts_key][inv_type]     = total unit count
    # working_counts[mrbts_key][inv_type] = working-state unit count
    # fam_counter[inv_type][fam_type]     = occurrence count
    #   → used to pick a representative vendorUnitFamilyType for sorting

    all_counts     = defaultdict(lambda: defaultdict(int))
    working_counts = defaultdict(lambda: defaultdict(int))
    fam_counter    = defaultdict(lambda: defaultdict(int))

    seen_mrbts  = set()
    mrbts_order = []   # insertion-order list; sorted numerically afterward

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
            mrbts_order.append(mrbts_key)

        all_counts[mrbts_key][inv_type] += 1
        fam_counter[inv_type][fam_type] += 1

        if _is_working(state):
            working_counts[mrbts_key][inv_type] += 1

    # ── Determine column order for inventoryUnitType ─────────────────────────
    # Representative family = most common vendorUnitFamilyType for each inv type
    fam_map = {
        inv_type: max(fam_counts, key=fam_counts.get)
        for inv_type, fam_counts in fam_counter.items()
    }

    all_inv_types = sorted(
        fam_counter.keys(),
        key=lambda t: (_group_key(fam_map[t]), t.upper())
    )

    # Sort MRBTS numerically (fall back to string sort for non-numeric IDs)
    all_mrbts = sorted(seen_mrbts, key=lambda x: (int(x) if x.isdigit() else x))

    # ── Write workbook ───────────────────────────────────────────────────────
    wb = xlsxwriter.Workbook(output_path)
    hdr_fmt, str_fmt, num_fmt, zero_fmt, int_fmt = _make_formats(wb)

    def _write_site_sheet(ws_name, count_dict):
        ws = wb.add_worksheet(ws_name)
        ws.freeze_panes(1, 2)
        ws.set_zoom(85)
        ws.set_default_row(15)

        headers = ['MRBTS', 'Site Name'] + all_inv_types
        for c, h in enumerate(headers):
            ws.write(0, c, h, hdr_fmt)
        ws.set_row(0, 45)

        ws.set_column(0, 0, 10)                     # MRBTS id
        ws.set_column(1, 1, 32)                     # Site Name
        ws.set_column(2, len(headers) - 1, 9)       # count columns

        for r, mrbts in enumerate(all_mrbts, start=1):
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

    for c, h in enumerate(['inventoryUnitType', 'Working', 'Total']):
        ws_ov.write(0, c, h, hdr_fmt)
    ws_ov.set_row(0, 20)
    ws_ov.set_column(0, 0, 34)
    ws_ov.set_column(1, 2, 14)

    for r, inv_type in enumerate(all_inv_types, start=1):
        total_w = sum(working_counts[m].get(inv_type, 0) for m in all_mrbts)
        total_a = sum(all_counts[m].get(inv_type, 0)     for m in all_mrbts)
        ws_ov.write(r, 0, inv_type, str_fmt)
        ws_ov.write(r, 1, total_w,  num_fmt)
        ws_ov.write(r, 2, total_a,  num_fmt)

    wb.close()
    return len(all_mrbts)
