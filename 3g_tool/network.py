"""
network.py
3G WCDMA network model. Reads RNC, WBTS, WCEL and WNCEL sheets, fills
missing Dist_Names, and builds indexed lookups used by report builders.

Hierarchy:
    RNC
    └── WBTS  (Node B)
        └── WCEL  (cell)

WNCEL lives under the MRBTS hierarchy (multiradio) and is joined to WCEL
by: WNCEL.WNCEL == WCEL.WCEL  AND  WNCEL.MRBTS == WBTS.SBTSId
"""

from collections import defaultdict


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get(rec, *keys, default=''):
    """Flexible getter: tries each key in order, returns first non-empty value."""
    for key in keys:
        val = rec.get(key, '')
        if val is None:
            val = ''
        s = str(val).strip()
        if s and s.lower() not in ('nan', 'none', ''):
            if s.endswith('.0'):
                try:
                    return str(int(float(s)))
                except Exception:
                    pass
            return s
    return default


def to_num(val, default=0):
    """Safe conversion to int (preferred) or float. Returns default on failure."""
    if val == '' or val is None:
        return default
    try:
        f = float(str(val).strip())
        return int(f) if f == int(f) else f
    except (ValueError, TypeError):
        return default


# ---------------------------------------------------------------------------
# DN builders
# ---------------------------------------------------------------------------

def _rnc_dn(rnc):
    return f'PLMN-PLMN/RNC-{rnc}'

def _wbts_dn(rnc, wbts):
    return f'PLMN-PLMN/RNC-{rnc}/WBTS-{wbts}'

def _wcel_dn(rnc, wbts, wcel):
    return f'PLMN-PLMN/RNC-{rnc}/WBTS-{wbts}/WCEL-{wcel}'


# ---------------------------------------------------------------------------
# Network model
# ---------------------------------------------------------------------------

class Network:
    NEEDED_SHEETS = ['RNC', 'WBTS', 'WCEL', 'WNCEL']

    def __init__(self, sheets):
        for sn in self.NEEDED_SHEETS:
            if sn in sheets:
                _fill_dist_names(sn, sheets[sn])

        # RNC
        self.rnc_by_dn = {}
        for r in sheets.get('RNC', []):
            dn = r.get('Dist_Name', '')
            if dn and dn not in self.rnc_by_dn:
                self.rnc_by_dn[dn] = r

        # WBTS
        self.wbts_by_dn = {}
        for r in sheets.get('WBTS', []):
            dn = r.get('Dist_Name', '')
            if dn and dn not in self.wbts_by_dn:
                self.wbts_by_dn[dn] = r

        # WCEL
        self.wcel_list = sheets.get('WCEL', [])
        self.wcel_by_dn = {}
        for r in self.wcel_list:
            dn = r.get('Dist_Name', '')
            if dn and dn not in self.wcel_by_dn:
                self.wcel_by_dn[dn] = r

        # WNCEL — keyed by (MRBTS, WNCEL) where MRBTS=SBTSId and WNCEL=WCEL ID
        self.wncel_by_mrbts_wcel = {}
        for r in sheets.get('WNCEL', []):
            mrbts = get(r, 'MRBTS')
            wncel = get(r, 'WNCEL')
            if not (mrbts and wncel):
                continue
            key = (mrbts, wncel)
            if key not in self.wncel_by_mrbts_wcel:
                self.wncel_by_mrbts_wcel[key] = r


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

_DN_PATTERNS = {
    'RNC':   lambda r: f"PLMN-PLMN/RNC-{get(r, 'RNC')}",
    'WBTS':  lambda r: f"PLMN-PLMN/RNC-{get(r, 'RNC')}/WBTS-{get(r, 'WBTS')}",
    'WCEL':  lambda r: f"PLMN-PLMN/RNC-{get(r, 'RNC')}/WBTS-{get(r, 'WBTS')}/WCEL-{get(r, 'WCEL')}",
    'WNCEL': lambda r: f"PLMN-PLMN/MRBTS-{get(r, 'MRBTS')}/WNCEL-{get(r, 'WNCEL')}",
}


def _fill_dist_names(sheet_name, records):
    """Synthesise Dist_Name for rows where it is missing or blank."""
    pattern = _DN_PATTERNS.get(sheet_name)
    if pattern is None:
        return
    for rec in records:
        dn = str(rec.get('Dist_Name', '')).strip()
        if dn and dn.lower() not in ('nan', 'none', ''):
            continue
        if sheet_name == 'WNCEL':
            if not get(rec, 'MRBTS') or not get(rec, 'WNCEL'):
                continue
        else:
            rnc = get(rec, 'RNC')
            if not rnc:
                continue
            if sheet_name in ('WBTS', 'WCEL') and not get(rec, 'WBTS'):
                continue
            if sheet_name == 'WCEL' and not get(rec, 'WCEL'):
                continue
        rec['Dist_Name'] = pattern(rec)
