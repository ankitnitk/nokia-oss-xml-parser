"""
network.py
5G NR network model. Reads NRBTS and NRCELL sheets, fills missing
Dist_Names, and builds indexed lookups used by report builders.

Hierarchy:
    MRBTS
    └── NRBTS  (gNB)
        └── NRCELL  (cell)

Band for an NRBTS is derived from its cells' `nrarfcn` (NRCELL) -- there is
no band field on NRBTS itself.
"""

from collections import defaultdict


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get(rec, *keys, default=''):
    """Flexible getter: tries each key in order, returns first non-empty value.
    Trailing '.0' is stripped -- calamine reads Nokia integer columns as floats,
    so MRBTS 901001 arrives as '901001.0'."""
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

def _nrbts_dn(mrbts, nrbts):
    return f'PLMN-PLMN/MRBTS-{mrbts}/NRBTS-{nrbts}'


def _nrcell_dn(mrbts, nrbts, nrcell):
    return f'PLMN-PLMN/MRBTS-{mrbts}/NRBTS-{nrbts}/NRCELL-{nrcell}'


# ---------------------------------------------------------------------------
# Network model
# ---------------------------------------------------------------------------

class Network:
    NEEDED_SHEETS = ['NRBTS', 'NRCELL']

    def __init__(self, sheets):
        for sn in self.NEEDED_SHEETS:
            if sn in sheets:
                _fill_dist_names(sn, sheets[sn])

        # ---- NRBTS --------------------------------------------------------
        self.nrbts_list = sheets.get('NRBTS', [])
        self.nrbts_by_dn = {}
        for r in self.nrbts_list:
            dn = get(r, 'Dist_Name')
            if dn and dn not in self.nrbts_by_dn:
                self.nrbts_by_dn[dn] = r

        # ---- NRCELL -------------------------------------------------------
        self.nrcell_list = sheets.get('NRCELL', [])
        self.nrcell_by_dn = {}
        self.nrcell_list_by_nrbts_dn = defaultdict(list)
        for r in self.nrcell_list:
            dn = get(r, 'Dist_Name')
            if dn and dn not in self.nrcell_by_dn:
                self.nrcell_by_dn[dn] = r
            k = _key_nrbts(r)
            if k:
                self.nrcell_list_by_nrbts_dn[k].append(r)

    # -----------------------------------------------------------------------
    # Convenience accessors
    # -----------------------------------------------------------------------

    def nrcell_for_nrbts(self, nrbts_dn):
        return self.nrcell_list_by_nrbts_dn.get(nrbts_dn, [])

    def nrarfcns_for_nrbts(self, nrbts_dn):
        """Sorted unique NRARFCN integers across the NRBTS's cells."""
        seen = set()
        for r in self.nrcell_for_nrbts(nrbts_dn):
            v = to_num(get(r, 'nrarfcn'), default=None)
            if v:
                seen.add(int(v))
        return sorted(seen)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _key_nrbts(rec):
    """NRBTS DN for indexing a cell-level record, or '' if IDs are missing."""
    mrbts = get(rec, 'MRBTS')
    nrbts = get(rec, 'NRBTS')
    if not mrbts or not nrbts:
        return ''
    return _nrbts_dn(mrbts, nrbts)


_DN_PATTERNS = {
    'NRBTS':  lambda r: f"PLMN-PLMN/MRBTS-{get(r, 'MRBTS')}/NRBTS-{get(r, 'NRBTS')}",
    'NRCELL': lambda r: (f"PLMN-PLMN/MRBTS-{get(r, 'MRBTS')}/NRBTS-{get(r, 'NRBTS')}"
                         f"/NRCELL-{get(r, 'NRCELL')}"),
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
        if not get(rec, 'MRBTS') or not get(rec, 'NRBTS'):
            continue
        if sheet_name == 'NRCELL' and not get(rec, 'NRCELL'):
            continue
        rec['Dist_Name'] = pattern(rec)
