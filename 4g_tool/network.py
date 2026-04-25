"""
network.py
4G LTE network model. Reads all sheets, fills missing Dist_Names,
builds indexes used by report builders.
"""

from collections import defaultdict


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get(rec, col, default=''):
    """Safe getter: returns string value, '' for None/NaN/missing.
    Strips leading apostrophe (Excel text-forced numbers like '03)."""
    val = rec.get(col, default)
    if val is None:
        return default
    s = str(val).strip()
    if s.lower() in ('none', 'nan', ''):
        return default
    if s.startswith("'"):
        s = s[1:]
    return s


def to_num(val, default=0):
    """Safe conversion to int (preferred) or float. Returns default on failure."""
    if val == '' or val is None:
        return default
    try:
        f = float(val)
        return int(f) if f == int(f) else f
    except (ValueError, TypeError):
        return default


# ---------------------------------------------------------------------------
# DN builders  (match the format that appears in the dump)
# ---------------------------------------------------------------------------

def _lnbts_dn(mrbts, lnbts):
    return f'PLMN-PLMN/MRBTS-{mrbts}/LNBTS-{lnbts}'


def _lncel_dn(mrbts, lnbts, lncel):
    return f'PLMN-PLMN/MRBTS-{mrbts}/LNBTS-{lnbts}/LNCEL-{lncel}'


# ---------------------------------------------------------------------------
# Network model
# ---------------------------------------------------------------------------

class Network:
    """
    Holds all 4G LTE records and provides indexed lookups.

    Hierarchy (equivalent 2G object in brackets):
        MRBTS  [BSC]
        └── LNBTS  [BCF]
            ├── LNBTS_FDD-0   (FDD-specific LNBTS params)
            ├── LNBTS_TDD-0   (TDD-specific LNBTS params)
            └── LNCEL-{n}  [Segment/BTS]
                ├── LNCEL_FDD-0  (FDD cell params: earfcnDL, earfcnUL, dlChBw…)
                ├── LNCEL_TDD-0  (TDD cell params: earfcn, chBw…)
                ├── IRFIM-{n}    (inter-freq measurement config)
                └── LNHOIF-{n}   (inter-freq handover config)
    """

    NEEDED_SHEETS = [
        'LNBTS', 'LNBTS_FDD', 'LNBTS_TDD',
        'LNCEL', 'LNCEL_FDD', 'LNCEL_TDD',
        'IRFIM', 'LNHOIF', 'SIB', 'REDRT', 'CAPR',
    ]

    def __init__(self, sheets):
        # Fill any missing Dist_Names first
        for sn in self.NEEDED_SHEETS:
            if sn in sheets:
                _fill_dist_names(sn, sheets[sn])

        # ---- LNBTS --------------------------------------------------------
        self.lnbts_list = sheets.get('LNBTS', [])
        self.lnbts_by_dn = {}
        for r in self.lnbts_list:
            dn = r.get('Dist_Name', '')
            if dn and dn not in self.lnbts_by_dn:
                self.lnbts_by_dn[dn] = r

        # ---- LNBTS_FDD / LNBTS_TDD ---------------------------------------
        self.lnbts_fdd_by_lnbts_dn = {}
        for r in sheets.get('LNBTS_FDD', []):
            k = _key_lnbts(r)
            if k and k not in self.lnbts_fdd_by_lnbts_dn:
                self.lnbts_fdd_by_lnbts_dn[k] = r

        self.lnbts_tdd_by_lnbts_dn = {}
        for r in sheets.get('LNBTS_TDD', []):
            k = _key_lnbts(r)
            if k and k not in self.lnbts_tdd_by_lnbts_dn:
                self.lnbts_tdd_by_lnbts_dn[k] = r

        # ---- LNCEL --------------------------------------------------------
        self.lncel_by_dn = {}
        self.lncel_list_by_lnbts_dn = defaultdict(list)
        for r in sheets.get('LNCEL', []):
            dn = r.get('Dist_Name', '')
            if dn and dn not in self.lncel_by_dn:
                self.lncel_by_dn[dn] = r
            k = _key_lnbts(r)
            if k:
                self.lncel_list_by_lnbts_dn[k].append(r)

        # ---- LNCEL_FDD ----------------------------------------------------
        self.lncel_fdd_by_lncel_dn = {}
        self.lncel_fdd_list_by_lnbts_dn = defaultdict(list)
        for r in sheets.get('LNCEL_FDD', []):
            lnbts_k = _key_lnbts(r)
            lncel_k = _key_lncel(r)
            if lnbts_k:
                self.lncel_fdd_list_by_lnbts_dn[lnbts_k].append(r)
            if lncel_k and lncel_k not in self.lncel_fdd_by_lncel_dn:
                self.lncel_fdd_by_lncel_dn[lncel_k] = r

        # ---- LNCEL_TDD ----------------------------------------------------
        self.lncel_tdd_by_lncel_dn = {}
        self.lncel_tdd_list_by_lnbts_dn = defaultdict(list)
        for r in sheets.get('LNCEL_TDD', []):
            lnbts_k = _key_lnbts(r)
            lncel_k = _key_lncel(r)
            if lnbts_k:
                self.lncel_tdd_list_by_lnbts_dn[lnbts_k].append(r)
            if lncel_k and lncel_k not in self.lncel_tdd_by_lncel_dn:
                self.lncel_tdd_by_lncel_dn[lncel_k] = r

        # ---- LNHOIF -------------------------------------------------------
        self.lnhoif_list_by_lncel_dn = defaultdict(list)
        for r in sheets.get('LNHOIF', []):
            k = _key_lncel(r)
            if k:
                self.lnhoif_list_by_lncel_dn[k].append(r)

        # ---- IRFIM --------------------------------------------------------
        self.irfim_list_by_lncel_dn = defaultdict(list)
        for r in sheets.get('IRFIM', []):
            k = _key_lncel(r)
            if k:
                self.irfim_list_by_lncel_dn[k].append(r)

        # ---- SIB ----------------------------------------------------------
        self.sib_by_lncel_dn = {}
        for r in sheets.get('SIB', []):
            k = _key_lncel(r)
            if k and k not in self.sib_by_lncel_dn:
                self.sib_by_lncel_dn[k] = r

        # ---- REDRT --------------------------------------------------------
        self.redrt_list_by_lncel_dn = defaultdict(list)
        for r in sheets.get('REDRT', []):
            k = _key_lncel(r)
            if k:
                self.redrt_list_by_lncel_dn[k].append(r)

        # ---- CAPR --------------------------------------------------------
        self.capr_list_by_lncel_dn = defaultdict(list)
        for r in sheets.get('CAPR', []):
            k = _key_lncel(r)
            if k:
                self.capr_list_by_lncel_dn[k].append(r)

    # -----------------------------------------------------------------------
    # Convenience accessors
    # -----------------------------------------------------------------------

    def fdd_cells_for_lnbts(self, lnbts_dn):
        return self.lncel_fdd_list_by_lnbts_dn.get(lnbts_dn, [])

    def tdd_cells_for_lnbts(self, lnbts_dn):
        return self.lncel_tdd_list_by_lnbts_dn.get(lnbts_dn, [])

    def lncel_for_lnbts(self, lnbts_dn):
        return self.lncel_list_by_lnbts_dn.get(lnbts_dn, [])

    def lnhoif_count(self, lncel_dn):
        return len(self.lnhoif_list_by_lncel_dn.get(lncel_dn, []))

    def irfim_count(self, lncel_dn):
        return len(self.irfim_list_by_lncel_dn.get(lncel_dn, []))

    def lte_mode(self, lnbts_dn):
        has_fdd = bool(self.lncel_fdd_list_by_lnbts_dn.get(lnbts_dn))
        has_tdd = bool(self.lncel_tdd_list_by_lnbts_dn.get(lnbts_dn))
        if has_fdd and has_tdd:
            return 'FDD+TDD'
        if has_fdd:
            return 'FDD'
        if has_tdd:
            return 'TDD'
        return ''

    def earfcns_for_lnbts(self, lnbts_dn):
        """Returns sorted list of unique EARFCN integers for an LNBTS."""
        seen = set()
        for r in self.lncel_fdd_list_by_lnbts_dn.get(lnbts_dn, []):
            v = to_num(get(r, 'earfcnDL'))
            if v:
                seen.add(int(v))
        for r in self.lncel_tdd_list_by_lnbts_dn.get(lnbts_dn, []):
            v = to_num(get(r, 'earfcn'))
            if v:
                seen.add(int(v))
        return sorted(seen)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _key_lnbts(rec):
    """Returns the LNBTS DN string for indexing, or '' if IDs are missing."""
    mrbts = get(rec, 'MRBTS')
    lnbts = get(rec, 'LNBTS')
    if not mrbts or not lnbts:
        return ''
    return _lnbts_dn(mrbts, lnbts)


def _key_lncel(rec):
    """Returns the LNCEL DN string for indexing, or '' if IDs are missing."""
    mrbts = get(rec, 'MRBTS')
    lnbts = get(rec, 'LNBTS')
    lncel = get(rec, 'LNCEL')
    if not mrbts or not lnbts or not lncel:
        return ''
    return _lncel_dn(mrbts, lnbts, lncel)


_DN_PATTERNS = {
    'LNBTS':     lambda r: f"PLMN-PLMN/MRBTS-{get(r,'MRBTS')}/LNBTS-{get(r,'LNBTS')}",
    'LNBTS_FDD': lambda r: f"PLMN-PLMN/MRBTS-{get(r,'MRBTS')}/LNBTS-{get(r,'LNBTS')}/LNBTS_FDD-{get(r,'LNBTS_FDD')}",
    'LNBTS_TDD': lambda r: f"PLMN-PLMN/MRBTS-{get(r,'MRBTS')}/LNBTS-{get(r,'LNBTS')}/LNBTS_TDD-{get(r,'LNBTS_TDD')}",
    'LNCEL':     lambda r: f"PLMN-PLMN/MRBTS-{get(r,'MRBTS')}/LNBTS-{get(r,'LNBTS')}/LNCEL-{get(r,'LNCEL')}",
    'LNCEL_FDD': lambda r: f"PLMN-PLMN/MRBTS-{get(r,'MRBTS')}/LNBTS-{get(r,'LNBTS')}/LNCEL-{get(r,'LNCEL')}/LNCEL_FDD-{get(r,'LNCEL_FDD')}",
    'LNCEL_TDD': lambda r: f"PLMN-PLMN/MRBTS-{get(r,'MRBTS')}/LNBTS-{get(r,'LNBTS')}/LNCEL-{get(r,'LNCEL')}/LNCEL_TDD-{get(r,'LNCEL_TDD')}",
    'IRFIM':     lambda r: f"PLMN-PLMN/MRBTS-{get(r,'MRBTS')}/LNBTS-{get(r,'LNBTS')}/LNCEL-{get(r,'LNCEL')}/IRFIM-{get(r,'IRFIM')}",
    'LNHOIF':    lambda r: f"PLMN-PLMN/MRBTS-{get(r,'MRBTS')}/LNBTS-{get(r,'LNBTS')}/LNCEL-{get(r,'LNCEL')}/LNHOIF-{get(r,'LNHOIF')}",
    'SIB':       lambda r: f"PLMN-PLMN/MRBTS-{get(r,'MRBTS')}/LNBTS-{get(r,'LNBTS')}/LNCEL-{get(r,'LNCEL')}/SIB-{get(r,'SIB')}",
    'REDRT':     lambda r: f"PLMN-PLMN/MRBTS-{get(r,'MRBTS')}/LNBTS-{get(r,'LNBTS')}/LNCEL-{get(r,'LNCEL')}/REDRT-{get(r,'REDRT')}",
    'CAPR':      lambda r: f"PLMN-PLMN/MRBTS-{get(r,'MRBTS')}/LNBTS-{get(r,'LNBTS')}/LNCEL-{get(r,'LNCEL')}/CAPR-{get(r,'CAPR')}",
}


def _fill_dist_names(sheet_name, records):
    """Synthesise Dist_Name for rows where it is missing or blank."""
    pattern = _DN_PATTERNS.get(sheet_name)
    if pattern is None:
        return
    for rec in records:
        if rec.get('Dist_Name'):
            continue
        mrbts = get(rec, 'MRBTS')
        lnbts = get(rec, 'LNBTS')
        if not mrbts or not lnbts:
            continue
        # For LNCEL-level sheets, also need LNCEL id
        if sheet_name not in ('LNBTS', 'LNBTS_FDD', 'LNBTS_TDD'):
            if not get(rec, 'LNCEL'):
                continue
        rec['Dist_Name'] = pattern(rec)
