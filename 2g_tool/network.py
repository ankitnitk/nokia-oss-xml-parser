"""
network.py
Builds the 2G network object model from raw sheet records.
Hierarchy: Network -> BSC -> BCF -> BTS -> Segment
"""

import math
from collections import Counter


# ---------------------------------------------------------------------------
# DN helpers
# ---------------------------------------------------------------------------

def extract_id(dn, level):
    for part in str(dn).split('/'):
        if part.upper().startswith(level.upper() + '-'):
            return part[len(level) + 1:]
    return ''


def parent_dn(dn, level):
    parts = str(dn).split('/')
    result = []
    for part in parts:
        result.append(part)
        if part.upper().startswith(level.upper() + '-'):
            break
    return '/'.join(result)


def get(record, *keys, default=''):
    """
    Flexible getter: tries each key in order, returns first non-empty value.
    Converts whole-number floats to int string automatically.
    """
    for key in keys:
        val = record.get(key, '')
        if val is None:
            val = ''
        val_str = str(val).strip()
        if val_str and val_str not in ('nan', 'None', ''):
            if val_str.endswith('.0'):
                try:
                    return str(int(float(val_str)))
                except Exception:
                    pass
            return val_str
    return default


def to_num(val, default=0):
    """Convert a value to float, return default on failure."""
    try:
        return float(str(val).strip())
    except (ValueError, TypeError):
        return default


# ---------------------------------------------------------------------------
# Network model
# ---------------------------------------------------------------------------

class Network:

    # Per-sheet: ordered list of (DN_label, column_name) pairs that build Dist_Name.
    # e.g. BCF row has columns 'BSC' and 'BCF'; TRX row has 'BSC','BCF','BTS','TRX'.
    _DN_LEVELS = {
        'BSC':  [('BSC',  'BSC')],
        'BCF':  [('BSC',  'BSC'),  ('BCF',  'BCF')],
        'BTS':  [('BSC',  'BSC'),  ('BCF',  'BCF'),  ('BTS',  'BTS')],
        'TRX':  [('BSC',  'BSC'),  ('BCF',  'BCF'),  ('BTS',  'BTS'),  ('TRX',  'TRX')],
        'HOC':  [('BSC',  'BSC'),  ('BCF',  'BCF'),  ('BTS',  'BTS'),  ('HOC',  'HOC')],
        'POC':  [('BSC',  'BSC'),  ('BCF',  'BCF'),  ('BTS',  'BTS'),  ('POC',  'POC')],
        'ADCE': [('BSC',  'BSC'),  ('BCF',  'BCF'),  ('BTS',  'BTS'),  ('ADCE', 'ADCE')],
        'ADJW': [('BSC',  'BSC'),  ('BCF',  'BCF'),  ('BTS',  'BTS'),  ('ADJW', 'ADJW')],
        'ADJL': [('BSC',  'BSC'),  ('BCF',  'BCF'),  ('BTS',  'BTS'),  ('ADJL', 'ADJL')],
        'MAL':  [('BSC',  'BSC'),  ('MAL',  'MAL')],
    }

    @staticmethod
    def _fill_dist_names(records, levels):
        """
        For any record that is missing Dist_Name (or has it empty/null),
        synthesise it from the individual ID columns defined by `levels`.
        Operates in-place — no records are dropped.
        """
        for r in records:
            if str(r.get('Dist_Name', '')).strip() in ('', 'nan', 'None'):
                parts = ['PLMN-PLMN']
                ok = True
                for label, col in levels:
                    val = str(r.get(col, '')).strip()
                    if val.endswith('.0'):          # whole-float → int string
                        try:
                            val = str(int(float(val)))
                        except Exception:
                            pass
                    if not val or val in ('nan', 'None'):
                        ok = False
                        break
                    parts.append(f'{label}-{val}')
                if ok:
                    r['Dist_Name'] = '/'.join(parts)

    def __init__(self, sheets):
        self.sheets = sheets

        # --- Ensure every sheet has Dist_Name, synthesising from ID columns if absent ---
        for sheet_name, levels in self._DN_LEVELS.items():
            self._fill_dist_names(sheets.get(sheet_name, []), levels)

        self.bsc_by_dn   = self._index(sheets.get('BSC', []))
        self.bcf_by_dn   = self._index(sheets.get('BCF', []))
        self.bts_records = sheets.get('BTS', [])
        self.mal_by_dn   = self._index(sheets.get('MAL', []))

        # TRX: index by parent BTS DN -> list of TRX records
        self.trx_by_bts_dn = {}
        for r in sheets.get('TRX', []):
            dn = get(r, 'Dist_Name')
            if '/TRX-' in dn:
                bts_dn = dn[:dn.rindex('/TRX-')]
                self.trx_by_bts_dn.setdefault(bts_dn, []).append(r)

        # Adjacency sheets: index by parent BTS DN -> list of records
        self.adce_by_bts_dn = self._build_adj_index(sheets.get('ADCE', []), '/ADCE-')
        self.adjw_by_bts_dn = self._build_adj_index(sheets.get('ADJW', []), '/ADJW-')
        self.adjl_by_bts_dn = self._build_adj_index(sheets.get('ADJL', []), '/ADJL-')

        # One-way ADCE detection using CI+LAC matching (not targetCellDN).
        # _adce_neighbors_by_bts: src_bts_dn -> set{(tgt_ci, tgt_lac)} defined by that BTS
        # _adce_by_src_bts:       src_bts_dn -> [(adce_r, tgt_ci, tgt_lac)]
        # _adce_records_flat:     flat list   of (adce_r, src_bts_dn, tgt_ci, tgt_lac)
        self._adce_neighbors_by_bts = {}
        self._adce_by_src_bts       = {}
        self._adce_records_flat     = []
        for adce_r in sheets.get('ADCE', []):
            src_full = get(adce_r, 'Dist_Name')
            pos = src_full.upper().find('/ADCE-')
            if pos == -1:
                continue
            src_bts_dn = src_full[:pos]
            tgt_ci  = get(adce_r, 'adjacentCellIdCI')
            tgt_lac = get(adce_r, 'adjacentCellIdLac')
            if not (src_bts_dn and tgt_ci and tgt_lac):
                continue
            self._adce_neighbors_by_bts.setdefault(src_bts_dn, set()).add((tgt_ci, tgt_lac))
            self._adce_by_src_bts.setdefault(src_bts_dn, []).append((adce_r, tgt_ci, tgt_lac))
            self._adce_records_flat.append((adce_r, src_bts_dn, tgt_ci, tgt_lac))

        # HOC / POC: first record per parent BTS DN
        self.hoc_by_bts_dn = self._build_first_child_index(sheets.get('HOC', []), '/HOC-')
        self.poc_by_bts_dn = self._build_first_child_index(sheets.get('POC', []), '/POC-')

        self.segments = self._build_segments()

        # Reverse lookups built after segments are ready
        self._bts_dn_to_seg    = {}   # bts_dn -> seg dict
        self._bts_dn_to_record = {}   # bts_dn -> BTS record (for CI/LAC of source BTS)
        for seg in self.segments.values():
            for bts_r in seg['all_bts']:
                bts_dn = get(bts_r, 'Dist_Name')
                self._bts_dn_to_seg[bts_dn]    = seg
                self._bts_dn_to_record[bts_dn] = bts_r

        # (cellId, lac) -> first matching segment — for resolving ADCE targets
        self._ci_lac_to_seg = {}
        for seg in self.segments.values():
            m = seg.get('master') or {}
            ci  = get(m, 'cellId')
            lac = get(m, 'locationAreaIdLAC')
            if ci and lac and (ci, lac) not in self._ci_lac_to_seg:
                self._ci_lac_to_seg[(ci, lac)] = seg

        # cellId-only index for discrepancy checks.
        # First occurrence wins; _ci_duplicates tracks CIs seen in >1 segment.
        self._ci_to_seg     = {}
        self._ci_duplicates = set()
        for seg in self.segments.values():
            m  = seg.get('master') or {}
            ci = get(m, 'cellId')
            if ci:
                if ci in self._ci_to_seg:
                    self._ci_duplicates.add(ci)
                else:
                    self._ci_to_seg[ci] = seg

        # --- Per-result caches (populated lazily on first access) ---
        self._channel_cache          = {}   # seg_dn  -> channel counts + tch_freqs_str
        self._bcch_cache             = {}   # bts_dn  -> initialFrequency of BCCH TRX
        self._power_cache            = {}   # bts_dn  -> max trxRfPower in W
        self._bands_cache            = {}   # seg_dn  -> bands string
        self._hop_cache              = {}   # seg_dn  -> (hopping_str, mal_str)
        self._discrepant_adce_cache  = {}   # bts_dn  -> int count

        # BCF frequency pool: built once per BCF.
        # bcf_dn -> Counter{freq_int: count} of effective frequencies in that BCF.
        # RF hopping BTS: frequencies come from MAL sheet (not TRX initialFrequency).
        # Non-RF BTS: frequencies come from TRX initialFrequency as before.
        # Used by TCH Freq highlighting in cell_summary.
        self.bcf_freq_pool = {}
        for bts_r in self.bts_records:
            bts_dn = get(bts_r, 'Dist_Name')
            if '/BTS-' not in bts_dn:
                continue
            bcf_dn = bts_dn[:bts_dn.rindex('/BTS-')]
            pool   = self.bcf_freq_pool.setdefault(bcf_dn, Counter())
            if get(bts_r, 'hoppingMode') == '2':      # RF hopping → MAL frequencies
                for freq in self._get_mal_freqs(bts_r):
                    if isinstance(freq, int):
                        pool[freq] += 1
            else:                                      # None / BB → TRX initialFrequency
                for trx_r in self.trx_by_bts_dn.get(bts_dn, []):
                    freq = get(trx_r, 'initialFrequency')
                    if freq:
                        try:
                            pool[int(freq)] += 1
                        except (ValueError, TypeError):
                            pass

    # -----------------------------------------------------------------------

    @staticmethod
    def _build_first_child_index(records, marker):
        """Index first record per parent BTS DN (for HOC, POC)."""
        marker_up = marker.upper()
        idx = {}
        for r in records:
            dn = get(r, 'Dist_Name')
            pos = dn.upper().find(marker_up)
            if pos != -1:
                bts_dn = dn[:pos]
                if bts_dn not in idx:
                    idx[bts_dn] = r
        return idx

    @staticmethod
    def _build_adj_index(records, marker):
        marker_up = marker.upper()
        idx = {}
        for r in records:
            dn = get(r, 'Dist_Name')
            pos = dn.upper().find(marker_up)
            if pos != -1:
                bts_dn = dn[:pos]
                idx.setdefault(bts_dn, []).append(r)
        return idx

    @staticmethod
    def _index(records):
        idx = {}
        for r in records:
            dn = get(r, 'Dist_Name')
            if dn and dn not in idx:
                idx[dn] = r
        return idx

    def _build_segments(self):
        groups = {}
        for r in self.bts_records:
            dn     = get(r, 'Dist_Name')
            bsc_id = extract_id(dn, 'BSC')
            seg_id = get(r, 'segmentId')
            if not bsc_id or not seg_id:
                continue
            seg_dn = f'PLMN-PLMN/BSC-{bsc_id}/SEG-{seg_id}'
            groups.setdefault(seg_dn, []).append(r)

        segments = {}
        for seg_dn, bts_list in groups.items():
            master = None
            slaves = []
            for r in bts_list:
                if get(r, 'masterBcf') == '1' or get(r, 'primaryBcf') == '1':
                    master = r
                else:
                    slaves.append(r)
            if master is None and bts_list:
                master = bts_list[0]
                slaves = bts_list[1:]

            # Cache BCF DN once — avoids recomputing per column during report build
            bcf_dns = set()
            for r in bts_list:
                dn = get(r, 'Dist_Name')
                if '/BTS-' in dn:
                    bcf_dns.add(dn[:dn.rindex('/BTS-')])
            bcf_dn = bcf_dns.pop() if len(bcf_dns) == 1 else ''

            segments[seg_dn] = {
                'seg_dn':  seg_dn,
                'master':  master,
                'slaves':  slaves,
                'all_bts': bts_list,
                'bcf_dn':  bcf_dn,
            }
        return segments

    # -----------------------------------------------------------------------
    # Lookups
    # -----------------------------------------------------------------------

    def get_bsc(self, bsc_dn):
        return self.bsc_by_dn.get(bsc_dn, {})

    def get_bcf(self, bcf_dn):
        return self.bcf_by_dn.get(bcf_dn, {})

    def bcf_dn_for_segment(self, seg):
        return seg.get('bcf_dn', '')

    def band_label(self, freq_band_val):
        return {'0': '900', '1': '1800'}.get(str(freq_band_val).strip(), str(freq_band_val))

    def bands_for_segment(self, seg):
        seg_dn = seg['seg_dn']
        if seg_dn in self._bands_cache:
            return self._bands_cache[seg_dn]
        bands = set()
        for r in seg['all_bts']:
            b = self.band_label(get(r, 'frequencyBandInUse'))
            if b:
                bands.add(b)
        result = ' & '.join(sorted(bands))
        self._bands_cache[seg_dn] = result
        return result

    def trx_for_bts(self, bts_record):
        """Return list of TRX records for a given BTS record."""
        dn = get(bts_record, 'Dist_Name')
        return self.trx_by_bts_dn.get(dn, [])

    def bcch_freq(self, bts_record):
        """
        Return initialFrequency of the BCCH TRX (channel0Type == 4 or 7).
        Cached per BTS DN.
        """
        dn = get(bts_record, 'Dist_Name')
        if dn in self._bcch_cache:
            return self._bcch_cache[dn]
        result = ''
        for trx in self.trx_by_bts_dn.get(dn, []):
            if get(trx, 'channel0Type') in ('4', '7'):
                result = get(trx, 'initialFrequency')
                break
        self._bcch_cache[dn] = result
        return result

    def trx_count(self, bts_record):
        return len(self.trx_for_bts(bts_record))

    def tch_freqs(self, seg):
        """
        Comma-separated sorted initialFrequency values for all non-BCCH TRXs
        in the segment. Computed as part of channel_type_counts and cached.
        """
        seg_dn = seg['seg_dn']
        if seg_dn not in self._channel_cache:
            self.channel_type_counts(seg)
        return self._channel_cache[seg_dn]['tch_freqs_str']

    def adce_for_bts(self, bts_record):
        return self.adce_by_bts_dn.get(get(bts_record, 'Dist_Name'), [])

    def adjw_for_bts(self, bts_record):
        return self.adjw_by_bts_dn.get(get(bts_record, 'Dist_Name'), [])

    def adjl_for_bts(self, bts_record):
        return self.adjl_by_bts_dn.get(get(bts_record, 'Dist_Name'), [])

    def hoc_for_bts(self, bts_record):
        return self.hoc_by_bts_dn.get(get(bts_record, 'Dist_Name'), {})

    def poc_for_bts(self, bts_record):
        return self.poc_by_bts_dn.get(get(bts_record, 'Dist_Name'), {})

    def lar(self, bts_record):
        """LAR = -110 + nonBcchLayerAccessThr from HOC sheet."""
        val = get(self.hoc_for_bts(bts_record), 'nonBcchLayerAccessThr')
        if not val:
            return ''
        try:
            return -110 + int(float(val))
        except (ValueError, TypeError):
            return ''

    def ler(self, bts_record):
        """LER = -110 + rxLevel from HOC sheet."""
        val = get(self.hoc_for_bts(bts_record), 'rxLevel')
        if not val:
            return ''
        try:
            return -110 + int(float(val))
        except (ValueError, TypeError):
            return ''

    def diff_lac_adce_count(self, bts_record, lac):
        return sum(1 for r in self.adce_for_bts(bts_record)
                   if get(r, 'adjacentCellIdLac') != lac)

    def oneway_adce_for_bts(self, bts_record):
        """
        ADCE records from this BTS that are one-way.
        A neighbour X->Y is one-way when:
          - Y cannot be found in the dump (CI+LAC match fails), OR
          - Y exists but does not define X (by CI+LAC) in its own ADCE list.
        """
        bts_dn  = get(bts_record, 'Dist_Name')
        src_ci  = get(bts_record, 'cellId')
        src_lac = get(bts_record, 'locationAreaIdLAC')
        results = []
        for adce_r, tgt_ci, tgt_lac in self._adce_by_src_bts.get(bts_dn, []):
            tgt_seg = self._ci_lac_to_seg.get((tgt_ci, tgt_lac))
            if tgt_seg is None:
                results.append(adce_r)   # target not in dump → one-way
            else:
                tgt_master_dn = get(tgt_seg.get('master') or {}, 'Dist_Name')
                if (src_ci, src_lac) not in self._adce_neighbors_by_bts.get(tgt_master_dn, set()):
                    results.append(adce_r)   # target exists but doesn't define source → one-way
        return results

    def all_oneway_adce_rows(self):
        """
        Generator yielding one dict per one-way ADCE entry across the whole network:
          src_seg_name, src_bts_dn, tgt_ci, tgt_lac, tgt_seg_name, remark
        """
        for adce_r, src_bts_dn, tgt_ci, tgt_lac in self._adce_records_flat:
            tgt_seg = self._ci_lac_to_seg.get((tgt_ci, tgt_lac))
            if tgt_seg is None:
                remark       = 'Target cell not found in dump'
                tgt_seg_name = ''
            else:
                tgt_master_dn = get(tgt_seg.get('master') or {}, 'Dist_Name')
                src_bts_r     = self._bts_dn_to_record.get(src_bts_dn, {})
                src_ci  = get(src_bts_r, 'cellId')
                src_lac = get(src_bts_r, 'locationAreaIdLAC')
                if (src_ci, src_lac) not in self._adce_neighbors_by_bts.get(tgt_master_dn, set()):
                    remark       = ''
                    tgt_seg_name = get(tgt_seg.get('master') or {}, 'segmentName')
                else:
                    continue   # bidirectional — skip

            src_seg      = self._bts_dn_to_seg.get(src_bts_dn)
            src_seg_name = get(src_seg.get('master') or {}, 'segmentName') if src_seg else ''
            yield {
                'src_seg_name': src_seg_name,
                'src_bts_dn':   src_bts_dn,
                'tgt_ci':       tgt_ci,
                'tgt_lac':      tgt_lac,
                'tgt_seg_name': tgt_seg_name,
                'remark':       remark,
            }

    # ADCE field mapping: (adce_field, actual_bts_field, display_key)
    # BCCH is handled separately via bcch_freq() so it has no actual_bts_field.
    _ADCE_CHECKS = [
        ('adjacentCellIdLac', 'locationAreaIdLAC', 'LAC'),
        ('adjCellBsicNcc',    'bsIdentityCodeNCC', 'NCC'),
        ('adjCellBsicBcc',    'bsIdentityCodeBCC', 'BCC'),
        ('adjacentCellIdMCC', 'locationAreaIdMCC', 'MCC'),
        ('adjacentCellIdMNC', 'locationAreaIdMNC', 'MNC'),
    ]

    def discrepant_adce_count(self, bts_record):
        """
        Count of ADCE entries for this BTS where the neighbour CI parameters
        don't match the actual cell data.  Discrepant conditions:
          • Target CI not found in the network
          • Target CI exists in multiple segments (repeated)
          • Any of LAC/NCC/BCC/MCC/MNC/BCCH differs from actual cell values
        Result is cached per BTS DN.
        """
        bts_dn = get(bts_record, 'Dist_Name')
        if bts_dn in self._discrepant_adce_cache:
            return self._discrepant_adce_cache[bts_dn]

        count = 0
        for adce_r, tgt_ci, tgt_lac in self._adce_by_src_bts.get(bts_dn, []):
            tgt_seg = self._ci_to_seg.get(tgt_ci)
            if tgt_seg is None or tgt_ci in self._ci_duplicates:
                count += 1
                continue
            tgt_m = tgt_seg.get('master') or {}
            mismatch = any(
                get(adce_r, af) != get(tgt_m, bf)
                for af, bf, _ in self._ADCE_CHECKS
            ) or (get(adce_r, 'bcchFrequency') != self.bcch_freq(tgt_m))
            if mismatch:
                count += 1

        self._discrepant_adce_cache[bts_dn] = count
        return count

    def all_discrepant_adce_rows(self):
        """
        Generator yielding one dict per discrepant ADCE entry across the network.
        An entry is discrepant when the target CI is not found, is duplicated, or
        any of the 6 parameter fields (LAC/NCC/BCC/MCC/MNC/BCCH) don't match.

        Each dict contains:
          src_bsc_name, src_seg_dn, src_cell_name, tgt_ci,
          adce_vals   : {LAC/NCC/BCC/MCC/MNC/BCCH -> adce-defined value}
          actual_vals : {LAC/NCC/BCC/MCC/MNC/BCCH -> actual cell value ('' if not found)}
          mismatched  : set of keys where adce_vals != actual_vals
          remark      : '' | 'Target Cell ID is repeated' | 'Target Cell ID not found in network'
        """
        for adce_r, src_bts_dn, tgt_ci, tgt_lac in self._adce_records_flat:
            # ── Source cell info ─────────────────────────────────────────────
            src_seg       = self._bts_dn_to_seg.get(src_bts_dn)
            src_master    = (src_seg.get('master') or {}) if src_seg else {}
            src_seg_dn    = src_seg['seg_dn'] if src_seg else ''
            src_bsc_dn    = 'PLMN-PLMN/BSC-' + extract_id(src_bts_dn, 'BSC')
            src_bsc_name  = get(self.bsc_by_dn.get(src_bsc_dn, {}), 'name')
            src_seg_name  = get(src_master, 'segmentName')
            src_cell_name = get(src_master, 'name')

            # ── ADCE-defined values ──────────────────────────────────────────
            adce_vals = {
                'LAC':  get(adce_r, 'adjacentCellIdLac'),
                'NCC':  get(adce_r, 'adjCellBsicNcc'),
                'BCC':  get(adce_r, 'adjCellBsicBcc'),
                'MCC':  get(adce_r, 'adjacentCellIdMCC'),
                'MNC':  get(adce_r, 'adjacentCellIdMNC'),
                'BCCH': get(adce_r, 'bcchFrequency'),
            }

            # ── Look up actual cell ──────────────────────────────────────────
            tgt_seg      = self._ci_to_seg.get(tgt_ci)
            is_duplicate = tgt_ci in self._ci_duplicates

            if tgt_seg is None:
                actual_vals = {k: '' for k in adce_vals}
                mismatched  = set()
                remark      = 'Target Cell ID not found in network'
            else:
                tgt_m = tgt_seg.get('master') or {}
                actual_vals = {
                    'LAC':  get(tgt_m, 'locationAreaIdLAC'),
                    'NCC':  get(tgt_m, 'bsIdentityCodeNCC'),
                    'BCC':  get(tgt_m, 'bsIdentityCodeBCC'),
                    'MCC':  get(tgt_m, 'locationAreaIdMCC'),
                    'MNC':  get(tgt_m, 'locationAreaIdMNC'),
                    'BCCH': self.bcch_freq(tgt_m),
                }
                mismatched = {k for k in adce_vals
                              if adce_vals[k] != actual_vals[k]}

                if is_duplicate:
                    remark = 'Target Cell ID is repeated'
                elif mismatched:
                    remark = ''
                else:
                    continue   # fully matching, non-duplicate → not discrepant

            yield {
                'src_bsc_name':  src_bsc_name,
                'src_seg_dn':    src_seg_dn,
                'src_seg_name':  src_seg_name,
                'src_cell_name': src_cell_name,
                'tgt_ci':        tgt_ci,
                'adce_vals':     adce_vals,
                'actual_vals':   actual_vals,
                'mismatched':    mismatched,
                'remark':        remark,
            }

    def seg_name_for_bts_dn(self, bts_dn):
        """Return segmentName for the segment that owns bts_dn, or ''."""
        seg = self._bts_dn_to_seg.get(bts_dn)
        if seg:
            return get(seg.get('master') or {}, 'segmentName')
        return ''

    def channel_type_counts(self, seg):
        """
        Count channel slot types + collect TCH frequencies across all TRX
        of all BTS in the segment — one pass, fully cached per segment.
        """
        seg_dn = seg['seg_dn']
        if seg_dn in self._channel_cache:
            return self._channel_cache[seg_dn]

        sdcch = ccch = not_used = tch = gtch_master = gtch_slave = 0
        ch_keys = [f'channel{i}Type' for i in range(8)]
        master_dn     = get(seg.get('master') or {}, 'Dist_Name')
        gtch_per_bts  = {}   # bts_dn -> (bts_record, gtch_count)
        tch_freq_list = []   # effective TCH frequencies for this segment

        for bts_r in seg['all_bts']:
            bts_dn    = get(bts_r, 'Dist_Name')
            bts_gprs  = get(bts_r, 'gprsEnabled') == '1'
            is_master = bts_dn == master_dn
            bts_gtch  = 0
            is_rf     = get(bts_r, 'hoppingMode') == '2'

            # TCH frequencies:
            #   RF hopping → from MAL sheet (usedMobileAllocation lookup)
            #   None / BB  → from non-BCCH TRX initialFrequency (original logic)
            if is_rf:
                tch_freq_list.extend(self._get_mal_freqs(bts_r))

            for trx_r in self.trx_by_bts_dn.get(bts_dn, []):
                is_bcch  = get(trx_r, 'channel0Type') in ('4', '7')
                trx_gprs = get(trx_r, 'gprsEnabledTrx') == '1'

                # Non-RF: collect TCH freq from TRX initialFrequency (skip BCCH TRX)
                if not is_rf and not is_bcch:
                    freq = get(trx_r, 'initialFrequency')
                    if freq:
                        try:
                            tch_freq_list.append(int(freq))
                        except (ValueError, TypeError):
                            tch_freq_list.append(freq)

                # Channel-type counts always come from TRX regardless of hopping mode
                for key in ch_keys:
                    ct = get(trx_r, key)
                    if not ct:
                        continue
                    if ct in ('3', '8'):
                        sdcch += 1
                    elif ct == '6':
                        ccch += 1
                    elif ct == '9':
                        not_used += 1
                    elif ct in ('0', '1', '2'):
                        tch += 1
                        if ct in ('0', '2') and trx_gprs and bts_gprs:
                            bts_gtch += 1
                            if is_master:
                                gtch_master += 1
                            else:
                                gtch_slave += 1
            gtch_per_bts[bts_dn] = (bts_r, bts_gtch)

        # CDED / CDEF: floor(capacity% / 100 * bts_gtch), summed per master/slave
        cded_master = cdef_master = cded_slave = cdef_slave = 0
        for bts_dn, (bts_r, gtch) in gtch_per_bts.items():
            dedicated_cap = to_num(get(bts_r, 'dedicatedGPRScapacity')) / 100
            default_cap   = to_num(get(bts_r, 'defaultGPRScapacity'))   / 100
            cded = math.floor(dedicated_cap * gtch)
            cdef = math.floor(default_cap   * gtch)
            if bts_dn == master_dn:
                cded_master = cded
                cdef_master = cdef
            else:
                cded_slave += cded
                cdef_slave += cdef

        # Build sorted TCH freq string
        tch_freq_list.sort(key=lambda f: (isinstance(f, str), f))
        tch_freqs_str = ', '.join(str(f) for f in tch_freq_list)

        result = {'sdcch': sdcch, 'ccch': ccch, 'not_used': not_used,
                  'tch': tch, 'gtch_master': gtch_master, 'gtch_slave': gtch_slave,
                  'cded_master': cded_master, 'cdef_master': cdef_master,
                  'cded_slave': cded_slave,   'cdef_slave': cdef_slave,
                  'tch_freqs_str': tch_freqs_str}
        self._channel_cache[seg_dn] = result
        return result

    @staticmethod
    def _parse_list_field(val):
        """
        Parse a list-field value as stored in the xlsx output:
          'List;91;95'  ->  [91, 95]
        Values that convert cleanly to int are returned as int, else str.
        Returns [] for empty / missing values.
        """
        s = str(val).strip() if val else ''
        if not s or s in ('nan', 'None'):
            return []
        parts = s.split(';')
        if parts and parts[0].strip().lower() == 'list':
            parts = parts[1:]
        result = []
        for p in parts:
            p = p.strip()
            if not p:
                continue
            try:
                result.append(int(float(p)))
            except (ValueError, TypeError):
                result.append(p)
        return result

    def _get_mal_freqs(self, bts_r):
        """
        Return list of MAL frequencies (ints) for an RF hopping BTS.
        Looks up the MAL record via usedMobileAllocation + BSC ID,
        then parses the 'frequency' list field.
        """
        mal_id = get(bts_r, 'usedMobileAllocation')
        if not mal_id:
            return []
        bsc_id = extract_id(get(bts_r, 'Dist_Name'), 'BSC')
        if not bsc_id:
            return []
        mal_dn = f'PLMN-PLMN/BSC-{bsc_id}/MAL-{mal_id}'
        mal_r  = self.mal_by_dn.get(mal_dn, {})
        return self._parse_list_field(get(mal_r, 'frequency'))

    # Hopping mode translation table
    _HOP_MAP = {'0': 'None', '1': 'BB', '2': 'RF'}

    def hopping_mode_and_mal(self, seg):
        """Cached — returns same tuple on repeated calls for the same segment."""
        seg_dn = seg['seg_dn']
        if seg_dn in self._hop_cache:
            return self._hop_cache[seg_dn]
        result = self._compute_hopping(seg)
        self._hop_cache[seg_dn] = result
        return result

    def _compute_hopping(self, seg):
        """
        Returns (hopping_str, mal_str) for the segment.

        hopping_str : master first, then slaves, joined by ' & '.
                      hoppingMode values translated: 0→None, 1→BB, 2→RF.
        mal_str     : usedMobileAllocation for each BTS (master & slaves),
                      joined by ' & ', but only if at least one BTS uses RF
                      hopping.  Empty string when no BTS is RF.
        """
        bts_list = [b for b in ([seg.get('master')] + list(seg.get('slaves', [])))
                    if b is not None]

        hop_parts = []
        mal_parts = []
        any_rf    = False

        for bts_r in bts_list:
            raw = get(bts_r, 'hoppingMode')
            hop_parts.append(self._HOP_MAP.get(raw, raw))
            if raw == '2':
                mal_parts.append(get(bts_r, 'usedMobileAllocation'))
                any_rf = True
            else:
                mal_parts.append('NA')   # placeholder; only shown if any BTS is RF

        hop_str = ' & '.join(hop_parts)
        mal_str = ' & '.join(mal_parts) if any_rf else ''
        return hop_str, mal_str

    def max_trx_power_w(self, bts_record):
        """Max trxRfPower across all TRXs for a BTS, converted mW -> W. Cached per BTS DN."""
        dn = get(bts_record, 'Dist_Name')
        if dn in self._power_cache:
            return self._power_cache[dn]
        powers = [to_num(get(trx, 'trxRfPower'))
                  for trx in self.trx_by_bts_dn.get(dn, [])]
        if not powers:
            result = ''
        else:
            val = max(powers) / 1000
            result = int(val) if val == int(val) else round(val, 2)
        self._power_cache[dn] = result
        return result
