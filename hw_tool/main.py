"""
hw_tool/main.py  —  Standalone HW Inventory Report

Reads INVUNIT + MRBTS/LNBTS sheets from a parsed OSS XML dump (xlsx or xlsb)
and writes a 3-sheet HW inventory report:

  "Site wise (All)"     — per-MRBTS unit-type counts (all states)
  "Site wise (Working)" — per-MRBTS unit-type counts (state=working only)
  "Overall"             — total working vs total count per unit type

Usage (command-line):
    python hw_tool/main.py
    python hw_tool/main.py input.xlsx -o output.xlsx

Usage (called from main parser with pre_read):
    from hw_tool.main import run_hw_report
    run_hw_report(pre_read=sheets_dict, output_path='HW_Report.xlsx')
"""

import os
import sys
import time


# ---------------------------------------------------------------------------
# Re-use the xlsx_reader from 2g_tool (it is format-agnostic)
# ---------------------------------------------------------------------------

def _get_xlsx_reader():
    """Import read_xlsx from 2g_tool — prefer calamine, fall back to openpyxl."""
    # When running standalone the 2g_tool dir may not be on sys.path yet.
    _here   = os.path.dirname(os.path.abspath(__file__))
    _parent = os.path.dirname(_here)
    tool_2g = os.path.join(_parent, '2g_tool')

    if tool_2g not in sys.path:
        sys.path.insert(0, tool_2g)

    from xlsx_reader import read_xlsx   # noqa
    return read_xlsx


# ---------------------------------------------------------------------------
# Public API — callable from main parser
# ---------------------------------------------------------------------------

NEEDED_SHEETS = ['INVUNIT', 'MRBTS', 'LNBTS']


def run_hw_report(pre_read=None, input_file=None, output_path=None,
                  progress_fn=None):
    """
    Build the HW report.

    pre_read     — dict {sheet_name: [row_dicts]} from the main parser's
                   pre_read snapshot.  When supplied, input_file is ignored.
    input_file   — path to xlsx/xlsb (used when pre_read is None).
    output_path  — destination xlsx path.
    progress_fn  — optional callable(message: str) for progress printing.

    Returns number of MRBTS sites written, or 0 on error.
    """
    from report import build_hw_report   # hw_tool/report.py

    def _log(msg):
        if progress_fn:
            progress_fn(msg)
        else:
            print(msg)

    # ── Load sheets ──────────────────────────────────────────────────────────
    if pre_read is not None:
        sheets = pre_read
    else:
        if not input_file:
            raise ValueError('Either pre_read or input_file must be supplied.')
        read_xlsx = _get_xlsx_reader()
        _log(f'HW: Reading {os.path.basename(input_file)} …')
        sheets = read_xlsx(input_file, sheet_names=NEEDED_SHEETS,
                           progress_fn=lambda m: _log(f'  [HW] {m}'))

    # ── Validate ─────────────────────────────────────────────────────────────
    if not sheets.get('INVUNIT'):
        _log('HW: INVUNIT sheet not found or empty — skipping HW report.')
        return 0

    has_site = sheets.get('MRBTS') or sheets.get('LNBTS')
    if not has_site:
        _log('HW: Neither MRBTS nor LNBTS found — site names will be blank.')

    # ── Build ────────────────────────────────────────────────────────────────
    t0    = time.perf_counter()
    count = build_hw_report(sheets, output_path)
    elapsed = time.perf_counter() - t0

    _log(f'HW: Done — {count} sites → {os.path.basename(output_path)}'
         f'  ({elapsed:.1f}s)')
    return count


# ---------------------------------------------------------------------------
# Standalone entry point
# ---------------------------------------------------------------------------

def _standalone():
    import argparse
    import tkinter as tk
    from tkinter import filedialog

    ap = argparse.ArgumentParser(
        description='HW Inventory Report — Nokia OSS dump → Excel')
    ap.add_argument('input',  nargs='?', help='Input xlsx/xlsb file')
    ap.add_argument('-o', '--output', help='Output xlsx file')
    args = ap.parse_args()

    root = tk.Tk()
    root.withdraw()

    # Input file
    input_file = args.input
    if not input_file:
        input_file = filedialog.askopenfilename(
            title='Select parsed OSS dump (xlsx / xlsb)',
            filetypes=[('Excel files', '*.xlsx *.xlsb'), ('All files', '*.*')],
        )
    if not input_file:
        print('No input file selected.')
        return

    # Output file
    output_path = args.output
    if not output_path:
        base = os.path.splitext(os.path.basename(input_file))[0]
        output_path = filedialog.asksaveasfilename(
            title='Save HW report as',
            initialdir=os.path.dirname(input_file),
            initialfile=f'{base}_HW_Report.xlsx',
            defaultextension='.xlsx',
            filetypes=[('Excel Workbook', '*.xlsx')],
        )
    if not output_path:
        print('No output file selected.')
        return

    root.destroy()

    # Ensure hw_tool is on path so `from report import ...` works
    _here = os.path.dirname(os.path.abspath(__file__))
    if _here not in sys.path:
        sys.path.insert(0, _here)

    count = run_hw_report(input_file=input_file, output_path=output_path)
    if count:
        print(f'\nReport written: {output_path}')


if __name__ == '__main__':
    _standalone()
