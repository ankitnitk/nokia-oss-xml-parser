"""
main.py
Entry point for the 2G Network Parameter Tool.
Opens file dialogs for input/output, dispatches to report builders.
"""

import os
import sys
import time
import tkinter as tk
from tkinter import filedialog, messagebox


def ask_input_file():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    path = filedialog.askopenfilename(
        title='Select 2G Parameter Dump File',
        filetypes=[
            ('Excel files', '*.xlsb *.xlsx'),
            ('Binary Excel', '*.xlsb'),
            ('Excel Workbook', '*.xlsx'),
            ('All files', '*.*'),
        ]
    )
    root.destroy()
    return path


def ask_output_file(input_path):
    input_dir  = os.path.dirname(input_path)
    input_name = os.path.splitext(os.path.basename(input_path))[0]
    default_name = f'{input_name}_summary.xlsx'
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    path = filedialog.asksaveasfilename(
        title='Save Summary As',
        initialdir=input_dir,
        initialfile=default_name,
        defaultextension='.xlsx',
        filetypes=[('Excel Workbook', '*.xlsx')],
    )
    root.destroy()
    return path


def progress(msg):
    print(f'  {msg}')


def read_file(input_path, ext, needed_sheets, header_row=1):
    """
    Read the input file using the best available reader.

    For .xlsx  : calamine (fast) -> openpyxl fallback
    For .xlsb  : calamine (fast) -> custom binary parser fallback
    Both open the file as a read-only stream - no temp copy needed.
    """
    try:
        import python_calamine  # noqa: F401
        from xlsx_reader import read_xlsx
        return read_xlsx(input_path, sheet_names=needed_sheets,
                         progress_fn=progress, header_row=header_row)
    except ImportError:
        pass

    # calamine not available - use format-specific fallback
    if ext == '.xlsx':
        from xlsx_reader import read_xlsx
        return read_xlsx(input_path, sheet_names=needed_sheets,
                         progress_fn=progress, header_row=header_row)
    else:
        # .xlsb - use our custom BIFF12 parser
        progress('calamine not found, using built-in XLSB parser...')
        from xlsb_reader import read_xlsb
        return read_xlsb(input_path, sheet_names=needed_sheets,
                         progress_fn=progress, header_row=header_row)


def main():
    print('=' * 55)
    print('  2G Network Parameter Tool')
    print('=' * 55)
    print()

    print('Select input file...')
    input_path = ask_input_file()
    if not input_path:
        print('No input file selected. Exiting.')
        return

    print(f'Input : {input_path}')
    ext = os.path.splitext(input_path)[1].lower()
    if ext not in ('.xlsb', '.xlsx'):
        messagebox.showerror('Unsupported File',
                             f'File type "{ext}" is not supported.\n'
                             'Please select an .xlsb or .xlsx file.')
        return

    print('Select output location...')
    output_path = ask_output_file(input_path)
    if not output_path:
        print('No output location selected. Exiting.')
        return

    print(f'Output: {output_path}')
    print()

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    headers_row1 = messagebox.askyesno(
        'Header Row',
        'Are column headers in Row 1 of each sheet?\n\n'
        'Click Yes  → headers in Row 1\n'
        'Click No   → headers in Row 2  (default Nokia format)',
        default=messagebox.NO,
    )
    root.destroy()
    header_row = 0 if headers_row1 else 1   # 0-based index of the header row
    print(f'Header row: {"Row 1" if headers_row1 else "Row 2 (default)"}')
    print()

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    neighbour_checks = messagebox.askyesno(
        'Neighbour Checks',
        'Run neighbour checks?\n\n'
        '  • One-Way ADCE\n'
        '  • Co-Site Missing Neighbours\n\n'
        '(May add a few extra seconds for large dumps)',
        default=messagebox.NO,
    )
    root.destroy()
    print(f'Neighbour checks: {"Yes" if neighbour_checks else "No"}')
    print()

    t0 = time.time()
    needed_sheets = ['BSC', 'BCF', 'BTS', 'TRX', 'ADCE', 'ADJW', 'ADJL', 'HOC', 'POC']

    print('Reading file...')
    try:
        sheets = read_file(input_path, ext, needed_sheets, header_row=header_row)
    except PermissionError:
        messagebox.showerror(
            'File Locked',
            'Cannot read the file.\n\n'
            'Possible causes:\n'
            '  - File is open in Excel (please close it)\n'
            '  - OneDrive is still syncing the file\n\n'
            'Please close the file / wait for sync and try again.\n\n'
            f'File: {input_path}'
        )
        return

    print()

    print('Building network model...')
    from network import Network
    network = Network(sheets)
    bsc_count = len(set(
        s['seg_dn'].split('/')[1] for s in network.segments.values()
    ))
    print(f'  {len(network.segments):,} segments found across {bsc_count:,} BSCs')
    print()

    print('Generating Cell Details...')
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'reports'))
    from reports.cell_summary import build as build_cell_summary
    n = build_cell_summary(network, output_path, progress_fn=progress,
                           neighbour_checks=neighbour_checks)

    elapsed = time.time() - t0
    print()
    print('=' * 55)
    print(f'  Done.  {n:,} cells written in {elapsed:.1f}s')
    print(f'  Output: {output_path}')
    print('=' * 55)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        import traceback
        print()
        print('ERROR:', e)
        traceback.print_exc()
        input('\nPress Enter to exit...')
