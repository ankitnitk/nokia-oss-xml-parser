"""
main.py
Entry point for the 4G LTE Network Parameter Tool.
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
        title='Select 4G LTE Parameter Dump File',
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
    from xlsx_reader import read_xlsx
    return read_xlsx(input_path, sheet_names=needed_sheets,
                     progress_fn=progress, header_row=header_row)


def main():
    print('=' * 55)
    print('  4G LTE Network Parameter Tool')
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
    header_row = 0 if headers_row1 else 1
    print(f'Header row: {"Row 1" if headers_row1 else "Row 2 (default)"}')
    print()

    from network import Network
    needed_sheets = Network.NEEDED_SHEETS

    t0 = time.time()
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
    network = Network(sheets)
    print(f'  {len(network.lnbts_list):,} LNBTS records')
    print(f'  {len(network.lncel_by_dn):,} LNCEL records')
    fdd_total = sum(len(v) for v in network.lncel_fdd_list_by_lnbts_dn.values())
    tdd_total = sum(len(v) for v in network.lncel_tdd_list_by_lnbts_dn.values())
    print(f'  {fdd_total:,} FDD cells  |  {tdd_total:,} TDD cells')
    print()

    print('Generating report...')
    sys.path.insert(0, os.path.dirname(__file__))
    from reports.lnbts_summary import build
    n = build(network, output_path, progress_fn=progress)

    elapsed = time.time() - t0
    print()
    print('=' * 55)
    print(f'  Done.  {n:,} LNBTS written in {elapsed:.1f}s')
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
