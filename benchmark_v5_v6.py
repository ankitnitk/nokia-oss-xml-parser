#!/usr/bin/env python3
"""
benchmark_v5_v6.py  —  Compare V5.1 vs V6.0 OSS XML parser performance.

Runs both converters on the same input file(s), collects wall-clock time,
peak RAM, and average CPU, then prints a side-by-side summary table.

Requires psutil for RAM/CPU monitoring (pip install psutil).
Without psutil only timing is reported.

Usage:
    python benchmark_v5_v6.py  file.xml.gz
    python benchmark_v5_v6.py  file.zip --classes BCF,BTS,TRX
    python benchmark_v5_v6.py  file.gz  --runs 3
"""

import argparse
import os
import subprocess
import sys
import tempfile
import threading
import time
from datetime import datetime

try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False
    print('[WARN] psutil not found — RAM/CPU metrics unavailable.  pip install psutil')


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------

def _monitor_process(proc, interval=0.2):
    """
    Poll proc + children every `interval` seconds.
    Returns (peak_ram_bytes, avg_cpu_pct).
    Runs in its own thread; call .join() after proc finishes.
    """
    peak_ram = 0
    cpu_samples = []

    if not HAS_PSUTIL:
        return peak_ram, 0.0

    try:
        ps = psutil.Process(proc.pid)
    except psutil.NoSuchProcess:
        return peak_ram, 0.0

    while proc.poll() is None:
        try:
            children = ps.children(recursive=True)
            all_procs = [ps] + children

            ram = sum(p.memory_info().rss for p in all_procs
                      if p.is_running())
            if ram > peak_ram:
                peak_ram = ram

            # cpu_percent(interval=None) returns since last call — fast
            cpu = sum(p.cpu_percent(interval=None) for p in all_procs
                      if p.is_running())
            if cpu > 0:
                cpu_samples.append(cpu)

        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
        time.sleep(interval)

    avg_cpu = sum(cpu_samples) / len(cpu_samples) if cpu_samples else 0.0
    return peak_ram, avg_cpu


def run_converter(script_path, input_files, output_path, classes=None):
    """
    Run one converter script as a subprocess and collect metrics.
    Returns dict with keys: elapsed, peak_ram_mb, avg_cpu_pct, stdout, returncode.
    """
    cmd = [sys.executable, script_path] + input_files + ['-o', output_path]
    if classes:
        cmd += ['--classes', ','.join(sorted(classes))]

    t0 = time.perf_counter()

    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )

    # Kick off RAM/CPU monitor in a background thread
    peak_ram = [0]
    avg_cpu  = [0.0]

    def _mon():
        r, c = _monitor_process(proc)
        peak_ram[0] = r
        avg_cpu[0]  = c

    mon_thread = threading.Thread(target=_mon, daemon=True)
    mon_thread.start()

    stdout, _ = proc.communicate()
    elapsed   = time.perf_counter() - t0
    mon_thread.join(timeout=2)

    return {
        'elapsed':     elapsed,
        'peak_ram_mb': peak_ram[0] / (1024 * 1024) if peak_ram[0] else None,
        'avg_cpu_pct': avg_cpu[0] if avg_cpu[0] else None,
        'stdout':      stdout,
        'returncode':  proc.returncode,
    }


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

def _fmt_mb(mb):
    if mb is None:
        return 'N/A'
    if mb >= 1024:
        return f'{mb/1024:.1f} GB'
    return f'{mb:.0f} MB'

def _fmt_cpu(pct):
    if pct is None:
        return 'N/A'
    return f'{pct:.0f}%'

def _fmt_time(s):
    if s < 60:
        return f'{s:.1f}s'
    m, sec = divmod(int(s), 60)
    return f'{m}m {sec:02d}s'

def _pct_diff(old, new):
    """Return '−12.3%' or '+5.6%' relative improvement from old→new."""
    if old is None or new is None or old == 0:
        return ''
    diff = (new - old) / old * 100
    sign = '+' if diff > 0 else ''
    return f'{sign}{diff:.1f}%'

def _speedup(old_t, new_t):
    if old_t and new_t and new_t > 0:
        ratio = old_t / new_t
        if ratio >= 1:
            return f'{ratio:.2f}× faster'
        else:
            return f'{1/ratio:.2f}× slower'
    return ''


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    here = os.path.dirname(os.path.abspath(__file__))

    ap = argparse.ArgumentParser(
        description='Benchmark V5.1 vs V6.0 OSS XML converter'
    )
    ap.add_argument('inputs',    nargs='+', metavar='FILE',
                    help='Input XML / gz / zip file(s)')
    ap.add_argument('--classes', default='',
                    help='Comma-separated MO class filter (optional)')
    ap.add_argument('--runs',    type=int, default=1,
                    help='Number of benchmark runs per version (default 1)')
    ap.add_argument('--v51',     default=os.path.join(here, 'oss_xml_to_xlsx_v5.1.py'),
                    help='Path to V5.1 script')
    ap.add_argument('--v60',     default=os.path.join(here, 'oss_xml_to_xlsx_v6.0.py'),
                    help='Path to V6.0 script')
    args = ap.parse_args()

    classes = {c.strip() for c in args.classes.split(',') if c.strip()} or None

    for path in args.inputs:
        if not os.path.isfile(path):
            print(f'ERROR: Input file not found: {path}', file=sys.stderr)
            sys.exit(1)

    for script, label in [(args.v51, 'V5.1'), (args.v60, 'V6.0')]:
        if not os.path.isfile(script):
            print(f'ERROR: {label} script not found: {script}', file=sys.stderr)
            sys.exit(1)

    print(f'\nOSS XML Parser Benchmark  —  {datetime.now().strftime("%Y-%m-%d %H:%M")}')
    print(f'Input  : {", ".join(os.path.basename(f) for f in args.inputs)}')
    if classes:
        print(f'Classes: {", ".join(sorted(classes))}')
    print(f'Runs   : {args.runs} per version')
    print(f'psutil : {"yes" if HAS_PSUTIL else "no (install psutil for RAM/CPU)"}')
    print()

    results = {'V5.1': [], 'V6.0': []}
    scripts = {'V5.1': args.v51, 'V6.0': args.v60}

    for version in ('V5.1', 'V6.0'):
        script = scripts[version]
        print(f'─── {version} ────────────────────────────────────────────')
        for run_i in range(1, args.runs + 1):
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False,
                                             prefix=f'bench_{version}_') as tf:
                out_path = tf.name

            try:
                print(f'  Run {run_i}/{args.runs} → {os.path.basename(out_path)} ...',
                      end='', flush=True)
                r = run_converter(script, args.inputs, out_path, classes)
                file_mb = (os.path.getsize(out_path) / (1024*1024)
                           if os.path.isfile(out_path) else None)
                r['file_mb'] = file_mb
                results[version].append(r)

                status = 'OK' if r['returncode'] == 0 else f'FAIL (rc={r["returncode"]})'
                print(f'  {status}  {_fmt_time(r["elapsed"])}'
                      f'  RAM={_fmt_mb(r["peak_ram_mb"])}'
                      f'  CPU={_fmt_cpu(r["avg_cpu_pct"])}'
                      f'  out={_fmt_mb(file_mb)}')

                if r['returncode'] != 0:
                    print('  --- stdout/stderr ---')
                    for line in r['stdout'].splitlines()[-30:]:
                        print('  ' + line)
            finally:
                try:
                    os.remove(out_path)
                except OSError:
                    pass
        print()

    # ── Summary table ────────────────────────────────────────────────────────
    def _best(runs, key):
        vals = [r[key] for r in runs if r.get(key) is not None and r['returncode'] == 0]
        return min(vals) if vals else None

    def _avg(runs, key):
        vals = [r[key] for r in runs if r.get(key) is not None and r['returncode'] == 0]
        return sum(vals) / len(vals) if vals else None

    v51 = results['V5.1']
    v60 = results['V6.0']

    t51   = _avg(v51, 'elapsed')
    t60   = _avg(v60, 'elapsed')
    ram51 = _avg(v51, 'peak_ram_mb')
    ram60 = _avg(v60, 'peak_ram_mb')
    cpu51 = _avg(v51, 'avg_cpu_pct')
    cpu60 = _avg(v60, 'avg_cpu_pct')
    sz51  = _avg(v51, 'file_mb')
    sz60  = _avg(v60, 'file_mb')

    W = 14
    print('═' * 62)
    print(f'{"SUMMARY (avg over runs)":<22}  {"V5.1":>{W}}  {"V6.0":>{W}}  {"Change":>{W}}')
    print('─' * 62)
    print(f'{"Wall-clock time":<22}  '
          f'{_fmt_time(t51) if t51 else "N/A":>{W}}  '
          f'{_fmt_time(t60) if t60 else "N/A":>{W}}  '
          f'{_speedup(t51, t60):>{W}}')
    if HAS_PSUTIL:
        print(f'{"Peak RAM":<22}  '
              f'{_fmt_mb(ram51):>{W}}  '
              f'{_fmt_mb(ram60):>{W}}  '
              f'{_pct_diff(ram51, ram60):>{W}}')
        print(f'{"Avg CPU (all cores)":<22}  '
              f'{_fmt_cpu(cpu51):>{W}}  '
              f'{_fmt_cpu(cpu60):>{W}}  '
              f'{_pct_diff(cpu51, cpu60):>{W}}')
    print(f'{"Output file size":<22}  '
          f'{_fmt_mb(sz51):>{W}}  '
          f'{_fmt_mb(sz60):>{W}}  '
          f'{_pct_diff(sz51, sz60):>{W}}')
    print('═' * 62)

    if t51 and t60:
        saved = t51 - t60
        sign  = 'saved' if saved >= 0 else 'added'
        print(f'\nV6.0 is {_speedup(t51, t60)} — '
              f'{abs(saved):.1f}s {sign} per run.')
    if sz51 and sz60:
        sz_diff = sz51 - sz60
        sign    = 'smaller' if sz_diff >= 0 else 'larger'
        print(f'Output file is {abs(sz_diff):.1f} MB {sign} '
              f'({_pct_diff(sz51, sz60).strip()} size change).')


if __name__ == '__main__':
    main()
