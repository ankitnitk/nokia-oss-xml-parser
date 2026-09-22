# -*- mode: python ; coding: utf-8 -*-
#
# V6.5.9 -- slimmed + compatibility-hardened build. Functionally identical to
# V6.5.7; only the packaging changed. See EXCLUDES below for what was dropped
# and why (each was verified unused before removal).
#
from PyInstaller.utils.hooks import collect_all

datas = [('../2g_tool', '2g_tool'), ('../4g_tool', '4g_tool'),
         ('../3g_tool', '3g_tool'), ('../hw_tool', 'hw_tool')]
binaries = []
hiddenimports = []
tmp_ret = collect_all('xlsxwriter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('python_calamine')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

EXCLUDES = [
    # Heavy data/science libs -- never imported by the parser or any tool.
    'numpy', 'lxml', 'pandas', 'pyarrow', 'fastparquet',
    'fsspec', 'tqdm', 'urllib3', 'requests',

    # Pillow (~5.7 MB packed, _avif.pyd alone is 4.3 MB). openpyxl pulls it in
    # only for embedded-image support; nothing here reads or writes images.
    'PIL', 'Pillow',

    # pywin32's GUI/MFC layer (~3 MB, mostly mfc140u.dll). Excel automation
    # here is late-bound win32com.client.DispatchEx only -- never win32ui,
    # never gencache/EnsureDispatch. MFC redistributable DLLs are also a
    # frequent corporate-AV false-positive trigger.
    'win32ui', 'win32uiole', 'Pythonwin',

    # No network I/O anywhere in the codebase, so OpenSSL (libcrypto-3.dll +
    # libssl-3.dll, ~2.1 MB) is dead weight.
    'ssl', '_ssl',

    # Dev-only stdlib.
    'unittest', 'doctest', 'pdb', 'pydoc', 'pydoc_data',
]

a = Analysis(
    ['../oss_xml_to_xlsx_v6.5.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
    # No asserts or docstring-dependent code anywhere -> safe to strip both.
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='OSS_XML_Parser_v6.5.9',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # Explicitly OFF (was True, but silently a no-op since UPX isn't on PATH).
    # UPX-packed exes are one of the most common antivirus false positives, so
    # pinning this False keeps builds identical on machines that do have UPX.
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info_v659.txt',
)
