# -*- mode: python ; coding: utf-8 -*-
#
# V6.5.8 "folder" (onedir) build -- same code, same excludes as the onefile
# V6.5.8 spec, but distributed as a folder instead of a single exe.
#
# Why this variant exists:
#   A onefile exe unpacks its entire payload into %TEMP%\_MEIxxxxxx on every
#   single launch and loads its DLLs/.pyds from there. Managed Windows fleets
#   routinely block exactly that -- AppLocker / WDAC rules and Defender ASR
#   ("Block executable content from running unless it meets a prevalence,
#   age, or trusted list criterion") deny loading executable code out of a
#   user-writable temp path. That is the classic cause of an exe that runs
#   fine on one Windows 11 box and dies on the next one with identical specs.
#
#   This build never extracts anything at runtime, so it sidesteps that whole
#   class of policy failure -- and it starts several times faster because the
#   unpack step is gone.
#
# Distribute the whole dist_v658_folder\OSS_XML_Parser_v6.5.8 folder (zipped);
# users run OSS_XML_Parser_v6.5.8.exe inside it.
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
    'numpy', 'lxml', 'pandas', 'pyarrow', 'fastparquet',
    'fsspec', 'tqdm', 'urllib3', 'requests',
    'PIL', 'Pillow',
    'win32ui', 'win32uiole', 'Pythonwin',
    'ssl', '_ssl',
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
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,      # <-- the onedir switch
    name='OSS_XML_Parser_v6.5.8',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='version_info_v658.txt',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='OSS_XML_Parser_v6.5.8',
)
