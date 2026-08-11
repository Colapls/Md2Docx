# -*- mode: python ; coding: utf-8 -*-
import os
import sys

block_cipher = None

# ------------------------------------------------------------
# Locate Tcl/Tk resources (DLL + data files) automatically.
# This supports both Anaconda-style installs and standard Python.
# ------------------------------------------------------------

def _find_first_existing(paths):
    for p in paths:
        if p and os.path.exists(p):
            return p
    return None

conda_prefix = os.environ.get('CONDA_PREFIX', '') or sys.prefix
userprofile  = os.environ.get('USERPROFILE', '') or ''

candidates_root = [
    conda_prefix,
    os.path.join(userprofile, 'anaconda3'),
    os.path.join(userprofile, 'miniconda3'),
    'C:\\Users\\NHY\\anaconda3',
    'C:\\Users\\NHY\\miniconda3',
]

# 1) Tcl/Tk runtime DLLs (tcl86t.dll / tk86t.dll)
tk_bin_dir = _find_first_existing([
    os.path.join(r, 'Library', 'bin') for r in candidates_root
] + [
    os.path.join(r, 'DLLs') for r in candidates_root
] + [
    os.path.join(r, 'bin') for r in candidates_root
])

# 2) Tcl/Tk data dirs (tcl8.6/, tk8.6/)
tcl_lib_dir = _find_first_existing([
    os.path.join(r, 'Library', 'lib', 'tcl8.6') for r in candidates_root
] + [
    os.path.join(r, 'lib', 'tcl8.6') for r in candidates_root
] + [
    os.path.join(r, 'tcl', 'tcl8.6') for r in candidates_root
])

tk_lib_dir = _find_first_existing([
    os.path.join(r, 'Library', 'lib', 'tk8.6') for r in candidates_root
] + [
    os.path.join(r, 'lib', 'tk8.6') for r in candidates_root
] + [
    os.path.join(r, 'tcl', 'tk8.6') for r in candidates_root
])

print(f"[Tk] bin dir: {tk_bin_dir}")
print(f"[Tk] tcl data: {tcl_lib_dir}")
print(f"[Tk] tk  data: {tk_lib_dir}")

binaries = []
datas = []

if tk_bin_dir:
    binaries += [
        (os.path.join(tk_bin_dir, 'tcl86t.dll'), '.'),
        (os.path.join(tk_bin_dir, 'tk86t.dll'),  '.'),
    ]

if tcl_lib_dir:
    datas += [(tcl_lib_dir, 'tcl8.6')]
if tk_lib_dir:
    datas += [(tk_lib_dir, 'tk8.6')]

# 3) latex2mathml ships a symbols data file; PyInstaller's hook
# doesn't pick it up automatically. Bundle it explicitly so that
# the frozen app can read it at runtime.
import importlib
try:
    _l2m_dir = os.path.dirname(importlib.import_module('latex2mathml').__file__)
    datas += [(os.path.join(_l2m_dir, 'unimathsymbols.txt'), 'latex2mathml')]
except Exception as _e:
    print(f'[warn] could not locate latex2mathml data: {_e}')


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=[
        'markdown',
        'markdown.extensions',
        'markdown.extensions.extra',
        'markdown.extensions.codehilite',
        'markdown.extensions.tables',
        'markdown.extensions.fenced_code',
        'markdown.extensions.nl2br',
        'docx',
        'docx.oxml',
        'docx.shared',
        'docx.enum',
        'docx.enum.text',
        'docx.enum.style',
        'lxml',
        'lxml.etree',
        'tkinter',
        '_tkinter',
        # ttk is loaded by tkinter dynamically; declare explicitly to be safe
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.font',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[os.path.abspath('runtime_hook_tk.py')],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Markdown2Docx',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
