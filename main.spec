import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files
_vendor = Path.cwd() / '_vendor'
if _vendor.is_dir():
    sys.path.insert(0, str(_vendor))
# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[str(_vendor)] if _vendor.is_dir() else [],
    binaries=[],
    datas=collect_data_files('tkinterdnd2'),
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Horizon Chantier',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
if sys.platform == "darwin":
    app = BUNDLE(
        exe,
        name='Horizon Chantier.app',
        icon=None,
        bundle_identifier=None,
        info_plist={
            'CFBundleDevelopmentRegion': 'fr',
            'CFBundleLocalizations': ['fr'],
            'CFBundleAllowMixedLocalizations': True,
        },
    )
