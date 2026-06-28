# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


project_root = Path(SPECPATH).resolve()
asset_dir = project_root / "print_assist" / "assets"

excluded_optional_modules = [
    "IPython",
    "matplotlib",
    "numba",
    "numpy",
    "openpyxl",
    "pandas",
    "pyarrow",
    "pytest",
    "scipy",
    "setuptools",
]

a = Analysis(
    [str(project_root / "main.py")],
    pathex=[str(project_root)],
    binaries=[],
    datas=[(str(asset_dir), "print_assist/assets")],
    hiddenimports=["win32timezone"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excluded_optional_modules,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    name="PrintAssist",
    icon=str(asset_dir / "print-assist.ico"),
    debug=False,
    bootloader_ignore_signals=False,
    exclude_binaries=True,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="PrintAssist",
)
