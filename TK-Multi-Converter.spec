# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for the TK multi-country table converter.

Builds a single Windows .exe that supports ID / PH / TH — the user picks
the target country from a top-of-window radio in the GUI. The country-
specific assets (template xlsx + default size chart) are shipped as three
sibling folders next to the exe (assets-id, assets-ph, assets-th) and
the app picks them up at runtime via config.get_assets_dir(country).

Build:
    pyinstaller --noconfirm TK-Multi-Converter.spec
"""

from pathlib import Path
import sys

block_cipher = None
APP_NAME = "TK-Multi-Converter"
PROJECT_DIR = Path(SPECPATH).resolve()  # set by PyInstaller

# Make `app` importable as a package so relative imports inside
# app/main.py keep working in the frozen exe.
sys.path.insert(0, str(PROJECT_DIR))

# Entry: invoke the package's __main__ via -m flag.
# PyInstaller supports the standard '-c' argv form for this.
a = Analysis(
    [str(PROJECT_DIR / "app" / "__main__.py")],
    pathex=[str(PROJECT_DIR)],
    binaries=[],
    datas=[],
    hiddenimports=[
        # Force-include every module in the app package so relative imports
        # in the frozen exe always resolve.
        "app",
        "app.config",
        "app.presets",
        "app.source_reader",
        "app.converter",
        "app.product_pool",
        "app.tiktok_writer",
        # Optional drag-and-drop library. The GUI imports it lazily
        # inside a try/except, so if it's missing at runtime the exe
        # still works (just no DnD).
        "tkinterdnd2",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib", "scipy", "pandas", "numpy.testing",
        "pytest", "setuptools",
    ],
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
    name=APP_NAME,
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
    icon=None,
)
