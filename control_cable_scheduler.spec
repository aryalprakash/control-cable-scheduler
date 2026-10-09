# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for Control Cable Scheduler (macOS + Windows)."""
import sys
from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None
app_name = "Control Cable Scheduler"

pymupdf_datas, pymupdf_binaries, pymupdf_hidden = collect_all("pymupdf")
ezdxf_datas, ezdxf_binaries, ezdxf_hidden = collect_all("ezdxf")
openpyxl_hidden = collect_submodules("openpyxl")

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=pymupdf_binaries + ezdxf_binaries,
    datas=pymupdf_datas + ezdxf_datas,
    hiddenimports=(
        pymupdf_hidden
        + ezdxf_hidden
        + openpyxl_hidden
        + ["fitz", "pymupdf", "ezdxf", "openpyxl"]
    ),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
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
    [],
    exclude_binaries=True,
    name=app_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
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
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=app_name,
)

if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name=f"{app_name}.app",
        icon=None,
        bundle_identifier="com.controlcablescheduler.app",
        info_plist={
            "CFBundleName": app_name,
            "CFBundleDisplayName": app_name,
            "NSHighResolutionCapable": True,
        },
    )
