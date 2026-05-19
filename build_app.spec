# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for img-batch-paster macOS .app bundle.

Build:
    pyinstaller build_app.spec

Output:
    dist/img-batch-paster.app
"""
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# 把 package 內的 static / templates 帶進來
datas = []
datas += collect_data_files("img_batch_paster.web", includes=["static/**/*"])
datas += collect_data_files("img_batch_paster.templates", includes=["*.pptx"])

hidden = [
    "img_batch_paster.web.app",
    "img_batch_paster.web.template_render",
    "img_batch_paster.pptx_writer",
    "img_batch_paster.xlsx_writer",
    "img_batch_paster.keynote_export",
    "img_batch_paster.grouper",
    "openpyxl.cell._writer",
]

a = Analysis(
    ["src/img_batch_paster/app_bundle.py"],
    pathex=["src"],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "pytest"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="img-batch-paster",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
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
    upx=False,
    upx_exclude=[],
    name="img-batch-paster",
)

app = BUNDLE(
    coll,
    name="img-batch-paster.app",
    icon=None,
    bundle_identifier="com.zealzel.imgbatchpaster",
    info_plist={
        "CFBundleShortVersionString": "1.0.0",
        "CFBundleVersion": "1.0.0",
        "NSHighResolutionCapable": True,
        "LSMinimumSystemVersion": "11.0",
    },
)
