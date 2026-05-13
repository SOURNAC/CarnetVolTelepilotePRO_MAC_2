# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files


block_cipher = None

datas = [
    ("assets/tof_logo_header.png", "assets"),
    ("assets/lettre_mission_vierge.docx", "assets"),
    ("assets/cerfa_15476-04.pdf", "assets"),
    ("assets/guide_preparation_mission_complet.pdf", "assets"),
    ("assets/attestation_consentement_vierge.pdf", "assets"),
]
datas += collect_data_files("openpyxl")

a = Analysis(
    ["carnet_vol_pro_app.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="CarnetVolTelepilotePRO",
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
    name="CarnetVolTelepilotePRO",
)

app = BUNDLE(
    coll,
    name="CarnetVolTelepilotePRO.app",
    icon=None,
    bundle_identifier="fr.snc-drones.carnet-vol-telepilote-pro",
    info_plist={
        "CFBundleName": "CarnetVolTelepilotePRO",
        "CFBundleDisplayName": "CarnetVolTelepilotePRO",
        "CFBundleShortVersionString": "1.0.0",
        "CFBundleVersion": "1.0.0",
        "NSHighResolutionCapable": True,
    },
)
