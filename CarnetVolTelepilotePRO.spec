# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['carnet_vol_pro_app.py'],
    pathex=[],
    binaries=[],
    datas=[('assets\\tof_logo_header.png', 'assets'), ('assets\\lettre_mission_vierge.docx', 'assets'), ('assets\\cerfa_15476-04.pdf', 'assets'), ('assets\\guide_preparation_mission_complet.pdf', 'assets'), ('assets\\attestation_consentement_vierge.pdf', 'assets')],
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
    name='CarnetVolTelepilotePRO',
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
