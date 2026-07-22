# -*- mode: python ; coding: utf-8 -*-
# Build: pyinstaller desktop_app.spec
# Çıktı: dist/ReyartLeadSistemi/ReyartLeadSistemi.exe (klasör; output/ yanına oluşur)
# Tek dosya .exe istersen: onefile=True'ya çevir (aşağıdaki not).

a = Analysis(
    ['desktop_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['pandas', 'openpyxl'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ReyartLeadSistemi',
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
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ReyartLeadSistemi',
)
