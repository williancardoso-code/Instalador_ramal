# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['instalar_atendas.py'],
    pathex=[],
    binaries=[],
    datas=[('imagens_config', 'imagens_config')],
    hiddenimports=['requests', 'pywinauto', 'pyautogui'],
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
    name='Instalador_Atendas',
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
    icon=['imagens_config\\icone.ico'],
)
