# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = []
hiddenimports += collect_submodules('openpyxl')
hiddenimports += collect_submodules('holidays')


block_cipher = None


a = Analysis(
    ['generar_proyeccion.py'],
    pathex=[],
    binaries=[],
    datas=[('Inputs/Parametros.xlsx', 'Inputs'), ('Inputs/Puerto Chile.xlsx', 'Inputs'), ('Inputs/Venta - Plan.xlsx', 'Inputs'), ('Inputs/stock.xlsx', 'Inputs'), ('Inputs/ETA/Distribucion+Logistica - Pedidos AP-Confirmados.xlsx', 'Inputs/ETA'), ('Inputs/ETA/Logística - Pedidos Planta-Puerto-Embarcado.xlsx', 'Inputs/ETA'), ('Inputs/Asignaciones.xlsx', 'Inputs'), ('Inputs/Produccion faltante.xlsx', 'Inputs'), ('Img/Notice.png', 'Img'), ('Img/ico.ico', 'Img')],
    hiddenimports=hiddenimports,
    hookspath=['.'],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='generar_proyeccion',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['Img/ico.ico']
)
