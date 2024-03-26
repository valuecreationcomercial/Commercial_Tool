# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['DatabaseOrganization_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('icon.ico','.')],
    hiddenimports=[],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    name='Commercial_Tool',
    strip=False,
    upx=True,
    console=False,
    icon='icon.ico'
)
