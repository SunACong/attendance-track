# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files, copy_metadata
import streamlit
import os

block_cipher = None

# 获取streamlit安装路径
streamlit_path = os.path.dirname(streamlit.__file__)

datas = [
    (os.path.join(streamlit_path, "runtime"), "streamlit/runtime"),
]
datas += collect_data_files("streamlit")
datas += copy_metadata("streamlit")

a = Analysis(
    ['run.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'processPCKQ', 'processLGDJ', 'processQJDJ', 'processYDKQ', 'all', 'streamlit'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='run',
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
)
