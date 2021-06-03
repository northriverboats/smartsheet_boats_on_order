# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['smartsheet_boats_on_order.py'],
             pathex=['/home/fwarren/builds/smartsheet_boats_on_order'],
             binaries=[],
             datas=[
                 ('.env','.'),
                 ('templates','templates'),
                 ('templates/downloads','templates/downloads'),
             ],
             hiddenimports=['smartsheet.reports','emailer.emailer'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='smartsheet_boats_on_order',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True )
