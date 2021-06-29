# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(['D:/6.Scrpit/5.Python Projects/GitHub Project/Project-Searcher/Project_Search_App.py'],
             pathex=['D:\\6.Scrpit\\5.Python Projects\\GitHub Project\\Project-Searcher'],
             binaries=[],
             datas=[],
             hiddenimports=[],
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
          name='Project_Search_App',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False , icon='D:\\6.Scrpit\\5.Python Projects\\GitHub Project\\Project-Searcher\\Icon.ico')
