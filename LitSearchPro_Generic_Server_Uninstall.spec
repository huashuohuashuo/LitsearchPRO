# -*- mode: python ; coding: utf-8 -*-

a = Analysis(['LitSearchPro_Generic_Server_Uninstall.py'], pathex=[], binaries=[], datas=[],
             hiddenimports=[], hookspath=[], hooksconfig={}, runtime_hooks=[],
             excludes=[], noarchive=False, optimize=0)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [],
          name='LitSearchPro_Generic_Server_v22.1.21_Uninstall',
          debug=False, strip=False, upx=True, console=False,
          disable_windowed_traceback=False, icon=['generic_logo.ico'])
