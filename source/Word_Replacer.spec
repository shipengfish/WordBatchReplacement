# -*- mode: python ; coding: utf-8 -*-

# windows 使用该规则
block_cipher = None

a = Analysis(['advanced-word-replacer-app.py'],
             pathex=['./'],
             binaries=[],
             datas=[],
             hiddenimports=['PyQt6'],
             hookspath=[],
             hooksconfig={},
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
          name='多文档替换器',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          disable_windowed_traceback=False,
          target_arch='x86_64',
          codesign_identity=None,
          entitlements_file=None,
          icon='icon.ico')

# mac 使用以下规则

#block_cipher = None
#
#a = Analysis(['advanced-word-replacer-app.py'],
#             pathex=['./'],
#             binaries=[],
#             datas=[],
#             hiddenimports=['PyQt6'],
#             hookspath=[],
#             hooksconfig={},
#             runtime_hooks=[],
#             excludes=[],
#             win_no_prefer_redirects=False,
#             win_private_assemblies=False,
#             cipher=block_cipher,
#             noarchive=False)
#
#pyz = PYZ(a.pure, a.zipped_data,
#          cipher=block_cipher)
#
#exe = EXE(pyz,
#          a.scripts,
#          [],
#          exclude_binaries=True,
#          name='多文档替换器',
#          debug=False,
#          bootloader_ignore_signals=False,
#          strip=False,
#          upx=True,
#          console=False,
#          disable_windowed_traceback=False,
#          target_arch=None,
#          codesign_identity=None,
#          entitlements_file=None )
#
#coll = COLLECT(exe,
#               a.binaries,
#               a.zipfiles,
#               a.datas,
#               strip=False,
#               upx=True,
#               upx_exclude=[],
#               name='多文档替换器')
#
#app = BUNDLE(coll,
#             name='多文档替换器.app',
#             icon='icon.icns',  # 确保您有一个有效的 .icns 文件
#             bundle_identifier=None,
#             info_plist={
#                 'NSHighResolutionCapable': 'True',
#                 'NSPrincipalClass': 'NSApplication',
#                 'CFBundleShortVersionString': '1.0.0',
#             },
#            )