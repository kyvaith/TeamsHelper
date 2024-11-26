# teams_helper.spec
# Generated for building the Teams Helper application into a standalone executable.

# Import PyInstaller build spec functionalities
block_cipher = None

a = Analysis(
    ['main.py'],  # Main entry point of the application
    pathex=['.'],  # Include current directory
    binaries=[],
    datas=[],
    hiddenimports=['lameenc', 'pystray', 'PIL.Image', 'PIL.ImageDraw'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TeamsHelper',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Use False if you want a GUI app (no console window)
    icon='icon.ico',  # Optional: Path to your icon file
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='TeamsHelper',
)
