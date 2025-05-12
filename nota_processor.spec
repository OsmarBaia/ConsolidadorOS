# os_consolidador.spec

from PyInstaller.utils.hooks import collect_data_files
from pathlib import Path

tesseract_dir = Path("tesseract")
poppler_dir = Path("poppler/bin")
tessdata = list(tesseract_dir.glob("tessdata/*.traineddata"))
poppler_files = list(poppler_dir.glob("*.*"))

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        (str(tesseract_dir / "tesseract.exe"), "tesseract"),
        *[(str(td), f"tesseract/tessdata") for td in tessdata],
        *[(str(pf), "poppler/bin") for pf in poppler_files],
    ],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ConsolidadorOS',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='icon.ico' if Path('icon.ico').exists() else None
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='ConsolidadorOS'
)
