# PyInstaller 6: one analysis, two entry points, one shared dependency directory.
from pathlib import Path
from PyInstaller.utils.hooks import copy_metadata

root = Path(SPECPATH)
metadata = []
for distribution in ("pywin32", "xlwings", "openpyxl", "pandas", "Pillow", "loguru"):
    metadata += copy_metadata(distribution)
analysis = Analysis(
    [str(root / "launcher.py")],
    pathex=[str(root / "src")],
    binaries=[],
    datas=[(str(root / "config" / name), "config")
           for name in ("default.toml", "columns.toml", "writing.toml")] + metadata,
    hiddenimports=["win32timezone", "pythoncom", "pywintypes", "pandas", "xlwings"],
    hookspath=[], hooksconfig={}, runtime_hooks=[],
    excludes=["pytest", "mypy", "ruff", "matplotlib", "IPython", "jupyter"],
    noarchive=False,
)
archive = PYZ(analysis.pure)
version = str(root / "resources" / "windows-version.txt")
gui = EXE(archive, analysis.scripts, [], exclude_binaries=True,
          name="KuchenlandImporter", debug=False, bootloader_ignore_signals=False,
          strip=False, upx=False, console=False, version=version)
cli = EXE(archive, analysis.scripts, [], exclude_binaries=True,
          name="KuchenlandImporter-cli", debug=False, bootloader_ignore_signals=False,
          strip=False, upx=False, console=True, version=version)
collection = COLLECT(gui, cli, analysis.binaries, analysis.datas,
                     strip=False, upx=False, name="KuchenlandImporter")
