"""Build a versioned Windows x64 folder/ZIP and a SHA-256 file manifest."""

import hashlib
import importlib.metadata
import json
import os
import platform
import shutil
import struct
import subprocess
import sys
from pathlib import Path
from uuid import uuid4

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from kuchenland_importer import __version__  # noqa: E402


def main() -> None:
    if sys.platform != "win32" or sys.version_info[:2] != (3, 12) or struct.calcsize("P") != 8:
        raise SystemExit("Build requires Windows x64 and Python 3.12 x64.")
    if importlib.metadata.version("pyinstaller") != "6.22.0":
        raise SystemExit("Install requirements-build.txt first.")
    release_name = f"KuchenlandImporter-{__version__}-windows-x64"
    release = ROOT / "dist" / release_name
    archive = release.parent / (release.name + ".zip")
    if release.exists() or archive.exists():
        raise SystemExit(f"Release already exists; preserve it or choose a new version: {release}")
    work = ROOT / "build" / uuid4().hex
    work.mkdir(parents=True)
    environment = dict(os.environ, PYINSTALLER_CONFIG_DIR=str(work / "cache"))
    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--distpath",
        str(work / "dist"),
        "--workpath",
        str(work / "objects"),
        str(ROOT / "KuchenlandImporter.spec"),
    ]
    log = work / "pyinstaller.log"
    with log.open("w", encoding="utf-8") as stream:
        completed = subprocess.run(
            command, cwd=ROOT, env=environment, stdout=stream, stderr=subprocess.STDOUT, check=False
        )
    if completed.returncode:
        raise SystemExit(f"Build failed. Log: {log}")
    release.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(work / "dist" / "KuchenlandImporter"), release)
    shutil.copy2(ROOT / "docs" / "windows-user-guide.md", release / "Инструкция.md")
    for exe in ("KuchenlandImporter.exe", "KuchenlandImporter-cli.exe"):
        if not (release / exe).is_file():
            raise SystemExit(f"Missing executable: {exe}")
    licenses = release / "third-party-licenses"
    licenses.mkdir()
    for distribution in (
        "pywin32",
        "xlwings",
        "openpyxl",
        "pandas",
        "Pillow",
        "loguru",
        "numpy",
        "pytz",
        "python-dateutil",
        "et_xmlfile",
        "six",
        "colorama",
        "win32_setctime",
        "tzdata",
        "pyinstaller",
    ):
        package = importlib.metadata.distribution(distribution)
        for item in package.files or ():
            if any(
                word in item.name.lower() for word in ("license", "licence", "copying", "notice")
            ):
                source = Path(package.locate_file(item))
                if source.is_file():
                    destination = (
                        licenses / distribution / str(item).replace("/", "_").replace("\\", "_")
                    )
                    destination.parent.mkdir(exist_ok=True)
                    shutil.copy2(source, destination)
    python_license = Path(sys.base_prefix) / "LICENSE.txt"
    if python_license.is_file():
        shutil.copy2(python_license, licenses / "Python-LICENSE.txt")
    versions = {}
    for distribution in (
        "pyinstaller",
        "pyinstaller-hooks-contrib",
        "pywin32",
        "xlwings",
        "openpyxl",
        "pandas",
        "Pillow",
        "loguru",
        "numpy",
    ):
        versions[distribution] = importlib.metadata.version(distribution)
    manifest = {
        "version": __version__,
        "python": platform.python_version(),
        "platform": "windows-x64",
        "dependencies": versions,
        "source_files": {
            str(p.relative_to(ROOT)).replace("\\", "/"): hashlib.sha256(p.read_bytes()).hexdigest()
            for p in sorted((ROOT / "src").rglob("*.py"))
        },
        "files": {
            str(p.relative_to(release)).replace("\\", "/"): hashlib.sha256(
                p.read_bytes()
            ).hexdigest()
            for p in sorted(release.rglob("*"))
            if p.is_file()
        },
    }
    (release / "release-manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    shutil.make_archive(str(release), "zip", release.parent, release.name)
    digest = hashlib.sha256(archive.read_bytes()).hexdigest()
    archive.with_suffix(".zip.sha256").write_text(f"{digest}  {archive.name}\n", encoding="ascii")
    print(f"Release: {release}\nZIP: {archive}\nSHA256: {digest}\nBuild log: {log}")


if __name__ == "__main__":
    main()
