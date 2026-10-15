"""Verify release hashes and run both frozen entry points in an isolated profile."""

import argparse
import hashlib
import json
import os
import subprocess
from pathlib import Path
from uuid import uuid4


def main() -> None:
    parser = argparse.ArgumentParser(description="Verify a Windows Kuchenland release")
    parser.add_argument("package", type=Path)
    parser.add_argument("--source-root", type=Path)
    parser.add_argument("--work-dir", type=Path)
    args = parser.parse_args()
    package = args.package.resolve()
    manifest = json.loads((package / "release-manifest.json").read_text(encoding="utf-8"))
    for name, expected in manifest["files"].items():
        path = (package / name).resolve()
        if not path.is_relative_to(package):
            raise ValueError(f"Invalid manifest path: {name}")
        if hashlib.sha256(path.read_bytes()).hexdigest() != expected:
            raise ValueError(f"File hash mismatch: {name}")
    if args.source_root:
        source = args.source_root.resolve()
        for name, expected in manifest["source_files"].items():
            path = (source / name).resolve()
            if not path.is_relative_to(source):
                raise ValueError(f"Invalid source path: {name}")
            if hashlib.sha256(path.read_bytes()).hexdigest() != expected:
                raise ValueError(f"Source hash mismatch: {name}")
    work = (args.work_dir or Path(".runtime") / f"package-check-{uuid4().hex}").resolve()
    work.mkdir(parents=True, exist_ok=False)
    environment = dict(
        os.environ,
        LOCALAPPDATA=str(work / "profile"),
        PATH=str(Path(os.environ["SystemRoot"]) / "System32"),
    )
    for name in ("PYTHONPATH", "PYTHONHOME", "VIRTUAL_ENV", "TCL_LIBRARY", "TK_LIBRARY"):
        environment.pop(name, None)
    commands = [
        ("version", "KuchenlandImporter-cli.exe", ["--version"]),
        ("diagnose", "KuchenlandImporter-cli.exe", ["diagnose"]),
        ("gui", "KuchenlandImporter.exe", ["--check-startup", str(work / "gui.json")]),
    ]
    for label, exe, arguments in commands:
        completed = subprocess.run(
            [str(package / exe), *arguments],
            cwd=work,
            env=environment,
            capture_output=True,
            timeout=60,
            check=False,
        )
        (work / f"{label}.stdout.txt").write_bytes(completed.stdout)
        (work / f"{label}.stderr.txt").write_bytes(completed.stderr)
        if completed.returncode:
            raise RuntimeError(f"{label}: exit {completed.returncode}; inspect {work}")
    version = (work / "version.stdout.txt").read_text(encoding="utf-8").strip()
    gui = json.loads((work / "gui.json").read_text(encoding="utf-8"))
    if version != manifest["version"] or gui["version"] != version or gui["status"] != "ok":
        raise ValueError("Version or GUI verification failed.")
    result = {
        "version": version,
        "files_checked": len(manifest["files"]),
        "gui": gui,
        "status": "ok",
    }
    (work / "result.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(json.dumps(result))


if __name__ == "__main__":
    main()
