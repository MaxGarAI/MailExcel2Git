"""Shared frozen entry point: GUI by default, console in the -cli executable."""

import sys
from pathlib import Path

if __name__ == "__main__":
    if not getattr(sys, "frozen", False):
        sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
    if Path(sys.executable).stem.lower().endswith("-cli") or "--cli" in sys.argv[1:]:
        from kuchenland_importer.presentation.cli import main
        arguments = sys.argv[1:]
        if "--cli" in arguments:
            arguments.remove("--cli")
        raise SystemExit(main(arguments))
    from kuchenland_importer.presentation.gui import main

    raise SystemExit(main())
