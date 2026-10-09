"""Launch the application from a source checkout without installation."""

import sys
from pathlib import Path

if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))
    from kuchenland_importer.presentation.cli import main

    raise SystemExit(main())
