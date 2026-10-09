"""Ports for collecting mail and recording its provenance."""

from pathlib import Path
from typing import Protocol

from kuchenland_importer.domain.capture import MailCapture


class MailSource(Protocol):
    def capture(self, run_id: str, directory: Path, extensions: tuple[str, ...]) -> MailCapture: ...


class CaptureManifestStore(Protocol):
    def save(self, capture: MailCapture) -> Path: ...
