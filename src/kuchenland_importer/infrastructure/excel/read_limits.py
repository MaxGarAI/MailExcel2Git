"""Read budgets apply to OOXML structure, not the compressed workbook file size."""

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class ReadLimits:
    max_parts: int = 10_000
    max_uncompressed_bytes: int | None = 256 * 1024 * 1024
    max_xml_bytes: int = 16 * 1024 * 1024


SOURCE_LIMITS = ReadLimits()
# The local master accumulates photographs across many independent source files.
MASTER_LIMITS = ReadLimits(
    max_parts=100_000, max_uncompressed_bytes=None, max_xml_bytes=64 * 1024 * 1024
)
