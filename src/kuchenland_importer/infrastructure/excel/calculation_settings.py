"""Restore only calcPr attributes in an Excel-saved package, preserving all other bytes."""

import os
import re
from pathlib import Path
from uuid import uuid4
from zipfile import ZipFile


def restore_calculation_mode(path: Path, mode: str) -> None:
    if mode not in {"auto", "manual", "autoNoTable"}:
        raise ValueError("Неизвестный исходный режим расчёта общей книги.")
    partial = path.with_name(f".kl-calc-{uuid4().hex}.zip")
    try:
        with ZipFile(path) as source:
            xml = source.read("xl/workbook.xml")
            tags = list(re.finditer(rb"<(?:\w+:)?calcPr\b[^>]*>", xml))
            if len(tags) != 1:
                raise ValueError("В сохранённой Excel книге нет однозначного calcPr.")
            tag = tags[0]
            changed = tag.group()
            attributes = {"calcMode": mode}
            if mode != "manual":
                attributes.update({"fullCalcOnLoad": "1", "forceFullCalc": "1"})
            for key, value in attributes.items():
                attribute = key.encode("ascii")
                replacement = attribute + b'="' + value.encode("ascii") + b'"'
                pattern = rb"\b" + attribute + rb'="[^"]*"'
                if re.search(pattern, changed):
                    changed = re.sub(pattern, replacement, changed)
                else:
                    end = -2 if changed.endswith(b"/>") else -1
                    changed = changed[:end] + b" " + replacement + changed[end:]
            result = xml[: tag.start()] + changed + xml[tag.end() :]
            with ZipFile(partial, "x") as destination:
                destination.comment = source.comment
                for entry in source.infolist():
                    destination.writestr(
                        entry,
                        result
                        if entry.filename == "xl/workbook.xml"
                        else source.read(entry.filename),
                    )
        os.replace(partial, path)
    finally:
        partial.unlink(missing_ok=True)
