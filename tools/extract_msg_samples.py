"""QA utility for the supplied Unicode MSG samples, not a general MSG importer."""

import argparse
import hashlib
import json
import struct
from datetime import UTC, datetime, timedelta
from pathlib import Path

import pythoncom
import pywintypes


def read_stream(storage, name: str) -> bytes:
    stream = storage.OpenStream(name, None, 0x10, 0)
    chunks = []
    while chunk := stream.Read(1024 * 1024):
        chunks.append(chunk)
    return b"".join(chunks)


def text(storage, tag: str) -> str:
    try:
        return read_stream(storage, f"__substg1.0_{tag}001F").decode("utf-16-le").rstrip("\0")
    except pywintypes.com_error:
        return ""


def delivery_time(storage) -> datetime:
    properties = read_stream(storage, "__properties_version1.0")
    for offset in range(32, len(properties) - 15, 16):
        if struct.unpack_from("<I", properties, offset)[0] == 0x0E060040:
            ticks = struct.unpack_from("<Q", properties, offset + 8)[0]
            return datetime(1601, 1, 1, tzinfo=UTC) + timedelta(microseconds=ticks // 10)
    raise ValueError("MSG sample has no delivery time; sent date cannot replace it")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    data = {
        "schema_version": 1,
        "stage": "mail_capture",
        "run_id": "msg-qa-samples",
        "selected_count": 0,
        "mails": [],
        "attachments": [],
        "issues": [],
    }
    for index, path in enumerate(sorted(args.source.glob("*.msg")), 1):
        original = path.read_bytes()
        digest = hashlib.sha256(original).hexdigest()
        storage = pythoncom.StgOpenStorage(str(path.resolve()), None, 0x20, None, 0)
        data["mails"].append(
            {
                "selection_index": index,
                "entry_id": "msg:" + digest,
                "store_id": "offline-msg-samples",
                "subject": text(storage, "0037"),
                "received_at": delivery_time(storage).isoformat(),
                "sender_name": text(storage, "0C1A"),
                "sender_address": text(storage, "5D01") or text(storage, "0C1F"),
            }
        )
        folder = output / f"mail-{index:04d}"
        folder.mkdir()
        for name, kind, *_ in storage.EnumElements():
            if kind != 1 or not name.startswith("__attach_version1.0_"):
                continue
            attachment = storage.OpenStorage(name, None, 0x10, None, 0)
            filename = text(attachment, "3707") or text(attachment, "3704")
            suffix = Path(filename).suffix.lower()
            if suffix not in {".xlsx", ".xlsm", ".xls", ".xlsb"}:
                continue
            content = read_stream(attachment, "__substg1.0_37010102")
            target = folder / f"attachment-{len(data['attachments']) + 1:04d}{suffix}"
            target.write_bytes(content)
            data["attachments"].append(
                {
                    "mail_index": index - 1,
                    "attachment_index": int(name.rsplit("#", 1)[1], 16) + 1,
                    "original_name": filename,
                    "path": target.relative_to(output).as_posix(),
                    "sha256": hashlib.sha256(content).hexdigest(),
                }
            )
        storage = None
        assert hashlib.sha256(path.read_bytes()).hexdigest() == digest
    data["selected_count"] = len(data["mails"])
    (output / "manifest.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"Extracted {len(data['attachments'])} Excel attachments from {len(data['mails'])} MSG")


if __name__ == "__main__":
    main()
