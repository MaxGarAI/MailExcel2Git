from contextlib import contextmanager
from datetime import UTC, datetime, timedelta, timezone
from hashlib import sha256
from pathlib import Path
from types import SimpleNamespace
from typing import cast
from unittest.mock import Mock

import pytest
import pywintypes

from kuchenland_importer.application.capture_mail import CaptureMail
from kuchenland_importer.domain.errors import AttachmentError, CaptureStorageError, OutlookError
from kuchenland_importer.infrastructure.capture_manifest import JsonCaptureManifestStore
from kuchenland_importer.infrastructure.outlook import session, source
from kuchenland_importer.infrastructure.outlook.attachment_writer import save_attachment
from kuchenland_importer.infrastructure.outlook.com_types import (
    MailItem,
    OutlookApplication,
    Selection,
)
from kuchenland_importer.infrastructure.outlook.metadata import read_metadata
from kuchenland_importer.infrastructure.outlook.source import capture_selection


def com_failure() -> pywintypes.com_error:
    return pywintypes.com_error(-2147221021, "test COM failure", None, None)


class FakeAttachment:
    def __init__(self, name: str, content: bytes = b"Excel bytes", fail: bool = False):
        self.FileName = name
        self.content = content
        self.fail = fail

    def SaveAsFile(self, path: str) -> None:
        Path(path).write_bytes(self.content)
        if self.fail:
            raise com_failure()


class FakeCollection:
    def __init__(self, items: list[object]):
        self.items = items
        self.Count = len(items)

    def Item(self, index: int) -> object:
        item = self.items[index - 1]
        if isinstance(item, Exception):
            raise item
        return item


def mail(
    entry: str = "entry", attachments: list[object] | None = None, store: str = "store"
) -> SimpleNamespace:
    return SimpleNamespace(
        Class=43,
        EntryID=entry,
        Parent=SimpleNamespace(StoreID=store),
        Subject="Заказник — весна",
        SenderName="Закупщик",
        SenderEmailAddress="sender@example.test",
        PropertyAccessor=SimpleNamespace(GetProperty=lambda _: datetime(2026, 10, 9, 10)),
        Attachments=FakeCollection(attachments or []),
    )


def collect(tmp_path: Path, items: list[object]):
    return capture_selection(cast(Selection, FakeCollection(items)), "run", tmp_path, (".xlsx",))


def test_mixed_selection_and_untrusted_names(tmp_path: Path) -> None:
    capture = collect(
        tmp_path,
        [
            SimpleNamespace(Class=26),
            mail(attachments=[FakeAttachment(r"..\..\CON.XLSX"), FakeAttachment("ignore.pdf")]),
            mail("second", [FakeAttachment(r"..\..\CON.XLSX")]),
        ],
    )
    assert capture.selected_count == 3
    assert len(capture.mails) == 2
    assert capture.selection_positions == (2, 3)
    assert len(capture.attachments) == 2
    paths = [item.path for item in capture.attachments]
    assert paths[0] != paths[1]
    assert all(path.is_relative_to(tmp_path) for path in paths)
    assert all(path.name == "attachment-0001.xlsx" for path in paths)
    assert capture.attachments[0].sha256 == sha256(b"Excel bytes").hexdigest()
    assert capture.attachments[0].original_name == r"..\..\CON.XLSX"
    assert capture.issues[0].code == "NOT_MAIL"


def test_one_failed_attachment_does_not_block_other_files(tmp_path: Path) -> None:
    capture = collect(
        tmp_path,
        [
            mail(
                attachments=[
                    FakeAttachment("bad.xlsx", fail=True),
                    FakeAttachment("empty.xlsx", b""),
                    FakeAttachment("good.xlsx"),
                ]
            )
        ],
    )
    assert len(capture.attachments) == 1
    assert capture.attachments[0].attachment_index == 3
    assert capture.error_count == 2
    assert not list(tmp_path.rglob("*.part"))


def test_duplicate_identity_and_same_id_in_another_store(tmp_path: Path) -> None:
    capture = collect(tmp_path, [mail(), mail(), mail(store="another-store")])
    assert len(capture.mails) == 2
    assert [issue.code for issue in capture.issues].count("DUPLICATE_MAIL") == 1


def test_snapshot_item_failure_and_bad_metadata_continue(tmp_path: Path) -> None:
    broken = mail("bad")
    broken.PropertyAccessor = SimpleNamespace(GetProperty=lambda _: None)
    capture = collect(tmp_path, [com_failure(), broken, mail("good")])
    assert len(capture.mails) == 1
    assert capture.error_count == 2
    assert {issue.code for issue in capture.issues} >= {
        "SELECTION_ITEM_FAILED",
        "MAIL_READ_FAILED",
        "NO_EXCEL",
    }


def test_empty_selection_is_actionable(tmp_path: Path) -> None:
    with pytest.raises(OutlookError, match="не выделены"):
        collect(tmp_path, [])


def test_utc_delivery_time_is_not_converted_twice(tmp_path: Path) -> None:
    item = mail()
    # PT_SYSTIME is UTC wall-clock even if a wrapper adds a tzinfo label.
    item.PropertyAccessor = SimpleNamespace(
        GetProperty=lambda _: datetime(2026, 10, 9, 10, tzinfo=timezone(timedelta(hours=3)))
    )
    assert collect(tmp_path, [item]).mails[0].received_at == datetime(2026, 10, 9, 10, tzinfo=UTC)


def test_com_cleanup_after_connection_failure(monkeypatch: pytest.MonkeyPatch) -> None:
    initialize, uninitialize = Mock(), Mock()
    monkeypatch.setattr(session.pythoncom, "CoInitializeEx", initialize)
    monkeypatch.setattr(session.pythoncom, "CoUninitialize", uninitialize)
    monkeypatch.setattr(session.win32com.client, "GetActiveObject", Mock(side_effect=com_failure()))
    with pytest.raises(OutlookError, match="Классический Outlook"):
        with session.outlook_session():
            pytest.fail("Unavailable Outlook must not yield")
    initialize.assert_called_once()
    uninitialize.assert_called_once()


def test_com_initialization_failure_is_not_uninitialized(monkeypatch: pytest.MonkeyPatch) -> None:
    uninitialize = Mock()
    monkeypatch.setattr(session.pythoncom, "CoInitializeEx", Mock(side_effect=com_failure()))
    monkeypatch.setattr(session.pythoncom, "CoUninitialize", uninitialize)
    with pytest.raises(OutlookError, match="инициализировать COM"):
        with session.outlook_session():
            pytest.fail("Failed initialization must not yield")
    uninitialize.assert_not_called()


def test_com_cleanup_after_body_failure(monkeypatch: pytest.MonkeyPatch) -> None:
    uninitialize = Mock()
    monkeypatch.setattr(session.pythoncom, "CoInitializeEx", Mock())
    monkeypatch.setattr(session.pythoncom, "CoUninitialize", uninitialize)
    monkeypatch.setattr(session.win32com.client, "GetActiveObject", Mock(return_value=object()))
    with pytest.raises(RuntimeError, match="body failure"):
        with session.outlook_session():
            raise RuntimeError("body failure")
    uninitialize.assert_called_once()


def test_attachment_never_overwrites_existing_file(tmp_path: Path) -> None:
    existing = tmp_path / "attachment-0001.xlsx"
    existing.write_bytes(b"original")
    with pytest.raises(AttachmentError, match="перезапись запрещена"):
        save_attachment(
            FakeAttachment("new.xlsx"),
            read_metadata(cast(MailItem, mail())),
            1,
            tmp_path,
            "new.xlsx",
        )
    assert existing.read_bytes() == b"original"


def test_manifest_failure_keeps_saved_attachment(tmp_path: Path) -> None:
    capture = collect(tmp_path, [mail(attachments=[FakeAttachment("good.xlsx")])])
    report = tmp_path / "manifest.json"
    report.write_text("existing report", encoding="utf-8")
    with pytest.raises(CaptureStorageError):
        JsonCaptureManifestStore().save(capture)
    assert report.read_text(encoding="utf-8") == "existing report"
    assert capture.attachments[0].path.read_bytes() == b"Excel bytes"


def test_no_explorer_is_actionable(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    @contextmanager
    def fake_session():
        yield cast(OutlookApplication, SimpleNamespace(ActiveExplorer=lambda: None))

    monkeypatch.setattr(source, "outlook_session", fake_session)
    with pytest.raises(OutlookError, match="основное окно"):
        source.OutlookMailSource().capture("run", tmp_path, (".xlsx",))


def test_use_case_writes_manifest_and_preserves_original_name(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import json

    selection = FakeCollection([mail(attachments=[FakeAttachment("Поставщик.xlsx")])])
    app = SimpleNamespace(ActiveExplorer=lambda: SimpleNamespace(Selection=selection))

    @contextmanager
    def fake_session():
        yield cast(OutlookApplication, app)

    monkeypatch.setattr(source, "outlook_session", fake_session)
    service = CaptureMail(source.OutlookMailSource(), JsonCaptureManifestStore())
    first = service.execute(tmp_path, (".xlsx",))
    second = service.execute(tmp_path, (".xlsx",))
    assert first.capture.directory != second.capture.directory
    manifest = json.loads(first.manifest.read_text(encoding="utf-8"))
    assert manifest["stage"] == "mail_capture"
    assert manifest["mails"][0]["received_at"].endswith("+00:00")
    assert manifest["attachments"][0]["original_name"] == "Поставщик.xlsx"
    assert manifest["attachments"][0]["path"] == "mail-0001/attachment-0001.xlsx"
    assert not list(tmp_path.rglob("*.part"))


def test_cli_capture_does_not_require_destination_workbook(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    from kuchenland_importer.presentation.cli import main

    @contextmanager
    def fake_session():
        selection = FakeCollection([mail(attachments=[FakeAttachment("товары.xlsx")])])
        yield cast(
            OutlookApplication,
            SimpleNamespace(ActiveExplorer=lambda: SimpleNamespace(Selection=selection)),
        )

    monkeypatch.setattr(source, "outlook_session", fake_session)
    config = Path(__file__).resolve().parents[2] / "config" / "default.toml"
    missing_book = tmp_path / "does-not-exist.xlsx"
    assert (
        main(
            [
                "capture-outlook",
                "--config",
                str(config),
                "--data-dir",
                str(tmp_path / "data"),
                "--workbook",
                str(missing_book),
            ]
        )
        == 0
    )
    assert "Сохранено Excel-вложений: 1" in capsys.readouterr().out
    assert not missing_book.exists()
