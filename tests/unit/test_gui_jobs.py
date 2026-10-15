from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pytest

from kuchenland_importer.application.job_control import JobControl
from kuchenland_importer.domain.errors import ExcelInputError
from kuchenland_importer.domain.report import ImportIssue, ImportReport, Severity
from kuchenland_importer.infrastructure.import_runner import ImportOutcome
from kuchenland_importer.presentation.gui_jobs import import_job


def test_invalid_target_does_not_contact_outlook():
    app = SimpleNamespace(settings=SimpleNamespace(workbook=None), capture_outlook=Mock())
    with pytest.raises(ExcelInputError, match="общую книгу"):
        import_job(app, Path("config/default.toml"), JobControl(lambda event: None))
    app.capture_outlook.assert_not_called()


def test_capture_without_attachments_reports_manifest_and_does_not_write(tmp_path):
    target = tmp_path / "master.xlsx"
    target.write_bytes(b"exists")
    issue = ImportIssue("ATTACHMENT_FAILED", "Вложение недоступно", Severity.ERROR, "selection:1")
    capture = SimpleNamespace(attachments=(), issues=(issue,))
    app = SimpleNamespace(
        settings=SimpleNamespace(workbook=target),
        capture_outlook=Mock(
            return_value=SimpleNamespace(capture=capture, manifest=tmp_path / "capture.json")
        ),
    )
    with patch("kuchenland_importer.presentation.gui_jobs.run_import") as writer:
        with pytest.raises(ExcelInputError, match="Вложение недоступно"):
            import_job(app, Path("config/default.toml"), JobControl(lambda event: None))
        writer.assert_not_called()


def test_partial_capture_issues_are_kept_alongside_committed_result(tmp_path):
    target = tmp_path / "master.xlsx"
    target.write_bytes(b"exists")
    issue = ImportIssue("ITEM_FAILED", "Письмо недоступно", Severity.ERROR, "selection:2")
    capture = SimpleNamespace(attachments=(object(),), issues=(issue,))
    app = SimpleNamespace(
        settings=SimpleNamespace(workbook=target),
        capture_outlook=Mock(
            return_value=SimpleNamespace(capture=capture, manifest=tmp_path / "capture.json")
        ),
    )
    saved = ImportOutcome(ImportReport("run", products_added=1), None, None, True)
    with patch("kuchenland_importer.presentation.gui_jobs.run_import", return_value=saved):
        result = import_job(app, Path("config/default.toml"), JobControl(lambda event: None))
    assert result.outcome.committed
    assert result.capture_issues == (issue,)
    assert result.capture_manifest == tmp_path / "capture.json"
