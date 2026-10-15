from threading import Event

import pytest

from kuchenland_importer.application.job_control import JobCancelled, JobControl, Progress
from kuchenland_importer.domain.errors import ApplicationError
from kuchenland_importer.domain.report import ImportReport
from kuchenland_importer.infrastructure.import_runner import ImportOutcome
from kuchenland_importer.presentation.gui_jobs import GuiResult
from kuchenland_importer.presentation.gui_worker import Completed, Failed, GuiWorker


def result():
    return GuiResult(ImportOutcome(ImportReport("test"), None, None, False))


def finish(worker):
    worker.thread.join(timeout=5)
    assert not worker.thread.is_alive()
    return worker.drain()


def test_cancellation_before_write_and_refusal_after_write_boundary():
    control = JobControl(lambda event: None)
    assert control.request_cancel()
    with pytest.raises(JobCancelled):
        control.begin_write()
    control = JobControl(lambda event: None)
    control.begin_write()
    assert not control.request_cancel()
    assert not control.cancelled.is_set()


def test_worker_rejects_parallel_job_until_completion_consumed():
    worker = GuiWorker()
    ready, release = Event(), Event()

    def task(control):
        control.checkpoint("Reading")
        ready.set()
        assert release.wait(5)
        return result()

    worker.start(task)
    assert ready.wait(5)
    with pytest.raises(RuntimeError):
        worker.start(task)
    release.set()
    events = finish(worker)
    assert isinstance(events[0], Progress)
    assert isinstance(events[-1], Completed)
    assert not worker.busy
    worker.start(lambda control: result())
    assert isinstance(finish(worker)[-1], Completed)


def test_worker_cancel_while_preparing_delivers_cancelled_event():
    worker = GuiWorker()
    ready, release = Event(), Event()

    def task(control):
        ready.set()
        assert release.wait(5)
        control.checkpoint("Before transaction")
        return result()

    worker.start(task)
    assert ready.wait(5)
    worker.cancel()
    release.set()
    event = finish(worker)[-1]
    assert isinstance(event, Failed) and event.cancelled


@pytest.mark.parametrize(
    "error,expected",
    [
        (ApplicationError("Outlook недоступен"), "Outlook недоступен"),
        (RuntimeError("internal private data"), "Непредвиденная ошибка"),
    ],
)
def test_worker_delivers_actionable_errors_without_unexpected_details(error, expected):
    worker = GuiWorker()

    def task(control):
        raise error

    worker.start(task)
    event = finish(worker)[-1]
    assert isinstance(event, Failed) and expected in event.message
    assert "internal private data" not in event.message
