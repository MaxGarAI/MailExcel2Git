"""One owned non-daemon worker with a Tk-independent event queue."""

from collections.abc import Callable
from dataclasses import dataclass
from queue import Empty, Queue
from threading import Thread

from loguru import logger

from kuchenland_importer.application.job_control import JobCancelled, JobControl, Progress
from kuchenland_importer.domain.errors import ApplicationError
from kuchenland_importer.domain.field_review import FieldChoice, FieldDecision
from kuchenland_importer.presentation.gui_field_review import ReviewRequest
from kuchenland_importer.presentation.gui_jobs import GuiResult


@dataclass(frozen=True, slots=True)
class Completed:
    result: GuiResult


@dataclass(frozen=True, slots=True)
class Failed:
    message: str
    cancelled: bool = False


type WorkerEvent = Progress | Completed | Failed | ReviewRequest


class GuiWorker:
    def __init__(self) -> None:
        self.events: Queue[WorkerEvent] = Queue()
        self.control = JobControl(self.events.put)
        self.busy = False
        self.cancellable = False
        self.thread: Thread | None = None

    def start(self, task: Callable[[JobControl], GuiResult]) -> None:
        if self.busy:
            raise RuntimeError("Другой импорт ещё выполняется.")
        self.busy = self.cancellable = True
        self.control = JobControl(self.events.put)
        self.control.review_fields = self._review
        self.thread = Thread(
            target=self._run, args=(task,), name="procurement-import", daemon=False
        )
        self.thread.start()

    def _run(self, task: Callable[[JobControl], GuiResult]) -> None:
        try:
            result = task(self.control)
        except JobCancelled as error:
            self.events.put(Failed(str(error), True))
        except ApplicationError as error:
            logger.warning("Ошибка GUI-импорта: {}", error)
            self.events.put(Failed(str(error)))
        except Exception:
            logger.exception("Непредвиденная ошибка GUI-импорта")
            self.events.put(Failed("Непредвиденная ошибка. Подробности в журнале приложения."))
        else:
            self.events.put(Completed(result))

    def _review(self, choice: FieldChoice) -> FieldDecision:
        request = ReviewRequest(choice)
        self.events.put(request)
        while not self.control.cancelled.is_set():
            try:
                decision = request.response.get(timeout=0.2)
            except Empty:
                continue
            if decision is None:
                raise JobCancelled("Импорт отменён при согласовании полей.")
            return decision
        raise JobCancelled("Импорт отменён при согласовании полей.")

    def cancel(self) -> None:
        if self.busy and self.cancellable:
            self.control.request_cancel()

    def drain(self) -> list[WorkerEvent]:
        events: list[WorkerEvent] = []
        while True:
            try:
                event = self.events.get_nowait()
            except Empty:
                return events
            if isinstance(event, Progress):
                self.cancellable = event.cancellable
            elif not isinstance(event, ReviewRequest):
                self.busy = self.cancellable = False
            events.append(event)
