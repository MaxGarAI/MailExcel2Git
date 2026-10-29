"""Cooperative cancellation and progress without dependencies on Tk or Office."""

from collections.abc import Callable
from dataclasses import dataclass, field
from threading import Event, Lock

from kuchenland_importer.domain.errors import ApplicationError
from kuchenland_importer.domain.field_review import FieldChoice, FieldDecision


class JobCancelled(ApplicationError):
    """Cancelled at a boundary before destination writes."""


@dataclass(frozen=True, slots=True)
class Progress:
    message: str
    cancellable: bool = True


@dataclass(slots=True)
class JobControl:
    publish: Callable[[Progress], None]
    cancelled: Event = field(default_factory=Event)

    _lock: Lock = field(default_factory=Lock)
    _writing: bool = False
    review_fields: Callable[[FieldChoice], FieldDecision] | None = None

    def request_cancel(self) -> bool:
        with self._lock:
            if self._writing:
                return False
            self.cancelled.set()
            return True

    def checkpoint(self, message: str) -> None:
        if self.cancelled.is_set():
            raise JobCancelled("Импорт отменён до записи общей книги.")
        self.publish(Progress(message))

    def begin_write(self) -> None:
        with self._lock:
            if self.cancelled.is_set():
                raise JobCancelled("Импорт отменён до записи общей книги.")
            self._writing = True
        self.publish(Progress("Запись и проверка книги. Дождитесь завершения.", False))
