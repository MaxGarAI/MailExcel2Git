"""Errors that can be translated into actionable user messages."""


class ApplicationError(Exception):
    """Base class for expected application failures."""


class ConfigurationError(ApplicationError):
    """Configuration is missing, inconsistent, or malformed."""


class StartupError(ApplicationError):
    """Application resources cannot be initialized."""


class OutlookError(ApplicationError):
    """Classic Outlook is unavailable or its selection cannot be obtained."""


class AttachmentError(ApplicationError):
    """An attachment could not be safely saved and verified."""


class CaptureStorageError(ApplicationError):
    """Source capture directories or manifest could not be written."""


class ExcelInputError(ApplicationError):
    """Excel preview inputs or report storage are invalid or unavailable."""


class SeasonRoutingError(ApplicationError):
    """A row has no unambiguous configured seasonal destination."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code
