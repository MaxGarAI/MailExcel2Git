"""Initialize COM in the calling thread and attach only to a running Outlook."""

import sys
from collections.abc import Iterator
from contextlib import contextmanager
from typing import cast

import pythoncom
import pywintypes
import win32com.client

from kuchenland_importer.domain.errors import OutlookError
from kuchenland_importer.infrastructure.outlook.com_types import OutlookApplication


@contextmanager
def outlook_session() -> Iterator[OutlookApplication]:
    if sys.platform != "win32":
        raise OutlookError("Подключение к Outlook поддерживается только в Windows.")
    app: OutlookApplication | None = None
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    except pywintypes.com_error as error:
        raise OutlookError("Не удалось инициализировать COM в текущем потоке.") from error
    try:
        try:
            app = cast(OutlookApplication, win32com.client.GetActiveObject("Outlook.Application"))
        except pywintypes.com_error as error:
            raise OutlookError(
                "Классический Outlook недоступен. Откройте Outlook с настроенной учётной "
                "записью и выделите письма. Новый Outlook этим способом не поддерживается."
            ) from error
        yield app
    finally:
        app = None
        pythoncom.CoUninitialize()
