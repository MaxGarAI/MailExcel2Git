"""Small typed boundary around Outlook's dynamically dispatched objects."""

from typing import Protocol


class PropertyAccessor(Protocol):
    def GetProperty(self, name: str) -> object: ...


class ComAttachment(Protocol):
    @property
    def FileName(self) -> str: ...

    def SaveAsFile(self, path: str) -> None: ...


class Attachments(Protocol):
    @property
    def Count(self) -> int: ...

    def Item(self, index: int) -> ComAttachment: ...


class Folder(Protocol):
    @property
    def StoreID(self) -> str: ...


class OutlookItem(Protocol):
    @property
    def Class(self) -> int: ...


class MailItem(OutlookItem, Protocol):
    @property
    def EntryID(self) -> str: ...

    @property
    def Parent(self) -> Folder: ...

    @property
    def Subject(self) -> str: ...

    @property
    def SenderName(self) -> str: ...

    @property
    def SenderEmailAddress(self) -> str: ...

    @property
    def PropertyAccessor(self) -> PropertyAccessor: ...

    @property
    def Attachments(self) -> Attachments: ...


class Selection(Protocol):
    @property
    def Count(self) -> int: ...

    def Item(self, index: int) -> object: ...


class Explorer(Protocol):
    @property
    def Selection(self) -> Selection: ...


class OutlookApplication(Protocol):
    def ActiveExplorer(self) -> Explorer | None: ...
