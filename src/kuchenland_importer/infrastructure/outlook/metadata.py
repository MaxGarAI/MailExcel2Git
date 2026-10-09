"""Read identity and delivery time; the MAPI PT_SYSTIME value is UTC."""

from datetime import UTC, datetime

from kuchenland_importer.domain.mail import MailMetadata
from kuchenland_importer.infrastructure.outlook.com_types import MailItem

DELIVERY_TIME = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"


def read_metadata(item: MailItem) -> MailMetadata:
    received = item.PropertyAccessor.GetProperty(DELIVERY_TIME)
    if not isinstance(received, datetime):
        raise ValueError("Outlook не предоставил дату получения письма.")
    # PropertyAccessor returns UTC wall-clock values, unlike the local ReceivedTime property.
    received_utc = received.replace(tzinfo=UTC)
    return MailMetadata(
        entry_id=item.EntryID,
        store_id=item.Parent.StoreID,
        subject=item.Subject or "",
        received_at=received_utc,
        sender_name=item.SenderName or "",
        sender_address=item.SenderEmailAddress or "",
    )
