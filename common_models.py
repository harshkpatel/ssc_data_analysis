# common_models.py
# Defines common data structures used across the scraper.

from dataclasses import dataclass
from datetime import datetime
from typing import Optional

@dataclass
class Email:
    """
    Represents a single email with its relevant details.
    """
    message_id: str  # Unique ID from Outlook for the message
    subject: Optional[str] = None
    body: Optional[str] = None
    sender_address: Optional[str] = None
    received_date: Optional[datetime] = None # Python datetime object