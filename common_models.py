#!/usr/bin/env python3
from dataclasses import dataclass
from typing import List, Dict, Optional
from datetime import datetime


@dataclass
class Email:
    """Data model for an email message."""
    subject: str
    content: str
    received: str

    def to_dict(self) -> Dict[str, str]:
        """Convert to dictionary for CSV export."""
        return {
            "subject": self.subject,
            "content": self.content,
            "received": self.received
        }