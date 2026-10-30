"""Office Privacy Cleaner. No network calls or document-content logging."""

__version__ = "1.0.0"

from .engine import clean_file
from .models import CleanError, CleanReport

__all__ = ["clean_file", "CleanError", "CleanReport"]
