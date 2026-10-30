from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol


class CleanError(Exception):
    """A file was not safely processed; the destination was not replaced."""


@dataclass
class CleanReport:
    source: str
    output: str
    changes: dict[str, int] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)

    def count(self, category: str, amount: int = 1) -> None:
        self.changes[category] = self.changes.get(category, 0) + amount

    def warn(self, message: str) -> None:
        if message not in self.warnings:
            self.warnings.append(message)


class Cleaner(Protocol):
    extensions: frozenset[str]

    def clean(self, source: Path, report: CleanReport) -> bytes: ...
