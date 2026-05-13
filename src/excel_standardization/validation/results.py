"""Shared validation result structures."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List


@dataclass
class ValidationResult:
    """A single validation finding for one field in one row."""

    field_name: str
    message: str
    severity: str = "error"  # "error" | "warning"

    def __str__(self) -> str:
        return f"[{self.severity.upper()}] {self.field_name}: {self.message}"


@dataclass
class RowValidationResult:
    """Aggregated validation results for a single data row."""

    row_index: int
    row_uid: str | None
    findings: List[ValidationResult] = field(default_factory=list)

    @property
    def is_valid(self) -> bool:
        return not any(f.severity == "error" for f in self.findings)

    @property
    def has_warnings(self) -> bool:
        return any(f.severity == "warning" for f in self.findings)

    def add(self, field_name: str, message: str, severity: str = "error") -> None:
        self.findings.append(ValidationResult(field_name=field_name, message=message, severity=severity))

    def status_summary(self) -> str:
        """Return a pipe-separated summary of all messages, or empty string."""
        return " | ".join(f.message for f in self.findings)
