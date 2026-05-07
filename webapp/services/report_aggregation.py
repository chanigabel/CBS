"""Aggregation helpers for processing reports."""

from __future__ import annotations

from collections import Counter
from typing import Iterable

from webapp.models.processing_report import (
    InvalidIdentifierValue,
    MissingRequiredExportField,
    MissingRequiredFieldSummary,
    SummaryCount,
)


def aggregate_missing_required_fields(
    fields: Iterable[MissingRequiredExportField],
) -> list[MissingRequiredFieldSummary]:
    counts = Counter()
    for field in fields:
        counts[field.field_name] += field.rows_affected
    return [
        MissingRequiredFieldSummary(field=field_name, count=count)
        for field_name, count in sorted(counts.items())
    ]


def aggregate_validation_messages(messages: Iterable[str]) -> list[SummaryCount]:
    counts = Counter(message for message in messages if message)
    return [
        SummaryCount(message=message, count=count)
        for message, count in sorted(counts.items())
    ]


def aggregate_identifier_messages(
    details: Iterable[InvalidIdentifierValue],
    is_real_identifier_issue,
) -> list[SummaryCount]:
    counts = Counter()
    seen_rows = set()
    for item in details:
        if not is_real_identifier_issue(item.status_message):
            continue
        row_key = item.row_uid if item.row_uid is not None else item.row_number
        key = (item.sheet_name, row_key, item.status_message)
        if key in seen_rows:
            continue
        seen_rows.add(key)
        counts[item.status_message] += 1
    return [
        SummaryCount(message=message, count=count)
        for message, count in sorted(counts.items())
    ]

