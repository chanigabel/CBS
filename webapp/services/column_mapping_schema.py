"""JSON-backed schema for manual column header mapping."""

from __future__ import annotations

import json
import re
from copy import deepcopy
from pathlib import Path
from typing import Dict, List

from fastapi import HTTPException


DEFAULT_COLUMN_MAPPING_SCHEMA: Dict[str, List[str]] = {
    "first_name": ["first name", "firstname", "שם פרטי", "שם"],
    "last_name": ["last name", "lastname", "surname", "family name", "שם משפחה"],
    "father_name": ["father name", "fathername", "שם האב"],
    "gender": ["gender", "sex", "מין"],
    "id_number": ["id number", "id", "מספר זהות", "תעודת זהות", "ת.ז"],
    "passport": ["passport", "מספר דרכון", "דרכון"],
    "birth_date": ["birth date", "date of birth", "dob", "תאריך לידה"],
    "birth_year": ["birth year", "שנת לידה"],
    "birth_month": ["birth month", "חודש לידה"],
    "birth_day": ["birth day", "יום לידה"],
    "entry_date": ["entry date", "admission date", "תאריך כניסה"],
    "entry_year": ["entry year", "שנת כניסה"],
    "entry_month": ["entry month", "חודש כניסה"],
    "entry_day": ["entry day", "יום כניסה"],
}


def _normalize_label(value: str) -> str:
    value = (value or "").strip().lower()
    value = value.replace("_", " ")
    value = re.sub(r"\s+", " ", value)
    return value


class ColumnMappingSchemaService:
    """Stores canonical target fields and accepted synonyms in JSON."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self.ensure_exists()

    def ensure_exists(self) -> None:
        if self.path.exists():
            return
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(
            json.dumps(
                {
                    "version": "2026.05.14-001",
                    "fields": DEFAULT_COLUMN_MAPPING_SCHEMA,
                },
                indent=2,
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )

    def load(self) -> Dict[str, object]:
        self.ensure_exists()
        raw = json.loads(self.path.read_text(encoding="utf-8"))
        fields = raw.get("fields")
        if not isinstance(fields, dict):
            raise HTTPException(status_code=500, detail="Invalid column mapping schema.")
        return raw

    def reload(self) -> Dict[str, object]:
        """Reload the current JSON schema from disk."""
        return self.load()

    def save(self, raw: Dict[str, object]) -> None:
        fields = raw.get("fields")
        if not isinstance(fields, dict):
            raise HTTPException(status_code=400, detail="Column mapping schema must contain fields.")
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(raw, indent=2, ensure_ascii=False), encoding="utf-8")

    def fields(self) -> List[str]:
        return sorted((self.load().get("fields") or {}).keys())

    def mappings(self) -> Dict[str, List[str]]:
        return deepcopy(self.load().get("fields") or {})

    def suggestions(self) -> List[str]:
        values: List[str] = []
        seen: set[str] = set()
        for field, synonyms in self.mappings().items():
            for value in [field, *synonyms]:
                if value not in seen:
                    values.append(value)
                    seen.add(value)
        return values

    def resolve(self, value: str) -> str:
        candidate = (value or "").strip()
        if not candidate:
            raise HTTPException(status_code=400, detail="Column target name is required.")

        fields = self.mappings()
        if candidate in fields:
            return candidate

        normalized = _normalize_label(candidate)
        for field, synonyms in fields.items():
            if _normalize_label(field) == normalized:
                return field
            if any(_normalize_label(synonym) == normalized for synonym in synonyms):
                return field

        raise HTTPException(
            status_code=400,
            detail=f"'{candidate}' is not a supported standardized field name or synonym.",
        )

    def add_mapping(self, standard_name: str, synonym: str) -> Dict[str, List[str]]:
        standard_name = (standard_name or "").strip()
        synonym = (synonym or "").strip()
        if not standard_name or not synonym:
            raise HTTPException(status_code=400, detail="standard_name and synonym are required.")
        raw = self.load()
        fields = raw.setdefault("fields", {})
        synonyms = fields.setdefault(standard_name, [])
        if synonym not in synonyms and synonym != standard_name:
            synonyms.append(synonym)
        self.save(raw)
        return self.mappings()

    def remove_mapping(self, standard_name: str, synonym: str) -> Dict[str, List[str]]:
        standard_name = (standard_name or "").strip()
        synonym = (synonym or "").strip()
        raw = self.load()
        fields = raw.get("fields") or {}
        if standard_name not in fields:
            raise HTTPException(status_code=404, detail=f"Standard field '{standard_name}' not found.")
        if synonym == standard_name:
            fields.pop(standard_name)
        else:
            fields[standard_name] = [item for item in fields[standard_name] if item != synonym]
        self.save(raw)
        return self.mappings()
