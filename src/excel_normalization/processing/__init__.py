"""Processing layer for field normalization.

This package contains the FieldProcessor abstract base class and its
concrete implementations for different field types (names, gender, dates,
identifiers). It also contains the NormalizationPipeline for applying
engines to JSON rows.
"""

from .field_processor import FieldProcessor
from .name_processor import NameFieldProcessor
from .gender_processor import GenderFieldProcessor
from .date_processor import DateFieldProcessor
from .identifier_processor import IdentifierFieldProcessor
from .normalization_pipeline import NormalizationPipeline

__all__ = [
    "FieldProcessor",
    "NameFieldProcessor",
    "GenderFieldProcessor",
    "DateFieldProcessor",
    "IdentifierFieldProcessor",
    "NormalizationPipeline",
]
