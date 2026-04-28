"""Processing layer for field standardization.

This package contains the FieldProcessor abstract base class and its
concrete implementations for different field types (names, gender, dates,
identifiers). It also contains the standardizationPipeline for applying
engines to JSON rows.
"""

from .field_processor import FieldProcessor
from .name_processor import NameFieldProcessor
from .gender_processor import GenderFieldProcessor
from .date_processor import DateFieldProcessor
from .identifier_processor import IdentifierFieldProcessor
from .standardization_pipeline import standardizationPipeline

__all__ = [
    "FieldProcessor",
    "NameFieldProcessor",
    "GenderFieldProcessor",
    "DateFieldProcessor",
    "IdentifierFieldProcessor",
    "standardizationPipeline",
]
