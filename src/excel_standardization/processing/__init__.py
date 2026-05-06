"""Active processing layer for dataset standardization.

Legacy direct Excel field processors were moved to ``archive_legacy/`` and are
not part of the active Web/Dataset runtime.
"""

from .standardization_pipeline import standardizationPipeline

__all__ = ["standardizationPipeline"]
