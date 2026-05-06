"""Tests proving public legacy direct Excel entry points are disabled."""

import pytest

from src.excel_standardization.orchestrator import standardizationOrchestrator


LEGACY_DISABLED_MESSAGE = "Disabled legacy direct Excel path"


def test_normalize_workbook_legacy_direct_path_is_disabled():
    orch = standardizationOrchestrator()

    with pytest.raises(RuntimeError, match=LEGACY_DISABLED_MESSAGE):
        orch.normalize_workbook("input.xlsx")


def test_process_workbook_json_legacy_direct_path_is_disabled():
    orch = standardizationOrchestrator()

    with pytest.raises(RuntimeError, match=LEGACY_DISABLED_MESSAGE):
        orch.process_workbook_json("input.xlsx", "output.xlsx")


def test_export_from_processors_legacy_direct_path_is_disabled():
    orch = standardizationOrchestrator()

    with pytest.raises(RuntimeError, match=LEGACY_DISABLED_MESSAGE):
        orch.export_vba_parity_workbook_from_processors("input.xlsx", "output.xlsx")
