"""Testes da extração de metadados das abas (core.sheet_inspector)."""

from __future__ import annotations

from openpyxl import Workbook

from core.sheet_inspector import (
    inspect_sheet,
    inspect_workbook,
    workbook_summary,
)


def test_inspect_sheet_counts(template):
    md = inspect_sheet(template["Dados"])
    assert md.name == "Dados"
    assert md.formula_count == 1  # apenas B4
    assert md.merged_cells_count == 1
    assert md.data_validation_count == 1
    assert md.conditional_format_count == 1
    assert md.chart_count == 1
    assert md.hidden is False


def test_inspect_sheet_counts_calc_formulas(template):
    # 'Calc' tem 5 fórmulas (A1..A5; A3 é uma fórmula que devolve texto).
    md = inspect_sheet(template["Calc"])
    assert md.formula_count == 5


def test_summary_line_lists_relevant_attributes(template):
    line = inspect_sheet(template["Dados"]).summary_line()
    assert "fórmula" in line
    assert "mesclada" in line


def test_summary_line_for_empty_sheet():
    wb = Workbook()
    md = inspect_sheet(wb.active)
    assert md.summary_line() == "(sem fórmulas)"


def test_inspect_workbook_lists_all_sheets(template):
    names = [m.name for m in inspect_workbook(template)]
    assert names == ["Dados", "Calc"]


def test_workbook_summary(template):
    summary = workbook_summary(template)
    assert summary.sheet_count == 2
    assert "Dados" in summary.sheet_names
