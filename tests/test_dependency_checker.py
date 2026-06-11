"""Testes da análise de dependências (core.dependency_checker)."""

from __future__ import annotations

from openpyxl import Workbook

from core.dependency_checker import analyze_dependencies, extract_sheet_references


# --------------------------------------------------------------------------- #
# Extração de referências de abas em fórmulas
# --------------------------------------------------------------------------- #
def test_extract_unquoted_references():
    refs = extract_sheet_references("=FATURAS!A1 + Dados!B2")
    assert refs == {"FATURAS", "Dados"}


def test_extract_quoted_reference_with_space_and_accent():
    assert extract_sheet_references("='Análise Econômica'!B2") == {"Análise Econômica"}


def test_extract_ignores_text_literals():
    assert extract_sheet_references('="texto com ! interno"') == set()
    assert extract_sheet_references('=A1&"x!y"&FATURAS!B1') == {"FATURAS"}


def test_extract_handles_doubled_quote_in_name():
    assert extract_sheet_references("='O''Brien'!A1") == {"O'Brien"}


def test_extract_returns_empty_when_no_reference():
    assert extract_sheet_references("") == set()
    assert extract_sheet_references("=A1+B2*3") == set()


# --------------------------------------------------------------------------- #
# Análise contra a planilha destino
# --------------------------------------------------------------------------- #
def test_detects_broken_dependencies(template, destination):
    report = analyze_dependencies(template, destination, ["Calc"])
    assert report.has_broken
    assert "FATURAS" in report.broken_sheets
    assert "Análise Econômica" in report.broken_sheets


def test_satisfied_when_target_already_has_sheet(template, destination):
    # 'Calc' referencia 'Dados', que existe na destino.
    report = analyze_dependencies(template, destination, ["Calc"])
    satisfied = {dep.referenced_sheet for dep in report.satisfied}
    assert "Dados" in satisfied


def test_self_reference_is_ignored(template, destination):
    report = analyze_dependencies(template, destination, ["Calc"])
    referenced = {dep.referenced_sheet for dep in report.satisfied + report.broken}
    assert "Calc" not in referenced


def test_satisfied_when_referenced_sheet_also_selected(template):
    dest = Workbook()
    dest.active.title = "Solo"  # não tem 'Dados' nem 'Calc'

    # Copiando só 'Calc': 'Dados' está quebrada.
    broken_report = analyze_dependencies(template, dest, ["Calc"])
    assert "Dados" in broken_report.broken_sheets

    # Copiando 'Calc' + 'Dados': 'Dados' será criada → satisfeita.
    ok_report = analyze_dependencies(template, dest, ["Calc", "Dados"])
    satisfied = {dep.referenced_sheet for dep in ok_report.satisfied}
    assert "Dados" in satisfied
    assert "Dados" not in ok_report.broken_sheets


def test_dependencies_are_deduplicated_per_pair(template, destination):
    # 'FATURAS' é referenciada em A1 e A5; deve haver só uma entrada quebrada.
    report = analyze_dependencies(template, destination, ["Calc"])
    faturas = [dep for dep in report.broken if dep.referenced_sheet == "FATURAS"]
    assert len(faturas) == 1
    assert faturas[0].source_sheet == "Calc"
    assert faturas[0].example_formula  # guarda um exemplo de fórmula
