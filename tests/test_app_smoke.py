"""Smoke test do app Streamlit.

Usa o framework ``AppTest`` para executar ``app.py`` num harness (sem navegador)
e garantir que ele inicializa sem exceções — pega erros de import, de sintaxe e
de uso indevido da API do Streamlit. A interação com upload de arquivos não é
coberta aqui (limitação do ``AppTest``); o fluxo completo é validado pelos
testes do núcleo + o teste manual E2E do README.
"""

from __future__ import annotations

from streamlit.testing.v1 import AppTest


def test_app_boots_without_exception():
    at = AppTest.from_file("app.py", default_timeout=60).run()
    assert not at.exception
    # O cabeçalho é injetado via st.markdown (HTML); confere a marca do produto.
    assert any("Sheet Copier" in m.value for m in at.markdown)


# Exercita em runtime os componentes dos passos 2-6, que o teste acima não
# alcança (o app para no passo 1 por falta de uploads).
_COMPONENTS_SCRIPT = """
import streamlit as st
from core.sheet_inspector import SheetMetadata
from core.dependency_checker import DependencyReport, SheetDependency
from core.sheet_copier import CopyReport, SheetCopyResult
from ui.components import (
    app_header, badge, icon, render_copy_report,
    render_dependency_report, render_sheet_list, step_card,
)
from ui.styles import inject_styles

inject_styles()
app_header()

with step_card(1, "Teste", "description"):
    st.write("conteudo do card")

render_sheet_list([
    SheetMetadata(name="A", formula_count=5, merged_cells_count=2, chart_count=1),
    SheetMetadata(name="B", hidden=True),
])

render_dependency_report(DependencyReport(
    satisfied=[SheetDependency("X", "Dados", "B1", "=Dados!A1")],
    broken=[SheetDependency("X", "FATURAS", "B2", "=FATURAS!A1")],
))
render_dependency_report(DependencyReport(
    satisfied=[SheetDependency("X", "Dados", "B1", "=Dados!A1")], broken=[]
))

render_copy_report(CopyReport(
    results=[SheetCopyResult(
        sheet_name="A", created=False, mode="replace",
        cells_copied=10, charts=1, warnings=["grafico: falhou"],
    )],
    added_named_ranges=["ListaSimNao"],
))

st.markdown(badge("ok", "success") + icon("check_circle"), unsafe_allow_html=True)
"""


def test_components_render_without_exception():
    at = AppTest.from_string(_COMPONENTS_SCRIPT, default_timeout=60).run()
    assert not at.exception
