"""Componentes Streamlit reutilizáveis (visual corporativo + Material Icons).

Duas formas de ícone convivem por necessidade técnica:

* :func:`icon` / :func:`render_icon` — ``<span class="material-icons">`` para uso
  dentro de blocos HTML injetados (``st.markdown(..., unsafe_allow_html=True)``);
* parâmetro nativo ``icon=":material/..."`` — usado nos componentes que **não**
  aceitam HTML no rótulo (``st.success``, ``st.button``, ``st.download_button``…),
  diretamente no ``app.py``.
"""

from __future__ import annotations

import streamlit as st

from core.dependency_checker import DependencyReport
from core.sheet_copier import CopyReport
from core.sheet_inspector import SheetMetadata

PRIMARY = "#1A73E8"
MUTED = "#9AA0A6"


# --------------------------------------------------------------------------- #
# Ícones
# --------------------------------------------------------------------------- #
def icon(name: str, size: str = "18px", color: str = "inherit", style: str = "") -> str:
    """Retorna o HTML inline de um Google Material Icon.

    Para uso dentro de strings passadas a ``st.markdown(..., unsafe_allow_html=True)``.
    """
    return (
        f'<span class="material-icons" '
        f'style="font-size:{size};color:{color};vertical-align:middle;'
        f'margin-right:4px;{style}">{name}</span>'
    )


def render_icon(name: str, size: str = "18px", color: str = "inherit") -> None:
    """Renderiza um ícone Material isolado via ``st.markdown``."""
    st.markdown(icon(name, size, color), unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# Estruturas visuais
# --------------------------------------------------------------------------- #
def app_header() -> None:
    """Renderiza o cabeçalho principal da aplicação."""
    st.markdown(
        f"""
        <div class="app-header">
          <span class="material-icons" style="font-size:32px;color:{PRIMARY}">content_copy</span>
          <div>
            <h1>Sheet Copier</h1>
            <div class="subtitle">Cópia de abas Excel com preservação total de propriedades</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def step_card(step_number: int, title: str, icon_name: str):
    """Cria um *card* delimitado para um passo e devolve o container.

    Usa ``st.container(border=True)`` (um ``<div>`` HTML injetado não envolve
    widgets nativos do Streamlit). Renderiza o rótulo/título dentro do card e
    retorna o container — use com ``with step_card(...):`` para inserir os
    widgets do passo dentro dele.

    Args:
        step_number: Número do passo.
        title: Título do passo.
        icon_name: Nome do Material Icon exibido ao lado do título.

    Returns:
        O container Streamlit do card (utilizável como gerenciador de contexto).
    """
    card = st.container(border=True)
    card.markdown(
        f'<div class="step-label">Passo {step_number}</div>'
        f'<div class="step-title">{icon(icon_name, "20px", PRIMARY)}{title}</div>',
        unsafe_allow_html=True,
    )
    return card


def badge(text: str, kind: str = "neutral") -> str:
    """Retorna o HTML de um badge colorido (``success``/``warning``/``error``/``info``/``neutral``)."""
    return f'<span class="badge badge-{kind}">{text}</span>'


# --------------------------------------------------------------------------- #
# Listagem de abas
# --------------------------------------------------------------------------- #
def sheet_row_html(md: SheetMetadata) -> str:
    """Retorna o HTML de uma linha de aba com seus metadados."""
    parts: list[str] = []
    if md.formula_count:
        parts.append(f'{icon("functions", "14px", MUTED)}{md.formula_count} fórmulas')
    if md.data_validation_count:
        parts.append(f'{icon("rule", "14px", MUTED)}{md.data_validation_count} validações')
    if md.merged_cells_count:
        parts.append(f'{icon("merge", "14px", MUTED)}{md.merged_cells_count} mescladas')
    if md.conditional_format_count:
        parts.append(f'{icon("format_color_fill", "14px", MUTED)}{md.conditional_format_count} cond.')
    if md.chart_count:
        parts.append(f'{icon("bar_chart", "14px", MUTED)}{md.chart_count} gráficos')
    meta = " &nbsp;·&nbsp; ".join(parts) if parts else "sem metadados"
    hidden = f" {badge('oculta', 'neutral')}" if md.hidden else ""
    return (
        '<div class="sheet-row">'
        f'<span class="material-icons" style="color:{PRIMARY};font-size:18px">tab</span>'
        f'<span class="sheet-name">{md.name}{hidden}</span>'
        f'<span class="sheet-meta">{meta}</span>'
        "</div>"
    )


def render_sheet_list(metadata: list[SheetMetadata]) -> None:
    """Renderiza a listagem de abas do modelo (apenas exibição)."""
    st.markdown("".join(sheet_row_html(m) for m in metadata), unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# Relatórios
# --------------------------------------------------------------------------- #
def render_dependency_report(report: DependencyReport) -> bool:
    """Mostra as dependências e devolve se a cópia pode prosseguir.

    Returns:
        ``True`` se não há quebras, ou se o usuário marcou que entende o risco.
    """
    if not report.satisfied and not report.broken:
        st.info(
            "Nenhuma referência entre abas foi detectada nas abas selecionadas.",
            icon=":material/info:",
        )
        return True

    if report.broken:
        linhas = []
        for nome in report.broken_sheets:
            exemplos = [d for d in report.broken if d.referenced_sheet == nome][:2]
            ex = "; ".join(
                f"`{e.source_sheet}!{e.example_cell}` -> {e.example_formula}" for e in exemplos
            )
            linhas.append(f"- **{nome}** nao existe na destino - ex.: {ex}")
        st.warning(
            "**Dependencias quebradas:**\n\n" + "\n".join(linhas),
            icon=":material/warning:",
        )
        if report.satisfied:
            st.caption(f"{len(report.satisfied)} dependencia(s) atendida(s).")
        return st.checkbox("Entendo o risco e quero continuar mesmo assim", key="ack_risk")

    st.success(
        f"Todas as {len(report.satisfied)} dependencia(s) estao atendidas.",
        icon=":material/check_circle:",
    )
    return True


def render_copy_report(report: CopyReport) -> None:
    """Renderiza o resumo final da cópia com badges e detalhes por aba."""
    badges = " ".join(
        [
            badge(f"{report.sheets_copied} copiadas", "success"),
            badge(f"{report.sheets_created} criadas", "info"),
            badge(
                f"{report.total_warnings} avisos",
                "warning" if report.total_warnings else "neutral",
            ),
        ]
    )
    st.markdown(f'<div style="margin:6px 0 4px">{badges}</div>', unsafe_allow_html=True)

    if report.added_named_ranges:
        st.caption("Named ranges adicionados: " + ", ".join(report.added_named_ranges))

    with st.expander("Detalhes por aba", expanded=report.total_warnings > 0):
        for r in report.results:
            tag = badge("criada", "info") if r.created else badge("substituida", "neutral")
            st.markdown(
                f'{icon("tab", "16px", PRIMARY)}<b>{r.sheet_name}</b> &nbsp;{tag}<br>'
                f'<span class="sheet-meta">{r.cells_copied} células &nbsp;·&nbsp; '
                f'{r.merged_ranges} mescladas &nbsp;·&nbsp; {r.data_validations} validações '
                f'&nbsp;·&nbsp; {r.conditional_formats} cond. &nbsp;·&nbsp; {r.charts} gráficos '
                f'&nbsp;·&nbsp; {r.images} imagens</span>',
                unsafe_allow_html=True,
            )
            for w in r.warnings:
                st.caption(f"aviso: {w}")
