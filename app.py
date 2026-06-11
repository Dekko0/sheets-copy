"""Entrypoint Streamlit — Sheet Copier.

Fluxo em 5 passos: escolher modelo -> enviar destino -> selecionar abas ->
analisar dependencias -> executar e baixar.

Os arquivos enviados sao mantidos apenas como *bytes* em ``st.session_state`` e
em cache; os workbooks sao recarregados a cada operacao para que a copia (que
modifica o destino in-place) nunca contamine o estado entre execucoes.

Execute com:  ``streamlit run app.py``  (ou ``py -m streamlit run app.py``)
"""

from __future__ import annotations

import hashlib
import logging
from datetime import datetime
from pathlib import Path

import streamlit as st

from config.defaults import (
    ALLOWED_EXTENSIONS,
    APP_VERSION,
    COPY_MODES,
    DEFAULT_SELECTED_SHEETS,
    DEFAULT_TEMPLATE_PATH,
    default_template_exists,
)
from core.dependency_checker import analyze_dependencies
from core.exceptions import SheetCopierError
from core.sheet_copier import copy_sheets
from core.sheet_inspector import inspect_workbook
from core.workbook_loader import load_workbook_from_bytes, save_workbook_to_bytes
from ui.components import (
    app_header,
    icon,
    render_copy_report,
    render_dependency_report,
    render_sheet_list,
    step_card,
)
from ui.styles import inject_styles

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
logger = logging.getLogger("app")

_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_XLSM_MIME = "application/vnd.ms-excel.sheet.macroEnabled.12"
_UPLOAD_TYPES = [ext.lstrip(".") for ext in ALLOWED_EXTENSIONS]
_PRIMARY = "#1A73E8"
_MUTED = "#9AA0A6"

st.set_page_config(
    page_title="Sheet Copier",
    page_icon=":material/content_copy:",
    layout="wide",
    initial_sidebar_state="expanded",
)
inject_styles()


# --------------------------------------------------------------------------- #
# Helpers com cache (recarregam o workbook apenas quando os bytes mudam)
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner=False)
def cached_inspect(data: bytes, filename: str):
    """Retorna ``(sheetnames, metadata)`` do modelo, em cache pelos bytes."""
    wb = load_workbook_from_bytes(data, filename=filename)
    return list(wb.sheetnames), inspect_workbook(wb)


@st.cache_data(show_spinner=False)
def cached_dest_sheetnames(data: bytes, filename: str) -> list[str]:
    """Retorna os nomes das abas da planilha destino, em cache pelos bytes."""
    return list(load_workbook_from_bytes(data, filename=filename).sheetnames)


@st.cache_data(show_spinner=False)
def cached_dependencies(t_data: bytes, t_name: str, d_data: bytes, d_name: str, selected: tuple):
    """Analisa dependencias (em cache por modelo + destino + selecao)."""
    tpl = load_workbook_from_bytes(t_data, filename=t_name)
    dst = load_workbook_from_bytes(d_data, filename=d_name)
    return analyze_dependencies(tpl, dst, list(selected))


def _short_hash(data: bytes) -> str:
    """Hash curto e estavel dos bytes (usado em chaves de widgets)."""
    return hashlib.md5(data, usedforsecurity=False).hexdigest()[:10]


def _mime_for(extension: str) -> str:
    return _XLSM_MIME if extension.lower() == ".xlsm" else _XLSX_MIME


# --------------------------------------------------------------------------- #
# Cabecalho + Sidebar
# --------------------------------------------------------------------------- #
app_header()

with st.sidebar:
    st.markdown(
        f'{icon("layers", "18px", _PRIMARY)}'
        f'<b style="font-size:14px;color:#E8EAED;">Sheet Copier</b>'
        f'<span style="font-size:12px;color:{_MUTED};"> &nbsp;v{APP_VERSION}</span>',
        unsafe_allow_html=True,
    )
    st.divider()
    model_slot = st.empty()  # preenchido apos o Passo 1
    st.divider()
    st.markdown(f'<b>{icon("help_outline", "18px", _MUTED)}Como usar</b>', unsafe_allow_html=True)
    st.caption("1. Selecione ou importe um modelo")
    st.caption("2. Importe a planilha destino")
    st.caption("3. Escolha as abas")
    st.caption("4. Analise as dependencias")
    st.caption("5. Execute e baixe o resultado")
    st.divider()
    st.caption("Os arquivos originais nunca sao alterados — tudo ocorre em memoria.")


# --------------------------------------------------------------------------- #
# Passo 1 — Modelo
# --------------------------------------------------------------------------- #
with step_card(1, "Modelo (template)", "description"):
    origem = st.radio(
        "Origem do modelo",
        ["Usar modelo padrao (RD_MODELO.xlsx)", "Subir outro modelo"],
        horizontal=True,
        label_visibility="collapsed",
    )

    template_bytes: bytes | None = None
    template_name: str | None = None
    is_default = False

    if origem.startswith("Usar"):
        if default_template_exists():
            template_bytes = DEFAULT_TEMPLATE_PATH.read_bytes()
            template_name = DEFAULT_TEMPLATE_PATH.name
            is_default = True
        else:
            st.error(
                f"Modelo padrao nao encontrado em `{DEFAULT_TEMPLATE_PATH}`. "
                "Coloque o RD_MODELO.xlsx na pasta `assets/` ou escolha "
                "**Subir outro modelo**.",
                icon=":material/error:",
            )
    else:
        upload = st.file_uploader(
            "Arquivo do modelo (.xlsx / .xlsm)", type=_UPLOAD_TYPES, key="tpl_upload"
        )
        if upload is not None:
            template_bytes = upload.getvalue()
            template_name = upload.name

    if not template_bytes or not template_name:
        st.stop()

    try:
        template_sheets, template_meta = cached_inspect(template_bytes, template_name)
    except SheetCopierError as exc:
        st.error(str(exc), icon=":material/error:")
        st.stop()

    st.success(
        f"Modelo ativo: **{template_name}** — {len(template_sheets)} abas · "
        f"{len(template_bytes) / 1024:.0f} KB",
        icon=":material/check_circle:",
    )

# Preenche o slot "Modelo ativo" na sidebar (agora que o nome e conhecido).
with model_slot.container():
    st.markdown(f'<b>{icon("description", "18px", _MUTED)}Modelo ativo</b>', unsafe_allow_html=True)
    st.caption(f"{template_name} · {len(template_sheets)} abas")


# --------------------------------------------------------------------------- #
# Passo 2 — Planilha destino
# --------------------------------------------------------------------------- #
with step_card(2, "Planilha destino", "table_chart"):
    dest_upload = st.file_uploader(
        "Arquivo destino (.xlsx / .xlsm)", type=_UPLOAD_TYPES, key="dest_upload"
    )
    if dest_upload is None:
        st.info("Envie a planilha destino para continuar.", icon=":material/upload_file:")
        st.stop()

    dest_bytes = dest_upload.getvalue()
    dest_name = dest_upload.name
    try:
        dest_sheets = cached_dest_sheetnames(dest_bytes, dest_name)
    except SheetCopierError as exc:
        st.error(str(exc), icon=":material/error:")
        st.stop()

    st.success(
        f"Destino: **{dest_name}** — {len(dest_sheets)} abas",
        icon=":material/check_circle:",
    )
    with st.expander("Abas existentes na destino"):
        st.write(", ".join(dest_sheets))


# --------------------------------------------------------------------------- #
# Passo 3 — Selecao de abas
# --------------------------------------------------------------------------- #
with step_card(3, "Selecione as abas para copiar", "layers"):
    render_sheet_list(template_meta)

    preselected = (
        [s for s in DEFAULT_SELECTED_SHEETS if s in template_sheets] if is_default else []
    )
    selected = st.multiselect(
        "Abas a copiar",
        options=template_sheets,
        default=preselected,
        key=f"sel_{_short_hash(template_bytes)}",
    )
    if not selected:
        st.info("Selecione ao menos uma aba para continuar.", icon=":material/info:")
        st.stop()

    to_create = [s for s in selected if s not in dest_sheets]
    if to_create:
        st.caption("Serao criadas na destino: " + ", ".join(to_create))


# --------------------------------------------------------------------------- #
# Passo 4 — Dependencias
# --------------------------------------------------------------------------- #
with step_card(4, "Analise de dependencias", "account_tree"):
    dep_report = cached_dependencies(
        template_bytes, template_name, dest_bytes, dest_name, tuple(selected)
    )
    can_proceed = render_dependency_report(dep_report)


# --------------------------------------------------------------------------- #
# Modo de copia + execucao
# --------------------------------------------------------------------------- #
with step_card(5, "Executar copia", "sync"):
    mode_label = st.selectbox(
        "Modo de copia (quando a aba ja existe na destino)",
        options=list(COPY_MODES.values()),
        index=0,
        help="Substituir: recria a aba do zero. Sobrescrever: mantem celulas do "
        "destino fora do intervalo do modelo.",
    )
    mode = next(key for key, label in COPY_MODES.items() if label == mode_label)

    executar = st.button(
        "Executar copia",
        type="primary",
        icon=":material/sync:",
        disabled=not can_proceed,
    )

if executar:
    try:
        with st.status("Executando copia...", expanded=True) as status:
            bar = st.progress(0.0)

            def _progress(current: int, total: int, name: str) -> None:
                bar.progress(current / total)
                st.write(f":material/check_circle: **{name}** ({current}/{total})")

            template_wb = load_workbook_from_bytes(template_bytes, filename=template_name)
            dest_wb = load_workbook_from_bytes(dest_bytes, filename=dest_name)
            report = copy_sheets(
                template_wb,
                dest_wb,
                selected,
                mode=mode,
                copy_named_ranges=True,
                progress_callback=_progress,
            )
            output = save_workbook_to_bytes(dest_wb).getvalue()
            status.update(label="Copia concluida", state="complete")

        extension = Path(dest_name).suffix or ".xlsx"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.session_state["result"] = {
            "report": report,
            "data": output,
            "name": f"{Path(dest_name).stem}_atualizado_{timestamp}{extension}",
            "mime": _mime_for(extension),
        }
    except SheetCopierError as exc:
        st.error(f"Erro na copia: {exc}", icon=":material/error:")
    except Exception as exc:  # noqa: BLE001
        st.error("Ocorreu um erro inesperado durante a copia.", icon=":material/error:")
        with st.expander("Detalhes tecnicos"):
            st.exception(exc)


# --------------------------------------------------------------------------- #
# Passo 6 — Resultado
# --------------------------------------------------------------------------- #
result = st.session_state.get("result")
if result:
    with step_card(6, "Resultado", "download"):
        render_copy_report(result["report"])
        st.download_button(
            "Baixar planilha resultante",
            data=result["data"],
            file_name=result["name"],
            mime=result["mime"],
            icon=":material/download:",
        )
