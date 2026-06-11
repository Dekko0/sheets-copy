"""Estilos globais da aplicação (paleta corporativa Google Blue + Material Icons).

Carrega as fontes (Material Icons, Google Sans, Roboto Mono) e injeta o CSS
global. Chamar :func:`inject_styles` uma única vez, logo após
``st.set_page_config``.
"""

from __future__ import annotations

import streamlit as st

_HEAD = """
<link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&family=Roboto+Mono:wght@400;500&display=swap" rel="stylesheet">
"""

_CSS = """
<style>
/* Paleta Material *dark*. Os papeis dos tokens neutros sao mantidos:
   --neutral-900 = texto primario, --neutral-100 = realce sutil. Em modo
   escuro esses valores se invertem (texto claro / superficie escura). */
:root {
    --primary:       #1A73E8;
    --primary-dark:  #1557B0;
    --primary-light: #8AB4F8;   /* azul claro para realces sobre fundo escuro */
    --success:       #81C995;
    --warning:       #FDD663;
    --danger:        #F28B82;
    --neutral-900:   #E8EAED;   /* texto primario (claro) */
    --neutral-700:   #9AA0A6;   /* texto secundario / muted */
    --neutral-300:   #3C4043;   /* bordas e divisores */
    --neutral-100:   #303134;   /* realce sutil / hover */
    --surface:       #2D2E31;   /* superficie elevada (cards, linhas) */
    --sidebar-bg:    #28292B;
}

/* ── Fonte global ── */
html, body, [class*="css"], .stMarkdown, .stButton, .stTextInput, .stSelectbox {
    font-family: 'Google Sans', 'Segoe UI', sans-serif;
}

/* Largura de conteúdo confortável mesmo em layout wide */
.block-container { max-width: 1080px; padding-top: 1.5rem; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background-color: var(--sidebar-bg);
    border-right: 1px solid var(--neutral-300);
}
[data-testid="stSidebar"] .stMarkdown p { font-size: 13px; color: var(--neutral-700); }
[data-testid="stSidebar"] .stCaption, [data-testid="stSidebar"] small { color: var(--neutral-700); }

/* ── Cabeçalho principal ── */
.app-header {
    display: flex; align-items: center; gap: 14px;
    padding: 8px 0 14px;
    border-bottom: 2px solid var(--primary);
    margin-bottom: 22px;
}
.app-header h1 { font-size: 22px; font-weight: 600; color: var(--neutral-900); margin: 0; line-height: 1.2; }
.app-header .subtitle { font-size: 13px; color: var(--neutral-700); }

/* ── Step cards (st.container(border=True)) ── */
[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 8px;
    margin-bottom: 6px;
    transition: box-shadow .15s ease;
}
[data-testid="stVerticalBlockBorderWrapper"]:hover {
    box-shadow: 0 1px 6px rgba(60,64,67,0.15);
}
.step-label {
    font-size: 11px; font-weight: 600; letter-spacing: 0.8px;
    text-transform: uppercase; color: var(--primary); margin-bottom: 2px;
}
.step-title {
    font-size: 15px; font-weight: 500; color: var(--neutral-900); margin-bottom: 6px;
}

/* ── Badges ── */
.badge {
    display: inline-flex; align-items: center; gap: 4px;
    padding: 2px 10px; border-radius: 12px; font-size: 12px; font-weight: 500;
}
.badge-success { background:#1E3A28; color:#81C995; }
.badge-warning { background:#3D3420; color:#FDD663; }
.badge-error   { background:#3D2422; color:#F28B82; }
.badge-info    { background:#1F3A5F; color:#8AB4F8; }
.badge-neutral { background:#303134; color:#9AA0A6; }

/* ── Linhas de aba ── */
.sheet-row {
    display: flex; align-items: center; gap: 12px;
    padding: 9px 12px; border-radius: 6px;
    border: 1px solid var(--neutral-300); margin-bottom: 6px; background: var(--surface);
}
.sheet-row:hover { background: var(--neutral-100); }
.sheet-name { font-size: 14px; font-weight: 500; color: var(--neutral-900); flex: 1; }
.sheet-meta { font-size: 12px; color: var(--neutral-700); }

/* ── Botão de download (verde sucesso) ── */
[data-testid="stDownloadButton"] button {
    background-color: #1E8E3E !important;
    color: #fff !important; border: none !important; border-radius: 4px;
    font-family: 'Google Sans', sans-serif; font-weight: 500;
}
[data-testid="stDownloadButton"] button:hover { background-color: #137333 !important; }

/* ── Alertas ── */
.stAlert { border-radius: 8px; border-left-width: 4px; }

/* ── Expander ── */
[data-testid="stExpander"] { border: 1px solid var(--neutral-300); border-radius: 8px; }

/* ── File uploader ── */
[data-testid="stFileUploaderDropzone"] {
    border: 2px dashed var(--neutral-300); border-radius: 8px; background: var(--neutral-100);
}

/* ── Divider ── */
hr { border-color: var(--neutral-300); margin: 16px 0; }

/* ── Código inline / monoespaçado ── */
code, kbd { font-family: 'Roboto Mono', monospace; font-size: 12px; }
</style>
"""


def inject_styles() -> None:
    """Carrega fontes e injeta o CSS global (chamar uma vez, no topo do app)."""
    st.markdown(_HEAD + _CSS, unsafe_allow_html=True)
