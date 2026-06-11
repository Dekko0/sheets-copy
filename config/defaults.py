"""Valores padrão e constantes de configuração da aplicação.

Centraliza o caminho do modelo embutido, as abas pré-selecionadas e as
constantes compartilhadas entre o núcleo e a interface, evitando "números
mágicos" espalhados pelo código.
"""

from __future__ import annotations

from pathlib import Path

# --------------------------------------------------------------------------- #
# Identificação da aplicação
# --------------------------------------------------------------------------- #
APP_NAME: str = "Sistema de Cópia de Abas Excel"
APP_VERSION: str = "1.0.0"

# --------------------------------------------------------------------------- #
# Modelo (template) padrão embutido
# --------------------------------------------------------------------------- #
DEFAULT_TEMPLATE_PATH: Path = Path(__file__).parent.parent / "assets" / "RD_MODELO.xlsx"
DEFAULT_TEMPLATE_NAME: str = "RD_MODELO (padrão)"

# Abas pré-selecionadas quando o modelo padrão é usado.
# Quando o usuário sobe um modelo customizado, nenhuma aba vem pré-selecionada.
DEFAULT_SELECTED_SHEETS: list[str] = [
    "Tabela de Transferência",
    "UC_DADOS",
    "TABELAS_RESUMO_2",
    "TABELAS_RESUMO",
    "CONSUMO_EQP",
    "Análise Econômica",
    "EQUIPAMENTOS_EFICIENTES (BDD)",
]

# --------------------------------------------------------------------------- #
# Extensões aceitas
# --------------------------------------------------------------------------- #
ALLOWED_EXTENSIONS: tuple[str, ...] = (".xlsx", ".xlsm")

# --------------------------------------------------------------------------- #
# Modos de cópia (quando a aba já existe na destino)
# --------------------------------------------------------------------------- #
COPY_MODE_REPLACE: str = "replace"  # limpa a aba destino antes de colar (default)
COPY_MODE_OVERLAY: str = "overlay"  # sobrescreve apenas as células do modelo

COPY_MODES: dict[str, str] = {
    COPY_MODE_REPLACE: "Substituir completamente",
    COPY_MODE_OVERLAY: "Sobrescrever apenas células do modelo",
}


def default_template_exists() -> bool:
    """Indica se o modelo padrão (RD_MODELO.xlsx) está presente em ``assets/``.

    Returns:
        ``True`` se o arquivo existir e for um arquivo regular, senão ``False``.
    """
    return DEFAULT_TEMPLATE_PATH.is_file()
