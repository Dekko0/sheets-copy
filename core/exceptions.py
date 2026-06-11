"""Exceções customizadas do domínio de cópia de abas.

Todas herdam de :class:`SheetCopierError`, permitindo que chamadores capturem
qualquer erro do domínio com um único ``except`` enquanto ainda distinguem
casos específicos quando necessário.
"""

from __future__ import annotations


class SheetCopierError(Exception):
    """Classe base para todos os erros do domínio."""


class InvalidWorkbookError(SheetCopierError):
    """O arquivo enviado não é um workbook Excel válido ou não pôde ser lido."""


class SheetNotFoundError(SheetCopierError):
    """Uma aba esperada não existe no workbook de origem."""


class DependencyBrokenError(SheetCopierError):
    """Há dependências quebradas e a operação foi abortada sem confirmação."""


class CopyOperationError(SheetCopierError):
    """Falha inesperada durante a cópia de uma aba."""
