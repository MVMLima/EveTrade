"""
eve_utils.py — Helpers NaN-safe compartilhados entre eve_trade, eve_market e eve_analyst.
"""

import math


def safe_pct(v, fmt="+.1f") -> str:
    """Formata float como porcentagem; retorna 'N/D' se None, NaN ou Inf."""
    if v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))):
        return "N/D"
    return f"{v:{fmt}}%"


def safe_int(v, default=0) -> int:
    """Converte para int com segurança contra NaN/None/Inf."""
    if v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))):
        return default
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def safe_vol(v) -> str:
    """Formata volume/quantidade com segurança contra NaN."""
    n = safe_int(v, -1)
    return f"{n:,.0f}" if n >= 0 else "N/D"


def safe_dos(v) -> str:
    """Formata Dias de Oferta com segurança contra NaN."""
    if v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))):
        return "N/D"
    return f"{v:.1f}d"


def safe_float(v, fmt=".2f", suffix="") -> str:
    """Formata float com segurança contra NaN/None/Inf."""
    if v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))):
        return "N/D"
    return f"{v:{fmt}}{suffix}"


# Aliases com prefixo _ para compatibilidade com código legado
_safe_pct   = safe_pct
_safe_int   = safe_int
_safe_vol   = safe_vol
_safe_dos   = safe_dos
_safe_float = safe_float
