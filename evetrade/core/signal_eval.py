"""Avaliação prospectiva de sinais (Conviction Score / COMPRAR-VENDER).

Dado um registro de sinais com data e o histórico diário de preços, mede o
retorno real H dias depois e compara grupos. Módulo puro: os preços entram
por ``price_at(type_id, date_iso) -> float | None``.
"""
from __future__ import annotations

import statistics
from collections.abc import Callable, Iterable
from datetime import date, timedelta

PriceAt = Callable[[int, str], "float | None"]


def _ranks(xs: list[float]) -> list[float]:
    order = sorted(range(len(xs)), key=xs.__getitem__)
    ranks = [0.0] * len(xs)
    i = 0
    while i < len(order):
        j = i
        while j + 1 < len(order) and xs[order[j + 1]] == xs[order[i]]:
            j += 1
        for k in range(i, j + 1):
            ranks[order[k]] = (i + j) / 2 + 1
        i = j + 1
    return ranks


def spearman(xs: list[float], ys: list[float]) -> float | None:
    """Correlação de postos; None se houver menos de 3 pares ou variância zero."""
    if len(xs) < 3:
        return None
    rx, ry = _ranks(xs), _ranks(ys)
    try:
        return statistics.correlation(rx, ry)
    except statistics.StatisticsError:
        return None


def forward_returns(signals: Iterable[dict], price_at: PriceAt, horizon: int,
                    today: date) -> list[dict]:
    """Retorno em `horizon` dias de cada sinal maduro (data + horizon <= today).

    signals: dicts com date (YYYY-MM-DD), type_id, score, acao.
    """
    out = []
    for s in signals:
        d0 = date.fromisoformat(s["date"])
        d1 = d0 + timedelta(days=horizon)
        if d1 > today:
            continue
        p0 = price_at(int(s["type_id"]), d0.isoformat())
        p1 = price_at(int(s["type_id"]), d1.isoformat())
        if not p0 or not p1:
            continue
        out.append({**s, "ret": p1 / p0 - 1.0})
    return out


def summarize_signal(rets: list[dict]) -> dict:
    """Correlação score×retorno, retorno médio por ação e top-bottom quintil."""
    if not rets:
        return {"n": 0}
    scores = [float(r["score"]) for r in rets]
    returns = [r["ret"] for r in rets]
    by_acao: dict[str, list[float]] = {}
    for r in rets:
        by_acao.setdefault(r["acao"], []).append(r["ret"])
    ranked = sorted(rets, key=lambda r: float(r["score"]))
    q = max(1, len(ranked) // 5)
    top = statistics.mean(r["ret"] for r in ranked[-q:])
    bottom = statistics.mean(r["ret"] for r in ranked[:q])
    return {
        "n": len(rets),
        "spearman": spearman(scores, returns),
        "mean_ret_by_acao": {k: (statistics.mean(v), len(v)) for k, v in by_acao.items()},
        "top_minus_bottom": top - bottom,
    }
