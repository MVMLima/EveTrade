"""Registro de sinais do Market Analyst para avaliação futura (P1.3).

O Conviction Score depende do livro de ofertas do momento, que não fica
guardado no histórico diário. Por isso não dá para testá-lo no passado: cada
execução do Analyst grava aqui os sinais do dia, e
``tools/backtest_conviction.py`` mede o retorno real quando eles amadurecem.
"""
from __future__ import annotations

import csv
import logging
from collections.abc import Iterable
from datetime import date
from pathlib import Path

logger = logging.getLogger(__name__)

FIELDS = ["date", "region_id", "type_id", "score", "acao", "n_fontes", "sell_min"]


def append_signals(path: Path, region_id: int, rows: Iterable[dict],
                   day: date | None = None) -> int:
    """Acrescenta os sinais do dia (substitui um registro anterior do mesmo dia/região)."""
    day_s = (day or date.today()).isoformat()
    existing: list[dict] = []
    if path.exists():
        with path.open(newline="", encoding="utf-8") as f:
            existing = [r for r in csv.DictReader(f)
                        if not (r["date"] == day_s and r["region_id"] == str(region_id))]
    new = [{"date": day_s, "region_id": region_id, "type_id": int(r["type_id"]),
            "score": r.get("conviction_score", ""), "acao": r.get("acao", ""),
            "n_fontes": r.get("n_fontes", ""), "sell_min": r.get("sell_min", "")}
           for r in rows if r.get("conviction_score") not in (None, "")]
    tmp = path.with_suffix(".tmp")
    with tmp.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(existing)
        w.writerows(new)
    tmp.replace(path)
    logger.info("conviction log: %d sinais de %s gravados em %s", len(new), day_s, path)
    return len(new)


def load_signals(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open(newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))
