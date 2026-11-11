"""Backtest prospectivo do Conviction Score (eve_analyst.py).

Lê eve_data/conviction_log.csv (gravado a cada execução do Market Analyst) e
mede o retorno real do preço médio diário 7 e 30 dias depois de cada sinal.

Uso:
    python tools/backtest_conviction.py --out docs/BACKTEST_CONVICTION.md

Critério de aprovação (por horizonte): correlação de postos score×retorno
> 0,10 **e** quintil de maior score rende mais que o de menor score.
"""
from __future__ import annotations

import argparse
import logging
import sqlite3
import sys
from datetime import date, timedelta
from functools import cache
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from evetrade.core.signal_eval import forward_returns, summarize_signal  # noqa: E402
from evetrade.snapshots import load_signals  # noqa: E402

logger = logging.getLogger("backtest_conviction")
DATA = ROOT / "eve_data"
HORIZONS = (7, 30)
TOLERANCE_DAYS = 2   # aceita o dia útil mais próximo se a data exata faltar


def make_price_at(db: Path, region_id: int):
    con = sqlite3.connect(f"file:{db}?mode=ro", uri=True)

    @cache
    def price_at(type_id: int, day: str) -> float | None:
        d = date.fromisoformat(day)
        lo = (d - timedelta(days=TOLERANCE_DAYS)).isoformat()
        hi = (d + timedelta(days=TOLERANCE_DAYS)).isoformat()
        row = con.execute(
            "SELECT average FROM history WHERE region_id=? AND type_id=? "
            "AND date BETWEEN ? AND ? ORDER BY ABS(julianday(date) - julianday(?)) LIMIT 1",
            (region_id, type_id, lo, hi, day)).fetchone()
        return row[0] if row and row[0] else None
    return price_at


def render(results: dict, n_signals: int, days: list[str]) -> str:
    lines = ["# Backtest prospectivo — Conviction Score", "",
             f"Gerado em {date.today().isoformat()} por `tools/backtest_conviction.py`.", "",
             f"- Sinais registrados: {n_signals} em {len(days)} dia(s) "
             f"({days[0] if days else '—'} a {days[-1] if days else '—'})",
             "- Retorno = preço médio diário em D+H / preço médio em D − 1", ""]
    if all(r["n"] == 0 for r in results.values()):
        lines += ["**Ainda sem sinais maduros.** O primeiro resultado aparece 7 dias depois",
                  "da primeira execução do Market Analyst com esta versão; o de 30 dias,",
                  "um mês depois. Até lá o Conviction Score continua marcado como",
                  "experimental."]
        return "\n".join(lines) + "\n"
    lines += ["| Horizonte | N | Spearman score×retorno | Top − bottom quintil | "
              "Retorno médio por ação |", "|---|---|---|---|---|"]
    for h, r in results.items():
        if not r["n"]:
            lines.append(f"| {h} d | 0 | — | — | — |")
            continue
        acao = ", ".join(f"{k}: {m * 100:+.2f}% (n={n})"
                         for k, (m, n) in sorted(r["mean_ret_by_acao"].items()))
        rho = "—" if r["spearman"] is None else f"{r['spearman']:+.3f}"
        lines.append(f"| {h} d | {r['n']} | {rho} | {r['top_minus_bottom'] * 100:+.2f} pp | "
                     f"{acao} |")
    ok = [h for h, r in results.items()
          if r["n"] and (r["spearman"] or 0) > 0.10 and r["top_minus_bottom"] > 0]
    lines += ["", "**Veredito:** " + (
        f"sinal útil em h={ok}." if ok else
        "sem evidência de que score alto anteceda retorno maior.")]
    return "\n".join(lines) + "\n"


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--log", type=Path, default=DATA / "conviction_log.csv")
    ap.add_argument("--db", type=Path, default=DATA / "market_history.db")
    ap.add_argument("--out", type=Path)
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(message)s")

    signals = load_signals(args.log)
    days = sorted({s["date"] for s in signals})
    results = {}
    for h in HORIZONS:
        rets: list[dict] = []
        for rid in sorted({int(s["region_id"]) for s in signals}):
            price_at = make_price_at(args.db, rid)
            subset = [s for s in signals if int(s["region_id"]) == rid]
            rets += forward_returns(subset, price_at, h, date.today())
        results[h] = summarize_signal(rets)
    report = render(results, len(signals), days)
    logger.info("%s", report)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(report, encoding="utf-8")
        logger.info("relatório: %s", args.out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
