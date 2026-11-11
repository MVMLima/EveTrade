"""Backtest walk-forward da aba Previsão (eve_trade.predict_prices).

Pergunta: a previsão erra menos que o baseline ingênuo "preço futuro = preço
de hoje" (passeio aleatório)? Só dados até a data de corte entram no modelo.

Uso:
    python tools/backtest_forecast.py                # Jita, 200 itens
    python tools/backtest_forecast.py --items 50 --out docs/BACKTEST_PREVISAO.md

Saída: tabela no terminal e (com --out) um relatório Markdown.
Determinístico: sem aleatoriedade; a amostra depende só do banco local.
"""
from __future__ import annotations

import argparse
import logging
import sqlite3
import statistics
import sys
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import eve_trade as et  # noqa: E402

logger = logging.getLogger("backtest_forecast")
HORIZONS = (1, 7, 14)


@dataclass(frozen=True)
class Config:
    region_id: int = 10000002
    n_items: int = 200
    min_days: int = 250
    n_cutoffs: int = 12
    step_days: int = 14


@dataclass
class Stats:
    err_model: list = field(default_factory=list)
    err_naive: list = field(default_factory=list)
    model_better: int = 0
    dir_hits: int = 0
    dir_total: int = 0
    band_hits: int = 0
    n: int = 0


def load_sample(db: Path, cfg: Config) -> dict[int, list[dict]]:
    """Itens mais negociados (ISK/dia) com histórico longo."""
    con = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
    rows = con.execute(
        """SELECT type_id FROM history WHERE region_id = ?
           GROUP BY type_id HAVING COUNT(*) >= ?
           ORDER BY AVG(volume * average) DESC LIMIT ?""",
        (cfg.region_id, cfg.min_days, cfg.n_items)).fetchall()
    sample = {}
    for (tid,) in rows:
        recs = con.execute(
            "SELECT date, average, highest, lowest, volume FROM history "
            "WHERE region_id = ? AND type_id = ? ORDER BY date", (cfg.region_id, tid)).fetchall()
        sample[tid] = [{"date": d, "average": a, "highest": h, "lowest": lo, "volume": v}
                       for d, a, h, lo, v in recs if a and a > 0]
    con.close()
    return sample


def evaluate(sample: dict[int, list[dict]], cfg: Config) -> dict[int, Stats]:
    stats = {h: Stats() for h in HORIZONS}
    hmax = max(HORIZONS)
    for hist in sample.values():
        by_date = {r["date"]: r["average"] for r in hist}
        n = len(hist)
        for k in range(cfg.n_cutoffs):
            cut = n - hmax - 1 - k * cfg.step_days
            if cut < 60:
                break
            train = hist[:cut + 1]
            fc = et.predict_prices(train, horizon=hmax)
            if not fc or not fc.get("fcast_central"):
                continue
            last = train[-1]["average"]
            for h in HORIZONS:
                target = fc["dates_fcast"][h - 1]
                actual = by_date.get(target)
                if not actual:
                    continue
                pred = fc["fcast_central"][h - 1]
                s = stats[h]
                em, en = abs(pred - actual) / actual, abs(last - actual) / actual
                s.err_model.append(em)
                s.err_naive.append(en)
                s.model_better += em < en
                if actual != last and pred != last:
                    s.dir_total += 1
                    s.dir_hits += (pred > last) == (actual > last)
                lo, up = fc["fcast_lower"][h - 1], fc["fcast_upper"][h - 1]
                s.band_hits += lo <= actual <= up
                s.n += 1
    return stats


def summarize(stats: dict[int, Stats]) -> list[dict]:
    out = []
    for h, s in stats.items():
        if not s.n:
            continue
        mae_m, mae_n = statistics.mean(s.err_model), statistics.mean(s.err_naive)
        out.append({
            "h": h, "n": s.n,
            "mae_model": mae_m * 100, "mae_naive": mae_n * 100,
            "med_model": statistics.median(s.err_model) * 100,
            "med_naive": statistics.median(s.err_naive) * 100,
            "skill": (1 - mae_m / mae_n) * 100 if mae_n else 0.0,
            "win_rate": s.model_better / s.n * 100,
            "dir_acc": s.dir_hits / s.dir_total * 100 if s.dir_total else float("nan"),
            "band_cov": s.band_hits / s.n * 100,
        })
    return out


def verdict(rows: list[dict]) -> str:
    beats = [r for r in rows if r["skill"] > 5 and r["dir_acc"] > 55]
    if len(beats) == len(rows):
        return "APROVADO: bate o baseline ingênuo em todos os horizontes."
    if beats:
        return f"PARCIAL: bate o baseline só em h={[r['h'] for r in beats]}."
    return "REPROVADO: não bate o baseline ingênuo (preço futuro = preço de hoje)."


def render(rows: list[dict], cfg: Config, n_items: int) -> str:
    lines = [
        "# Backtest — aba Previsão (`predict_prices`)", "",
        f"Gerado em {date.today().isoformat()} por `tools/backtest_forecast.py`.", "",
        f"- Região: {cfg.region_id} · itens: {n_items} mais negociados em ISK/dia "
        f"com ≥ {cfg.min_days} dias de histórico",
        f"- Walk-forward: {cfg.n_cutoffs} cortes a cada {cfg.step_days} dias; o modelo só vê "
        "dados até o corte",
        "- Baseline: passeio aleatório (preço futuro = último preço médio diário)",
        "- Critério de aprovação: skill > 5% **e** acerto de direção > 55%", "",
        "| Horizonte | N | MAE modelo | MAE baseline | Skill | Modelo melhor | "
        "Acerto direção | Cobertura banda ±1σ |",
        "|---|---|---|---|---|---|---|---|",
    ]
    for r in rows:
        lines.append(
            f"| {r['h']} d | {r['n']} | {r['mae_model']:.2f}% | {r['mae_naive']:.2f}% | "
            f"{r['skill']:+.1f}% | {r['win_rate']:.1f}% | {r['dir_acc']:.1f}% | "
            f"{r['band_cov']:.1f}% |")
    lines += ["", f"**Veredito:** {verdict(rows)}", "",
              "Skill = 1 − MAE_modelo / MAE_baseline (positivo = modelo melhor). "
              "Acerto de direção de 50% equivale a cara-ou-coroa. "
              "Uma banda ±1σ bem calibrada cobriria ~68%."]
    return "\n".join(lines) + "\n"


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--items", type=int, default=Config.n_items)
    ap.add_argument("--out", type=Path)
    args = ap.parse_args()
    logger.setLevel(logging.INFO)   # eve_trade fixa o root em WARNING
    cfg = Config(n_items=args.items)
    sample = load_sample(Path(et.HIST_DB), cfg)
    logger.info("amostra: %d itens", len(sample))
    rows = summarize(evaluate(sample, cfg))
    report = render(rows, cfg, len(sample))
    logger.info("\n%s", report)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(report, encoding="utf-8")
        logger.info("relatório: %s", args.out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
