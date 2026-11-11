"""Max Margem: maiores margens líquidas de station trading + EDP.

Módulo puro: sem Tkinter, sem rede, sem estado global. O histórico entra
por ``hist_provider(type_id) -> list`` (ESI/SQLite ficam com o chamador).
"""

import concurrent.futures
from collections.abc import Callable

from .fees import calc_fee_rates
from .history import history_stats

# Melhor compra precisa valer ao menos esta fração da melhor venda.
# Sem isso, ordens simbólicas (0,01 ISK) geram margem ~852% (skills 0) e
# dominam o ranking. 0.10 => margem bruta máxima de 90%.
MM_MIN_BUY_RATIO = 0.10

HistProvider = Callable[[int], list]


def margin_candidates(
    idx: dict,
    item_cache: dict,
    region_name: str,
    broker_rel: int = 0,
    accounting: int = 0,
    min_margin: float = 0.0,
    max_sell_price: float = 0.0,
    min_buy_ratio: float = MM_MIN_BUY_RATIO,
    progress_cb=None,
) -> list:
    """Itens com margem líquida positiva, ordenados por margem (desc)."""
    rb, rt = calc_fee_rates(broker_rel, accounting)
    candidates = []
    total = len(idx)
    chunk = max(1, total // 40)
    for i, (tid, entry) in enumerate(idx.items()):
        if progress_cb and i % chunk == 0:
            progress_cb((i, total, f"Calculando margens... {i:,}/{total:,}"))
        sells, buys = entry["sell"], entry["buy"]
        if not sells or not buys:
            continue
        ps = float(sells[0]["price"])
        pb = float(buys[0]["price"])
        if ps <= pb or ps <= 0:
            continue
        if max_sell_price > 0 and ps > max_sell_price:
            continue
        if pb / ps < min_buy_ratio:
            continue
        cost = pb + pb * rb + ps * rb + ps * rt
        pi = ps - cost
        if pi <= 0:
            continue
        net_margin = pi / cost * 100
        if net_margin < min_margin:
            continue
        info = item_cache.get(tid, {})
        candidates.append(
            {
                "type_id": tid,
                "item": info.get("name", str(tid)),
                "ps": round(ps, 2),
                "pb": round(pb, 2),
                "gross_margin": round((ps - pb) / ps * 100, 2),
                "net_margin": round(net_margin, 2),
                "profit_unit": round(pi, 2),
                "sell_vol": sum(int(o.get("volume_remain", 0)) for o in sells),
                "buy_vol": sum(int(o.get("volume_remain", 0)) for o in buys),
                "edp": None,
                "avg_vol_90": None,
                "region": region_name,
            }
        )
    candidates.sort(key=lambda x: x["net_margin"], reverse=True)
    return candidates


def select_by_volume(
    candidates: list,
    hist_provider: HistProvider | None,
    min_vol: float = 1,
    top_n: int = 200,
    progress_cb=None,
) -> list:
    """Aplica min_vol (média diária 90d) ANTES do corte em top_n.

    Percorre os candidatos em lotes, por margem decrescente, até juntar top_n
    aprovados. Sem hist_provider não há histórico: só devolve algo se min_vol <= 0.
    Resultado ordenado por EDP (lucro estimado por dia).
    """
    results: list = []
    if hist_provider is not None:

        def _stats(r):
            return history_stats(hist_provider(r["type_id"]))

        batch = max(top_n, 50)
        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as pool:
            for start in range(0, len(candidates), batch):
                if len(results) >= top_n:
                    break
                chunk = candidates[start : start + batch]
                if progress_cb:
                    progress_cb(
                        f"Historico/volume: {start + len(chunk):,}/"
                        f"{len(candidates):,} candidatos, "
                        f"{len(results)} aprovados..."
                    )
                for r, st in zip(chunk, pool.map(_stats, chunk), strict=True):
                    if not st or st["avg_vol_90"] < min_vol:
                        continue
                    r["avg_vol_90"] = round(st["avg_vol_90"], 1)
                    r["edp"] = round(r["profit_unit"] * st["avg_vol_90"], 2)
                    results.append(r)
                    if len(results) >= top_n:
                        break
    elif min_vol <= 0:
        results = candidates[:top_n]

    results.sort(key=lambda x: x["edp"] or 0, reverse=True)
    return results
