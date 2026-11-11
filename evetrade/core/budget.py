"""Orçamento: carteira gulosa de station trading dentro de um capital.

Módulo puro: sem Tkinter, sem rede, sem estado global. O histórico entra
por ``hist_provider(type_id) -> list``.
"""

import concurrent.futures
from collections.abc import Callable

from .fees import calc_fee_rates
from .history import history_stats

HistProvider = Callable[[int], list]

RISK_DAYS = {"Quick Flip (1d)": 1, "Balanced (3d)": 3, "Deep Stock (7d)": 7}


def allocate_budget(
    idx: dict,
    item_cache: dict,
    region_name: str,
    budget: float,
    broker_rel: int,
    accounting: int,
    risk_mode: str = "Balanced (3d)",
    min_vol: int = 1,
    hist_provider: HistProvider | None = None,
    progress_cb=None,
):
    """Greedy knapsack: maximiza lucro dado budget disponivel.

    idx: build_order_index(...) já sem blacklist. Retorna (basket, stats).
    """
    rb, rt = calc_fee_rates(broker_rel, accounting)
    days = RISK_DAYS.get(risk_mode, 3)

    candidates = []
    total = len(idx)
    chunk = max(1, total // 40)
    for i, (tid, entry) in enumerate(idx.items()):
        if progress_cb and i % chunk == 0:
            progress_cb((i, total, f"Triagem... {i:,}/{total:,}"))
        sells = entry["sell"]
        buys = entry["buy"]
        if not sells or not buys:
            continue
        ps = float(sells[0]["price"])
        pb = float(buys[0]["price"])
        if pb < 1000 or ps <= pb:
            continue
        gross = (ps - pb) / ps * 100
        if gross > 80:
            continue
        cost_unit = pb + pb * rb + ps * rb + ps * rt
        pi_net = ps - cost_unit
        if pi_net <= 0:
            continue
        candidates.append(
            {
                "type_id": tid,
                "ps": ps,
                "pb": pb,
                "cost_unit": cost_unit,
                "pi_net": pi_net,
                "efficiency": pi_net / cost_unit,
                "gross_margin": round(gross, 2),
            }
        )

    candidates.sort(key=lambda x: x["efficiency"], reverse=True)

    # 2E — busca historico so dos candidatos que podem caber no budget
    # estimativa conservadora: top candidatos por custo unitario
    budget_est = budget * 1.2  # folga de 20%
    hist_subset = []
    acc_cost = 0.0
    for c in candidates:
        if acc_cost >= budget_est:
            break
        hist_subset.append(c)
        acc_cost += c["cost_unit"]
    # garante ao menos os top-150 (limite de ordens)
    hist_subset = candidates[: max(len(hist_subset), 150)]

    if hist_provider is not None and hist_subset:
        if progress_cb:
            progress_cb(f"Buscando historico de {len(hist_subset)} candidatos...")

        def _fetch_hist(c):
            st = history_stats(hist_provider(c["type_id"]))
            c["avg_vol_90"] = round(st["avg_vol_90"], 1) if st else 0.0

        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as pool:
            list(pool.map(_fetch_hist, hist_subset))
        # candidatos fora do subset ficam com avg_vol_90=0 (descartados pelo min_vol)
        fetched = {c["type_id"] for c in hist_subset}
        for c in candidates:
            if c["type_id"] not in fetched:
                c["avg_vol_90"] = 0.0

    basket = []
    remaining = budget
    orders_used = 0
    total_spent = total_profit = 0.0

    for c in candidates:
        if remaining <= 0 or orders_used >= 150:
            break
        if c.get("avg_vol_90", 0) < min_vol:
            continue
        u_max = max(1, int(c["avg_vol_90"] * days))
        u_afford = int(remaining / c["cost_unit"])
        if u_afford <= 0:
            continue
        qty = min(u_max, u_afford)
        total_buy = qty * c["cost_unit"]
        total_prof = qty * c["pi_net"]
        info = item_cache.get(c["type_id"], {})
        basket.append(
            {
                "type_id": c["type_id"],
                "item": info.get("name", str(c["type_id"])),
                "ps": round(c["ps"], 2),
                "pb": round(c["pb"], 2),
                "qty": qty,
                "u_max": u_max,
                "cost_unit": round(c["cost_unit"], 2),
                "pi_net": round(c["pi_net"], 2),
                "total_buy": round(total_buy, 2),
                "total_profit": round(total_prof, 2),
                "net_margin": round(c["pi_net"] / c["cost_unit"] * 100, 2),
                "avg_vol_90": c.get("avg_vol_90", 0),
                "region": region_name,
            }
        )
        remaining -= total_buy
        total_spent += total_buy
        total_profit += total_prof
        orders_used += 1

    stats = {
        "n_items": len(basket),
        "budget": budget,
        "total_spent": round(total_spent, 2),
        "total_profit": round(total_profit, 2),
        "orders_used": orders_used,
        "utilization": round(total_spent / budget * 100, 1) if budget > 0 else 0,
        "limit_hit": orders_used >= 150,
    }
    return basket, stats
