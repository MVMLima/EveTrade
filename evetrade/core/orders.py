"""Índice de ordens por type_id e utilitários de ordens.

Módulo puro: sem Tkinter, sem rede, sem estado global.
"""

from collections import defaultdict


def build_order_index(orders):
    # Normaliza is_buy_order str->bool UMA VEZ e ja pre-ordena as listas.
    # sells ASC por preco, buys DESC -- match_orders_granular nao precisa re-sort.
    raw = defaultdict(lambda: {"buy": [], "sell": []})
    for o in orders:
        is_buy = o["is_buy_order"]
        if isinstance(is_buy, str):
            is_buy = is_buy.strip().lower() == "true"
            o["is_buy_order"] = is_buy
        raw[int(o["type_id"])]["buy" if is_buy else "sell"].append(o)
    for entry in raw.values():
        entry["sell"].sort(key=lambda o: float(o["price"]))
        entry["buy"].sort(key=lambda o: float(o["price"]), reverse=True)
    return raw


def order_min_volume(order) -> int:
    """min_volume de uma ordem de compra (CSV traz string; ausente => 1)."""
    try:
        return max(1, int(order.get("min_volume") or 1))
    except (TypeError, ValueError):
        return 1


def filter_blacklist(idx: dict, blacklist: set) -> dict:
    """Remove type_ids na blacklist do indice de ordens. Retorna novo dict."""
    if not blacklist:
        return idx
    return {tid: v for tid, v in idx.items() if tid not in blacklist}
