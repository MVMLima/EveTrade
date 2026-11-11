"""
eve_market.py — API pública de análise de mercado do EVE Trade Suite.

Ponto de entrada para consumidores externos (eve_analyst.py, scripts,
notebooks) que precisam das funções de análise sem depender da UI do
eve_trade.py.

Duas camadas:
  * cálculo puro (sem Tkinter, sem rede) — re-exportado de ``evetrade.core``.
    Use quando você já tem as ordens e o histórico em mãos (notebooks, testes);
  * funções ``analyze_*`` abaixo — também buscam nomes, histórico (SQLite/ESI)
    e ordens; ainda delegam a eve_trade até a camada de dados ser extraída.

Uso:
    import eve_market as em
    resultados = em.analyze_max_margin(orders, region, item_cache)
    # ou, sem I/O:
    idx = em.build_order_index(orders)
    cands = em.margin_candidates(idx, item_cache, region)
"""

import eve_trade as _et
from evetrade.core import (  # noqa: F401 — API pública re-exportada
    MM_MIN_BUY_RATIO,
    RISK_DAYS,
    allocate_budget,
    arbitrage_opportunities,
    build_order_index,
    calc_fee_rates,
    calc_lucro_dia,
    calculate_industry_cost,
    filter_blacklist,
    find_routes,
    history_stats,
    margin_candidates,
    match_orders_granular,
    order_min_volume,
    select_by_volume,
)

# ── Análises de mercado ────────────────────────────────────────────────────────

def analyze_max_margin(orders, region_name, item_cache,
                       broker_rel=0, accounting=0,
                       min_margin=0.0, max_sell_price=0.0,
                       min_vol=1, top_n=200,
                       blacklist=None, progress_cb=None,
                       min_buy_ratio=_et.MM_MIN_BUY_RATIO):
    """Maiores margens líquidas + EDP. Veja eve_trade.analyze_max_margin."""
    return _et.analyze_max_margin(
        orders, region_name, item_cache,
        broker_rel=broker_rel, accounting=accounting,
        min_margin=min_margin, max_sell_price=max_sell_price,
        min_vol=min_vol, top_n=top_n,
        blacklist=blacklist, progress_cb=progress_cb,
        min_buy_ratio=min_buy_ratio,
    )


def analyze_supply_shortage(orders, region_name, item_cache,
                             max_dos=30.0, min_margin=0.0,
                             min_vol=5, top_n=200,
                             blacklist=None, progress_cb=None):
    """Itens em escassez de suprimento (DoS baixo). Veja eve_trade."""
    return _et.analyze_supply_shortage(
        orders, region_name, item_cache,
        max_dos=max_dos, min_margin=min_margin,
        min_vol=min_vol, top_n=top_n,
        blacklist=blacklist, progress_cb=progress_cb,
    )


def analyze_demand_supply(orders, region_name, item_cache,
                           min_margin=0.0, min_vol=5, top_n=200,
                           blacklist=None, progress_cb=None):
    """Desequilíbrio demanda/oferta (DSIF). Veja eve_trade."""
    return _et.analyze_demand_supply(
        orders, region_name, item_cache,
        min_margin=min_margin, min_vol=min_vol, top_n=top_n,
        blacklist=blacklist, progress_cb=progress_cb,
    )


def analyze_price_movement(region_name, item_cache,
                            min_change_pct=5.0, min_vol=5, top_n=200,
                            blacklist=None, progress_cb=None):
    """Variação de preço nos últimos 30 dias. Veja eve_trade."""
    return _et.analyze_price_movement(
        region_name, item_cache,
        min_change_pct=min_change_pct, min_vol=min_vol, top_n=top_n,
        blacklist=blacklist, progress_cb=progress_cb,
    )


def analyze_budget(orders, region_name, item_cache,
                   budget=1_000_000_000, broker_rel=0, accounting=0,
                   risk_mode="Balanced (3d)", top_n=50,
                   blacklist=None, progress_cb=None):
    """Alocação de capital por risco. Retorna (basket, stats). Veja eve_trade."""
    return _et.analyze_budget(
        orders, region_name, item_cache,
        budget=budget, broker_rel=broker_rel, accounting=accounting,
        risk_mode=risk_mode, top_n=top_n,
        blacklist=blacklist, progress_cb=progress_cb,
    )


def analyze_market_making(orders, region_name, item_cache,
                           broker_rel=0, accounting=0,
                           min_spread=1.0, min_vol=10, top_n=100,
                           blacklist=None, progress_cb=None):
    """Oportunidades de market making. Veja eve_trade."""
    return _et.analyze_market_making(
        orders, region_name, item_cache,
        broker_rel=broker_rel, accounting=accounting,
        min_spread=min_spread, min_vol=min_vol, top_n=top_n,
        blacklist=blacklist, progress_cb=progress_cb,
    )


# ── Utilitários de dados ───────────────────────────────────────────────────────

def load_orders(region_name: str) -> list:
    """Carrega ordens do CSV local (ou baixa da ESI se ausente)."""
    orders = _et.load_orders_from_csv(region_name)
    if not orders:
        region_id = _et.REGIONS.get(region_name, 10000002)
        orders = _et.download_region_orders(region_id, region_name)
    return orders or []


def warm_history(region_name: str, progress_cb=None) -> int:
    """Pré-aquece cache de histórico para a região. Retorna nº de itens."""
    region_id = _et.REGIONS.get(region_name, 10000002)
    return _et._warm_hist_cache(region_id, progress_cb)


# ── Constantes re-exportadas ───────────────────────────────────────────────────
REGIONS     = _et.REGIONS
HUB_REGIONS = getattr(_et, "HUB_REGIONS", {
    "Jita":    10000002,
    "Amarr":   10000043,
    "Dodixie": 10000032,
    "Hek":     10000042,
    "Rens":    10000030,
})
