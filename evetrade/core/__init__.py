"""Cálculos puros do EVE Trade Suite (ver evetrade/__init__.py)."""

from .budget import RISK_DAYS, allocate_budget
from .fees import calc_fee_rates
from .history import calc_lucro_dia, history_stats
from .industry import calculate_industry_cost
from .margin import (
    MM_MAX_ASK_VS_AVG,
    MM_MIN_BID_VS_AVG,
    MM_MIN_BUY_RATIO,
    margin_candidates,
    select_by_volume,
)
from .orders import build_order_index, filter_blacklist, order_min_volume
from .routes import (
    MAX_INT32,
    arbitrage_opportunities,
    execution_depth,
    find_routes,
    match_orders_granular,
)

__all__ = [
    "MAX_INT32",
    "MM_MAX_ASK_VS_AVG",
    "MM_MIN_BID_VS_AVG",
    "MM_MIN_BUY_RATIO",
    "RISK_DAYS",
    "allocate_budget",
    "arbitrage_opportunities",
    "build_order_index",
    "calc_fee_rates",
    "calc_lucro_dia",
    "calculate_industry_cost",
    "execution_depth",
    "filter_blacklist",
    "find_routes",
    "history_stats",
    "margin_candidates",
    "match_orders_granular",
    "order_min_volume",
    "select_by_volume",
]
