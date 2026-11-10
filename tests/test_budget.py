"""analyze_budget — alocação gulosa respeitando orçamento e giro."""
import pytest

from conftest import JITA, et, history, order


def _orders():
    return [
        # item 1: pb=10k ps=13k (bruta 23%) — eficiente
        order(1, 13_000, 1000, buy=False), order(1, 10_000, 1000, buy=True),
        # item 2: pb=100k ps=125k (bruta 20%) — menos eficiente (custo 116.125)
        order(2, 125_000, 1000, buy=False), order(2, 100_000, 1000, buy=True),
        # item 3: pb < 1000 ISK — descartado pela triagem
        order(3, 900, 1000, buy=False), order(3, 500, 1000, buy=True),
        # item 4: margem bruta > 80% — descartado (suspeito)
        order(4, 100_000, 10, buy=False), order(4, 5_000, 10, buy=True),
    ]


def test_budget_respected_and_filters_applied(offline):
    offline.update({1: history(10), 2: history(10), 3: history(10), 4: history(10)})
    basket, stats = et.analyze_budget(_orders(), JITA, {}, budget=1_000_000,
                                      broker_rel=0, accounting=0,
                                      risk_mode="Balanced (3d)")
    tids = [b["type_id"] for b in basket]
    assert 3 not in tids and 4 not in tids
    assert stats["total_spent"] <= 1_000_000
    for b in basket:
        assert b["qty"] <= b["u_max"] == int(10 * 3)        # giro 10/dia * 3 dias
        assert b["total_buy"] == pytest.approx(b["qty"] * b["cost_unit"], rel=1e-6)


def test_most_efficient_item_first(offline):
    offline.update({1: history(10), 2: history(10)})
    basket, _ = et.analyze_budget(_orders(), JITA, {}, budget=10_000_000,
                                  broker_rel=0, accounting=0)
    assert basket[0]["type_id"] == 1
    # item 1, skills 0: custo = 10000 + 300 + 390 + 975 = 11665 ; lucro = 1335
    assert basket[0]["cost_unit"] == pytest.approx(11_665)
    assert basket[0]["pi_net"] == pytest.approx(1_335)


def test_min_vol_excludes_illiquid(offline):
    offline.update({1: history(0.5), 2: history(10)})
    basket, _ = et.analyze_budget(_orders(), JITA, {}, budget=10_000_000,
                                  broker_rel=0, accounting=0, min_vol=1)
    assert [b["type_id"] for b in basket] == [2]
