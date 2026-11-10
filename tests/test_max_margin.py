"""analyze_max_margin — margens com resposta calculada à mão."""
import pytest

from conftest import JITA, et, history, order


def test_fee_rates_by_skill():
    assert et._calc_fee_rates(0, 0) == pytest.approx((0.03, 0.075))
    # Broker Relations 5: 3% - 5*0.3% = 1.5%; Accounting 5: 7.5% * (1 - 0.55)
    assert et._calc_fee_rates(5, 5) == pytest.approx((0.015, 0.075 * 0.45))


def test_net_margin_known_value(offline):
    # ps=200, pb=100, skills 0: custo = 100 + 3 + 6 + 15 = 124; lucro = 76
    offline[1] = history(50)
    orders = [order(1, 200, 10, buy=False), order(1, 100, 10, buy=True)]
    rows = et.analyze_max_margin(orders, JITA, {}, min_vol=1)
    assert len(rows) == 1
    r = rows[0]
    assert r["profit_unit"] == pytest.approx(76.0)
    assert r["net_margin"] == pytest.approx(76 / 124 * 100, abs=0.01)
    assert r["gross_margin"] == pytest.approx(50.0)
    assert r["avg_vol_90"] == pytest.approx(50.0)
    assert r["edp"] == pytest.approx(76.0 * 50)


def test_symbolic_bid_is_rejected(offline):
    # Caso real da revisão: venda 550M contra compra de 0,03 ISK (margem ~852%)
    offline[2] = history(100)
    orders = [order(2, 550_000_000, 1, buy=False), order(2, 0.03, 100, buy=True)]
    assert et.analyze_max_margin(orders, JITA, {}, min_vol=1) == []


def test_min_buy_ratio_zero_disables_guard(offline):
    offline[2] = history(100)
    orders = [order(2, 550_000_000, 1, buy=False), order(2, 0.03, 100, buy=True)]
    rows = et.analyze_max_margin(orders, JITA, {}, min_vol=1, min_buy_ratio=0)
    assert len(rows) == 1


def test_volume_filter_runs_before_top_n(offline):
    # Dois itens de margem alta e volume baixo; um de margem menor e volume alto.
    # Com top_n=1 e min_vol=20 a resposta certa é o item de volume alto.
    # (O código antigo cortava em top_n antes do filtro e devolvia lista vazia.)
    offline.update({10: history(1), 11: history(2), 12: history(500)})
    orders = [
        order(10, 300, 5, buy=False), order(10, 100, 5, buy=True),
        order(11, 290, 5, buy=False), order(11, 100, 5, buy=True),
        order(12, 160, 5, buy=False), order(12, 100, 5, buy=True),
    ]
    rows = et.analyze_max_margin(orders, JITA, {}, min_vol=20, top_n=1)
    assert [r["type_id"] for r in rows] == [12]


def test_unprofitable_after_fees_is_excluded(offline):
    # ps=110, pb=100, skills 0: custo = 100+3+3.3+8.25 = 114.55 > 110
    offline[3] = history(100)
    orders = [order(3, 110, 5, buy=False), order(3, 100, 5, buy=True)]
    assert et.analyze_max_margin(orders, JITA, {}, min_vol=1) == []
