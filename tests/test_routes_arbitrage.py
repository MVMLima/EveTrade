"""find_routes e match_orders_granular — lucro, volume e min_volume."""
import pytest

from conftest import et, order


def test_find_routes_known_profit():
    src = [order(1, 100, 10, buy=False)]
    dst = [order(1, 150, 4, buy=True)]
    routes = et.find_routes(src, dst, "A", "B", min_roi=10,
                            item_cache={1: {"name": "X", "volume": 2.5}})
    assert len(routes) == 1
    r = routes[0]
    assert r["profit_unit"] == 50
    assert r["roi"] == pytest.approx(50.0)
    assert r["vol_realizavel"] == 4          # limitado pela compra no destino
    assert r["lucro_total"] == 200
    assert r["capital"] == 1000              # compra todo o lote de origem
    assert r["total_volume"] == pytest.approx(10.0)


def test_find_routes_respects_buy_min_volume():
    src = [order(1, 100, 10, buy=False)]
    dst = [order(1, 150, 400, buy=True, min_volume=20)]   # exige >= 20 un
    assert et.find_routes(src, dst, "A", "B", min_roi=10) == []


def test_find_routes_min_lucro_total_filters_tiny_trades():
    # Caso real (Festival Launcher): compra 0,02, vende 2,06, 1 unidade
    src = [order(1, 0.02, 1, buy=False)]
    dst = [order(1, 2.06, 743, buy=True)]
    assert len(et.find_routes(src, dst, "A", "B", min_roi=10)) == 1
    assert et.find_routes(src, dst, "A", "B", min_roi=10,
                          min_lucro_total=1_000_000) == []


def _book(sells, buys):
    idx = et.build_order_index(
        [order(1, p, v, buy=False) for p, v in sells] +
        [order(1, p, v, buy=True, min_volume=m) for p, v, m in buys])
    return idx[1]["sell"], idx[1]["buy"]


def test_granular_walks_the_book_without_fees():
    sells, buys = _book(sells=[(100, 5), (110, 5)], buys=[(150, 3, 1), (120, 10, 1)])
    trades = et.match_orders_granular(sells, buys, 1.0, 0, 0, 0, 0)
    assert [(t["qty"], t["price_orig"], t["price_dest"]) for t in trades] == [
        (3, 100, 150), (2, 100, 120), (5, 110, 120)]
    assert sum(t["lucro_liq_total"] for t in trades) == pytest.approx(50*3 + 20*2 + 10*5)


def test_granular_applies_fees_and_freight():
    # custo = 100*1.025 + 2 m3 * 10 = 122.5 ; receita = 150*(1-0.075) = 138.75
    sells, buys = _book(sells=[(100, 1)], buys=[(150, 1, 1)])
    t, = et.match_orders_granular(sells, buys, 2.0, 2.5, 3.5, 4.0, 10.0)
    assert t["custo_unit"] == pytest.approx(122.5)
    assert t["lucro_liq_unit"] == pytest.approx(138.75 - 122.5)


def test_granular_stops_when_unprofitable():
    sells, buys = _book(sells=[(100, 5)], buys=[(100, 5, 1)])
    assert et.match_orders_granular(sells, buys, 1.0, 0, 0, 0, 0) == []


def test_granular_skips_unreachable_min_volume():
    # A melhor compra exige 50 un, mas só há 5 à venda: usa a segunda compra
    sells, buys = _book(sells=[(100, 5)], buys=[(200, 100, 50), (150, 10, 1)])
    trades = et.match_orders_granular(sells, buys, 1.0, 0, 0, 0, 0)
    assert [(t["qty"], t["price_dest"]) for t in trades] == [(5, 150)]
