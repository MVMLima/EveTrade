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


def test_execution_depth_walks_book_to_break_even():
    # Origem: 5 @100, 5 @140, 10 @200. Destino: compra 3 @150, 10 @130.
    # Executável: 3 un (100->150) + 2 un (100->130); 140 >= 130 interrompe.
    src = [order(1, 100, 5, buy=False), order(1, 140, 5, buy=False),
           order(1, 200, 10, buy=False)]
    dst = [order(1, 150, 3, buy=True), order(1, 130, 10, buy=True)]
    r, = et.find_routes(src, dst, "A", "B", min_roi=10, sales_tax_pct=0)
    assert r["vol_realizavel"] == 3                    # só o topo (comportamento antigo)
    assert r["unid_executavel"] == 5
    assert r["lucro_garantido"] == pytest.approx(3 * 50 + 2 * 30)
    assert r["capital_executavel"] == pytest.approx(5 * 100)
    assert r["preco_venda_medio"] == pytest.approx((3 * 150 + 2 * 130) / 5)


def test_execution_depth_reflects_thin_destination_book():
    # Caso do usuário: venda mostrada a 60.990, mas o livro só absorve 2.000 un
    src = [order(1, 50_000, 10_000, buy=False)]
    dst = [order(1, 60_990, 2_000, buy=True), order(1, 49_000, 50_000, buy=True)]
    r, = et.find_routes(src, dst, "A", "B", min_roi=10, sales_tax_pct=0)
    assert r["unid_executavel"] == 2_000
    assert r["lucro_garantido"] == pytest.approx(2_000 * 10_990)


def test_execution_depth_stops_at_break_even_after_sales_tax():
    # Caso real (Complex Organisms): margem bruta < 1 ISK/un some com o imposto.
    # compra 100, vende 101 (bruto +1); com 7,5% recebe 93,425 -> prejuízo
    src = [order(1, 100, 1_000_000, buy=False)]
    dst = [order(1, 101, 1_000_000, buy=True), order(1, 130, 10, buy=True)]
    r, = et.find_routes(src, dst, "A", "B", min_roi=10)       # imposto padrão 7,5%
    assert r["unid_executavel"] == 10                          # só a compra a 130
    assert r["lucro_garantido"] == pytest.approx(10 * (130 * 0.925 - 100))


def test_lots_per_day_uses_destination_turnover(offline):
    from conftest import history
    src = [order(1, 100, 5, buy=False)]
    dst = [order(1, 150, 5, buy=True)]
    routes = et.find_routes(src, dst, "A", "B", min_roi=10)
    offline[1] = history(20)                        # destino negocia 20 un/dia
    et.enrich_routes_lucro_dia(routes, 10000043)
    assert routes[0]["reposicoes_dia"] == pytest.approx(20 / 5)


def test_execution_depth_caps_by_cargo():
    # 100 un a 2 m3/un e porao de 100 m3: so 50 un cabem
    src = [order(1, 100, 100, buy=False)]
    dst = [order(1, 150, 100, buy=True)]
    route = et.find_routes(src, dst, "A", "B", min_roi=10, sales_tax_pct=0,
                           max_volume=100,
                           item_cache={1: {"name": "X", "volume": 2.0}})[0]
    assert route["unid_executavel"] == 50
    assert route["m3_executavel"] == pytest.approx(100.0)
    assert route["lucro_garantido"] == pytest.approx(50 * 50)
    assert route["capital_executavel"] == pytest.approx(50 * 100)


def test_volume_limit_caps_units_instead_of_dropping_the_route():
    # Antes: rota descartada por total_volume > max_volume. Agora: fica, capada.
    src = [order(1, 100, 100, buy=False)]
    dst = [order(1, 150, 100, buy=True)]
    cache = {1: {"name": "X", "volume": 100.0}}
    route, = et.find_routes(src, dst, "A", "B", min_roi=10, sales_tax_pct=0,
                            max_volume=1000, item_cache=cache)
    assert route["unid_executavel"] == 10
    assert route["m3_executavel"] == pytest.approx(1000.0)


def test_route_dropped_when_a_single_unit_does_not_fit():
    src = [order(1, 100, 10, buy=False)]
    dst = [order(1, 150, 10, buy=True)]
    assert et.find_routes(src, dst, "A", "B", min_roi=10, sales_tax_pct=0,
                          max_volume=1000,
                          item_cache={1: {"name": "X", "volume": 2000.0}}) == []


def test_volume_limit_zero_means_no_limit():
    # O campo "Vol max (m3)" vazio chega como 0
    src = [order(1, 100, 100, buy=False)]
    dst = [order(1, 150, 100, buy=True)]
    route, = et.find_routes(src, dst, "A", "B", min_roi=10, sales_tax_pct=0,
                            max_volume=0,
                            item_cache={1: {"name": "X", "volume": 500.0}})
    assert route["unid_executavel"] == 100
