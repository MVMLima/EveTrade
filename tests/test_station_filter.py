"""Preço por estação: ordem de citadel não pode virar preço de produção.

O livro regional mistura estação NPC e estrutura privada. Numa comparação real
o mesmo material saiu a 2,45 na região e a 4,19 em Jita 4-4 — a margem do item
ia de 17% para 45% só pela ordem errada entrar na conta.
"""
import pytest

from conftest import STATION, et


def _ordem(location_id: int, price: float, buy: bool = False, volume: int = 1) -> dict:
    return {"location_id": location_id, "price": price, "is_buy_order": buy,
            "volume_remain": volume}


def test_filtro_mantem_so_a_estacao():
    ordens = [_ordem(STATION, 10.0), _ordem(1023000000001, 5.0),
              _ordem(60003761, 7.0)]
    r = et._filtrar_livro_estacao(ordens, STATION)
    assert [o["location_id"] for o in r] == [STATION]


def test_filtro_none_devolve_regiao_inteira():
    ordens = [_ordem(STATION, 10.0), _ordem(1023000000001, 5.0)]
    assert et._filtrar_livro_estacao(ordens, None) == ordens


def test_fetch_live_price_ignora_ordem_fora_da_estacao(monkeypatch):
    """Citadel com preço melhor não entra: só a ordem da estação conta."""
    ordens = [
        _ordem(1023000000001, 2.45, buy=False, volume=500),   # citadel, isca
        _ordem(STATION, 4.19, buy=False, volume=100),
        _ordem(1023000000001, 9.99, buy=True, volume=500),    # citadel, compra
        _ordem(STATION, 3.80, buy=True, volume=100),
    ]

    class _Resp:
        status_code = 200

        def json(self):
            return ordens

    class _Sessao:
        def get(self, url, params=None, timeout=None):
            return _Resp()

    monkeypatch.setattr(et, "_get_session", lambda: _Sessao())
    monkeypatch.setattr(et, "_get_hist_cached", lambda *a, **k: [])

    p = et.fetch_live_price(10000002, 34, station_id=STATION)
    assert p is not None
    assert p["sell_min"] == pytest.approx(4.19)
    assert p["buy_max"] == pytest.approx(3.80)

    sem_filtro = et.fetch_live_price(10000002, 34)
    assert sem_filtro is not None
    assert sem_filtro["sell_min"] == pytest.approx(2.45)
    assert sem_filtro["buy_max"] == pytest.approx(9.99)


def test_fetch_live_price_devolve_none_sem_ordem_na_estacao(monkeypatch):
    class _Resp:
        status_code = 200

        def json(self):
            return [_ordem(1023000000001, 5.0, buy=False, volume=10),
                    _ordem(1023000000001, 4.0, buy=True, volume=10)]

    class _Sessao:
        def get(self, url, params=None, timeout=None):
            return _Resp()

    monkeypatch.setattr(et, "_get_session", lambda: _Sessao())
    monkeypatch.setattr(et, "_get_hist_cached", lambda *a, **k: [])
    assert et.fetch_live_price(10000002, 34, station_id=STATION) is None
