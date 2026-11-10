"""Fixtures comuns: ordens sintéticas e isolamento de rede/disco.

Nenhum teste acessa a ESI, o SQLite de histórico ou o cache de nomes:
tudo que o motor buscaria fora é substituído por fakes determinísticos.
"""
import os
import sys
from datetime import datetime, timedelta, timezone

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import eve_trade as et  # noqa: E402

JITA = "The Forge (Jita)"
STATION = 60003760


def order(type_id: int, price: float, volume: int, buy: bool,
          location_id: int = STATION, min_volume: int = 1) -> dict:
    """Ordem no mesmo formato que load_orders_from_csv produz."""
    return {
        "type_id": type_id, "price": float(price), "volume_remain": int(volume),
        "is_buy_order": buy, "location_id": location_id,
        "min_volume": str(min_volume),
    }


def history(daily_volume: float, days: int = 90, price: float = 100.0) -> list:
    """Histórico diário constante dentro da janela de 90 dias."""
    today = datetime.now(timezone.utc).date()
    return [{"date": (today - timedelta(days=d)).strftime("%Y-%m-%d"),
             "volume": daily_volume, "average": price}
            for d in range(1, days + 1)]


@pytest.fixture
def offline(monkeypatch):
    """Desliga rede e disco do motor. Retorna dict type_id -> histórico.

    Preencha o dict no teste; type_id ausente => histórico vazio.
    Valor que for uma Exception é lançado (simula falha de histórico).
    """
    hist_by_tid: dict = {}

    def fake_hist(region_id, type_id):
        h = hist_by_tid.get(type_id, [])
        if isinstance(h, Exception):
            raise h
        return h

    monkeypatch.setattr(et, "_get_hist_cached", fake_hist)
    monkeypatch.setattr(et, "_warm_hist_cache", lambda *a, **k: 0)
    monkeypatch.setattr(et, "_ensure_names", lambda *a, **k: None)

    def no_network(*a, **k):
        raise AssertionError("teste tentou acessar a rede")
    monkeypatch.setattr(et, "esi_get", no_network)
    monkeypatch.setattr(et, "esi_post", no_network)
    return hist_by_tid
