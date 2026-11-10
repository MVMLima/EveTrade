"""P0.3 — falha de histórico aparece como ERRO, não como "sem dado"."""
import pytest

from conftest import et, order

RID, TID = 10000002, 34


@pytest.fixture
def isolated_hist(monkeypatch):
    """Banco local vazio e ESI controlável; zera o registro de falhas."""
    monkeypatch.setattr(et, "_db_max_date", lambda r, t: None)
    monkeypatch.setattr(et, "_db_read", lambda r, t: [])
    monkeypatch.setattr(et, "_db_write", lambda r, t, recs: None)
    monkeypatch.setattr(et, "_hist_fetch_errors", set())
    return monkeypatch


def test_esi_failure_is_recorded(isolated_hist):
    isolated_hist.setattr(et, "esi_get", lambda *a, **k: None)      # rede falhou
    assert et.fetch_history_incremental(RID, TID) == []
    assert et.hist_failed(RID, TID)
    assert "ERRO" in et._hist_err_note(RID, [TID])


def test_item_without_trades_is_not_an_error(isolated_hist):
    isolated_hist.setattr(et, "esi_get", lambda *a, **k: [])        # 200 com lista vazia
    assert et.fetch_history_incremental(RID, TID) == []
    assert not et.hist_failed(RID, TID)
    assert et._hist_err_note(RID, [TID]) == ""


def test_success_clears_previous_error(isolated_hist):
    isolated_hist.setattr(et, "esi_get", lambda *a, **k: None)
    et.fetch_history_incremental(RID, TID)
    isolated_hist.setattr(et, "esi_get", lambda *a, **k: [{"date": "2026-11-09",
                                                           "volume": 1, "average": 1}])
    et.fetch_history_incremental(RID, TID)
    assert not et.hist_failed(RID, TID)


def test_routes_flag_rows_with_history_error(isolated_hist, monkeypatch):
    monkeypatch.setattr(et, "_warm_hist_cache", lambda *a, **k: 0)
    monkeypatch.setattr(et, "_get_hist_cached", lambda r, t: [])
    monkeypatch.setattr(et, "_hist_fetch_errors", {(RID, 1)})
    routes = et.find_routes([order(1, 100, 10, buy=False), order(2, 100, 10, buy=False)],
                            [order(1, 150, 10, buy=True), order(2, 150, 10, buy=True)],
                            "A", "B", min_roi=10)
    msgs = []
    et.enrich_routes_lucro_dia(routes, RID, progress_cb=msgs.append)
    by_tid = {r["type_id"]: r for r in routes}
    assert by_tid[1]["hist_erro"] is True and by_tid[1]["lucro_dia"] is None
    assert by_tid[2]["hist_erro"] is False
    assert et.HIST_ERROR_MARK in msgs[-1]
