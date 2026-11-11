"""P1.3 — avaliação prospectiva de sinais e registro em CSV."""
from datetime import date

import pytest

from evetrade.core.signal_eval import forward_returns, spearman, summarize_signal
from evetrade.snapshots import append_signals, load_signals


def test_spearman_known_values():
    assert spearman([1, 2, 3, 4], [10, 20, 30, 40]) == pytest.approx(1.0)
    assert spearman([1, 2, 3, 4], [40, 30, 20, 10]) == pytest.approx(-1.0)
    assert spearman([1, 2], [1, 2]) is None            # poucos pares


def test_forward_returns_only_mature_signals():
    prices = {(1, "2026-01-01"): 100.0, (1, "2026-01-08"): 110.0,
              (2, "2026-01-01"): 100.0, (2, "2026-01-08"): 90.0}
    sig = [{"date": "2026-01-01", "type_id": 1, "score": 80, "acao": "COMPRAR"},
           {"date": "2026-01-01", "type_id": 2, "score": 20, "acao": "MONITORAR"},
           {"date": "2026-01-05", "type_id": 1, "score": 50, "acao": "COMPRAR"}]
    rets = forward_returns(sig, lambda t, d: prices.get((t, d)), 7, date(2026, 1, 9))
    assert [(r["type_id"], round(r["ret"], 4)) for r in rets] == [(1, 0.10), (2, -0.10)]


def test_summary_rewards_informative_score():
    rets = [{"score": s, "ret": s / 1000, "acao": "COMPRAR" if s > 50 else "MONITORAR"}
            for s in range(0, 100, 10)]
    out = summarize_signal(rets)
    assert out["spearman"] == pytest.approx(1.0)
    assert out["top_minus_bottom"] > 0
    assert set(out["mean_ret_by_acao"]) == {"COMPRAR", "MONITORAR"}


def test_append_signals_replaces_same_day(tmp_path):
    log = tmp_path / "conviction_log.csv"
    rows = [{"type_id": 34, "conviction_score": 70.0, "acao": "COMPRAR", "n_fontes": 3,
             "sell_min": 5.1}, {"type_id": 35, "conviction_score": None}]
    assert append_signals(log, 10000002, rows, date(2026, 1, 1)) == 1
    assert append_signals(log, 10000002, rows, date(2026, 1, 1)) == 1   # sem duplicar
    append_signals(log, 10000002, rows, date(2026, 1, 2))
    got = load_signals(log)
    assert [(r["date"], r["type_id"], r["acao"]) for r in got] == [
        ("2026-01-01", "34", "COMPRAR"), ("2026-01-02", "34", "COMPRAR")]
