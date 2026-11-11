"""P1.1 — evetrade.core é puro: importa sem Tkinter e calcula sem I/O."""
import subprocess
import sys
from pathlib import Path

from conftest import history, order

ROOT = Path(__file__).resolve().parent.parent


def test_core_imports_without_tkinter_or_network():
    code = (
        "import sys\n"
        "sys.modules['tkinter'] = None\n"          # qualquer import de tkinter falha
        "sys.modules['requests'] = None\n"         # e de rede também
        "import evetrade.core as c\n"
        "assert 'eve_trade' not in sys.modules\n"
        "print(len(c.__all__))\n"
    )
    r = subprocess.run([sys.executable, "-c", code], cwd=ROOT,
                       capture_output=True, text=True, check=False)
    assert r.returncode == 0, r.stderr
    assert int(r.stdout.strip()) >= 16


def test_margin_with_injected_history():
    from evetrade.core import build_order_index, margin_candidates, select_by_volume
    idx = build_order_index([order(1, 200, 10, buy=False), order(1, 100, 10, buy=True),
                             order(2, 160, 10, buy=False), order(2, 100, 10, buy=True)])
    cands = margin_candidates(idx, {}, "R")
    assert [c["type_id"] for c in cands] == [1, 2]          # ordenado por margem
    vols = {1: history(1), 2: history(50)}
    rows = select_by_volume(cands, vols.__getitem__, min_vol=20, top_n=5)
    assert [r["type_id"] for r in rows] == [2]


def test_budget_with_injected_history():
    from evetrade.core import allocate_budget, build_order_index
    idx = build_order_index([order(1, 13_000, 100, buy=False), order(1, 10_000, 100, buy=True)])
    basket, stats = allocate_budget(idx, {}, "R", budget=100_000, broker_rel=0,
                                    accounting=0, hist_provider=lambda t: history(10))
    assert basket[0]["qty"] == 8                             # int(100000 / 11665)
    assert stats["total_spent"] <= 100_000


def test_no_history_provider_means_no_volume_claims():
    from evetrade.core import build_order_index, margin_candidates, select_by_volume
    idx = build_order_index([order(1, 200, 10, buy=False), order(1, 100, 10, buy=True)])
    cands = margin_candidates(idx, {}, "R")
    assert select_by_volume(cands, None, min_vol=1) == []
    assert len(select_by_volume(cands, None, min_vol=0)) == 1
