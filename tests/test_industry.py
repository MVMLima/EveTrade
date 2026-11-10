"""calculate_industry_cost — fixa o comportamento ATUAL da fórmula.

Atenção (a verificar contra o jogo, ver plan/PLANO_MELHORIA.md):
  * o código SOMA ME do blueprint e bônus de material da estrutura;
    a CCP documenta os bônus como multiplicativos;
  * a facility tax incide sobre job_gross; no cliente atual ela parece
    incidir sobre o EIV.
Se a fórmula for corrigida, estes valores esperados mudam de propósito.
"""
import pytest

from conftest import et


def _cost(**kw):
    base = dict(materials=[{"type_id": 34, "quantity": 1000, "unit_price": 5.0}],
                product_qty=1, facility_tax=0.10, system_cost_index=0.05,
                adjusted_prices={34: 4.0}, me_level=10, runs=1)
    base.update(kw)
    return et.calculate_industry_cost(**base)


def test_known_breakdown():
    r = _cost()
    # ME10: 1000 * 0.9 = 900 un ; materiais 900*5 = 4500 ; EIV 900*4 = 3600
    assert r["materials"][0]["quantity"] == 900
    assert r["total_mat_cost"] == pytest.approx(4500)
    assert r["eiv"] == pytest.approx(3600)
    assert r["job_cost"] == pytest.approx(3600 * 0.05)          # 180
    assert r["scc_surcharge"] == pytest.approx(3600 * 0.04)     # 144 (SCC fixo 4%)
    assert r["facility_cost"] == pytest.approx(180 * 0.10)      # 18
    assert r["total_cost"] == pytest.approx(4500 + 180 + 144 + 18)
    assert r["cost_per_unit"] == pytest.approx(4842)


def test_minimum_one_unit_per_run():
    r = _cost(materials=[{"type_id": 34, "quantity": 1, "unit_price": 5.0}], runs=10)
    assert r["materials"][0]["quantity"] == 10   # max(runs, ceil(1*0.9*10)=9)


def test_structure_job_bonus_reduces_only_job_gross():
    r = _cost(structure_job_bonus=0.03)
    assert r["job_cost"] == pytest.approx(180 * 0.97)
    assert r["scc_surcharge"] == pytest.approx(144)   # SCC não recebe bônus


def test_current_me_and_structure_bonus_are_additive():
    # Comportamento atual: 1 - (0.10 + 0.01) = 0.89 -> 890 un.
    # Multiplicativo (CCP) daria ceil(1000*0.9*0.99) = 891.
    r = _cost(structure_mat_bonus=0.01)
    assert r["materials"][0]["quantity"] == 890


def test_cost_per_unit_divides_by_total_output():
    r = _cost(product_qty=10, runs=2)
    assert r["product_qty_total"] == 20
    assert r["cost_per_unit"] == pytest.approx(r["total_cost"] / 20, abs=0.01)
