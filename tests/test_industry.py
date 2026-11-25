"""calculate_industry_cost — fórmula CCP pós-Viridian.

Fontes da fórmula:
  * Viridian Expansion Notes (2023-06-09): TIF = EIV × ((SCI × bônus)
    + FacilityTax + SCC + AlphaClone) — taxas incidem sobre o EIV.
  * EVE University Wiki/Manufacturing: SCC 4% desde o patch 21.06 (2024-02-01),
    facility tax 0,25% em NPC, teto de 10%; EIV = materiais do BP em ME0.
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
    # ME10: materiais 1000*0.9 = 900 un -> 900*5 = 4500
    assert r["materials"][0]["quantity"] == 900
    assert r["total_mat_cost"] == pytest.approx(4500)
    # EIV usa a quantidade base (ME0): 1000*4 = 4000
    assert r["eiv"] == pytest.approx(4000)
    assert r["job_cost"] == pytest.approx(4000 * 0.05)          # 200
    assert r["scc_surcharge"] == pytest.approx(4000 * 0.04)     # 160
    assert r["facility_cost"] == pytest.approx(4000 * 0.10)     # 400 (sobre EIV)
    assert r["total_cost"] == pytest.approx(4500 + 200 + 160 + 400)
    assert r["cost_per_unit"] == pytest.approx(5260)


def test_eiv_does_not_use_me_or_structure_materials():
    """EIV é o valor do job em ME0 sem bônus — ME não pode baratear as taxas."""
    r0 = _cost(me_level=0, structure_mat_bonus=0.0)
    r1 = _cost(me_level=10, structure_mat_bonus=0.02)
    assert r0["eiv"] == r1["eiv"] == pytest.approx(4000)
    assert r1["total_mat_cost"] < r0["total_mat_cost"]


def test_minimum_one_unit_per_run():
    r = _cost(materials=[{"type_id": 34, "quantity": 1, "unit_price": 5.0}], runs=10)
    assert r["materials"][0]["quantity"] == 10   # max(runs, ceil(round(1*0.9*10,2)=9))


def test_structure_job_bonus_reduces_only_job_gross():
    r = _cost(structure_job_bonus=0.03)
    assert r["job_cost"] == pytest.approx(200 * 0.97)
    assert r["scc_surcharge"] == pytest.approx(160)   # SCC não recebe bônus


def test_me_and_structure_bonus_are_multiplicative():
    # 1000 * 0.90 * 0.99 = 891 (aditivo daria 890)
    r = _cost(structure_mat_bonus=0.01)
    assert r["materials"][0]["quantity"] == 891


def test_facility_tax_scales_with_eiv_not_with_sci():
    """Facility tax é % do EIV: mudar o índice do sistema não altera a taxa."""
    low = _cost(system_cost_index=0.01, facility_tax=0.05)
    high = _cost(system_cost_index=0.20, facility_tax=0.05)
    assert low["facility_cost"] == high["facility_cost"] == pytest.approx(200)


def test_cost_per_unit_divides_by_total_output():
    r = _cost(product_qty=10, runs=2)
    assert r["product_qty_total"] == 20
    assert r["materials"][0]["quantity"] == 1800
    assert r["cost_per_unit"] == pytest.approx(r["total_cost"] / 20, abs=0.01)


def test_alpha_clone_tax_is_off_by_default():
    """Conta Omega não paga a taxa: default tem de ser zero."""
    r = _cost()
    assert r["alpha_clone_cost"] == 0
    assert r["alpha_clone_tax_pct"] == 0
    assert r["total_job_cost"] == pytest.approx(200 + 160 + 400)


def test_alpha_clone_tax_is_quarter_percent_of_eiv():
    r = _cost(alpha_clone_tax=0.0025)
    assert r["alpha_clone_cost"] == pytest.approx(4000 * 0.0025)   # 10
    assert r["alpha_clone_tax_pct"] == pytest.approx(0.25)
    assert r["total_job_cost"] == pytest.approx(760 + 10)


def test_matches_client_breakdown_of_a_real_job():
    """Job real (Mjolnir, 100 runs) — o print do cliente fecha linha a linha.

    EIV 2.237.536 · SCI 2,91833% · bônus de estrutura 3% · alpha 0,25%
    → gross 63.340 · SCC 89.501 · alpha 5.594 · total 158.435.
    """
    r = et.calculate_industry_cost(
        materials=[{"type_id": 1, "quantity": 1, "unit_price": 0.0}],
        product_qty=100, runs=100, facility_tax=0.0,
        system_cost_index=0.0291833, adjusted_prices={1: 22375.36},
        me_level=0, structure_job_bonus=0.03, alpha_clone_tax=0.0025,
    )
    assert r["eiv"] == pytest.approx(2237536, abs=1)
    assert r["job_cost"] == pytest.approx(63340, abs=1)
    assert r["scc_surcharge"] == pytest.approx(89501, abs=1)
    assert r["alpha_clone_cost"] == pytest.approx(5594, abs=1)
    assert r["total_job_cost"] == pytest.approx(158435, abs=2)
