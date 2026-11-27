"""Tempo de produção e ISK/hora — as duas colunas novas do IND. SCANNER.

O tempo base vem do SDE (``activities.manufacturing.time``, uma run em TE0).
A conta é ``base × runs × (1 − TE/100) × (1 − bônus)``; skills, rig e
estrutura entram agregados no termo ``bônus``.
"""
import pytest

from conftest import et


def test_no_te_no_bonus_is_base_times_runs():
    # 10 min por run, 3 runs
    assert et.production_time_seconds(600, runs=3) == 1800


def test_te_reduces_one_percent_per_level():
    # TE10 corta 10% de 3600 s
    assert et.production_time_seconds(3600, te_level=10) == 3240
    # TE20 corta 20%
    assert et.production_time_seconds(3600, te_level=20) == 2880


def test_bonus_multiplies_after_te():
    # 3600 s, TE10 (3240) e mais 15% de estrutura/skills -> 2754
    assert et.production_time_seconds(3600, te_level=10, time_bonus=0.15) == 2754


def test_runs_multiply_before_discounts():
    assert et.production_time_seconds(600, runs=10, te_level=10, time_bonus=0.15) == 4590


def test_te_is_clamped_to_0_20():
    assert et.production_time_seconds(1000, te_level=50) == 800
    assert et.production_time_seconds(1000, te_level=-7) == 1000


def test_ceil_and_floor_at_one_second():
    # 3 s com TE10 = 2,7 -> 3 (nunca arredonda para baixo)
    assert et.production_time_seconds(3, te_level=10) == 3
    assert et.production_time_seconds(1, time_bonus=0.99) == 1


def test_unknown_base_returns_zero_not_zero_hours():
    assert et.production_time_seconds(0) == 0
    assert et.production_time_seconds(-500, runs=5) == 0


def test_isk_per_hour_is_profit_over_job_hours():
    # 1 bi de lucro num job de 2 h -> 500 M/h
    assert et.isk_per_hour(1_000_000_000, 7200) == pytest.approx(500_000_000)


def test_isk_per_hour_without_duration_is_zero():
    assert et.isk_per_hour(1_000_000, 0) == 0.0
