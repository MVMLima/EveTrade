"""P0.4 — recursos sem backtest ficam desligados por padrão."""
import pytest

from conftest import et


def test_off_by_default():
    assert et._CONFIG_DEFAULTS["experimental"] == "0"
    assert et.experimental_enabled(dict(et._CONFIG_DEFAULTS)) is False


@pytest.mark.parametrize("value,expected", [
    ("1", True), ("true", True), ("True", True), ("0", False), ("", False), ("no", False)])
def test_flag_parsing(value, expected):
    assert et.experimental_enabled({"experimental": value}) is expected


def test_lab_off_by_default():
    assert et._CONFIG_DEFAULTS["lab"] == "0"
    assert et.lab_enabled(dict(et._CONFIG_DEFAULTS)) is False
    assert et.lab_enabled({"lab": "1"}) is True


def test_forecast_notice_points_to_backtest_report():
    assert "docs/BACKTEST_PREVISAO.md" in et.FORECAST_RETIRED_NOTICE
