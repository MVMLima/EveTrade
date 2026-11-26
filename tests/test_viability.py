"""plan_lot / mineral_opportunity — dimensionamento por demanda e custo de
oportunidade do mineral.

Valores calculados à mão no docstring de cada teste.
"""
import pytest

from conftest import et

# ── plan_lot ─────────────────────────────────────────────────────────────────

def test_corta_lote_na_fatia_do_volume_diario():
    # Epithal do vídeo: 25 un/dia; pedir 200 com fatia de 15% -> 25*0.15 = 3.75 -> 3
    r = et.plan_lot(25, 200, max_share=0.15)
    assert r["lote_teto"] == 3
    assert r["lote_recomendado"] == 3
    assert r["excede"] is True
    assert r["confiavel"] is True
    assert r["dias_escoamento"] == pytest.approx(3 / 25)


def test_lote_menor_que_o_teto_passa_intacto():
    r = et.plan_lot(1000, 50, max_share=0.15)   # teto 150, pedido 50
    assert r["lote_recomendado"] == 50
    assert r["excede"] is False
    assert r["aviso"].startswith("Lote dentro")


def test_volume_desconhecido_nao_zera_o_lote():
    """0/None = sem dado, não 'mercado sem demanda'.

    Zerar aqui repetiria o bug do 'Vol max (m3) = 0', que derrubava a análise
    inteira quando o campo vinha vazio.
    """
    for vol in (0, None):
        r = et.plan_lot(vol, 200)
        assert r["lote_recomendado"] == 200
        assert r["excede"] is False
        assert r["confiavel"] is False
        assert r["lote_teto"] is None


def test_volume_irrisorio_ainda_produz_uma_unidade():
    # 5 un/dia * 15% = 0.75 -> teto 0, mas o lote mínimo é 1 e o aviso aparece
    r = et.plan_lot(5, 10)
    assert r["lote_teto"] == 0
    assert r["lote_recomendado"] == 1
    assert "menos de 1 unidade" in r["aviso"]


def test_pedido_zero_nao_gera_lote():
    r = et.plan_lot(100, 0)
    assert r["lote_recomendado"] == 0
    assert r["excede"] is False


def test_share_fora_da_faixa_e_saneado():
    assert et.plan_lot(100, 100, max_share=0)["max_share"] == 0.01
    assert et.plan_lot(100, 100, max_share=5)["max_share"] == 1.0


# ── mineral_opportunity ──────────────────────────────────────────────────────

MATS = [
    {"type_id": 34, "quantity": 1000, "unit_price": 10.0, "subtotal": 10000.0},
    {"type_id": 35, "quantity": 500, "unit_price": 4.0, "subtotal": 2000.0},
]
PRECOS = {34: 9.0, 35: 3.5}      # venda direta (buy max) abaixo da compra


def test_vender_mineral_rende_mais_que_fabricar():
    """Caso Venture do vídeo: job cost + mineral > receita do produto."""
    # mineral = 1000*9 + 500*3.5 = 9000 + 1750 = 10750
    # valor_agregado = 12000 - 2500 - 10750 = -1250
    r = et.mineral_opportunity(MATS, PRECOS, gross_revenue=12000,
                               job_cost=2500, units=2)
    assert r["mineral_value"] == pytest.approx(10750)
    assert r["net_revenue"] == pytest.approx(12000)
    assert r["value_added"] == pytest.approx(-1250)
    assert r["value_added_per_unit"] == pytest.approx(-625)
    assert r["verdict"] == "VENDER MINERAL"
    assert r["coverage"] == 1.0


def test_fabricar_quando_agrega_valor():
    # 15000 - 1500 - 10750 = +2750
    r = et.mineral_opportunity(MATS, PRECOS, gross_revenue=15000,
                               job_cost=1500, units=2)
    assert r["value_added"] == pytest.approx(2750)
    assert r["verdict"] == "CONSTRUIR"
    assert r["return_on_mineral"] == pytest.approx(25.58)


def test_sem_preco_de_um_material_nao_veredicta():
    """Dado faltando não pode virar 'CONSTRUIR' — o mineral estaria subestimado."""
    r = et.mineral_opportunity(MATS, {34: 9.0}, gross_revenue=99000,
                               job_cost=0, units=1)
    assert r["missing_prices"] == [35]
    assert r["coverage"] == pytest.approx(10000 / 12000, abs=1e-4)
    assert r["verdict"] == "INCOMPLETO"


def test_imposto_incide_nos_dois_lados():
    # 5% sobre produto e sobre mineral: 14250 - 0 - 10212.5 = +4037.5
    r = et.mineral_opportunity(MATS, PRECOS, gross_revenue=15000,
                               job_cost=0, units=1, sales_tax=0.05)
    assert r["net_revenue"] == pytest.approx(14250)
    assert r["mineral_value"] == pytest.approx(10212.5)
    assert r["value_added"] == pytest.approx(4037.5)


def test_units_zero_nao_divide_por_zero():
    r = et.mineral_opportunity(MATS, PRECOS, gross_revenue=1, job_cost=0, units=0)
    assert r["units"] == 1
    assert isinstance(r["value_added_per_unit"], float)


def test_lista_vazia_de_materiais_nao_quebra():
    r = et.mineral_opportunity([], {}, gross_revenue=100, job_cost=10, units=1)
    assert r["mineral_value"] == 0
    assert r["value_added"] == pytest.approx(90)
    assert r["verdict"] == "INCOMPLETO"      # cobertura 0: nada de green light
