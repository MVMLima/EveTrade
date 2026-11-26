"""Dimensionamento de lote por demanda e custo de oportunidade do mineral.

Dois erros de industrialista que o motor de custo não vê sozinho:

  1. Dimensionar o lote pela margem unitária e ignorar o volume diário do hub
     — o excedente empurra o preço para baixo e o estoque fica parado.
  2. Tratar mineral próprio como custo zero — o mineral tem valor de venda
     própria, e construir só agrega se a receita líquida do produto superar
     job cost + o que o mineral renderia vendido direto.

Módulo puro: sem Tkinter, sem rede, sem estado global. Quem chama passa os
números (volume médio, preços) já coletados.
"""


def plan_lot(daily_volume, want_qty, max_share=0.15):
    """Limita o lote produzido à fatia do volume diário que o hub absorve.

    Parâmetros:
      daily_volume — unidades/dia do produto no hub de venda (avg_vol_90)
      want_qty     — unidades que se pretende produzir
      max_share    — fração do volume diário aceita como sua fatia (0–1)

    Volume desconhecido (None/0) NÃO limito a produção: devolvo o lote pedido
    com ``confiavel=False``. Zerar aqui repetiria o bug do "Vol max (m3) = 0",
    em que a falta de dado derrubava a análise inteira em silêncio.
    """
    want = max(0, int(want_qty))
    vol = float(daily_volume or 0)
    share = min(1.0, max(0.01, float(max_share)))

    if vol <= 0:
        return {
            "daily_volume": 0.0,
            "max_share": share,
            "lote_teto": None,
            "lote_recomendado": want,
            "dias_escoamento": None,
            "excede": False,
            "confiavel": False,
            "aviso": "Sem volume diário no histórico do hub — dimensionamento não aplicado.",
        }

    teto = int(vol * share)                     # floor: fração de unidade não vende
    lote = max(1, min(want, teto)) if want else 0
    dias = lote / vol
    excede = want > lote

    if teto < 1:
        aviso = (f"Mercado absorve {vol:,.1f} un/dia: {share*100:.0f}% disso é menos "
                 f"de 1 unidade. Lote de 1 unidade já é agressivo.")
    elif excede:
        aviso = (f"Lote pedido ({want:,} un) é {want/vol:,.1f} dias de volume do hub. "
                 f"Cortado para {lote:,} un ({share*100:.0f}% de um dia).")
    else:
        aviso = f"Lote dentro da fatia de {share*100:.0f}% do volume diário."

    return {
        "daily_volume": vol,
        "max_share": share,
        "lote_teto": teto,
        "lote_recomendado": lote,
        "dias_escoamento": dias,
        "excede": excede,
        "confiavel": True,
        "aviso": aviso,
    }


def mineral_opportunity(mat_lines, mineral_prices, gross_revenue, job_cost,
                        units, sales_tax=0.0):
    """Compara construir contra vender o mineral direto no mercado.

    Parâmetros:
      mat_lines       — ``result["materials"]`` de calculate_industry_cost
                        (usa ``quantity`` final, já com ME, e ``subtotal``)
      mineral_prices  — {type_id: preço de venda direta} (buy max do hub)
      gross_revenue   — receita bruta do produto (preço × unidades)
      job_cost        — ``result["total_job_cost"]``
      units           — unidades produzidas (para o valor agregado por unidade)
      sales_tax       — imposto de venda, aplicado aos DOIS lados (0–1)

    O comparativo é: receita líquida do produto − job cost − valor líquido do
    mineral vendido direto. Positivo = fabricar agrega valor sobre o mineral.
    Negativo = o mercado paga mais pelo mineral cru do que pelo produto pronto.

    Se faltar preço de algum material, o valor do mineral fica subestimado e o
    veredicto vira ``INCOMPLETO`` — nunca ``CONSTRUIR`` sobre dado faltando.
    """
    tax = min(1.0, max(0.0, float(sales_tax)))
    units = max(1, int(units))

    valor_mineral = 0.0
    custo_coberto = 0.0
    custo_total_mat = 0.0
    sem_preco = []

    for line in mat_lines:
        subtotal = float(line.get("subtotal", 0.0))
        custo_total_mat += subtotal
        preco = mineral_prices.get(line["type_id"])
        if preco is None or preco <= 0:
            sem_preco.append(line["type_id"])
            continue
        valor_mineral += float(line["quantity"]) * float(preco)
        custo_coberto += subtotal

    cobertura = (custo_coberto / custo_total_mat) if custo_total_mat > 0 else 0.0
    receita_liquida = float(gross_revenue) * (1.0 - tax)
    mineral_liquido = valor_mineral * (1.0 - tax)
    valor_agregado = receita_liquida - float(job_cost) - mineral_liquido

    if cobertura < 0.999:
        verdict = "INCOMPLETO"
        aviso = (f"Sem preço de venda para {len(sem_preco)} material(is) — "
                 f"{cobertura*100:.0f}% do custo de material coberto. "
                 f"Valor do mineral subestimado; veredicto suspenso.")
    elif valor_agregado > 0:
        verdict = "CONSTRUIR"
        aviso = (f"Fabricar rende {valor_agregado:+,.0f} ISK acima de vender o "
                 f"mineral cru ({valor_agregado/units:+,.0f} ISK/un).")
    else:
        verdict = "VENDER MINERAL"
        aviso = (f"Vender o mineral rende {abs(valor_agregado):,.0f} ISK mais que "
                 f"fabricar ({abs(valor_agregado)/units:,.0f} ISK/un).")

    return {
        "mineral_value": round(mineral_liquido, 2),
        "mineral_value_gross": round(valor_mineral, 2),
        "net_revenue": round(receita_liquida, 2),
        "job_cost": round(float(job_cost), 2),
        "value_added": round(valor_agregado, 2),
        "value_added_per_unit": round(valor_agregado / units, 2),
        "return_on_mineral": (round(valor_agregado / mineral_liquido * 100, 2)
                              if mineral_liquido > 0 else None),
        "coverage": round(cobertura, 4),
        "missing_prices": sem_preco,
        "sales_tax": tax,
        "units": units,
        "verdict": verdict,
        "aviso": aviso,
    }
