"""Taxas de mercado (broker fee e sales tax) por nível de skill.

Módulo puro: sem Tkinter, sem rede, sem estado global.
"""


def calc_fee_rates(broker_rel: int, accounting: int):
    """Retorna (rb, rt) com base nos niveis de skill (1-5)."""
    rb = max(0.0, 0.03 - 0.003 * broker_rel)
    rt = 0.075 * (1.0 - 0.11 * accounting)
    return rb, rt
