"""EVE Trade Suite — motor de cálculo sem interface gráfica.

``evetrade.core`` reúne os cálculos que movem ISK (taxas, margens, rotas,
arbitragem, indústria, orçamento). Nenhum módulo daqui importa Tkinter,
acessa a rede ou guarda estado global: dados externos entram por parâmetro.
"""

__all__ = ["core"]
