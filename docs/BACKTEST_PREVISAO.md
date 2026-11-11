# Backtest — aba Previsão (`predict_prices`)

Gerado em 2026-11-10 por `tools/backtest_forecast.py`.

- Região: 10000002 · itens: 200 mais negociados em ISK/dia com ≥ 250 dias de histórico
- Walk-forward: 12 cortes a cada 14 dias; o modelo só vê dados até o corte
- Baseline: passeio aleatório (preço futuro = último preço médio diário)
- Critério de aprovação: skill > 5% **e** acerto de direção > 55%

| Horizonte | N | MAE modelo | MAE baseline | Skill | Modelo melhor | Acerto direção | Cobertura banda ±1σ |
|---|---|---|---|---|---|---|---|
| 1 d | 2380 | 2.81% | 2.39% | -17.7% | 38.2% | 56.7% | 73.2% |
| 7 d | 2381 | 4.94% | 4.48% | -10.3% | 43.7% | 54.7% | 85.3% |
| 14 d | 2385 | 6.77% | 5.87% | -15.4% | 42.3% | 54.6% | 84.9% |

**Veredito:** REPROVADO: não bate o baseline ingênuo (preço futuro = preço de hoje).

Skill = 1 − MAE_modelo / MAE_baseline (positivo = modelo melhor). Acerto de direção de 50% equivale a cara-ou-coroa. Uma banda ±1σ bem calibrada cobriria ~68%.
