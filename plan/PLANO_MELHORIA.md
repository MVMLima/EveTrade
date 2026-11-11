# Plano de melhoria — EVE Trade Suite

Data: 2026-11-10 · Base: revisão externa + medições no código atual (`main` @ 061dde9 + correções não commitadas).

**Objetivo:** transformar o app de "protótipo de pesquisa" em "ferramenta de decisão": todo número que pode
mover ISK precisa ter teste com resposta conhecida, erro visível e origem de dados datada.

---

## 1. Conferência da revisão

| Afirmação da revisão | Medido | Situação |
|---|---|---|
| ~23k linhas, `eve_trade.py` gigante | 23.171 linhas no total; `eve_trade.py` = 17.037 linhas, 543 `def` | Procede |
| UI e cálculo no mesmo módulo, dependentes de globals | `analyze_*` usam `REGIONS`, `_hist_cache` e a sessão ESI do módulo | Procede |
| `eve_market.py` é só re-export | 122 linhas de wrappers que chamam `eve_trade` | Procede |
| Zero testes, sem CI, sem LICENSE, sem pyproject | Confirmado | Procede (o `requirements.txt` já foi criado) |
| ~155 `except Exception` + 5 `except:` | 155 `except Exception` (122 em `eve_trade.py`) + 5 `except:` | Procede |
| Fallback do keyring vai para texto puro | `eve_trade.py:55-67` faz `set_keyring(PlaintextKeyring())`. **Além disso**, com `keyrings.alt` instalado, o `ChainerBackend` padrão já inclui `PlaintextKeyring` (prioridade 0.5) logo depois do SecretService. Se o SecretService falhar no `set_password`, a chave cai em arquivo sem aviso. | **Procede e é pior do que o descrito.** O `requirements.txt` criado nesta sessão instala `keyrings.alt`. |
| N+1 em `fetch_type_ids_bulk` | Confirmado | **Corrigido** (tabela SDE; 0 chamadas ESI para 23.103 itens) |
| Max Margem: filtra depois de truncar; lixo de 0,03 ISK | Confirmado (0 linhas com `min_vol=20`) | **Corrigido** (103 linhas; `min_buy_ratio=0.10`) |
| Sem data do CSV na UI | Só aparecia na aba Dados & Cache | **Corrigido** (9 abas) |
| 27 abas | 31 chamadas `nb.add` (inclui sub-abas) | Procede |
| Previsão / Conviction Score sem backtest | Nenhum backtest publicado para os dois. `backtest_strategy` existe, mas não valida esses sinais. | Procede |

---

## Status (2026-11-10, branch `fix/p0-hardening`)

| Item | Commit | Situação |
|---|---|---|
| Bugs 1–3 (N+1, Max Margem, data do CSV) | `945dcec` | feito |
| P0.1 keyring | `0588d78` | feito; `keyrings.alt` removido do venv |
| P0.2 testes de ouro | `0588d78` (+ `9c60167`, `19395a8`) | 43 testes offline; 5 falham no código antigo |
| P0.5 Rotas/Arbitragem | `0588d78` | `min_volume` respeitado; "Lucro Total min" padrão = 1 M ISK. Não havia ordens falsas: o lixo era lucro irrisório no topo do ranking por ROI |
| P0.3 erros visíveis | `9c60167` | falha de histórico ≠ "sem dado"; barra de status em vermelho; "ERRO" em Lucro/Dia; 0 `except:` puro. Os ~150 `except Exception` de UI/I/O restantes ficam para a P1 |
| P0.4 Experimental | `19395a8` | Previsão e AI Analista escondidos por padrão; aviso fixo no Analyst e nos relatórios |

### P1 (branch `refactor/p1`)

| Item | Commit | Situação |
|---|---|---|
| P1.2 pyproject + ruff + CI + LICENSE | `39fa9be` | feito; MIT; CI em 3.12 e 3.14 (roda após o push) |
| P1.1 motor sem Tkinter | `fe2fdc0` | **parcial**: `evetrade/core` com taxas, ordens, histórico, Max Margem, rotas, arbitragem, indústria e orçamento; saídas idênticas byte a byte ao código anterior com dados reais. **Falta**: Escassez, Demanda/Oferta, Mov. Preços, Market Making, Station Trader e a camada de dados (ESI, SQLite, CSV, SDE). `eve_trade.py` tem 16,7k linhas (meta < 10k não atingida) |
| P1.3 backtests | `5462138` | Previsão **reprovada** (skill −10% a −18% vs. passeio aleatório). Conviction: não testável no passado; o registro prospectivo começou (`eve_data/conviction_log.csv`) |
| P1.4 enxugar abas | `16014f7` | padrão 28 → 20 abas; Laboratório e Experimental opt-in; Previsão removida da UI (código mantido) |

### Revisão externa nº 2 (branch `fix/bait-and-route-depth`)

| Item | Commit | Situação |
|---|---|---|
| Isca a ~10% da venda (teto de 380,77%) | `c439af0` | compra < 0,6× média 30d **e** venda > 1,67× média 30d são descartadas; a revisão propôs só o lado da compra, mas aí a isca mudava de lado (200 de 200) |
| Profundidade na aba Rotas | `ff3a677` | Unid. exec., Lucro exec. (líq. de imposto), Lotes/dia |
| Ordenação de tabelas (bug desde a v14) | `ff3a677` | 12 tabelas: travava (iid=type_id) ou mostrava ordem diferente da lista (clique abria outro item) |
| Log do keyring headless | `ff3a677` | rebaixado para debug |

**Achados novos (a verificar antes de mudar):**
- Station Trader: `effective_sell` (limite de venda configurado) é calculado e nunca usado; a receita usa sempre o menor preço de venda atual.

**Achado anterior (a verificar no jogo antes de mudar):** `calculate_industry_cost` soma o ME do blueprint e o
bônus de material da estrutura (a CCP documenta como multiplicativo) e aplica a facility tax sobre
`job_gross` (no cliente atual parece ser sobre o EIV). `tests/test_industry.py` fixa o comportamento atual
e marca esses dois pontos.

---

## 2. Prioridades

### P0 — Segurança e números corretos (fazer primeiro)

**P0.1 Chave da API nunca em texto puro**
- Remover `keyrings.alt` do `requirements.txt` e desinstalar do venv.
- `eve_trade.py:55-67`: apagar o fallback `PlaintextKeyring`. Aceitar só backends seguros (rejeitar `keyrings.alt.*` e `keyring.backends.fail`).
- Sem backend seguro: ler `ANTHROPIC_API_KEY` do ambiente, ou guardar a chave só em memória durante a sessão. Mostrar isso na aba Configurações.
- Verificar se já existe um arquivo em texto puro (`~/.local/share/python_keyring/`). Se existir, avisar e oferecer apagar.
- **Aceite:** um teste com backend falho confirma que nenhum arquivo é escrito; o README volta a ser verdadeiro.

**P0.2 Testes de ouro dos cálculos que movem ISK** (antes de qualquer refatoração)
- Criar `tests/` com pytest e ordens sintéticas de resposta calculada à mão:
  1. `analyze_max_margin`: taxas, ordem filtro→corte, `min_buy_ratio`, `min_vol`.
  2. `find_routes` / `match_orders_granular`: lucro, volume realizável, m³.
  3. Custo industrial (fórmula CCP: ME, EIV, SCC 4%, facility tax).
  4. `analyze_budget` ou arbitragem: alocação respeitando o orçamento.
- O histórico e a ESI são substituídos por fakes (`monkeypatch` de `_get_hist_cached`, `esi_get`, `esi_post`). Nenhum teste acessa a rede.
- **Aceite:** `pytest` passa em menos de 10 s; reverter o bug do Max Margem faz pelo menos um teste falhar.

**P0.3 Erro visível ≠ "sem dado"**
- Auditar os 155 `except Exception` em três classes: (a) UI/cosmético: mantém com `log.debug`; (b) caminho de cálculo: deve marcar a linha ou o resultado como erro; (c) I/O de rede: retry ou erro contado.
- Nos `analyze_*`, uma falha de histórico vira campo `status="erro_hist"` em vez de `None`. A barra de status mostra "N itens com erro".
- Trocar os 5 `except:` por exceções específicas.
- **Aceite:** com um fake de histórico que lança exceção, a tabela mostra "erro" e o contador aparece.

**P0.4 Rebaixar o que parece rigoroso e não é**
- Abas Previsão, Conviction Score (`eve_analyst.py`) e AI Analista passam a ficar atrás de um interruptor "Experimental" (desligado por padrão), com aviso fixo: "sem validação estatística".
- **Aceite:** a instalação padrão não mostra esses números sem o aviso.

**P0.5 Mesmo filtro de sanidade em Rotas e Arbitragem**
- Verificar se `find_routes` e a arbitragem também aceitam ordens simbólicas (a revisão viu "topo da lista é lixo" em Rotas).
- Se aceitarem, aplicar a regra de `min_buy_ratio` / profundidade mínima e cobrir com teste do P0.2.

### P1 — Estrutura que permite testar e mudar com segurança

**P1.1 Extrair o motor de cálculo para um pacote sem Tkinter** (incremental, uma função por commit)
```
evetrade/
  core/      fees.py, orders.py (build_order_index), margin.py, routes.py, industry.py, budget.py
  data/      esi.py (sessão, ETag), history_db.py (SQLite), orders_csv.py, sde.py
  analysis/  shortage.py, demand_supply.py, market_making.py, cycles.py (experimental)
ui/          eve_trade.py passa a só montar abas e chamar evetrade.*
```
- O histórico é injetado como parâmetro (`hist_provider`), em vez de vir do `_hist_cache` global.
- `eve_market.py` vira a API real, ou é substituído por `evetrade/__init__.py`.
- Ordem: `fees` → `orders` → `margin` → `routes` → `industry` → resto. Os testes do P0.2 protegem cada passo.
- **Aceite:** `python -c "import evetrade.core"` funciona sem `tkinter`; `eve_trade.py` fica abaixo de 10k linhas no fim da P1.

**P1.2 Projeto mínimo profissional**
- `pyproject.toml` (deps, ruff, pytest), `ruff check` sem erros nos módulos novos.
- GitHub Actions: ruff + pytest em cada push.
- `LICENSE`: **decisão sua** (MIT, GPL ou "todos os direitos reservados").

**P1.3 Validar ou remover os sinais experimentais**
- Backtest walk-forward da Previsão contra um baseline ingênuo (preço de amanhã = preço de hoje). Métricas: MAE e acerto de direção.
- Conviction Score: mostrar se scores altos tiveram retorno maior nos 7 e 30 dias seguintes, usando o histórico do SQLite.
- Regra: se não bater o baseline com margem clara, a aba sai do app. Publicar os números em `docs/`.

**P1.4 Enxugar a interface**
- Proposta de núcleo (9 abas): Rotas, Arbitragem, Max Margem, Station Trader, Market Making, Indústria, Watchlist, Dados & Cache, Configurações.
- O restante vai para um menu "Laboratório" ou é removido. **Decisão sua:** quais você realmente usa?

### P2 — Qualidade de vida

- Expor `min_buy_ratio` e profundidade mínima (ISK na ponta de compra) na aba Max Margem.
- Histórico de git: commits pequenos (Conventional Commits) e um `docs/DECISOES.md` que transforme os comentários `FIX:` e as medições espalhadas no código em registro de decisões.
- `CHANGELOG.md` a partir da v15.
- Painel de log de erros na UI (os últimos N `log.warning`/`log.error`).

---

## 3. Ordem de execução sugerida

| # | Item | Esforço | Depende de |
|---|---|---|---|
| 0 | Commit das 3 correções já feitas | 5 min | — |
| 1 | P0.1 keyring | 1 h | — |
| 2 | P0.2 testes de ouro (4 cálculos) | 4–6 h | — |
| 3 | P0.5 sanidade em Rotas/Arbitragem | 1–2 h | 2 |
| 4 | P0.3 erros visíveis (caminho de cálculo primeiro) | 3–4 h | 2 |
| 5 | P0.4 interruptor Experimental | 1 h | — |
| 6 | P1.2 pyproject + ruff + CI | 1–2 h | 2 |
| 7 | P1.1 extração do motor (incremental) | 2–4 dias | 2, 6 |
| 8 | P1.3 backtests | 1–2 dias | 7 |
| 9 | P1.4 enxugar abas | 2–4 h | 8 + decisão sua |
| 10 | P2 | contínuo | — |

## 4. Fora de escopo agora

- Trocar Tkinter por outro framework: não resolve nenhum dos problemas acima.
- Reescrever do zero: o motor de dados (ETag, SQLite, warm cache, fórmula CCP) está bom; a dívida está nas camadas de cima.
- Otimizações de desempenho adicionais: o gargalo real (N+1) já saiu.

## 5. Decisões pendentes (suas)

1. Licença do projeto.
2. Quais abas você usa de fato (para o P1.4).
3. Se Previsão/Conviction devem ser escondidos já (P0.4) ou só depois do backtest (P1.3).
