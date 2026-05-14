# EVE Trade Suite — Resumo de Desenvolvimento

> Três aplicações Python/tkinter para trading no EVE Online.
> Desenvolvidas em sessões iterativas com correções, funcionalidades e otimizações.

---

## Estatísticas Gerais

| Métrica | Valor |
|---|---|
| Linhas totais | 19.285 |
| Métodos e funções | 574 |
| Abas (eve_trade.py) | 27 |
| Testes automatizados | 41/41 passando |

---

## eve_trade.py

**16.230 linhas · 2 classes · 415 métodos · 109 funções de módulo · 27 abas**

### Abas disponíveis

1. Rotas
2. Arbitragem
3. Otimizador
4. Max Margem
5. Escassez
6. Demanda/Oferta
7. Mov. Preços
8. Orçamento
9. Mercado Live
10. Watchlist
11. Paper Trading
12. AI Analista
13. Comparador
14. Configurações
15. Dashboard
16. Multi-Hop
17. Correlação
18. Spread Hist.
19. Ciclos
20. Previsão
21. Market Making
22. Indústria
23. Backtesting
24. Heatmap
25. Station Trader
26. Catálogo
27. Dados & Cache

---

### Correções de Bug

| Bug | Descrição |
|---|---|
| `_pop_arb` — "Item already exists" | `run_arbitrage` gerava múltiplas entradas por `type_id` (pares de estações). Corrigido usando `iid=str(i)` (índice), `_a_on_select` armazena `_a_cur_idx`, `_a_get_selected` usa índice direto |
| HEATMAP grupos | Expandido de 6 → 18 grupos (~110 type_ids). `_draw_heatmap` reescrito com `tk.Canvas` + scrollbars. Bug `_hm_drill_down .keys()` em lista corrigido |
| SPREAD HIST. truncada | `_build_tab_sph` tinha apenas 11 linhas. Reconstruída com todos os widgets, botões BACKTEST/GRÁFICO/PREVER/ANALISAR IA, `_bind_tid_lookup` e `_sph_open_*` implementados. `_cat_jump` corrigido para chamar `_start_sph` |
| `import re` faltando | Adicionado `import re` na linha 17 do arquivo |
| `exc` no lambda | `_worker_ind_auto` capturava `exc` em variável `_err = str(exc)` antes do lambda para evitar `NameError: cannot access free variable 'exc'` |
| Fuzzwork `materialTypeID` vs `typeid` | Campo correto é `materialTypeID` (não `typeid` que é o blueprint). Corrigido em `_parse_typematerials` |
| `blueprint±1` fallback incorreto | Removido — mapeamento produto→blueprint no EVE é arbitrário (Nova Heavy Missile=209, Blueprint=807, diferença=598) |
| `prod_tid` hardcoded errado | `_HARDCODED_BP` tinha `807: {"prod": 209}` mas type_id 209 é Scourge Heavy Missile. Todos os `"prod"` foram removidos do hardcoded |
| `_get_product_tid` ordem incorreta | Reordenado — dedução por nome (layer 1, mais confiável) antes de `bp_data["prod_tid"]` (layer 2) |
| `analyze_supply_shortage` parâmetros inválidos | Removidos parâmetros `broker_rel`/`accounting` inválidos |
| `analyze_demand_supply` parâmetros inválidos | Removidos parâmetros `broker_rel`/`accounting` inválidos |
| `analyze_budget` parâmetro errado | `dias_risco` → `risk_mode="Balanced (3d)"`, desempacotamento seguro com `isinstance` |
| `NaN→int` no Excel | Adicionados `_safe_int`, `_safe_vol`, `_safe_dos`, `_safe_pct`, `_safe_float` |
| Margem 1951% | Filtro `pb > 1`, `net_margin < 500%`, `min_vol=5` |
| Conviction score uniforme 35/100 | Normalizado por `max_possible` (análises disponíveis), graduado por nível |
| Hoboleaks 404 | URL inexistente. Substituído por sistema multi-fonte |
| `_cat_jump` ind sem auto-start | Adicionado `self.after(80, self._ind_auto_load)` |

---

### Funcionalidades Novas

#### Aba INDÚSTRIA — AUTO-PREENCHER completo

- Campo "Qtd desejada" + campo "Região" + botão azul "⟳ AUTO-PREENCHER MATERIAIS E PREÇOS"
- `_ind_auto_load` → `_worker_ind_auto`: resolve nome, busca blueprint, calcula runs, busca preços em paralelo, preenche tabela, dispara cálculo
- `_get_product_tid`: 3 camadas — nome no item_cache (primário), `bp_data["prod_tid"]`, hoboleaks_products SQLite
- Preço de referência: `buy_max` (maior buy order) do PRODUTO fabricado (não do blueprint)
- Tabela com coluna "Qtd total" (base × runs)

#### Fórmula CCP correta em `calculate_industry_cost`

```
me_combined  = ME_blueprint% + structure_mat_bonus%
qty_final    = ceil(qty_base × (1 - me_combined) × runs)
EIV          = Σ(adjusted_price × qty_final)
job_gross    = EIV × system_cost_index × (1 - structure_job_bonus)
SCC          = EIV × 0.04  ← taxa fixa CCP (4% do EIV, sempre cobrado)
facility     = job_gross × facility_tax
total_job    = job_gross + SCC + facility
total_cost   = materiais + total_job
cost/unit    = total_cost ÷ (qty_per_run × runs)
```

Campos novos na UI: `Estrutura Mat%` e `Estrutura Job%`.

Novos campos no retorno: `job_gross_before`, `job_struct_saving`, `scc_surcharge`, `total_job_cost`, `structure_job_bonus`, `structure_mat_bonus`.

#### Sistema de blueprint materials — 5 camadas

1. Cache SQLite individual (TTL 24h) — O(1), sem rede
2. `_load_local_sql_bz2`: lê `eve_data/industryActivityMaterials.sql.bz2`, extrai 36.193 tuplas com regex, indexa no SQLite
3. Fuzzwork `typematerials.php` (campo correto: `materialTypeID`)
4. Fuzzwork `blueprint/api/blueprint/` (endpoint alternativo)
5. Dados hardcoded (materiais apenas, sem `prod_tid`)

Constante: `_SDE_SQL_BZ2_PATH = os.path.join(DATA_DIR, "industryActivityMaterials.sql.bz2")`

SQLite schema (`blueprint_cache.db`):
- `bp_materials` (lookup_tid, activity_id, fetched_at, materials, qty_per_run, bp_tid)
- `hoboleaks_idx` (bp_tid, activity_id, mat_tid, quantity)
- `hoboleaks_products` (bp_tid, activity_id, prod_tid, qty_per_run)
- `hoboleaks_sync` (key, synced_at)

#### Função `_get_product_tid`

Resolve o type_id do produto a partir do blueprint em 3 camadas:

1. Dedução por nome no `item_cache` — "Nova Heavy Missile Blueprint" → remove " Blueprint" → busca "Nova Heavy Missile" (mais confiável)
2. `bp_data["prod_tid"]` — da Fuzzwork blueprint/api
3. `hoboleaks_products` SQLite

Nota: `_HARDCODED_BP` foi removido desta lógica pois os type_ids hardcoded estavam incorretos para alguns itens.

#### Outras funcionalidades

- HEATMAP: 18 grupos de mercado, scroll bidirecional, drill-down por item
- CATÁLOGO: botão "Abrir em Indústria (calcular custo)" com auto-start
- `fetch_material_prices_bulk`: busca paralela de preços com `ThreadPoolExecutor(max_workers=8)`
- `_ind_fetch_price`: usa `buy_max` ao invés de `sell_min` como referência de preço
- Relatório INDÚSTRIA com breakdown completo: EIV, Job Gross, SCC, Bônus Estrutura, Receita vs Custo, ROI

---

### Performance

| Otimização | Impacto |
|---|---|
| 11 `_sort_*` com `reorder_only=True` | 10-50× mais rápido ao ordenar colunas — usa `move()` em vez de delete+reinsert |
| `load_orders_from_csv` com cache RAM 30min | Não relê CSV a cada execução se carregado recentemente |
| `_warm_hist_cache` O(1) para cache quente | Uma query SQL inicial aquece o cache RAM para todas as subsequentes |
| `_bp_cache` SQLite TTL 24h | Blueprints consultados uma vez ficam disponíveis offline por 24h |
| `hoboleaks_idx` SQLite TTL 7 dias | Após primeiro download do SDE, todos os lookups são offline |
| Fonte aumentada em +2pt (349 ocorrências) | Legibilidade melhorada sem custo de performance |

---

## eve_dashboard.py

**1.155 linhas · 13 classes · 52 métodos · Dashboard Bloomberg-style**

### Classes

| Classe | Descrição |
|---|---|
| `PanelFrame` | Container base para painéis com borda e título |
| `BlinkLabel` | Label com animação de piscar para alertas |
| `ScrollableFrame` | Frame com scrollbar vertical integrada |
| `TickerBar` | Ticker animado horizontal com 14 itens |
| `DBStatusPanel` | Status do banco de dados histórico por hub |
| `WatchlistPanel` | Watchlist live com buy_max e alertas |
| `PaperTradingPanel` | P&L do paper trading com posições abertas |
| `OpportunitiesPanel` | TOP oportunidades via analyze_max_margin |
| `CycleSignalPanel` | Sinais de ciclo via detect_cycle_item |
| `HubComparePanel` | Comparação de 5 hubs × 4 referências |
| `SystemMetricsPanel` | Métricas do sistema (watchlist, DB, config) |
| `StatusBar` | Barra de status inferior com timestamp |
| `EVETerminal` | Janela principal integrando todos os painéis |

### Painéis

- **Ticker animado** — 14 itens com preço live rolando horizontalmente
- **TOP OPORTUNIDADES** — `analyze_max_margin` em background thread
- **WATCHLIST LIVE** — preços buy_max com alertas visuais
- **HUB COMPARISON** — 5 hubs (Jita, Amarr, Dodixie, Rens, Hek) × 4 referências
- **CICLOS DE MERCADO** — `detect_cycle_item` com sinal compra/venda
- **PAPER TRADING P&L** — saldo total e posições abertas
- **DATABASE STATUS** — TTL e cobertura de dados por hub
- **SYSTEM METRICS** — watchlist, configurações, estado do DB

Tecla `F5` abre `eve_trade.py` em processo separado via `subprocess.Popen`.

---

## eve_analyst.py

**1.972 linhas · 5 classes · 40 métodos · Análise multi-aba + IA + PDF/Excel**

### Classes

| Classe | Descrição |
|---|---|
| `MarketDataCollector` | Coleta e agrega dados das 5 análises do eve_trade |
| `AIAnalyst` | Consulta Claude API com system prompt especializado EVE |
| `ExcelReportGenerator` | Gera relatório Excel com openpyxl (8 abas) |
| `PDFReportGenerator` | Gera relatório PDF com ReportLab (A4 landscape) |
| `EVEAnalystApp` | Interface tkinter multi-aba |

### Análises executadas

1. MAX MARGEM — itens com maior margem líquida
2. ESCASSEZ — itens com supply crítico vs demanda
3. DEMANDA/OFERTA — desequilíbrios de mercado
4. MOV. PREÇOS — variações de preço nos últimos 30 dias
5. ORÇAMENTO — alocação de capital por risco

### Conviction Score

```
conviction = (fontes_confirmando / max_possible) × 100
```

- Normalizado por número de análises disponíveis (evita score uniforme)
- Graduado por nível: ALTO ≥ 70, MÉDIO 40-69, BAIXO < 40
- Ação recomendada: COMPRAR / VENDER / MONITORAR

### Relatórios

**Excel (openpyxl):**
- 8 abas com dados das análises individuais
- Formatação condicional por margem e score
- Auto-width em todas as colunas
- Helpers NaN-safe: `_safe_int`, `_safe_pct`, `_safe_vol`, `_safe_dos`, `_safe_float`

**PDF (ReportLab — A4 landscape):**
- Capa com metadados da análise
- Tabelas com alternância de cor por linha
- Seção de análise IA formatada

### Correções

| Bug | Correção |
|---|---|
| `NaN→int` no Excel (7 ocorrências) | Helpers `_safe_*` em todas as conversões |
| Margem 1951% | Filtro `pb > 1`, `net_margin < 500%`, `min_vol=5` |
| Conviction score uniforme 35/100 | Normalizado por `max_possible` |
| `analyze_supply_shortage` broker_rel inválido | Parâmetros removidos |
| `analyze_demand_supply` broker_rel inválido | Parâmetros removidos |
| `analyze_budget` usa `dias_risco` inválido | `risk_mode="Balanced (3d)"` |

---

## Constantes e Configurações Relevantes

```python
DATA_DIR           = os.path.join(os.path.dirname(os.path.abspath(__file__)), "eve_data")
HIST_DB            = os.path.join(DATA_DIR, "market_history.db")
_BP_DB_PATH        = os.path.join(DATA_DIR, "blueprint_cache.db")
_SDE_SQL_BZ2_PATH  = os.path.join(DATA_DIR, "industryActivityMaterials.sql.bz2")
_ORDERS_TTL        = 1800        # 30 min — ordens mudam com frequência
_BP_MATERIALS_TTL  = 86_400     # 24h — blueprints mudam raramente
_HOBOLEAKS_TTL     = 7 * 86_400 # 7 dias — SDE muda só com patches CCP
SCC_RATE           = 0.04       # 4% EIV — taxa fixa CCP, não configurável
```

Regiões disponíveis: `REGIONS` dict com todas as regiões do EVE.

Hubs de referência: `HUB_REGIONS = {"Jita": 10000002, "Amarr": 10000043, ...}`

---

*Gerado automaticamente — EVE Trade Suite*
