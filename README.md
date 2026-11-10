# EVE Trade Suite

Suite de três aplicações Python para análise de mercado e trading no EVE Online.

---

## Aplicações

### `eve_trade.py` — Trade Finder (aplicação principal)
Interface principal com 27 abas cobrindo todas as estratégias de mercado:

| Aba | Função |
|-----|--------|
| Rotas | Encontra rotas lucrativas entre regiões |
| Arbitragem | Oportunidades de compra/venda entre hubs |
| Otimizador | Otimiza cargo com OR-Tools |
| Max Margem | Itens com maior margem líquida |
| Escassez | Itens com supply crítico vs demanda |
| Demanda/Oferta | Desequilíbrios de mercado (DSIF score) |
| Mov. Preços | Variações de preço nos últimos 30 dias |
| Orçamento | Alocação de capital por perfil de risco |
| Mercado Live | Preços em tempo real via ESI |
| Watchlist | Monitoramento de itens com alertas |
| Paper Trading | Simulação de trades sem ISK real |
| AI Analista | Análise de indicadores técnicos com IA |
| Comparador | Comparação de preços entre hubs |
| Configurações | Skills, taxas e preferências |
| Dashboard | Visão geral integrada |
| Multi-Hop | Rotas com múltiplas escalas |
| Correlação | Correlação de preços entre itens |
| Spread Hist. | Histórico de spread por item |
| Ciclos | Detecção de ciclos de mercado |
| Previsão | Previsão de preços (WMA + regressão linear + sazonalidade) |
| Market Making | Oportunidades de formação de mercado |
| Indústria | Cálculo de custo de manufatura com fórmula CCP |
| Backtesting | Teste de estratégias em dados históricos |
| Heatmap | Mapa de calor de 18 grupos de mercado |
| Station Trader | Scanner de oportunidades em estação |
| Catálogo | Catálogo completo de itens |
| Dados & Cache | Gerenciamento de downloads e banco histórico |

### `eve_dashboard.py` — Market Terminal
Dashboard estilo Bloomberg com atualização em tempo real:
- Ticker animado com 14 itens de referência (PLEX, minerais, isótopos)
- Top oportunidades via `analyze_max_margin`
- Watchlist live com alertas visuais
- Comparação de preços entre 5 hubs (Jita, Amarr, Dodixie, Rens, Hek)
- Sinais de ciclo de mercado
- P&L do Paper Trading
- Status do banco de dados histórico

### `eve_analyst.py` — Market Analyst
Análise consolidada multi-fonte com relatórios exportáveis:
- Executa as 5 análises principais em paralelo
- **Conviction Score** (0–100): mede convergência entre análises
- Recomendação: **COMPRAR / VENDER / MONITORAR**
- Exportação para **Excel** (8 abas com formatação condicional)
- Exportação para **PDF** (A4 landscape com ReportLab)
- Integração com **Claude API** para análise em linguagem natural

---

## Requisitos

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

- Python 3.10+ (testado com 3.14)
- `tkinter` do sistema (no Ubuntu/Debian: `sudo apt install python3-tk`)
- Conexão com internet para a [EVE ESI API](https://esi.evetech.net)

> Se o Python do sistema for atualizado, o `venv` pode parar de funcionar (`ModuleNotFoundError`). Nesse caso, apague a pasta `venv/` e repita os comandos acima.

---

## Uso

```bash
source venv/bin/activate

# Aplicação principal
python3 eve_trade.py

# Dashboard (requer eve_trade.py na mesma pasta)
python3 eve_dashboard.py

# Analista com relatórios
python3 eve_analyst.py
```

---

## Estrutura

```
EveTrade/
├── eve_trade.py         # UI principal + motor de análise (27 abas)
├── eve_dashboard.py     # Dashboard Bloomberg-style
├── eve_analyst.py       # Análise multi-fonte + relatórios PDF/Excel
├── eve_market.py        # API pública de análise de mercado
├── eve_utils.py         # Helpers NaN-safe compartilhados
├── eve_route_map.py     # Mapa de rotas
├── eve_shopping_route.py# Rota de compras otimizada
├── requirements.txt     # Dependências Python
└── eve_data/
    ├── watchlist.json   # Itens monitorados
    ├── paper_trades.json# Histórico de paper trades
    └── blacklist.json   # Itens excluídos das análises
```

Os arquivos de cache (`.db`, `.csv`, `cache_items.json`) e o `config.json` são gerados localmente e não estão no repositório.

---

## Configuração da API Key (Claude)

A chave da API Anthropic nunca é gravada em arquivo de texto. Ordem de leitura:

1. variável de ambiente `ANTHROPIC_API_KEY`;
2. keyring **seguro** do sistema operacional (Secret Service/GNOME Keyring, KWallet, macOS Keychain, Windows Credential Locker);
3. sem keyring seguro, a chave digitada fica só na memória da sessão.

Backends inseguros (`keyrings.alt`, incluindo o `PlaintextKeyring`) são recusados. Uma chave antiga encontrada no `config.json` é migrada para o keyring e apagada do arquivo. Configure na aba **AI Analista**.

---

## Cálculo de Custo Industrial

Implementa a fórmula oficial CCP:

```
me_combined  = ME_blueprint% + structure_mat_bonus%
qty_final    = ceil(qty_base × (1 - me_combined) × runs)
EIV          = Σ(adjusted_price × qty_final)
job_gross    = EIV × system_cost_index × (1 - structure_job_bonus)
SCC          = EIV × 0.04   ← 4% fixo CCP
facility     = job_gross × facility_tax
total_cost   = materiais + job_gross + SCC + facility
```

Materiais resolvidos em 5 camadas: cache SQLite → SDE local (`.bz2`) → Fuzzwork typematerials → Fuzzwork blueprint API → hardcoded.

---

## Performance

| Otimização | Impacto |
|------------|---------|
| Cache de ordens em RAM (30 min) | Evita releitura de CSV a cada execução |
| ETag HTTP na ESI | Reduz tráfego em ~60–80% para endpoints estáveis |
| `_warm_hist_cache` O(1) | Uma query SQL aquece o cache RAM para toda a região |
| Cache SQLite de blueprints (24h) | Lookups offline após primeira consulta |
| SDE indexado localmente (7 dias) | 36.193 materiais disponíveis sem rede |
| Sort com `reorder_only=True` | 10–50× mais rápido ao ordenar colunas |
