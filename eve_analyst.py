"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  EVE MARKET ANALYST  ·  Análise Multi-Aba + Relatório IA                   ║
║                                                                              ║
║  Analisa as abas MAX MARGEM, ESCASSEZ, DEMANDA/OFERTA, MOV. PREÇOS          ║
║  e ORÇAMENTO do eve_trade.py e gera um relatório PDF + Excel completo       ║
║  com recomendações de COMPRA/VENDA fundamentadas por IA (Claude API).       ║
║                                                                              ║
║  Uso:  python3 eve_analyst.py                                                ║
║  Req:  eve_trade.py na mesma pasta                                           ║
║        pip install pandas reportlab openpyxl requests                        ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

# ─────────────────────────────────────────────────────────────────────────────
#  Imports standard
# ─────────────────────────────────────────────────────────────────────────────
import sys
import os
import json
import time
import math
import threading
import concurrent.futures
from datetime import datetime, timezone, timedelta
from pathlib import Path

# ─────────────────────────────────────────────────────────────────────────────
#  Resolve o módulo eve_trade (deve estar na mesma pasta)
# ─────────────────────────────────────────────────────────────────────────────
_HERE = Path(__file__).parent.resolve()
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

try:
    import eve_trade as et
    import eve_market as em
except ImportError as _e:
    print(f"[ERRO] Não foi possível importar eve_trade.py / eve_market.py: {_e}")
    print("       Execute eve_analyst.py na mesma pasta que eve_trade.py")
    sys.exit(1)

# ─────────────────────────────────────────────────────────────────────────────
#  Imports de análise e relatório
# ─────────────────────────────────────────────────────────────────────────────
try:
    import pandas as pd
except ImportError:
    print("[ERRO] pandas não instalado. Execute: pip install pandas")
    sys.exit(1)

try:
    import requests
except ImportError:
    print("[ERRO] requests não instalado. Execute: pip install requests")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                                  GradientFill)
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, Reference
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    print("[AVISO] openpyxl não instalado. Relatório Excel desativado.")

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm, mm
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                     Table, TableStyle, PageBreak, HRFlowable,
                                     KeepTogether)
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False
    print("[AVISO] reportlab não instalado. Relatório PDF desativado.")
    print("        Execute: pip install reportlab")

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ─────────────────────────────────────────────────────────────────────────────
#  Constantes
# ─────────────────────────────────────────────────────────────────────────────
VERSION      = "1.0"
APP_TITLE    = "EVE Market Analyst"
ANTHROPIC_URL= "https://api.anthropic.com/v1/messages"
CLAUDE_MODEL = "claude-sonnet-4-6"

# Paleta visual consistente com eve_trade
COLORS = {
    "bg":        "#0a0e1a",
    "panel":     "#0f1624",
    "border":    "#1e2d45",
    "accent":    "#00aaff",
    "green":     "#00ff88",
    "red":       "#ff4455",
    "yellow":    "#ffd700",
    "orange":    "#ff8c00",
    "text":      "#c8d8e8",
    "subtext":   "#5a7a9a",
    "header":    "#7ecfff",
    "row_odd":   "#0d1520",
    "row_even":  "#111d2e",
}

# Cores para relatório PDF
PDF_DARK   = colors.HexColor("#0a0e1a")
PDF_PANEL  = colors.HexColor("#0f1624")
PDF_ACCENT = colors.HexColor("#00aaff")
PDF_GREEN  = colors.HexColor("#00cc66")
PDF_RED    = colors.HexColor("#ff4455")
PDF_YELLOW = colors.HexColor("#cc9900")
PDF_TEXT   = colors.HexColor("#1a2a3a")
PDF_GRAY   = colors.HexColor("#4a5a6a")
PDF_LIGHT  = colors.HexColor("#e8f0f8")
PDF_MID    = colors.HexColor("#c0d0e0")


# ─────────────────────────────────────────────────────────────────────────────
#  1. COLETA E PROCESSAMENTO DE DADOS
#     Executa as 5 análises e consolida em DataFrames pandas
# ─────────────────────────────────────────────────────────────────────────────

from eve_utils import (safe_pct  as _safe_pct,
                       safe_int  as _safe_int,
                       safe_vol  as _safe_vol,
                       safe_dos  as _safe_dos,
                       safe_float as _safe_float)


class MarketDataCollector:
    """
    Responsável por coletar e consolidar dados de mercado do eve_trade.py.

    Executa as 5 análises (MAX MARGEM, ESCASSEZ, DEMANDA/OFERTA, MOV. PREÇOS,
    ORÇAMENTO) para o hub selecionado e retorna DataFrames padronizados.
    Todas as chamadas às funções de análise são feitas em thread separada
    para não bloquear a UI.
    """

    def __init__(self, region_name: str, budget: float = 1_000_000_000,
                 progress_cb=None):
        """
        Parâmetros:
          region_name   — Nome da região como em REGIONS dict do eve_trade
          budget        — Capital disponível para análise de orçamento (ISK)
          progress_cb   — Callback de progresso: fn(msg: str) ou fn(pct, total, msg)
        """
        self.region_name = region_name
        self.region_id   = et.REGIONS.get(region_name, 10000002)
        self.budget      = budget
        self._progress   = progress_cb or (lambda *a: None)

        # Carrega configurações e caches locais
        self.cfg        = et.load_config()
        self.item_cache = et.load_item_cache()
        self.br         = int(self.cfg.get("g_brel", 0))
        self.ac         = int(self.cfg.get("g_acc",  0))

        # DataFrames resultantes
        self.df_mm:  pd.DataFrame = pd.DataFrame()   # MAX MARGEM
        self.df_sc:  pd.DataFrame = pd.DataFrame()   # ESCASSEZ
        self.df_ds:  pd.DataFrame = pd.DataFrame()   # DEMANDA/OFERTA
        self.df_pm:  pd.DataFrame = pd.DataFrame()   # MOV. PREÇOS
        self.df_bud: pd.DataFrame = pd.DataFrame()   # ORÇAMENTO
        self.df_master: pd.DataFrame = pd.DataFrame()# Consolidado

        # Metadados
        self.timestamp = datetime.now(timezone.utc)
        self.errors: list = []

    # ── Passo 1: carrega ordens do CSV ou API ─────────────────────────────────
    def _load_orders(self) -> list:
        self._progress("Carregando ordens de mercado...")
        orders = et.load_orders_from_csv(self.region_name)
        if not orders:
            self._progress(f"CSV não encontrado — baixando da ESI para {self.region_name}...")
            try:
                orders = em.load_orders(self.region_name)
            except Exception as e:
                self.errors.append(f"Erro ao baixar ordens: {e}")
                return []
        self._progress(f"  {len(orders):,} ordens carregadas.")
        return orders

    # ── Passo 2: executa as 5 análises ────────────────────────────────────────
    def collect(self):
        """
        Ponto de entrada principal.
        Coleta dados de todas as 5 abas e consolida no df_master.
        """
        orders = self._load_orders()
        if not orders and not et._db_item_count(self.region_id):
            self.errors.append(
                "Sem dados disponíveis. Baixe ordens e sincronize o banco "
                "histórico na aba DADOS & CACHE do EVE Trade Finder.")
            return

        # Pré-aquece cache de histórico (1 query SQL para toda a região)
        self._progress("Pré-aquecendo cache de histórico...")
        try:
            et._warm_hist_cache(self.region_id, self._progress)
        except Exception as e:
            self.errors.append(f"Aviso cache: {e}")

        # ── MAX MARGEM ────────────────────────────────────────────────────────
        self._progress("Analisando MAX MARGEM...")
        try:
            # min_vol=5: filtra itens com menos de 5 unidades de buy/sell
            # min_margin=2.0: filtra margem líquida mínima real (evita ruído)
            # Nota: analyze_max_margin já filtra if not sells or not buys,
            # mas itens com pb muito baixo (1 ISK) geram margens irreais ~1951%.
            # max_sell_price=0 = sem limite superior.
            mm_raw = em.analyze_max_margin(
                orders, self.region_name, self.item_cache,
                broker_rel=self.br, accounting=self.ac,
                min_margin=2.0,
                min_vol=5,          # filtra baixíssima liquidez
                top_n=100)
            # Pós-filtro: remove entradas com pb=0 ou pb muito baixo
            # (itens com buy_orders que têm preço < 0.01% do sell são ruído)
            mm_raw = [r for r in mm_raw
                      if r.get("pb", 0) > 1
                      and r.get("net_margin", 0) < 500   # margem > 500% = dado inválido
                      and r.get("buy_vol", 0) >= 1]
            self.df_mm = pd.DataFrame(mm_raw)
            if not self.df_mm.empty:
                # Seleciona e renomeia colunas relevantes
                cols = {
                    "type_id":     "type_id",
                    "item":        "item",
                    "ps":          "sell_min",
                    "pb":          "buy_max",
                    "gross_margin":"gross_margin_pct",
                    "net_margin":  "net_margin_pct",
                    "profit_unit": "lucro_por_unidade",
                    "sell_vol":    "vol_sell",
                    "buy_vol":     "vol_buy",
                    "edp":         "edp",
                    "avg_vol_90":  "vol_90d_dia",
                }
                self.df_mm = self.df_mm[[c for c in cols if c in self.df_mm.columns]].rename(columns=cols)
                self.df_mm["fonte"] = "MAX_MARGEM"
            self._progress(f"  MAX MARGEM: {len(self.df_mm)} itens")
        except Exception as e:
            self.errors.append(f"Erro MAX MARGEM: {e}")

        # ── ESCASSEZ ──────────────────────────────────────────────────────────
        self._progress("Analisando ESCASSEZ...")
        try:
            # analyze_supply_shortage NÃO aceita broker_rel/accounting —
            # usa apenas critério de volume vs oferta disponível (DoS)
            sc_raw = em.analyze_supply_shortage(
                orders, self.region_name, self.item_cache,
                max_dos=30.0, min_vol=1, top_n=100)
            self.df_sc = pd.DataFrame(sc_raw)
            if not self.df_sc.empty:
                cols = {
                    "type_id":    "type_id",
                    "item":       "item",
                    "ps":         "sell_min",
                    "pb":         "buy_max",
                    "margin":     "margem_pct",
                    "sell_vol":   "vol_sell",
                    "avg_vol_90": "vol_90d_dia",
                    "dos":        "dias_oferta",
                    "score":      "score_escassez",
                }
                self.df_sc = self.df_sc[[c for c in cols if c in self.df_sc.columns]].rename(columns=cols)
                self.df_sc["fonte"] = "ESCASSEZ"
            self._progress(f"  ESCASSEZ: {len(self.df_sc)} itens")
        except Exception as e:
            self.errors.append(f"Erro ESCASSEZ: {e}")

        # ── DEMANDA/OFERTA ────────────────────────────────────────────────────
        self._progress("Analisando DEMANDA/OFERTA...")
        try:
            # analyze_demand_supply NÃO aceita broker_rel/accounting —
            # calcula DSIF score baseado em volume buy vs sell e preços
            ds_raw = em.analyze_demand_supply(
                orders, self.region_name, self.item_cache,
                min_vol=1, top_n=100)
            self.df_ds = pd.DataFrame(ds_raw)
            if not self.df_ds.empty:
                cols = {
                    "type_id":    "type_id",
                    "item":       "item",
                    "ps":         "sell_min",
                    "pb":         "buy_max",
                    "margin":     "margem_pct",
                    "buy_vol":    "vol_buy",
                    "sell_vol":   "vol_sell",
                    "ratio":      "ratio_demand_supply",
                    "avg_vol_90": "vol_90d_dia",
                    "dsif":       "dsif_score",
                }
                self.df_ds = self.df_ds[[c for c in cols if c in self.df_ds.columns]].rename(columns=cols)
                self.df_ds["fonte"] = "DEMANDA_OFERTA"
            self._progress(f"  DEMANDA/OFERTA: {len(self.df_ds)} itens")
        except Exception as e:
            self.errors.append(f"Erro DEMANDA/OFERTA: {e}")

        # ── MOV. PREÇOS ───────────────────────────────────────────────────────
        self._progress("Analisando MOV. PREÇOS (banco histórico)...")
        try:
            # Usa o banco histórico diretamente (mais rico que CSV)
            db_count = et._db_item_count(self.region_id)
            if db_count > 0:
                conn     = et._get_db()
                rows     = conn.execute(
                    "SELECT DISTINCT type_id FROM history WHERE region_id=?",
                    (self.region_id,)).fetchall()
                tids_db  = [r[0] for r in rows]
                pm_raw   = em.analyze_price_movement(
                    self.region_name, self.item_cache,
                    min_vol=3, min_margin=0.0, top_n=100,
                    orders=None, type_ids=tids_db)
            else:
                pm_raw   = em.analyze_price_movement(
                    self.region_name, self.item_cache,
                    min_vol=3, min_margin=0.0, top_n=100,
                    orders=orders)
            self.df_pm = pd.DataFrame(pm_raw)
            if not self.df_pm.empty:
                cols = {
                    "type_id":     "type_id",
                    "item":        "item",
                    "avg_p_week":  "preco_7d",
                    "avg_p_month": "preco_30d",
                    "price_chg":   "var_preco_pct",
                    "avg_v_week":  "vol_7d_dia",
                    "avg_v_month": "vol_30d_dia",
                    "vol_chg":     "var_vol_pct",
                    "margin":      "margem_pct",
                    "correlation": "correlacao",
                    "avg_vol_90":  "vol_90d_dia",
                    "score":       "score_pm",
                }
                self.df_pm = self.df_pm[[c for c in cols if c in self.df_pm.columns]].rename(columns=cols)
                self.df_pm["fonte"] = "MOV_PRECOS"
            self._progress(f"  MOV. PREÇOS: {len(self.df_pm)} itens")
        except Exception as e:
            self.errors.append(f"Erro MOV. PREÇOS: {e}")

        # ── ORÇAMENTO ─────────────────────────────────────────────────────────
        self._progress("Analisando ORÇAMENTO...")
        try:
            # analyze_budget retorna (basket, stats) — tupla de 2 elementos.
            # Parâmetro correto: risk_mode="Balanced (3d)" (string, não dias_risco int).
            # Opções: "Quick Flip (1d)", "Balanced (3d)", "Deep Stock (7d)"
            bud_result = em.analyze_budget(
                orders, self.region_name, self.item_cache,
                budget=self.budget,
                broker_rel=self.br,
                accounting=self.ac,
                risk_mode="Balanced (3d)",
                min_vol=1)
            # analyze_budget retorna TUPLA (basket: list, stats: dict)
            if isinstance(bud_result, tuple) and len(bud_result) == 2:
                bud_raw, bud_stats = bud_result
            else:
                bud_raw, bud_stats = bud_result, {}
            self.df_bud = pd.DataFrame(bud_raw)
            if not self.df_bud.empty:
                cols = {
                    "type_id":     "type_id",
                    "item":        "item",
                    "pb":          "buy_max",
                    "ps":          "sell_min",
                    "qty":         "qtd_recomendada",
                    "u_max":       "unidades_max",
                    "cost_unit":   "custo_unitario",
                    "pi_net":      "lucro_unitario_liq",
                    "net_margin":  "net_margin_pct",
                    "total_buy":   "custo_total",
                    "total_profit":"lucro_total_est",
                    "avg_vol_90":  "vol_90d_dia",
                }
                self.df_bud = self.df_bud[[c for c in cols if c in self.df_bud.columns]].rename(columns=cols)
                self.df_bud["fonte"] = "ORCAMENTO"
                # Bud stats como metadata
                self.bud_stats = bud_stats
            else:
                self.bud_stats = {}
            self._progress(f"  ORÇAMENTO: {len(self.df_bud)} itens")
        except Exception as e:
            self.errors.append(f"Erro ORÇAMENTO: {e}")
            self.bud_stats = {}

        # ── CONSOLIDAÇÃO ──────────────────────────────────────────────────────
        self._progress("Consolidando dados...")
        self._build_master()
        self._progress(f"Coleta concluída. {len(self.df_master)} itens únicos no master.")

    def _build_master(self):
        """
        Cria DataFrame consolidado com 1 linha por item, combinando
        dados de todas as 5 análises via merge em type_id.

        Métricas de síntese calculadas:
          conviction_score — pontuação 0-100 baseada em quantas abas
                             convergem na mesma recomendação para o item.
          acao_recomendada — COMPRAR / VENDER / MONITORAR
        """
        if self.df_mm.empty and self.df_sc.empty:
            return

        # Base: todos os itens que aparecem em pelo menos 1 análise
        all_tids = set()
        for df in [self.df_mm, self.df_sc, self.df_ds, self.df_pm, self.df_bud]:
            if not df.empty and "type_id" in df.columns:
                all_tids.update(df["type_id"].tolist())

        rows = []
        for tid in all_tids:
            name = self.item_cache.get(tid, {}).get("name", str(tid))
            row  = {"type_id": tid, "item": name}

            # Pega campos de cada análise
            mm_row  = self._get_row(self.df_mm,  tid)
            sc_row  = self._get_row(self.df_sc,  tid)
            ds_row  = self._get_row(self.df_ds,  tid)
            pm_row  = self._get_row(self.df_pm,  tid)
            bud_row = self._get_row(self.df_bud, tid)

            # Métricas principais (pega da fonte mais confiável)
            row["sell_min"]       = self._first(mm_row, sc_row, ds_row, bud_row, key="sell_min")
            row["buy_max"]        = self._first(mm_row, sc_row, ds_row, bud_row, key="buy_max")
            row["net_margin_pct"] = mm_row.get("net_margin_pct")
            row["gross_margin"]   = mm_row.get("gross_margin_pct")
            row["lucro_unitario"] = mm_row.get("lucro_por_unidade")
            row["edp"]            = mm_row.get("edp")
            row["vol_90d_dia"]    = self._first(mm_row, sc_row, ds_row, pm_row, bud_row,
                                                key="vol_90d_dia")
            row["dias_oferta"]    = sc_row.get("dias_oferta")
            row["score_escassez"] = sc_row.get("score_escassez")
            row["dsif_score"]     = ds_row.get("dsif_score")
            row["ratio_demand"]   = ds_row.get("ratio_demand_supply")
            row["var_preco_pct"]  = pm_row.get("var_preco_pct")
            row["var_vol_pct"]    = pm_row.get("var_vol_pct")
            row["preco_7d"]       = pm_row.get("preco_7d")
            row["preco_30d"]      = pm_row.get("preco_30d")
            row["qtd_orcamento"]  = bud_row.get("qtd_recomendada")
            row["custo_total_orc"]= bud_row.get("custo_total")
            row["lucro_total_orc"]= bud_row.get("lucro_total_est")

            # Quantas abas colocaram este item
            fontes = sum([
                1 if mm_row  else 0,
                1 if sc_row  else 0,
                1 if ds_row  else 0,
                1 if pm_row  else 0,
                1 if bud_row else 0,
            ])
            row["n_fontes"] = fontes

            # ── Conviction Score (0–100) ──────────────────────────────────────
            # Pontuação máxima teórica: 90 pts. Normaliza pelo número de
            # análises disponíveis para evitar scores uniformes quando
            # poucas abas retornam dados (ex: sem histórico BD → sem pm_row)
            score = 0.0
            max_possible = 0.0

            # +25 por margem líquida ≥ 5% (MAX MARGEM)
            if mm_row:
                max_possible += 35   # 25 + 10 (EDP)
                nm_val = mm_row.get("net_margin_pct", 0) or 0
                if not (isinstance(nm_val, float) and math.isnan(nm_val)):
                    if nm_val >= 10:   score += 35
                    elif nm_val >= 5:  score += 25
                    elif nm_val >= 2:  score += 10
                edp_val = mm_row.get("edp", 0) or 0
                if not (isinstance(edp_val, float) and math.isnan(edp_val)):
                    if edp_val > 50_000_000:   score += 10
                    elif edp_val > 10_000_000: score += 6
                    elif edp_val > 1_000_000:  score += 3

            # +20 por escassez alta (ESCASSEZ)
            if sc_row:
                max_possible += 20
                dos_val = sc_row.get("dias_oferta")
                if dos_val is not None and not (isinstance(dos_val, float) and math.isnan(dos_val)):
                    if dos_val <= 1:   score += 20
                    elif dos_val <= 3: score += 15
                    elif dos_val <= 7: score += 8

            # +20 por demanda > oferta (DEMANDA/OFERTA)
            if ds_row:
                max_possible += 20
                ratio_val = ds_row.get("ratio_demand_supply")
                dsif_val  = ds_row.get("dsif_score", 0) or 0
                if ratio_val is not None and not (isinstance(ratio_val, float) and math.isnan(ratio_val)):
                    if ratio_val > 2.0:  score += 20
                    elif ratio_val > 1.5: score += 14
                    elif ratio_val > 1.0: score += 7

            # +15 por preço subindo (MOV. PREÇOS)
            if pm_row:
                max_possible += 25   # 15 preço + 10 volume
                var_val = pm_row.get("var_preco_pct")
                if var_val is not None and not (isinstance(var_val, float) and math.isnan(var_val)):
                    if var_val > 10:   score += 15
                    elif var_val > 2:  score += 10
                    elif var_val < -5: score -= 15
                    elif var_val < -2: score -= 8
                vol_chg = pm_row.get("var_vol_pct")
                if vol_chg is not None and not (isinstance(vol_chg, float) and math.isnan(vol_chg)):
                    if vol_chg > 20:  score += 10
                    elif vol_chg > 5: score += 5

            # Normaliza para 0–100 baseado no máximo possível das análises disponíveis
            if max_possible > 0:
                score = min(100, max(0, score / max_possible * 100))
            else:
                score = 0

            row["conviction_score"] = round(score, 1)

            # ── Ação Recomendada ──────────────────────────────────────────────
            nm = mm_row.get("net_margin_pct", 0) or 0
            dos = sc_row.get("dias_oferta")
            var = pm_row.get("var_preco_pct", 0) or 0

            if nm >= 8 and score >= 50:
                row["acao"] = "COMPRAR"
                row["prioridade"] = "ALTA" if score >= 70 else "MEDIA"
            elif nm >= 4 and (dos is None or dos <= 5) and var >= -3:
                row["acao"] = "COMPRAR"
                row["prioridade"] = "MEDIA"
            elif var < -8 or (dos is not None and dos > 30):
                row["acao"] = "VENDER"
                row["prioridade"] = "ALTA" if var < -15 else "MEDIA"
            else:
                row["acao"] = "MONITORAR"
                row["prioridade"] = "BAIXA"

            rows.append(row)

        self.df_master = pd.DataFrame(rows)
        # Sanitiza NaN/Inf em colunas numéricas — evita "+nan%" no relatório
        numeric_cols = ["net_margin_pct", "gross_margin", "lucro_unitario", "edp",
                        "vol_90d_dia", "dias_oferta", "score_escassez", "dsif_score",
                        "ratio_demand", "var_preco_pct", "var_vol_pct", "preco_7d",
                        "preco_30d", "qtd_orcamento", "custo_total_orc",
                        "lucro_total_orc", "conviction_score"]
        for col in numeric_cols:
            if col in self.df_master.columns:
                self.df_master[col] = pd.to_numeric(
                    self.df_master[col], errors="coerce")

        # Ordena por conviction_score desc, depois por acao
        if not self.df_master.empty:
            self.df_master = self.df_master.sort_values(
                ["conviction_score", "net_margin_pct"],
                ascending=[False, False]).reset_index(drop=True)

    def _get_row(self, df: pd.DataFrame, tid: int) -> dict:
        """Retorna a linha de um DataFrame para um type_id, ou {} se ausente."""
        if df.empty or "type_id" not in df.columns:
            return {}
        matches = df[df["type_id"] == tid]
        return matches.iloc[0].to_dict() if len(matches) > 0 else {}

    def _first(self, *rows, key: str):
        """Retorna o primeiro valor não-None/NaN para uma chave nos dicts fornecidos."""
        for row in rows:
            v = row.get(key)
            if v is not None and not (isinstance(v, float) and math.isnan(v)):
                return v
        return None

    def summary_text(self) -> str:
        """Resumo textual dos dados coletados para usar no prompt da IA."""
        if self.df_master.empty:
            return "Sem dados disponíveis."

        # Top 30 itens por conviction_score
        top = self.df_master.head(30)

        lines = [
            f"ANÁLISE DE MERCADO — {self.region_name}",
            f"Data: {self.timestamp.strftime('%Y-%m-%d %H:%M UTC')}",
            f"Capital disponível: {self.budget/1e9:.2f}B ISK",
            f"Skills: Broker Relations {self.br} / Accounting {self.ac}",
            "",
            f"Total de itens analisados: {len(self.df_master)}",
            f"Itens COMPRAR: {(self.df_master['acao']=='COMPRAR').sum()}",
            f"Itens VENDER: {(self.df_master['acao']=='VENDER').sum()}",
            f"Itens MONITORAR: {(self.df_master['acao']=='MONITORAR').sum()}",
            "",
            "─" * 70,
            "TOP 30 ITENS (por Conviction Score):",
            "",
        ]

        for _, r in top.iterrows():
            nm  = _safe_pct(r.get('net_margin_pct'), '.1f')
            edp = r.get("edp")
            edp_str = f"{edp/1e6:.1f}M ISK/dia" if (edp and not (isinstance(edp, float) and math.isnan(edp))) else "N/D"
            dos = r.get("dias_oferta")
            dos_str = f"{dos:.1f}d" if (dos is not None and not (isinstance(dos, float) and math.isnan(dos))) else "N/D"
            var = r.get("var_preco_pct")
            var_str = f"{var:+.1f}%" if (var is not None and not (isinstance(var, float) and math.isnan(var))) else "N/D"
            vol = r.get("vol_90d_dia")
            vol_str = f"{vol:.0f}" if (vol is not None and not (isinstance(vol, float) and math.isnan(vol))) else "N/D"
            sell = r.get("sell_min")
            sell_str = f"{sell/1e6:.2f}M" if (sell and sell >= 1e6) else (f"{sell:,.0f}" if sell else "N/D")

            lines.append(
                f"[{r['acao']:<9}|Score:{_safe_int(r.get('conviction_score'))}|"
                f"Prio:{r.get('prioridade','?'):<6}] "
                f"{str(r['item'])[:30]:<30} | "
                f"Sell:{sell_str:<12} | Net:{nm:<7} | "
                f"EDP:{edp_str:<14} | DoS:{dos_str:<6} | "
                f"Var:{var_str:<7} | Vol/dia:{vol_str}"
            )

        # Métricas de mercado
        if not self.df_mm.empty and "net_margin_pct" in self.df_mm.columns:
            lines += [
                "",
                "─" * 70,
                "ESTATÍSTICAS DE MERCADO:",
                f"  Margem líquida média: {self.df_mm['net_margin_pct'].mean():.2f}%",
                f"  Margem líquida máx:  {self.df_mm['net_margin_pct'].max():.2f}%",
                f"  Itens c/ margem ≥10%: {(self.df_mm['net_margin_pct']>=10).sum()}",
            ]
        if not self.df_sc.empty and "dias_oferta" in self.df_sc.columns:
            lines += [
                f"  Itens c/ escassez ≤3d: {(self.df_sc['dias_oferta']<=3).sum()}",
                f"  DoS médio (escassez):  {self.df_sc['dias_oferta'].mean():.1f} dias",
            ]
        if not self.df_pm.empty and "var_preco_pct" in self.df_pm.columns:
            subindo = (self.df_pm["var_preco_pct"] > 0).sum()
            caindo  = (self.df_pm["var_preco_pct"] < 0).sum()
            lines += [
                f"  Preços subindo (7d→30d): {subindo} itens",
                f"  Preços caindo  (7d→30d): {caindo} itens",
            ]

        return "\n".join(lines)


# ─────────────────────────────────────────────────────────────────────────────
#  2. ANÁLISE IA — Claude API
# ─────────────────────────────────────────────────────────────────────────────

class AIAnalyst:
    """
    Consulta a Claude API com os dados consolidados para gerar análise
    qualitativa e recomendações estratégicas detalhadas.
    """

    SYSTEM_PROMPT = """Você é um analista especialista em mercado do jogo EVE Online.
Recebe dados quantitativos de mercado de um hub específico e gera análises
estratégicas detalhadas para traders humanos.

Seu relatório deve:
1. Identificar as MELHORES oportunidades de compra com justificativas quantitativas
2. Identificar itens que devem ser VENDIDOS ou evitados
3. Analisar tendências macro do mercado no hub
4. Recomendar estratégia de alocação do capital disponível
5. Destacar riscos e alertas específicos
6. Usar terminologia EVE Online corretamente (ISK, Station Trading, Market Making, EDP, DoS)

Seja objetivo, direto e use os dados fornecidos como base primária das recomendações.
Formato: texto estruturado com seções claramente demarcadas. Use ISK com abreviações
(M=milhões, B=bilhões). Limite a resposta a ~1500 palavras."""

    def __init__(self, api_key: str):
        self.api_key = api_key

    def analyze(self, market_summary: str, region_name: str,
                budget: float, progress_cb=None) -> str:
        """
        Envia os dados para a API Claude e retorna a análise completa.

        Parâmetros:
          market_summary — Texto gerado por MarketDataCollector.summary_text()
          region_name    — Nome do hub analisado
          budget         — Capital disponível em ISK
          progress_cb    — Callback de progresso opcional
        """
        if not self.api_key:
            return "[AVISO] API Key da Anthropic não configurada. Configure em eve_trade.py → aba CONFIGURAÇÕES → AI ANALISTA."

        if progress_cb:
            progress_cb("Enviando dados para Claude AI...")

        prompt = f"""Analise os dados de mercado a seguir do hub {region_name} no EVE Online
e gere um relatório estratégico completo de trading.

Capital disponível para investimento: {budget/1e9:.2f}B ISK

{market_summary}

---

Com base nesses dados, gere:

## RESUMO EXECUTIVO
(3-4 frases descrevendo o estado geral do mercado no hub)

## TOP 5 OPORTUNIDADES DE COMPRA
(Para cada item: nome, preço, margem líquida, EDP, justificativa, quantidade recomendada)

## ITENS PARA VENDER / EVITAR
(Itens com preço caindo, escassez excessiva de compradores ou margem negativa)

## ANÁLISE DE MERCADO MACRO
(Tendências, sazonalidade, pressão de preços no hub)

## ESTRATÉGIA DE ALOCAÇÃO DE CAPITAL
(Como distribuir o capital entre os melhores itens, com percentuais)

## ALERTAS E RISCOS
(Volatilidade alta, concentração de mercado, itens a monitorar)

## SCORE GERAL DO MERCADO
(Nota 0-100 com justificativa: oportunidades vs. riscos)"""

        try:
            resp = requests.post(
                ANTHROPIC_URL,
                headers={
                    "x-api-key":          self.api_key,
                    "anthropic-version":  "2023-06-01",
                    "content-type":       "application/json",
                },
                json={
                    "model":      CLAUDE_MODEL,
                    "max_tokens": 2000,
                    "system":     self.SYSTEM_PROMPT,
                    "messages":   [{"role": "user", "content": prompt}],
                },
                timeout=60,
            )
            if resp.status_code == 200:
                data = resp.json()
                text = data.get("content", [{}])[0].get("text", "")
                if progress_cb:
                    progress_cb("Análise IA concluída.")
                return text
            else:
                err = resp.json().get("error", {}).get("message", resp.text[:200])
                return f"[ERRO API] HTTP {resp.status_code}: {err}"
        except requests.Timeout:
            return "[ERRO] Timeout ao contatar a API Claude. Tente novamente."
        except Exception as e:
            return f"[ERRO] {e}"


# ─────────────────────────────────────────────────────────────────────────────
#  3. GERADOR DE RELATÓRIO EXCEL
# ─────────────────────────────────────────────────────────────────────────────

class ExcelReportGenerator:
    """
    Gera relatório Excel profissional com múltiplas abas:
      - Resumo Executivo
      - MAX MARGEM
      - ESCASSEZ
      - DEMANDA/OFERTA
      - MOV. PREÇOS
      - ORÇAMENTO
      - Análise IA
      - Dados Brutos
    """

    # Paleta EVE para Excel
    DARK_BG   = "0A0E1A"
    HEADER_BG = "0F1624"
    ACCENT    = "00AAFF"
    GREEN     = "00CC66"
    RED       = "FF4455"
    YELLOW    = "FFD700"
    ORANGE    = "FF8C00"
    TEXT_DARK = "1A2A3A"
    LIGHT_BG  = "E8F0F8"
    MID_BG    = "C8D8E8"
    WHITE     = "FFFFFF"

    def __init__(self, collector: 'MarketDataCollector',
                 ai_text: str = "", progress_cb=None):
        self.c          = collector
        self.ai_text    = ai_text
        self._progress  = progress_cb or (lambda *a: None)

    def _header_font(self, bold=True, size=11, color=None):
        return Font(bold=bold, size=size,
                    color=color or self.WHITE, name="Calibri")

    def _cell_font(self, bold=False, size=10, color=None):
        return Font(bold=bold, size=size,
                    color=color or self.TEXT_DARK, name="Calibri")

    def _fill(self, hex_color: str):
        return PatternFill("solid", fgColor=hex_color)

    def _border(self):
        thin = Side(style="thin", color="CCCCCC")
        return Border(left=thin, right=thin, top=thin, bottom=thin)

    def _acao_color(self, acao: str) -> str:
        return {"COMPRAR": self.GREEN, "VENDER": self.RED,
                "MONITORAR": self.YELLOW}.get(acao, self.MID_BG)

    def _fmt_isk(self, v, abbrev=True) -> str:
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return "--"
        if abbrev:
            if abs(v) >= 1e9: return f"{v/1e9:.2f}B"
            if abs(v) >= 1e6: return f"{v/1e6:.1f}M"
            if abs(v) >= 1e3: return f"{v/1e3:.0f}K"
        return f"{v:,.2f}"

    def _auto_width(self, ws):
        """Ajusta largura das colunas automaticamente."""
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    max_len = max(max_len, len(str(cell.value or "")))
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = min(50, max(10, max_len + 2))

    def generate(self, output_path: str):
        """Gera o arquivo Excel completo no caminho especificado."""
        if not HAS_OPENPYXL:
            raise ImportError("openpyxl não instalado.")

        wb = openpyxl.Workbook()
        wb.remove(wb.active)   # remove aba em branco padrão

        self._progress("Excel: criando aba Resumo...")
        self._sheet_summary(wb)

        if not self.c.df_mm.empty:
            self._progress("Excel: aba MAX MARGEM...")
            self._sheet_analysis(wb, self.c.df_mm, "MAX MARGEM",
                                  key_cols=["item","sell_min","buy_max",
                                            "net_margin_pct","lucro_por_unidade",
                                            "edp","vol_90d_dia"])

        if not self.c.df_sc.empty:
            self._progress("Excel: aba ESCASSEZ...")
            self._sheet_analysis(wb, self.c.df_sc, "ESCASSEZ",
                                  key_cols=["item","sell_min","buy_max",
                                            "margem_pct","vol_90d_dia",
                                            "dias_oferta","score_escassez"])

        if not self.c.df_ds.empty:
            self._progress("Excel: aba DEMANDA OFERTA...")
            self._sheet_analysis(wb, self.c.df_ds, "DEMANDA OFERTA",
                                  key_cols=["item","sell_min","buy_max",
                                            "margem_pct","vol_buy","vol_sell",
                                            "ratio_demand_supply","dsif_score"])

        if not self.c.df_pm.empty:
            self._progress("Excel: aba MOV PRECOS...")
            self._sheet_analysis(wb, self.c.df_pm, "MOV PRECOS",
                                  key_cols=["item","preco_7d","preco_30d",
                                            "var_preco_pct","vol_7d_dia",
                                            "vol_30d_dia","var_vol_pct","score_pm"])

        if not self.c.df_bud.empty:
            self._progress("Excel: aba ORCAMENTO...")
            self._sheet_analysis(wb, self.c.df_bud, "ORCAMENTO",
                                  key_cols=["item","buy_max","sell_min",
                                            "qtd_recomendada","custo_total",
                                            "lucro_total_est","net_margin_pct"])

        if not self.c.df_master.empty:
            self._progress("Excel: aba RECOMENDACOES...")
            self._sheet_recomendacoes(wb)

        if self.ai_text:
            self._progress("Excel: aba ANALISE IA...")
            self._sheet_ai(wb)

        self._progress("Excel: salvando arquivo...")
        wb.save(output_path)
        self._progress(f"Excel salvo: {output_path}")

    def _sheet_summary(self, wb):
        """Aba de resumo executivo com cards de métricas."""
        ws = wb.create_sheet("RESUMO")
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 2

        # Título
        ws.merge_cells("B1:J1")
        ws["B1"] = f"EVE MARKET ANALYST — {self.c.region_name}"
        ws["B1"].font   = Font(bold=True, size=16, color=self.WHITE, name="Calibri")
        ws["B1"].fill   = self._fill(self.DARK_BG)
        ws["B1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 36

        # Subtítulo
        ws.merge_cells("B2:J2")
        ws["B2"] = f"Gerado em {self.c.timestamp.strftime('%d/%m/%Y %H:%M UTC')} | Capital: {self._fmt_isk(self.c.budget)} ISK"
        ws["B2"].font   = Font(size=10, color=self.ACCENT, name="Calibri")
        ws["B2"].fill   = self._fill(self.HEADER_BG)
        ws["B2"].alignment = Alignment(horizontal="center")

        # Cards de métricas
        metrics = []
        if not self.c.df_master.empty:
            comprar  = (self.c.df_master["acao"] == "COMPRAR").sum()
            vender   = (self.c.df_master["acao"] == "VENDER").sum()
            monitorar= (self.c.df_master["acao"] == "MONITORAR").sum()
            metrics += [
                ("COMPRAR",   str(comprar),  self.GREEN),
                ("VENDER",    str(vender),   self.RED),
                ("MONITORAR", str(monitorar),self.YELLOW),
                ("TOTAL ITENS",str(len(self.c.df_master)), self.ACCENT),
            ]
        if not self.c.df_mm.empty and "net_margin_pct" in self.c.df_mm.columns:
            best_margin = self.c.df_mm["net_margin_pct"].max()
            metrics.append(("MELHOR MARGEM", f"{best_margin:.1f}%", self.GREEN))
        if not self.c.df_mm.empty and "edp" in self.c.df_mm.columns:
            best_edp = self.c.df_mm["edp"].max()
            metrics.append(("MELHOR EDP/DIA", self._fmt_isk(best_edp), self.ACCENT))

        row = 4
        col = 2
        for label, value, color in metrics:
            # Card com fundo colorido
            ws.merge_cells(start_row=row, start_column=col,
                           end_row=row, end_column=col+1)
            ws.cell(row=row, column=col, value=label).font = Font(
                bold=True, size=8, color=self.WHITE, name="Calibri")
            ws.cell(row=row, column=col).fill = self._fill(color)
            ws.cell(row=row, column=col).alignment = Alignment(horizontal="center")

            ws.merge_cells(start_row=row+1, start_column=col,
                           end_row=row+1, end_column=col+1)
            ws.cell(row=row+1, column=col, value=value).font = Font(
                bold=True, size=14, color=self.TEXT_DARK, name="Calibri")
            ws.cell(row=row+1, column=col).fill = self._fill(self.LIGHT_BG)
            ws.cell(row=row+1, column=col).alignment = Alignment(horizontal="center")

            col += 3
            if col > 14:
                col = 2; row += 4

        # Top 10 oportunidades
        row_start = row + 3
        if not self.c.df_master.empty:
            top10 = self.c.df_master[self.c.df_master["acao"] == "COMPRAR"].head(10)

            ws.cell(row=row_start, column=2, value="TOP 10 OPORTUNIDADES DE COMPRA")
            ws.cell(row=row_start, column=2).font = Font(bold=True, size=12,
                                                          color=self.WHITE, name="Calibri")
            ws.cell(row=row_start, column=2).fill = self._fill(self.HEADER_BG)
            ws.merge_cells(start_row=row_start, start_column=2, end_row=row_start, end_column=9)
            row_start += 1

            headers = ["#", "Item", "Preço Sell", "Net Margem%", "EDP/Dia",
                       "Vol/Dia 90d", "DoS", "Score"]
            for ci, h in enumerate(headers, 2):
                cell = ws.cell(row=row_start, column=ci, value=h)
                cell.font = self._header_font(size=9)
                cell.fill = self._fill(self.ACCENT)
                cell.alignment = Alignment(horizontal="center")
            row_start += 1

            for ri, (_, r) in enumerate(top10.iterrows(), 1):
                bg = self.LIGHT_BG if ri % 2 == 0 else self.WHITE
                vals = [
                    ri,
                    str(r.get("item", ""))[:35],
                    self._fmt_isk(r.get("sell_min")),
                    _safe_pct(r.get('net_margin_pct'), '.1f'),
                    self._fmt_isk(r.get("edp")),
                    _safe_vol(r.get('vol_90d_dia')),
                    _safe_dos(r.get('dias_oferta')),
                    f"{_safe_int(r.get('conviction_score'))}",
                ]
                for ci, v in enumerate(vals, 2):
                    cell = ws.cell(row=row_start, column=ci, value=v)
                    cell.font = self._cell_font(bold=(ci == 3))
                    cell.fill = self._fill(bg)
                    cell.alignment = Alignment(horizontal="right" if ci > 3 else "left")
                row_start += 1

    def _sheet_analysis(self, wb, df: pd.DataFrame, name: str, key_cols: list):
        """Cria uma aba de análise genérica com formatação."""
        ws = wb.create_sheet(name[:31])
        ws.sheet_view.showGridLines = False

        # Seleciona e ordena colunas
        available = [c for c in key_cols if c in df.columns]
        df_show   = df[available].copy()

        # Título
        ws.merge_cells(start_row=1, start_column=1,
                       end_row=1, end_column=len(available))
        ws.cell(row=1, column=1, value=f"{name} — {self.c.region_name}")
        ws.cell(row=1, column=1).font = Font(bold=True, size=13,
                                              color=self.WHITE, name="Calibri")
        ws.cell(row=1, column=1).fill = self._fill(self.DARK_BG)
        ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")
        ws.row_dimensions[1].height = 28

        # Cabeçalhos
        label_map = {
            "item": "Item", "sell_min": "Sell Min (ISK)", "buy_max": "Buy Max (ISK)",
            "net_margin_pct": "Margem Líq %", "gross_margin_pct": "Margem Bruta %",
            "lucro_por_unidade": "Lucro/Un (ISK)", "edp": "EDP/Dia (ISK)",
            "vol_90d_dia": "Vol/Dia 90d", "vol_sell": "Vol Sell",
            "vol_buy": "Vol Buy", "margem_pct": "Margem %",
            "dias_oferta": "Dias Oferta", "score_escassez": "Score Escassez",
            "dsif_score": "DSIF Score", "ratio_demand_supply": "Ratio D/S",
            "preco_7d": "Preço 7d", "preco_30d": "Preço 30d",
            "var_preco_pct": "Var Preço %", "vol_7d_dia": "Vol 7d/dia",
            "vol_30d_dia": "Vol 30d/dia", "var_vol_pct": "Var Vol %",
            "score_pm": "Score PM", "qtd_recomendada": "Qtd Recom.",
            "custo_total": "Custo Total (ISK)", "lucro_total_est": "Lucro Est. (ISK)",
        }

        for ci, col in enumerate(available, 1):
            cell = ws.cell(row=2, column=ci,
                           value=label_map.get(col, col.replace("_", " ").title()))
            cell.font  = Font(bold=True, size=10, color=self.WHITE, name="Calibri")
            cell.fill  = self._fill(self.HEADER_BG)
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.border = self._border()
        ws.row_dimensions[2].height = 32
        ws.freeze_panes = "A3"

        # Dados
        isk_cols = {"sell_min", "buy_max", "lucro_por_unidade", "edp",
                    "preco_7d", "preco_30d", "custo_total", "lucro_total_est"}
        pct_cols = {"net_margin_pct", "gross_margin_pct", "margem_pct",
                    "var_preco_pct", "var_vol_pct"}

        for ri, (_, row_data) in enumerate(df_show.iterrows(), 3):
            bg = self.LIGHT_BG if ri % 2 == 0 else self.WHITE
            for ci, col in enumerate(available, 1):
                val = row_data[col]
                # Formata valor
                if col in isk_cols:
                    display = self._fmt_isk(val)
                elif col in pct_cols:
                    display = f"{val:.2f}%" if (val is not None and not (isinstance(val, float) and math.isnan(val))) else "--"
                elif col == "item":
                    display = str(val)[:40]
                else:
                    display = val if (val is not None and not (isinstance(val, float) and math.isnan(val))) else "--"

                cell = ws.cell(row=ri, column=ci, value=display)
                cell.font   = self._cell_font(bold=(col == "item"))
                cell.fill   = self._fill(bg)
                cell.border = self._border()
                cell.alignment = Alignment(
                    horizontal="left" if col == "item" else "right")

                # Formatação condicional por cor
                if col == "net_margin_pct" and isinstance(val, (int, float)) and not math.isnan(val or 0):
                    if val >= 10: cell.fill = self._fill("CCFFDD")
                    elif val < 0: cell.fill = self._fill("FFDDDD")
                if col == "var_preco_pct" and isinstance(val, (int, float)) and not math.isnan(val or 0):
                    if val > 5: cell.fill = self._fill("CCFFDD")
                    elif val < -5: cell.fill = self._fill("FFDDDD")

        self._auto_width(ws)

    def _sheet_recomendacoes(self, wb):
        """Aba com todas as recomendações ordenadas por prioridade."""
        ws = wb.create_sheet("RECOMENDACOES")
        ws.sheet_view.showGridLines = False

        ws.merge_cells("A1:L1")
        ws.cell(row=1, column=1,
                value=f"RECOMENDAÇÕES — {self.c.region_name}")
        ws.cell(row=1, column=1).font = Font(bold=True, size=14,
                                              color=self.WHITE, name="Calibri")
        ws.cell(row=1, column=1).fill = self._fill(self.DARK_BG)
        ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")
        ws.row_dimensions[1].height = 30

        headers = ["Ação", "Prioridade", "Score", "Item",
                   "Sell Min", "Net Marg%", "EDP/Dia",
                   "DoS", "Var Preço%", "Vol/Dia",
                   "Qtd Orc.", "Fontes"]
        for ci, h in enumerate(headers, 1):
            cell = ws.cell(row=2, column=ci, value=h)
            cell.font  = Font(bold=True, size=10, color=self.WHITE, name="Calibri")
            cell.fill  = self._fill(self.HEADER_BG)
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.border = self._border()
        ws.row_dimensions[2].height = 32
        ws.freeze_panes = "A3"

        for ri, (_, r) in enumerate(self.c.df_master.iterrows(), 3):
            acao    = r.get("acao", "MONITORAR")
            bg_acao = {"COMPRAR": "CCFFDD", "VENDER": "FFDDDD",
                       "MONITORAR": "FFFACC"}.get(acao, "FFFFFF")

            vals = [
                acao,
                r.get("prioridade", ""),
                f"{_safe_int(r.get('conviction_score'))}",
                str(r.get("item", ""))[:40],
                self._fmt_isk(r.get("sell_min")),
                _safe_pct(r.get('net_margin_pct'), '.1f'),
                self._fmt_isk(r.get("edp")),
                _safe_dos(r.get('dias_oferta')),
                _safe_pct(r.get("var_preco_pct")),
                _safe_vol(r.get('vol_90d_dia')),
                str(_safe_int(r.get('qtd_orcamento'))),
                str(_safe_int(r.get('n_fontes'))),
            ]

            for ci, v in enumerate(vals, 1):
                cell = ws.cell(row=ri, column=ci, value=v)
                bg = bg_acao if ci == 1 else ("EEEEEE" if ri % 2 == 0 else "FFFFFF")
                cell.font   = self._cell_font(bold=(ci in (1, 4)))
                cell.fill   = self._fill(bg)
                cell.border = self._border()
                cell.alignment = Alignment(
                    horizontal="center" if ci in (1, 2, 3, 12) else
                    ("left" if ci == 4 else "right"))

        self._auto_width(ws)

    def _sheet_ai(self, wb):
        """Aba com análise gerada pela IA."""
        ws = wb.create_sheet("ANALISE IA")
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = 2
        ws.column_dimensions["B"].width = 120

        ws.merge_cells("B1:B1")
        ws.cell(row=1, column=2, value="ANÁLISE IA — Claude AI")
        ws.cell(row=1, column=2).font = Font(bold=True, size=14,
                                              color=self.WHITE, name="Calibri")
        ws.cell(row=1, column=2).fill = self._fill(self.DARK_BG)
        ws.row_dimensions[1].height = 30

        for ri, line in enumerate(self.ai_text.split("\n"), 3):
            cell = ws.cell(row=ri, column=2, value=line)
            is_heading = line.startswith("##")
            cell.font  = Font(bold=is_heading, size=10 if not is_heading else 11,
                               color=self.TEXT_DARK if not is_heading else "0A0E1A",
                               name="Calibri")
            cell.alignment = Alignment(wrap_text=True)
            if is_heading:
                cell.fill = self._fill(self.MID_BG)
            ws.row_dimensions[ri].height = 15


# ─────────────────────────────────────────────────────────────────────────────
#  4. GERADOR DE RELATÓRIO PDF
# ─────────────────────────────────────────────────────────────────────────────

class PDFReportGenerator:
    """
    Gera relatório PDF profissional com ReportLab usando o layout Bloomberg-like
    do EVE Trade Finder. Inclui:
      - Capa com métricas principais
      - Tabela de recomendações
      - Top oportunidades por análise
      - Análise da IA
    """

    def __init__(self, collector: 'MarketDataCollector',
                 ai_text: str = "", progress_cb=None):
        self.c         = collector
        self.ai_text   = ai_text
        self._progress = progress_cb or (lambda *a: None)

        if HAS_REPORTLAB:
            self._styles  = getSampleStyleSheet()
            self._setup_styles()

    def _setup_styles(self):
        """Define estilos personalizados para o relatório."""
        self._st = {}

        self._st["title"] = ParagraphStyle(
            "eve_title", parent=self._styles["Title"],
            fontSize=22, fontName="Helvetica-Bold",
            textColor=colors.HexColor("#00aaff"),
            spaceAfter=4)

        self._st["subtitle"] = ParagraphStyle(
            "eve_subtitle", parent=self._styles["Normal"],
            fontSize=10, fontName="Helvetica",
            textColor=colors.HexColor("#5a7a9a"),
            spaceAfter=2)

        self._st["h1"] = ParagraphStyle(
            "eve_h1", parent=self._styles["Heading1"],
            fontSize=14, fontName="Helvetica-Bold",
            textColor=colors.white,
            backColor=colors.HexColor("#0f1624"),
            spaceBefore=12, spaceAfter=4,
            leftIndent=6, rightIndent=6)

        self._st["h2"] = ParagraphStyle(
            "eve_h2", parent=self._styles["Heading2"],
            fontSize=11, fontName="Helvetica-Bold",
            textColor=colors.HexColor("#00aaff"),
            spaceBefore=8, spaceAfter=3)

        self._st["body"] = ParagraphStyle(
            "eve_body", parent=self._styles["Normal"],
            fontSize=9, fontName="Helvetica",
            textColor=colors.HexColor("#1a2a3a"),
            leading=13, spaceAfter=4)

        self._st["code"] = ParagraphStyle(
            "eve_code", parent=self._styles["Code"],
            fontSize=8, fontName="Courier",
            textColor=colors.HexColor("#1a2a3a"),
            backColor=colors.HexColor("#e8f0f8"),
            leading=11)

        self._st["ai"] = ParagraphStyle(
            "eve_ai", parent=self._styles["Normal"],
            fontSize=9, fontName="Helvetica",
            textColor=colors.HexColor("#1a2a3a"),
            leading=13, spaceAfter=3)

        self._st["ai_h"] = ParagraphStyle(
            "eve_ai_h", parent=self._styles["Heading3"],
            fontSize=10, fontName="Helvetica-Bold",
            textColor=colors.white,
            backColor=colors.HexColor("#1e2d45"),
            spaceBefore=6, spaceAfter=2,
            leftIndent=4)

    def _fmt_isk(self, v) -> str:
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return "--"
        if abs(v) >= 1e9: return f"{v/1e9:.2f}B"
        if abs(v) >= 1e6: return f"{v/1e6:.1f}M"
        if abs(v) >= 1e3: return f"{v/1e3:.0f}K"
        return f"{v:,.0f}"

    def _tbl_style(self, header_rows=1, stripe=True):
        """Estilo de tabela padrão EVE."""
        cmds = [
            # Cabeçalho
            ("BACKGROUND",   (0,0), (-1,header_rows-1), colors.HexColor("#0f1624")),
            ("TEXTCOLOR",    (0,0), (-1,header_rows-1), colors.white),
            ("FONTNAME",     (0,0), (-1,header_rows-1), "Helvetica-Bold"),
            ("FONTSIZE",     (0,0), (-1,header_rows-1), 8),
            ("ALIGN",        (0,0), (-1,header_rows-1), "CENTER"),
            # Corpo
            ("FONTNAME",     (0,header_rows), (-1,-1), "Helvetica"),
            ("FONTSIZE",     (0,header_rows), (-1,-1), 8),
            ("TEXTCOLOR",    (0,header_rows), (-1,-1), colors.HexColor("#1a2a3a")),
            # Bordas
            ("GRID",         (0,0), (-1,-1), 0.3, colors.HexColor("#c0c8d0")),
            ("TOPPADDING",   (0,0), (-1,-1), 3),
            ("BOTTOMPADDING",(0,0), (-1,-1), 3),
            ("LEFTPADDING",  (0,0), (-1,-1), 5),
            ("RIGHTPADDING", (0,0), (-1,-1), 5),
            ("ROWBACKGROUNDS",(0,header_rows), (-1,-1),
             [colors.white, colors.HexColor("#e8f0f8")] if stripe else [colors.white]),
        ]
        return TableStyle(cmds)

    def _acao_color(self, acao: str):
        return {"COMPRAR": colors.HexColor("#ccffdd"),
                "VENDER":  colors.HexColor("#ffdddd"),
                "MONITORAR": colors.HexColor("#fffacc")}.get(acao, colors.white)

    def generate(self, output_path: str):
        """Gera o arquivo PDF completo."""
        if not HAS_REPORTLAB:
            raise ImportError("reportlab não instalado.")

        self._progress("PDF: preparando documento...")
        doc = SimpleDocTemplate(
            output_path,
            pagesize=landscape(A4),
            leftMargin=1.5*cm, rightMargin=1.5*cm,
            topMargin=1.5*cm, bottomMargin=2*cm,
            title=f"EVE Market Analyst — {self.c.region_name}",
            author="EVE Market Analyst v" + VERSION,
        )

        story = []

        # ── CAPA ─────────────────────────────────────────────────────────────
        self._progress("PDF: capa...")
        story += self._build_cover()

        # ── RECOMENDAÇÕES ─────────────────────────────────────────────────────
        self._progress("PDF: tabela de recomendações...")
        story += self._build_recomendacoes()

        # ── POR ABA ───────────────────────────────────────────────────────────
        if not self.c.df_mm.empty:
            self._progress("PDF: MAX MARGEM...")
            story += self._build_analysis_section(
                "MAX MARGEM", self.c.df_mm,
                cols=["item","sell_min","buy_max","net_margin_pct",
                      "lucro_por_unidade","edp","vol_90d_dia"],
                headers=["Item","Sell Min","Buy Max","Margem Líq%",
                         "Lucro/Un","EDP/Dia","Vol/Dia"],
                n=15)

        if not self.c.df_sc.empty:
            self._progress("PDF: ESCASSEZ...")
            story += self._build_analysis_section(
                "ESCASSEZ", self.c.df_sc,
                cols=["item","sell_min","margem_pct","vol_90d_dia",
                      "dias_oferta","score_escassez"],
                headers=["Item","Sell Min","Margem%","Vol/Dia","DoS (dias)","Score"],
                n=15)

        if not self.c.df_ds.empty:
            self._progress("PDF: DEMANDA/OFERTA...")
            story += self._build_analysis_section(
                "DEMANDA/OFERTA", self.c.df_ds,
                cols=["item","sell_min","margem_pct","vol_buy",
                      "vol_sell","ratio_demand_supply","dsif_score"],
                headers=["Item","Sell","Margem%","Vol Compra","Vol Venda","Ratio D/S","DSIF"],
                n=15)

        if not self.c.df_pm.empty:
            self._progress("PDF: MOV. PREÇOS...")
            story += self._build_analysis_section(
                "MOV. PRECOS", self.c.df_pm,
                cols=["item","preco_7d","preco_30d","var_preco_pct",
                      "vol_7d_dia","var_vol_pct","score_pm"],
                headers=["Item","Preço 7d","Preço 30d","Var Preço%",
                         "Vol 7d/dia","Var Vol%","Score"],
                n=15)

        if not self.c.df_bud.empty:
            self._progress("PDF: ORÇAMENTO...")
            story += self._build_analysis_section(
                "ORCAMENTO", self.c.df_bud,
                cols=["item","buy_max","sell_min","qtd_recomendada",
                      "custo_total","lucro_total_est","net_margin_pct"],
                headers=["Item","Buy Max","Sell Min","Qtd","Custo Total",
                         "Lucro Est.","Margem%"],
                n=15)

        # ── ANÁLISE IA ────────────────────────────────────────────────────────
        if self.ai_text:
            self._progress("PDF: análise IA...")
            story += self._build_ai_section()

        # ── RODAPÉ / DISCLAIMER ───────────────────────────────────────────────
        story.append(Spacer(1, 0.5*cm))
        story.append(HRFlowable(width="100%", thickness=0.5,
                                 color=colors.HexColor("#1e2d45")))
        story.append(Paragraph(
            "Relatório gerado por EVE Market Analyst v" + VERSION + " | "
            "Dados de mercado obtidos via ESI (EVE Swagger Interface) | "
            "Este relatório é de uso exclusivo para fins de entretenimento no jogo EVE Online. "
            "Não constitui conselho financeiro real.",
            ParagraphStyle("disclaimer", fontSize=6, textColor=PDF_GRAY,
                           alignment=TA_CENTER)))

        self._progress("PDF: renderizando...")
        doc.build(story)
        self._progress(f"PDF salvo: {output_path}")

    def _build_cover(self) -> list:
        """Elementos da capa do relatório."""
        s = []
        s.append(Spacer(1, 0.5*cm))
        s.append(Paragraph("⬡  EVE MARKET ANALYST", self._st["title"]))
        s.append(Paragraph(
            f"Relatório de Análise de Mercado — {self.c.region_name}",
            self._st["subtitle"]))
        s.append(Paragraph(
            f"Gerado em {self.c.timestamp.strftime('%d/%m/%Y %H:%M UTC')} | "
            f"Capital: {self._fmt_isk(self.c.budget)} ISK | "
            f"Broker Rel: {self.c.br} | Accounting: {self.c.ac}",
            self._st["subtitle"]))
        s.append(Spacer(1, 0.4*cm))
        s.append(HRFlowable(width="100%", thickness=2,
                             color=colors.HexColor("#00aaff")))
        s.append(Spacer(1, 0.3*cm))

        # Cards de métricas em tabela horizontal
        if not self.c.df_master.empty:
            comprar   = (self.c.df_master["acao"] == "COMPRAR").sum()
            vender    = (self.c.df_master["acao"] == "VENDER").sum()
            monitorar = (self.c.df_master["acao"] == "MONITORAR").sum()
            total     = len(self.c.df_master)
            best_score= self.c.df_master["conviction_score"].max() if total else 0

            nm_max = self.c.df_mm["net_margin_pct"].max() if not self.c.df_mm.empty and "net_margin_pct" in self.c.df_mm.columns else 0
            edp_max= self.c.df_mm["edp"].max() if not self.c.df_mm.empty and "edp" in self.c.df_mm.columns else 0

            cards = [
                ["COMPRAR", "VENDER", "MONITORAR", "TOTAL", "MELHOR MARGEM",
                 "MELHOR EDP/DIA", "CONVICTION MAX"],
                [str(comprar), str(vender), str(monitorar), str(total),
                 f"{nm_max:.1f}%",
                 self._fmt_isk(edp_max),
                 f"{best_score:.0f}/100"],
            ]
            card_colors = [
                [colors.HexColor("#00cc66"), colors.HexColor("#ff4455"),
                 colors.HexColor("#cc9900"), colors.HexColor("#00aaff"),
                 colors.HexColor("#00cc66"), colors.HexColor("#00aaff"),
                 colors.HexColor("#7ecfff")] * 7,
            ]
            tbl = Table(cards, colWidths=[3.8*cm]*7)
            tbl.setStyle(TableStyle([
                ("BACKGROUND",   (0,0), (-1,0), colors.HexColor("#0f1624")),
                ("TEXTCOLOR",    (0,0), (-1,0), colors.white),
                ("FONTNAME",     (0,0), (-1,0), "Helvetica-Bold"),
                ("FONTSIZE",     (0,0), (-1,0), 8),
                ("ALIGN",        (0,0), (-1,-1), "CENTER"),
                ("FONTNAME",     (0,1), (-1,1), "Helvetica-Bold"),
                ("FONTSIZE",     (0,1), (-1,1), 14),
                ("TEXTCOLOR",    (0,1), (0,1),  colors.HexColor("#00aa44")),
                ("TEXTCOLOR",    (1,1), (1,1),  colors.HexColor("#cc2233")),
                ("TEXTCOLOR",    (2,1), (2,1),  colors.HexColor("#aa8800")),
                ("TEXTCOLOR",    (3,1), (3,1),  colors.HexColor("#00aaff")),
                ("BACKGROUND",   (0,1), (-1,1), colors.HexColor("#e8f0f8")),
                ("GRID",         (0,0), (-1,-1), 0.5, colors.HexColor("#1e2d45")),
                ("TOPPADDING",   (0,0), (-1,-1), 6),
                ("BOTTOMPADDING",(0,0), (-1,-1), 6),
                ("BOX",          (0,0), (-1,-1), 1.5, colors.HexColor("#00aaff")),
            ]))
            s.append(tbl)
            s.append(Spacer(1, 0.3*cm))

        # Erros/avisos
        if self.c.errors:
            s.append(Paragraph("⚠ AVISOS:", self._st["h2"]))
            for err in self.c.errors:
                s.append(Paragraph(f"• {err}", self._st["body"]))
            s.append(Spacer(1, 0.2*cm))

        return s

    def _build_recomendacoes(self) -> list:
        """Tabela completa de recomendações."""
        s = []
        s.append(PageBreak())
        s.append(Paragraph("RECOMENDAÇÕES DE TRADING", self._st["h1"]))
        s.append(Spacer(1, 0.2*cm))

        if self.c.df_master.empty:
            s.append(Paragraph("Sem dados disponíveis.", self._st["body"]))
            return s

        df = self.c.df_master.head(40)

        isk_cols = {"sell_min", "buy_max", "edp", "lucro_unitario",
                    "custo_total_orc", "lucro_total_orc"}

        headers = ["Ação", "Prio.", "Score", "Item",
                   "Sell Min", "Net%", "EDP/Dia",
                   "DoS", "Var%", "Vol/Dia", "Fontes"]

        tbl_data = [headers]
        tbl_cmds = list(self._tbl_style().getCommands())

        for ri, (_, r) in enumerate(df.iterrows(), 1):
            acao  = r.get("acao", "MONITORAR")
            vals  = [
                acao,
                r.get("prioridade","")[:4],
                f"{_safe_int(r.get('conviction_score'))}",
                str(r.get("item",""))[:32],
                self._fmt_isk(r.get("sell_min")),
                _safe_pct(r.get('net_margin_pct'), '.1f'),
                self._fmt_isk(r.get("edp")),
                _safe_dos(r.get('dias_oferta')),
                _safe_pct(r.get("var_preco_pct")),
                _safe_vol(r.get('vol_90d_dia')),
                str(_safe_int(r.get("n_fontes"))),
            ]
            tbl_data.append(vals)
            # Cor da linha por ação
            row_bg = self._acao_color(acao)
            if acao != "MONITORAR":
                tbl_cmds.append(("BACKGROUND", (0,ri), (0,ri), row_bg))

        col_ws = [1.7*cm, 1.2*cm, 1.2*cm, 7.5*cm,
                   2.5*cm, 1.5*cm, 2.5*cm, 1.5*cm, 1.5*cm, 1.8*cm, 1.2*cm]
        tbl = Table(tbl_data, colWidths=col_ws, repeatRows=1)
        tbl.setStyle(TableStyle(tbl_cmds))
        s.append(tbl)
        s.append(Spacer(1, 0.3*cm))

        # Legenda de cores
        s.append(Paragraph(
            "<b>Legenda:</b>  "
            "<font color='#008844'>■ COMPRAR</font>  "
            "<font color='#cc2233'>■ VENDER</font>  "
            "<font color='#886600'>■ MONITORAR</font>  |  "
            "Score = conviction score 0–100 | DoS = Dias de Oferta | "
            "EDP = lucro estimado por dia | Fontes = número de abas que confirmam o sinal",
            ParagraphStyle("legend", fontSize=7, textColor=PDF_GRAY,
                           leading=10)))
        return s

    def _build_analysis_section(self, title: str, df: pd.DataFrame,
                                 cols: list, headers: list, n: int = 15) -> list:
        """Seção de análise para cada aba."""
        s = []
        s.append(PageBreak())
        s.append(Paragraph(title, self._st["h1"]))
        s.append(Paragraph(
            f"Top {min(n, len(df))} itens por {title.lower()} | {self.c.region_name}",
            self._st["subtitle"]))
        s.append(Spacer(1, 0.2*cm))

        available = [c for c in cols if c in df.columns]
        hdrs      = [headers[cols.index(c)] for c in available]

        isk_cols = {"sell_min","buy_max","lucro_por_unidade","edp",
                    "preco_7d","preco_30d","custo_total","lucro_total_est",
                    "buy_max","cost_unit"}
        pct_cols = {"net_margin_pct","gross_margin_pct","margem_pct",
                    "var_preco_pct","var_vol_pct"}

        tbl_data = [hdrs]
        for _, row_data in df[available].head(n).iterrows():
            vals = []
            for c in available:
                v = row_data[c]
                if v is None or (isinstance(v, float) and math.isnan(v)):
                    vals.append("--")
                elif c in isk_cols:
                    vals.append(self._fmt_isk(v))
                elif c in pct_cols:
                    vals.append(f"{v:.2f}%")
                elif c == "item":
                    vals.append(str(v)[:35])
                else:
                    vals.append(str(v) if not isinstance(v, float) else f"{v:.2f}")
            tbl_data.append(vals)

        col_w = [7.5*cm] + [2.8*cm] * (len(available)-1)
        tbl = Table(tbl_data, colWidths=col_w, repeatRows=1)
        tbl.setStyle(self._tbl_style())
        s.append(tbl)
        return s

    def _build_ai_section(self) -> list:
        """Seção com análise gerada pela IA."""
        s = []
        s.append(PageBreak())
        s.append(Paragraph("ANÁLISE ESTRATÉGICA — CLAUDE AI", self._st["h1"]))
        s.append(Spacer(1, 0.2*cm))

        for line in self.ai_text.split("\n"):
            line = line.strip()
            if not line:
                s.append(Spacer(1, 3))
            elif line.startswith("## "):
                s.append(Paragraph(line[3:], self._st["ai_h"]))
            elif line.startswith("# "):
                s.append(Paragraph(line[2:], self._st["h2"]))
            elif line.startswith("**") and line.endswith("**"):
                s.append(Paragraph(f"<b>{line[2:-2]}</b>", self._st["ai"]))
            elif line.startswith("• ") or line.startswith("- ") or line.startswith("* "):
                s.append(Paragraph(f"&nbsp;&nbsp;&nbsp;• {line[2:]}",
                                    self._st["ai"]))
            else:
                s.append(Paragraph(line, self._st["ai"]))
        return s


# ─────────────────────────────────────────────────────────────────────────────
#  5. INTERFACE GRÁFICA
# ─────────────────────────────────────────────────────────────────────────────

class EVEAnalystApp(tk.Tk):
    """
    Interface gráfica do EVE Market Analyst.
    Permite configurar hub, capital, API key e iniciar a análise,
    acompanhar o progresso e gerar relatórios PDF e Excel.
    """

    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE}  v{VERSION}")
        self.configure(bg=COLORS["bg"])
        self.geometry("900x720")
        self.minsize(800, 620)

        self._collector: MarketDataCollector | None = None
        self._ai_text   = ""
        self._running   = False

        self._build_ui()

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        """Constrói toda a interface gráfica."""
        # ── Título ────────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=COLORS["panel"], height=48)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="⬡  EVE MARKET ANALYST",
                 bg=COLORS["panel"], fg=COLORS["accent"],
                 font=("Courier New", 16, "bold")).pack(side="left", padx=16)
        tk.Label(hdr, text=f"v{VERSION}  |  Análise Multi-Aba + IA",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier New", 9)).pack(side="right", padx=16)
        tk.Frame(self, bg=COLORS["accent"], height=2).pack(fill="x")

        # ── Painel de configuração ─────────────────────────────────────────────
        cfg = tk.LabelFrame(self, text=" CONFIGURAÇÃO ",
                             bg=COLORS["panel"], fg=COLORS["header"],
                             font=("Courier New", 9, "bold"),
                             highlightbackground=COLORS["border"],
                             highlightthickness=1)
        cfg.pack(fill="x", padx=10, pady=(8, 4))

        # Hub
        tk.Label(cfg, text="Hub de mercado:",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier New", 9)).grid(row=0, column=0, padx=10, pady=4, sticky="w")
        self._hub_var = tk.StringVar(value="The Forge (Jita)")
        self._hub_cb  = ttk.Combobox(cfg, textvariable=self._hub_var,
                                      values=list(et.REGIONS.keys()),
                                      state="readonly", width=30,
                                      font=("Courier New", 9))
        self._hub_cb.grid(row=0, column=1, padx=4, pady=4)

        # Capital
        tk.Label(cfg, text="Capital disponível (ISK):",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier New", 9)).grid(row=0, column=2, padx=10, sticky="w")
        self._cap_var = tk.StringVar(value="1000000000")
        tk.Entry(cfg, textvariable=self._cap_var,
                 bg=COLORS["border"], fg=COLORS["text"],
                 insertbackground=COLORS["text"],
                 font=("Courier New", 9), width=18,
                 relief="flat").grid(row=0, column=3, padx=4)

        # API Key
        tk.Label(cfg, text="API Key Claude:",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier New", 9)).grid(row=1, column=0, padx=10, pady=4, sticky="w")

        # Tenta carregar do keyring ou config
        default_key = et.keyring_load_key() or et.load_config().get("ai_key", "")
        self._key_var = tk.StringVar(value=default_key)
        self._key_entry = tk.Entry(cfg, textvariable=self._key_var,
                                    show="•", width=50,
                                    bg=COLORS["border"], fg=COLORS["text"],
                                    insertbackground=COLORS["text"],
                                    font=("Courier New", 9), relief="flat")
        self._key_entry.grid(row=1, column=1, columnspan=2, padx=4, pady=4)
        tk.Checkbutton(cfg, text="Mostrar key",
                       bg=COLORS["panel"], fg=COLORS["subtext"],
                       selectcolor=COLORS["border"],
                       activebackground=COLORS["panel"],
                       font=("Courier New", 8),
                       command=self._toggle_key).grid(row=1, column=3, padx=4)

        # Opções
        self._gen_pdf  = tk.BooleanVar(value=True)
        self._gen_xlsx = tk.BooleanVar(value=True)
        self._use_ai   = tk.BooleanVar(value=True)
        tk.Checkbutton(cfg, text="Gerar PDF",
                       variable=self._gen_pdf,
                       bg=COLORS["panel"], fg=COLORS["text"],
                       selectcolor=COLORS["border"],
                       activebackground=COLORS["panel"],
                       font=("Courier New", 9)).grid(row=2, column=0, padx=10)
        tk.Checkbutton(cfg, text="Gerar Excel",
                       variable=self._gen_xlsx,
                       bg=COLORS["panel"], fg=COLORS["text"],
                       selectcolor=COLORS["border"],
                       activebackground=COLORS["panel"],
                       font=("Courier New", 9)).grid(row=2, column=1, padx=4)
        tk.Checkbutton(cfg, text="Análise IA (Claude)",
                       variable=self._use_ai,
                       bg=COLORS["panel"], fg=COLORS["yellow"],
                       selectcolor=COLORS["border"],
                       activebackground=COLORS["panel"],
                       font=("Courier New", 9, "bold")).grid(row=2, column=2, padx=4)

        # Saída
        tk.Label(cfg, text="Pasta de saída:",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier New", 9)).grid(row=3, column=0, padx=10, pady=4, sticky="w")
        self._out_var = tk.StringVar(value=str(Path.home() / "Downloads"))
        tk.Entry(cfg, textvariable=self._out_var,
                 bg=COLORS["border"], fg=COLORS["text"],
                 insertbackground=COLORS["text"],
                 font=("Courier New", 9), width=46,
                 relief="flat").grid(row=3, column=1, columnspan=2, padx=4)
        tk.Button(cfg, text=" BROWSE ",
                  command=self._browse_output,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier New", 8), relief="flat",
                  cursor="hand2").grid(row=3, column=3, padx=4)

        # ── Botão principal ───────────────────────────────────────────────────
        self._run_btn = tk.Button(
            self, text="  ▶  ANALISAR E GERAR RELATÓRIO  ",
            command=self._start_analysis,
            bg=COLORS["orange"], fg="#000",
            font=("Courier New", 12, "bold"),
            relief="flat", cursor="hand2")
        self._run_btn.pack(pady=8)

        # ── Barra de progresso ────────────────────────────────────────────────
        prog_fr = tk.Frame(self, bg=COLORS["bg"])
        prog_fr.pack(fill="x", padx=10)
        self._prog_bar = ttk.Progressbar(prog_fr, mode="indeterminate", length=860)
        self._prog_bar.pack(side="left", fill="x", expand=True)
        self._prog_pct = tk.Label(prog_fr, text="", width=6,
                                   bg=COLORS["bg"], fg=COLORS["subtext"],
                                   font=("Courier New", 8))
        self._prog_pct.pack(side="right")

        # ── Log de progresso ──────────────────────────────────────────────────
        log_fr = tk.LabelFrame(self, text=" PROGRESSO ",
                                bg=COLORS["panel"], fg=COLORS["header"],
                                font=("Courier New", 9, "bold"),
                                highlightbackground=COLORS["border"],
                                highlightthickness=1)
        log_fr.pack(fill="both", expand=True, padx=10, pady=(4, 4))

        self._log_text = tk.Text(log_fr,
                                  bg=COLORS["bg"], fg=COLORS["text"],
                                  insertbackground=COLORS["text"],
                                  font=("Courier New", 9),
                                  relief="flat", state="disabled",
                                  wrap="word")
        vsb = tk.Scrollbar(log_fr, command=self._log_text.yview)
        vsb.pack(side="right", fill="y")
        self._log_text.config(yscrollcommand=vsb.set)
        self._log_text.pack(fill="both", expand=True, padx=4, pady=4)

        # Tags de cor para o log
        self._log_text.tag_config("ok",   foreground=COLORS["green"])
        self._log_text.tag_config("warn", foreground=COLORS["yellow"])
        self._log_text.tag_config("err",  foreground=COLORS["red"])
        self._log_text.tag_config("info", foreground=COLORS["accent"])
        self._log_text.tag_config("head", foreground=COLORS["header"])

        # ── Status bar ────────────────────────────────────────────────────────
        self._status = tk.Label(self, text="Pronto.",
                                 bg=COLORS["panel"], fg=COLORS["subtext"],
                                 font=("Courier New", 8), anchor="w")
        self._status.pack(fill="x", padx=10, pady=(0, 6))

        self._log("EVE Market Analyst iniciado.", "head")
        self._log(f"eve_trade.py importado de: {_HERE}", "info")
        self._log("Configure o hub e clique em ANALISAR.", "info")

    def _toggle_key(self):
        cur = self._key_entry.cget("show")
        self._key_entry.config(show="" if cur == "•" else "•")

    def _browse_output(self):
        path = filedialog.askdirectory(title="Pasta de saída")
        if path:
            self._out_var.set(path)

    # ── Logging ───────────────────────────────────────────────────────────────
    def _log(self, msg: str, tag: str = ""):
        """Adiciona linha ao log com timestamp e tag de cor."""
        ts  = datetime.now().strftime("%H:%M:%S")
        self._log_text.config(state="normal")
        self._log_text.insert("end", f"[{ts}] {msg}\n", tag or "")
        self._log_text.see("end")
        self._log_text.config(state="disabled")
        self._status.config(text=msg[:100])
        self.update_idletasks()

    def _progress_cb(self, msg):
        """Callback de progresso para os coletores de dados."""
        if isinstance(msg, tuple):
            done, total, txt = msg
            pct = int(done / total * 100) if total else 0
            self._prog_pct.config(text=f"{pct}%")
            self._log(f"  {txt}", "")
        else:
            self._log(str(msg), "")

    # ── Análise principal ─────────────────────────────────────────────────────
    def _start_analysis(self):
        """Inicia a análise em thread separada."""
        if self._running:
            return
        self._running = True
        self._run_btn.config(state="disabled")
        self._prog_bar.start(10)

        threading.Thread(target=self._run_analysis, daemon=True).start()

    def _run_analysis(self):
        """Thread de análise: coleta dados, chama IA, gera relatórios."""
        try:
            hub    = self._hub_var.get()
            api_key= self._key_var.get().strip()
            use_ai = self._use_ai.get()
            gen_pdf= self._gen_pdf.get()
            gen_xl = self._gen_xlsx.get()
            out_dir= Path(self._out_var.get())

            try:
                budget = float(self._cap_var.get().replace(",", "").replace(".", ""))
            except ValueError:
                budget = 1_000_000_000

            ts_str = datetime.now().strftime("%Y%m%d_%H%M")
            hub_slug = hub.replace("(", "").replace(")", "").replace(" ", "_")[:20]
            base_name = f"EVE_Analyst_{hub_slug}_{ts_str}"

            # ── 1. Coleta de dados ─────────────────────────────────────────────
            self.after(0, lambda: self._log(
                f"{'='*60}", "head"))
            self.after(0, lambda: self._log(
                f"Iniciando análise: {hub} | Capital: {budget/1e9:.2f}B ISK", "info"))
            self.after(0, lambda: self._log("─" * 40, ""))

            collector = MarketDataCollector(hub, budget, self._progress_cb)
            collector.collect()
            self._collector = collector

            if collector.errors:
                for err in collector.errors:
                    self.after(0, lambda e=err: self._log(f"⚠ {e}", "warn"))

            self.after(0, lambda: self._log(
                f"✓ Coleta concluída: {len(collector.df_master)} itens analisados", "ok"))

            # ── 2. Análise IA ──────────────────────────────────────────────────
            self._ai_text = ""
            if use_ai and api_key:
                self.after(0, lambda: self._log("Consultando Claude AI...", "info"))
                analyst = AIAnalyst(api_key)
                summary = collector.summary_text()
                self._ai_text = analyst.analyze(summary, hub, budget, self._progress_cb)
                if self._ai_text.startswith("[ERRO") or self._ai_text.startswith("[AVISO"):
                    self.after(0, lambda t=self._ai_text:
                               self._log(t, "warn"))
                else:
                    self.after(0, lambda: self._log(
                        f"✓ Análise IA concluída ({len(self._ai_text)} chars)", "ok"))
            elif use_ai and not api_key:
                self.after(0, lambda: self._log(
                    "⚠ API Key não configurada — análise IA pulada", "warn"))

            # ── 3. Geração de relatórios ───────────────────────────────────────
            reports_generated = []

            if gen_xl and HAS_OPENPYXL:
                xl_path = out_dir / f"{base_name}.xlsx"
                self.after(0, lambda: self._log(f"Gerando Excel: {xl_path.name}...", "info"))
                gen = ExcelReportGenerator(collector, self._ai_text, self._progress_cb)
                gen.generate(str(xl_path))
                reports_generated.append(str(xl_path))
                self.after(0, lambda: self._log(f"✓ Excel gerado", "ok"))
            elif gen_xl:
                self.after(0, lambda: self._log(
                    "⚠ openpyxl não instalado — Excel pulado", "warn"))

            if gen_pdf and HAS_REPORTLAB:
                pdf_path = out_dir / f"{base_name}.pdf"
                self.after(0, lambda: self._log(f"Gerando PDF: {pdf_path.name}...", "info"))
                gen_pdf_obj = PDFReportGenerator(collector, self._ai_text, self._progress_cb)
                gen_pdf_obj.generate(str(pdf_path))
                reports_generated.append(str(pdf_path))
                self.after(0, lambda: self._log(f"✓ PDF gerado", "ok"))
            elif gen_pdf:
                self.after(0, lambda: self._log(
                    "⚠ reportlab não instalado — PDF pulado", "warn"))

            # ── 4. Exibe resultado no log ──────────────────────────────────────
            self.after(0, lambda: self._log("─" * 60, ""))
            self.after(0, lambda: self._log("ANÁLISE CONCLUÍDA:", "head"))

            # Top 5 COMPRAR
            if not collector.df_master.empty:
                top_buy = collector.df_master[
                    collector.df_master["acao"] == "COMPRAR"].head(5)
                self.after(0, lambda: self._log(
                    f"TOP 5 OPORTUNIDADES DE COMPRA em {hub}:", "ok"))
                for _, r in top_buy.iterrows():
                    nm  = r.get("net_margin_pct", 0) or 0
                    edp = r.get("edp")
                    edp_s = f" | EDP:{edp/1e6:.1f}M/dia" if edp else ""
                    self.after(0, lambda item=r['item'], nm=nm, score=r['conviction_score'], edp_s=edp_s:
                               self._log(f"  ▸ {item[:35]:<35} Net:{nm:.1f}% | Score:{score:.0f}{edp_s}", "ok"))

                # Top 3 VENDER
                top_sell = collector.df_master[
                    collector.df_master["acao"] == "VENDER"].head(3)
                if len(top_sell):
                    self.after(0, lambda: self._log("VENDER / EVITAR:", "warn"))
                    for _, r in top_sell.iterrows():
                        var = r.get("var_preco_pct", 0) or 0
                        self.after(0, lambda item=r['item'], var=var:
                                   self._log(f"  ▸ {item[:35]:<35} Var Preço:{var:+.1f}%", "warn"))

            for rp in reports_generated:
                self.after(0, lambda p=rp: self._log(f"📄 Relatório: {p}", "info"))

            if not reports_generated:
                self.after(0, lambda: self._log(
                    "Nenhum relatório gerado (opções desmarcadas ou bibliotecas ausentes)", "warn"))

        except Exception as e:
            import traceback
            err_msg = traceback.format_exc()
            self.after(0, lambda e=str(e), tb=err_msg:
                       (self._log(f"[ERRO] {e}", "err"),
                        self._log(tb[:300], "err")))
        finally:
            self._running = False
            self.after(0, lambda: (
                self._prog_bar.stop(),
                self._prog_bar.config(value=0),
                self._run_btn.config(state="normal")))


# ─────────────────────────────────────────────────────────────────────────────
#  Entry point
# ─────────────────────────────────────────────────────────────────────────────

def main():
    """Ponto de entrada da aplicação."""
    app = EVEAnalystApp()
    app.mainloop()


if __name__ == "__main__":
    main()
