"""
EVE Online Trade Finder  v14
ABAS: 1=Rotas 2=Arbitragem 3=Otimizador 4=Max Margem 5=Escassez
      6=Demanda/Oferta 7=Mov.Precos 8=Orcamento 9=Mercado Live
     10=Watchlist 11=Paper Trading 12=AI Analista
     13=Comparador 14=Configuracoes 15=Dashboard 16=Multi-Hop
     17=Correlacao 18=Spread Hist. 19=Ciclos 20=Previsao
     21=Market Making 22=Industria 23=Backtesting 24=Heatmap
     25=Station Trader 26=Catalogo 27=Dados&Cache
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import concurrent.futures
import re
import requests
import requests.adapters
import csv
import os
import json
import time
import math
import sqlite3
import logging
from datetime import datetime, timedelta, timezone
from collections import defaultdict

logging.basicConfig(level=logging.WARNING,
                    format="%(asctime)s [%(levelname)s] %(message)s",
                    datefmt="%H:%M:%S")
log = logging.getLogger("EVETrade")

# ── Motor de cálculo (sem Tkinter) — evetrade/core ────────────────────────────
# Os nomes antigos (com "_") continuam válidos para a UI, eve_analyst,
# eve_dashboard e scripts externos.
from evetrade.core import (  # noqa: E402
    MAX_INT32,
    MM_MIN_BUY_RATIO,
    allocate_budget,
    arbitrage_opportunities,
    build_order_index,
    calc_lucro_dia,
    calculate_industry_cost,
    find_routes,
    margin_candidates,
    match_orders_granular,
    select_by_volume,
)
from evetrade.core import RISK_DAYS as _RISK_DAYS  # noqa: E402
from evetrade.core import calc_fee_rates as _calc_fee_rates  # noqa: E402
from evetrade.core import filter_blacklist as _filter_blacklist  # noqa: E402
from evetrade.core import history_stats as _history_stats  # noqa: E402
from evetrade.core import order_min_volume as _order_min_volume  # noqa: E402

try:
    import matplotlib
    matplotlib.use("TkAgg")
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    from ortools.linear_solver import pywraplp
    HAS_ORTOOLS = True
except ImportError:
    HAS_ORTOOLS = False

# ── API Key: só keyring SEGURO do SO, nunca texto puro ─────────────────────────
# Backends recusados: keyrings.alt (PlaintextKeyring/EncryptedKeyring em
# arquivo), fail/null e o chainer (que mistura backends seguros com o
# PlaintextKeyring quando keyrings.alt está instalado e cai nele sem aviso).
_INSECURE_KEYRING_MODULES = ("keyrings.alt", "keyring.backends.fail",
                             "keyring.backends.null", "keyring.backends.chainer")


def _is_secure_backend(backend) -> bool:
    return not type(backend).__module__.startswith(_INSECURE_KEYRING_MODULES)


def _select_secure_keyring():
    """Fixa o keyring no backend seguro de maior prioridade. None se não houver."""
    import keyring.backend
    cands = [b for b in keyring.backend.get_all_keyring()
             if _is_secure_backend(b) and b.priority > 0]
    if not cands:
        return None
    best = max(cands, key=lambda b: b.priority)
    keyring.set_keyring(best)
    return best


try:
    import keyring
    try:
        _kr_backend = _select_secure_keyring()
        if _kr_backend is not None:
            keyring.get_password("eve_trade_probe", "probe")   # backend funcional?
        HAS_KEYRING = _kr_backend is not None
    except Exception as _e:  # backend presente mas inutilizável (D-Bus, bloqueado…)
        log.warning("keyring indisponível: %s", _e)
        HAS_KEYRING = False
except ImportError:
    HAS_KEYRING = False

_KEYRING_SERVICE = "eve_trade_finder"
_KEYRING_USER    = "anthropic_api_key"
API_KEY_ENV      = "ANTHROPIC_API_KEY"


def api_key_storage_status() -> str:
    """Texto curto para a UI dizendo onde a chave fica guardada."""
    if os.environ.get(API_KEY_ENV):
        return f"lida da variável de ambiente {API_KEY_ENV}"
    if HAS_KEYRING:
        return f"salva no keyring do sistema ({type(keyring.get_keyring()).__name__})"
    return (f"sem keyring seguro: a chave fica só na memória desta sessão "
            f"(ou defina {API_KEY_ENV})")


def purge_plaintext_keyring_entry() -> bool:
    """Remove a chave deste app de um PlaintextKeyring legado. True se removeu."""
    try:
        from keyrings.alt.file import PlaintextKeyring
    except ImportError:
        return False
    try:
        pk = PlaintextKeyring()
        if pk.get_password(_KEYRING_SERVICE, _KEYRING_USER):
            pk.delete_password(_KEYRING_SERVICE, _KEYRING_USER)
            log.warning(
                "Chave da API removida do arquivo em texto puro %s", pk.file_path)
            return True
    except Exception as _e:
        log.warning("purge_plaintext_keyring_entry: %s", _e)
    return False


def keyring_save_key(key: str) -> bool:
    """Salva a API Key no keyring do SO. Retorna True se bem-sucedido."""
    if not HAS_KEYRING or not key:
        return False
    try:
        keyring.set_password(_KEYRING_SERVICE, _KEYRING_USER, key)
        return True
    except Exception as exc:
        log.warning("keyring_save_key: %s", exc)
        return False


def keyring_load_key() -> str:
    """Lê a API Key do keyring do SO. Retorna '' se não encontrada."""
    if not HAS_KEYRING:
        return ""
    try:
        return keyring.get_password(_KEYRING_SERVICE, _KEYRING_USER) or ""
    except Exception:
        return ""


def keyring_delete_key() -> bool:
    """Remove a API Key do keyring do SO."""
    if not HAS_KEYRING:
        return False
    try:
        keyring.delete_password(_KEYRING_SERVICE, _KEYRING_USER)
        return True
    except Exception:
        return False

# ── openpyxl — exportação Excel ───────────────────────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import (Font, PatternFill, Alignment,
                                  Border, Side, numbers)
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


def export_xlsx(path: str, sheet_name: str, headers: list,
                rows: list, col_widths: list = None) -> bool:
    """
    Exporta uma lista de dicts para Excel (.xlsx) com formatação profissional.

    headers    — lista de (col_key, col_label) OU só strings (usadas como label e key)
    rows       — lista de dicts (ou listas de valores, se headers for só strings)
    col_widths — larguras opcionais por coluna (auto se None)

    Retorna True se salvou com sucesso.
    """
    if not HAS_OPENPYXL:
        return False

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]   # Excel limita a 31 chars

    # Cores do tema EVE
    _HDR_FILL  = PatternFill("solid", fgColor="1a2030")
    _HDR_FONT  = Font(name="Calibri", bold=True, color="4db8ff", size=10)
    _ODD_FILL  = PatternFill("solid", fgColor="0f1520")
    _EVEN_FILL = PatternFill("solid", fgColor="141a28")
    _BODY_FONT = Font(name="Calibri", color="c8d0e0", size=9)
    _BORDER    = Border(
        bottom=Side(style="thin", color="1e2a3a"),
        right=Side(style="thin",  color="1e2a3a"))
    _ALIGN_L = Alignment(horizontal="left",   vertical="center", wrap_text=False)
    _ALIGN_R = Alignment(horizontal="right",  vertical="center")
    _ALIGN_C = Alignment(horizontal="center", vertical="center")

    # Normaliza headers para lista de (key, label)
    norm_headers = []
    for h in headers:
        if isinstance(h, (list, tuple)) and len(h) == 2:
            norm_headers.append(tuple(h))
        else:
            norm_headers.append((str(h), str(h)))

    keys   = [k for k, _ in norm_headers]
    labels = [l for _, l in norm_headers]

    # Linha de cabeçalho
    ws.row_dimensions[1].height = 22
    for col_idx, label in enumerate(labels, 1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.font      = _HDR_FONT
        cell.fill      = _HDR_FILL
        cell.alignment = _ALIGN_C
        cell.border    = _BORDER

    # Linhas de dados
    for row_idx, row in enumerate(rows, 2):
        fill = _ODD_FILL if row_idx % 2 == 1 else _EVEN_FILL
        ws.row_dimensions[row_idx].height = 18
        for col_idx, key in enumerate(keys, 1):
            if isinstance(row, dict):
                value = row.get(key, "")
            else:
                try: value = row[col_idx - 1]
                except (IndexError, TypeError): value = ""

            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font      = _BODY_FONT
            cell.fill      = fill
            cell.border    = _BORDER
            # Alinhamento por tipo
            if isinstance(value, (int, float)):
                cell.alignment = _ALIGN_R
            else:
                cell.alignment = _ALIGN_L

    # Larguras de coluna
    for col_idx, key in enumerate(keys, 1):
        letter = get_column_letter(col_idx)
        if col_widths and col_idx <= len(col_widths):
            ws.column_dimensions[letter].width = col_widths[col_idx - 1]
        else:
            # Auto: máximo entre label e valores
            max_len = len(labels[col_idx - 1])
            for row in rows[:50]:   # amostra
                v = row.get(key, "") if isinstance(row, dict) else ""
                max_len = max(max_len, len(str(v)))
            ws.column_dimensions[letter].width = min(max_len + 2, 40)

    # Freeze pane no cabeçalho
    ws.freeze_panes = "A2"

    # Auto-filtro
    ws.auto_filter.ref = (
        f"A1:{get_column_letter(len(keys))}{len(rows) + 1}")

    try:
        wb.save(path)
        return True
    except Exception as exc:
        log.error("export_xlsx(%s): %s", path, exc)
        return False

# ── Constantes ────────────────────────────────────────────────────────────────
ESI_BASE           = "https://esi.evetech.net/latest"
DATA_DIR           = os.path.join(os.path.dirname(os.path.abspath(__file__)), "eve_data")
STATION_CACHE_FILE = os.path.join(DATA_DIR, "cache_stations.csv")
ITEM_CACHE_FILE    = os.path.join(DATA_DIR, "cache_items.json")
CONFIG_FILE        = os.path.join(DATA_DIR, "config.json")
WATCHLIST_FILE     = os.path.join(DATA_DIR, "watchlist.json")
PAPER_FILE         = os.path.join(DATA_DIR, "paper_trades.json")
HIST_DB            = os.path.join(DATA_DIR, "market_history.db")
BLACKLIST_FILE     = os.path.join(DATA_DIR, "blacklist.json")
USER_AGENT         = "EVE-Trade-Finder/8.0 (enf.econ@gmail.com)"
_ESI_HEADERS       = {"User-Agent": USER_AGENT}
_etag_cache: dict  = {}   # {url_with_params: (etag_str, data)}  — ETag HTTP cache

os.makedirs(DATA_DIR, exist_ok=True)

# M5 — cache de ordens em memoria: {region_name: [orders]}
_orders_mem_cache: dict = {}   # {region_name: (timestamp_float, orders_list)}
_ORDERS_TTL = 1800             # 30 minutos — ordens mudam a cada 5min no EVE

REGIONS = {
    "The Forge (Jita)":       10000002,
    "Domain (Amarr)":         10000043,
    "Sinq Laison (Dodixie)":  10000032,
    "Metropolis (Hek)":       10000042,
    "Heimatar (Rens)":        10000030,
    "Lonetrek":               10000016,
    "The Citadel":            10000033,
    "Delve":                  10000060,
    "Pure Blind":             10000023,
    "Catch":                  10000014,
}

COLORS = {
    "bg":       "#0a0e1a",
    "panel":    "#0f1624",
    "border":   "#1e2d45",
    "accent":   "#00aaff",
    "green":    "#00ff88",
    "red":      "#ff4455",
    "yellow":   "#ffd700",
    "orange":   "#ff8c00",
    "text":     "#c8d8e8",
    "subtext":  "#5a7a9a",
    "header":   "#7ecfff",
    "row_odd":  "#0d1520",
    "row_even": "#111d2e",
    "tab_sel":  "#162030",
}

_thread_local = threading.local()

def _get_session():
    if not hasattr(_thread_local, "session"):
        s = requests.Session()
        s.headers.update({"User-Agent": USER_AGENT})
        s.mount("https://", requests.adapters.HTTPAdapter(max_retries=0))
        _thread_local.session = s
    return _thread_local.session


# ══════════════════════════════════════════════════════════════════════════════
#  CAMADA DE DADOS — ESI
# ══════════════════════════════════════════════════════════════════════════════

def esi_get(path, params=None, retries=3):
    """
    GET na ESI com suporte a ETag (RFC 7232 / If-None-Match).

    ETag cache: quando a ESI retorna ETag + 304 Not Modified,
    reutiliza o dado em memória sem parsear JSON novamente.
    Reduz tráfego e latência em ~60–80% para endpoints estáveis
    (histórico de mercado, preços ajustados, índices de sistema).
    """
    session = _get_session()
    url     = ESI_BASE + path
    # Monta chave de cache incluindo parâmetros de query
    cache_key = url + (str(sorted(params.items())) if params else "")

    headers = dict(_ESI_HEADERS)
    if cache_key in _etag_cache:
        etag, _ = _etag_cache[cache_key]
        headers["If-None-Match"] = etag

    for attempt in range(retries):
        try:
            r = session.get(url, params=params, headers=headers,
                            timeout=(30, 60))
            if r.status_code == 304:
                # Not Modified — dado em cache ainda válido
                return _etag_cache[cache_key][1]
            if r.status_code == 200:
                data = r.json()
                etag = r.headers.get("ETag")
                if etag:
                    _etag_cache[cache_key] = (etag, data)
                return data
            if r.status_code == 420:
                time.sleep(10)
                continue
            if r.status_code in (502, 503, 504):
                time.sleep(4 * (attempt + 1))
                continue
        except requests.exceptions.Timeout:
            time.sleep(4 * (attempt + 1))
        except Exception as exc:
            log.error("esi_get(%s): %s", path, exc)
            time.sleep(2)
    return None


def esi_post(path, payload):
    url = ESI_BASE + path
    try:
        r = _get_session().post(url, json=payload, timeout=15)
        if r.status_code == 200:
            return r.json()
    except Exception as exc:
        log.error("esi_post(%s): %s", path, exc)
    return []


# (region_id, type_id) cujo download de historico FALHOU e que ficaram sem
# nenhum dado local. Distingue "erro" de "item sem negociacao" na UI.
_hist_fetch_errors: set = set()


def fetch_market_history(region_id, type_id):
    """Historico ESI. [] = item sem negociacao; None = falha de rede/ESI."""
    return esi_get(f"/markets/{region_id}/history/", {"type_id": type_id})


def hist_failed(region_id, type_id) -> bool:
    """True se o historico deste item falhou e nao ha dado local."""
    return (region_id, type_id) in _hist_fetch_errors


def count_hist_failures(region_id, type_ids) -> int:
    return sum(1 for t in type_ids if (region_id, t) in _hist_fetch_errors)


HIST_ERROR_MARK = "⚠"


def _hist_err_note(region_id, type_ids) -> str:
    """Sufixo p/ mensagem final: quantos itens ficaram sem historico por ERRO."""
    n = count_hist_failures(region_id, type_ids)
    return (f"  {HIST_ERROR_MARK} {n} itens sem historico por ERRO de rede/ESI "
            f"(nao e falta de negociacao) — veja o log") if n else ""




# ══════════════════════════════════════════════════════════════════════════════
#  ANALISES DE MERCADO — ABAS 4-8
# ══════════════════════════════════════════════════════════════════════════════



# ── Helpers globais reutilizados por todas as analyze_*() ─────────────────────

def _region_id(region_name: str) -> int | None:
    """2A — lookup de region_id em uma linha. Era for-loop em 7 funcoes."""
    return REGIONS.get(region_name)


def _ensure_names(idx_or_tids, item_cache: dict, progress_cb=None):
    """2B — resolve nomes ausentes em batch e salva.
    Aceita dict (idx) ou list/set de type_ids.
    """
    keys     = idx_or_tids if isinstance(idx_or_tids, (list, set)) else idx_or_tids.keys()
    sem_nome = [t for t in keys if t not in item_cache]
    if sem_nome:
        if progress_cb:
            progress_cb(f"Resolvendo {len(sem_nome)} nomes...")
        item_cache.update(fetch_type_ids_bulk(sem_nome))
        save_item_cache(item_cache)






# ── ABA 4: Max Margem ─────────────────────────────────────────────────────────

# Melhor compra precisa valer ao menos esta fracao da melhor venda.
# Sem isso, ordens simbolicas (0,01 ISK) geram margem ~852% (skills 0) e
# dominam o ranking. 0.10 => margem bruta maxima de 90%.


def analyze_max_margin(orders, region_name, item_cache,
                        broker_rel=0, accounting=0,
                        min_margin=0.0, max_sell_price=0.0,
                        min_vol=1, top_n=200,
                        blacklist=None,
                        progress_cb=None,
                        min_buy_ratio=MM_MIN_BUY_RATIO):
    """Maiores margens liquidas + EDP (I/O aqui; calculo em evetrade.core.margin).

    O filtro de volume (min_vol) e aplicado ANTES do corte em top_n e
    min_buy_ratio descarta compras irrisorias — ver evetrade/core/margin.py.
    """
    idx       = _filter_blacklist(build_order_index(orders), blacklist or set())
    region_id = _region_id(region_name)          # 2A
    _ensure_names(idx, item_cache, progress_cb)  # 2B
    _warm_hist_cache(region_id, progress_cb)     # pré-aquece cache RAM
    candidates = margin_candidates(idx, item_cache, region_name,
                                   broker_rel, accounting, min_margin,
                                   max_sell_price, min_buy_ratio, progress_cb)
    hist = (lambda tid: _get_hist_cached(region_id, tid)) if region_id else None
    results = select_by_volume(candidates, hist, min_vol, top_n, progress_cb)
    if progress_cb:
        progress_cb(f"{len(results)} itens encontrados."
                    + _hist_err_note(region_id, [c["type_id"] for c in candidates]))
    return results


# ── ABA 5: Escassez de Suprimento ────────────────────────────────────────────

def analyze_supply_shortage(orders, region_name, item_cache,
                             max_dos=30.0, min_margin=0.0,
                             min_vol=5, top_n=200,
                             blacklist=None,
                             progress_cb=None):
    """Items In Low Supply: DoS = sell_vol / avg_vol_dia."""
    idx       = _filter_blacklist(build_order_index(orders), blacklist or set())
    region_id = _region_id(region_name)         # 2A
    _ensure_names(idx, item_cache, progress_cb) # 2B
    _warm_hist_cache(region_id, progress_cb)    # pré-aquece cache RAM

    candidates = []
    total = len(idx)
    chunk = max(1, total // 40)
    for i, (tid, entry) in enumerate(idx.items()):
        if progress_cb and i % chunk == 0:
            progress_cb((i, total, f"Filtrando... {i:,}/{total:,}"))
        sells = entry["sell"]; buys = entry["buy"]
        if not sells or not buys: continue
        ps = float(sells[0]["price"]); pb = float(buys[0]["price"])
        if ps <= pb or ps <= 0: continue
        margin = (ps - pb) / ps * 100
        if margin < min_margin: continue
        sell_vol = sum(int(o.get("volume_remain", 0)) for o in sells)
        info = item_cache.get(tid, {})
        candidates.append({
            "type_id": tid, "item": info.get("name", str(tid)),
            "ps": round(ps, 2), "pb": round(pb, 2),
            "margin": round(margin, 2), "sell_vol": sell_vol,
            "region": region_name,
        })

    candidates.sort(key=lambda x: x["sell_vol"])
    candidates = candidates[:top_n * 3]

    if region_id and candidates:
        if progress_cb:
            progress_cb(f"Buscando historico de {len(candidates)} itens...")
        def _fetch(c):
            h  = _get_hist_cached(region_id, c["type_id"])  # O1
            st = _history_stats(h)
            if st and st["avg_vol_90"] >= min_vol:
                vd = st["avg_vol_90"]
                c["avg_vol_90"] = round(vd, 1)
                c["dos"]   = round(c["sell_vol"] / vd if vd > 0 else 9999, 2)
                c["score"] = round(c["margin"] * math.log10(vd + 1) / (c["dos"] + 1), 4)
            else:
                c["avg_vol_90"] = 0.0; c["dos"] = 9999.0; c["score"] = 0.0
        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as pool:
            list(pool.map(_fetch, candidates))

    results = [c for c in candidates
               if c.get("dos", 9999) <= max_dos and c.get("avg_vol_90", 0) >= min_vol]
    results.sort(key=lambda x: x.get("score", 0), reverse=True)
    results = results[:top_n]
    if progress_cb: progress_cb(f"{len(results)} itens em escassez."
                                + _hist_err_note(region_id, idx.keys()))
    return results


# ── ABA 6: Demanda / Oferta (DSIF) ───────────────────────────────────────────

def analyze_demand_supply(orders, region_name, item_cache,
                           min_margin=0.0, min_vol=5, top_n=200,
                           blacklist=None,
                           progress_cb=None):
    """High Demand Low Supply — DSIF = (Bv/Sv)^2 * (Vd/Sv) * log10(Ps-Pb+1)."""
    idx       = _filter_blacklist(build_order_index(orders), blacklist or set())
    region_id = _region_id(region_name)         # 2A
    _ensure_names(idx, item_cache, progress_cb) # 2B
    _warm_hist_cache(region_id, progress_cb)    # pré-aquece cache RAM

    candidates = []
    total = len(idx)
    chunk = max(1, total // 40)
    for i, (tid, entry) in enumerate(idx.items()):
        if progress_cb and i % chunk == 0:
            progress_cb((i, total, f"Calculando DSIF... {i:,}/{total:,}"))
        sells = entry["sell"]; buys = entry["buy"]
        if not sells or not buys: continue
        ps = float(sells[0]["price"]); pb = float(buys[0]["price"])
        if ps <= pb or ps <= 0: continue
        margin = (ps - pb) / ps * 100
        if margin < min_margin: continue
        bv = sum(int(o.get("volume_remain", 0)) for o in buys)
        sv = sum(int(o.get("volume_remain", 0)) for o in sells)
        if sv == 0: continue
        info = item_cache.get(tid, {})
        candidates.append({
            "type_id": tid, "item": info.get("name", str(tid)),
            "ps": round(ps, 2), "pb": round(pb, 2),
            "margin": round(margin, 2),
            "buy_vol": bv, "sell_vol": sv,
            "ratio": round(bv / sv, 3),
            "region": region_name,
        })

    candidates.sort(key=lambda x: x["ratio"], reverse=True)
    candidates = candidates[:top_n * 3]

    if region_id and candidates:
        if progress_cb:
            progress_cb(f"Buscando historico de {len(candidates)} itens...")
        def _fetch(c):
            h  = _get_hist_cached(region_id, c["type_id"])  # O1
            st = _history_stats(h)
            vd = st["avg_vol_90"] if st else 1.0
            if vd < min_vol:
                c["dsif"] = 0.0; c["avg_vol_90"] = 0.0; return
            bv, sv = c["buy_vol"], c["sell_vol"]
            ps, pb = c["ps"], c["pb"]
            c["dsif"]       = round((bv/sv)**2 * (vd/sv) * math.log10(ps - pb + 1), 4)
            c["avg_vol_90"] = round(vd, 1)
        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as pool:
            list(pool.map(_fetch, candidates))

    results = [c for c in candidates if c.get("avg_vol_90", 0) >= min_vol]
    results.sort(key=lambda x: x.get("dsif", 0), reverse=True)
    results = results[:top_n]
    if progress_cb: progress_cb(f"{len(results)} itens com alta demanda."
                                + _hist_err_note(region_id, idx.keys()))
    return results


# ── ABA 7: Movimentação de Preços ─────────────────────────────────────────────

def analyze_price_movement(region_name, item_cache,
                            min_vol=5, min_margin=0.0, top_n=200,
                            orders=None, type_ids=None, progress_cb=None):
    """
    Price Movement: compara preco/volume da semana vs mes.
    type_ids: lista explicita de type_ids a analisar (opcional).
              Se None, usa orders (se fornecido) ou item_cache.
    """
    region_id = _region_id(region_name)
    if not region_id:
        return []

    if type_ids is not None:
        tids_raw = list(type_ids)
    elif orders:
        idx      = build_order_index(orders)
        tids_raw = list(idx.keys())
    else:
        tids_raw = list(item_cache.keys())

    total = len(tids_raw)

    # Pré-aquece o cache RAM com 1 query SQL antes das threads
    _warm_hist_cache(region_id, progress_cb)
    if progress_cb:
        progress_cb(f"Calculando movimentação de {total:,} itens...")

    done_cnt = [0]
    lock     = threading.Lock()
    chunk    = max(1, total // 40)   # ~40 atualizações de progresso

    def _fetch(tid):
        h  = _get_hist_cached(region_id, tid)   # hit no cache RAM — O(1)
        st = _history_stats(h)
        with lock:
            done_cnt[0] += 1
            if progress_cb and done_cnt[0] % chunk == 0:
                progress_cb((done_cnt[0], total,
                             f"Processando: {done_cnt[0]:,}/{total:,}"))
        if not st or st["avg_vol_90"] < min_vol:
            return None
        pw, pm = st["avg_p_week"], st["avg_p_month"]
        vw, vm = st["avg_v_week"], st["avg_v_month"]
        if pm <= 0 or vm <= 0:
            return None
        pc = (pw - pm) / pm * 100
        vc = (vw - vm) / vm * 100
        margin = 0.0
        if orders:
            sells = idx.get(tid, {}).get("sell", [])
            buys  = idx.get(tid, {}).get("buy",  [])
            if sells and buys:
                ps = float(sells[0]["price"]); pb = float(buys[0]["price"])
                if ps > 0:
                    margin = (ps - pb) / ps * 100
        if margin < min_margin:
            return None
        info = item_cache.get(tid, {})
        return {
            "type_id":     tid,
            "item":        info.get("name", str(tid)),
            "avg_p_week":  round(pw, 2),  "avg_p_month": round(pm, 2),
            "price_chg":   round(pc, 2),
            "avg_v_week":  round(vw, 1),  "avg_v_month": round(vm, 1),
            "vol_chg":     round(vc, 2),
            "margin":      round(margin, 2),
            "correlation": "Positiva" if (pc >= 0) == (vc >= 0) else "Negativa",
            "score":       round(math.log10(st["avg_vol_90"] + 1), 4),
            "avg_vol_90":  round(st["avg_vol_90"], 1),
            "region":      region_name,
        }

    results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as pool:
        for r in pool.map(_fetch, tids_raw):
            if r:
                results.append(r)

    results.sort(key=lambda x: x["score"], reverse=True)
    results = results[:top_n]
    if progress_cb:
        progress_cb(f"{len(results)} itens com movimentacao."
                    + _hist_err_note(region_id, tids_raw))
    return results


# ── ABA 8: Orçamento (Greedy Knapsack) ───────────────────────────────────────


def analyze_budget(orders, region_name, item_cache,
                   budget: float, broker_rel: int, accounting: int,
                   risk_mode: str = "Balanced (3d)",
                   min_vol: int = 1,
                   blacklist=None,
                   progress_cb=None):
    """Greedy knapsack (I/O aqui; calculo em evetrade.core.budget)."""
    idx       = _filter_blacklist(build_order_index(orders), blacklist or set())
    region_id = _region_id(region_name)         # 2A
    _ensure_names(idx, item_cache, progress_cb) # 2B
    _warm_hist_cache(region_id, progress_cb)    # pré-aquece cache RAM
    hist = (lambda tid: _get_hist_cached(region_id, tid)) if region_id else None
    basket, stats = allocate_budget(idx, item_cache, region_name, budget,
                                    broker_rel, accounting, risk_mode, min_vol,
                                    hist, progress_cb)
    if progress_cb:
        progress_cb(f"{len(basket)} itens no portfolio."
                    + _hist_err_note(region_id, idx.keys()))
    return basket, stats


# ══════════════════════════════════════════════════════════════════════════════
#  MERCADO LIVE — busca de preco atual por type_id e regiao
# ══════════════════════════════════════════════════════════════════════════════

# Itens populares de Jita para o quadro de mercado (type_id, nome)
MARKET_ITEMS = [
    (34,    "Tritanium"),
    (35,    "Pyerite"),
    (36,    "Mexallon"),
    (37,    "Isogen"),
    (38,    "Nocxium"),
    (39,    "Zydrine"),
    (40,    "Megacyte"),
    (11399, "Morphite"),
    (9848,  "Zydrine Ore"),
    (29668, "PLEX"),
    (40520, "Skill Injector"),
    (40519, "Skill Extractor"),
    (24698, "Compressed Tritanium"),
    (24702, "Compressed Pyerite"),
    (24704, "Compressed Mexallon"),
    (24706, "Compressed Isogen"),
    (24708, "Compressed Nocxium"),
    (24710, "Compressed Zydrine"),
    (24712, "Compressed Megacyte"),
    (587,   "Rifter"),
    (588,   "Merlin"),
    (582,   "Kestrel"),
    (583,   "Condor"),
    (590,   "Tristan"),
    (589,   "Incursus"),
    (605,   "Imicus"),
    (596,   "Breacher"),
    (32880, "Cerebral Accelerator"),
]

def fetch_live_price(region_id: int, type_id: int):
    """
    Retorna dict {sell_min, buy_max, spread_pct, vol_7d, net_pct, score}
    para um type_id numa regiao. Historico vem do cache com TTL 1h (O1/C1).
    """
    try:
        url = f"{ESI_BASE}/markets/{region_id}/orders/"
        params = {"datasource": "tranquility", "order_type": "all",
                  "type_id": type_id}
        r = _get_session().get(url, params=params, timeout=10)
        if r.status_code != 200:
            return None
        orders = r.json()
    except Exception:
        return None

    sells = sorted([float(o["price"]) for o in orders if not o["is_buy_order"]])
    buys  = sorted([float(o["price"]) for o in orders if  o["is_buy_order"]], reverse=True)
    if not sells or not buys:
        return None

    sell_min = sells[0]
    buy_max  = buys[0]
    spread   = (sell_min - buy_max) / sell_min * 100 if sell_min > 0 else 0

    # Volume 7 dias — via cache com TTL 1h (nao faz request extra a cada tick)
    hist    = _get_hist_cached(region_id, type_id)
    cutoff7 = (datetime.now(timezone.utc) - timedelta(days=7)).strftime("%Y-%m-%d")
    h7      = [d for d in hist if d.get("date","") >= cutoff7]
    vol7    = sum(d.get("volume",0) for d in h7) / max(len(h7), 1)

    # Margem liquida (sem skills = pior caso)
    rb, rt = _calc_fee_rates(0, 0)
    cost   = buy_max + buy_max*rb + sell_min*rb + sell_min*rt
    pi     = sell_min - cost
    net_pct= pi / cost * 100 if cost > 0 else 0

    # Score 0-10: ponderando margem liquida e giro
    score = min(10, max(0, round(
        (net_pct / 5) * 0.5 + (math.log10(vol7 + 1) / 9) * 5, 1)))

    return {
        "type_id":    type_id,
        "sell_min":   round(sell_min, 2),
        "buy_max":    round(buy_max,  2),
        "spread_pct": round(spread,   2),
        "vol_7d":     round(vol7,     0),
        "net_pct":    round(net_pct,  2),
        "score":      score,
    }


# ── Cache de historico ESI com TTL de 1 hora (O1) ─────────────────────────────
# Chave: (region_id, type_id)  →  (timestamp_float, list_of_history_dicts)
# ══════════════════════════════════════════════════════════════════════════════
#  BANCO DE HISTORICO LOCAL — SQLite incremental (market_history.db)
# ══════════════════════════════════════════════════════════════════════════════
#
#  Estrategia:
#   1. Primeira vez: baixa historico completo ESI e armazena no SQLite.
#   2. Dias seguintes: descobre a data maxima ja armazenada por item e
#      baixa apenas os dias novos — normalmente 1 registro por item/dia.
#   3. _get_hist_cached() le do SQLite (disco), com cache RAM de 1h por cima.
#   4. sync_history_db() e chamado pelo botao "SINCRONIZAR BD" na aba MOV. PRECOS.
#
#  Tabela unica: history(region_id, type_id, date, average, highest, lowest, volume)
#  Chave primaria: (region_id, type_id, date) — INSERT OR IGNORE e seguro e idempotente.

_hist_cache: dict = {}   # cache RAM: {(rid,tid): (timestamp, list)}
_HIST_TTL   = 3600       # segundos antes de re-ler do SQLite

# ── Conexoes thread-local (sqlite3 nao e thread-safe por padrao) ──────────────
_db_local = threading.local()

def _get_db() -> sqlite3.Connection:
    """Retorna conexao SQLite do thread atual, criando se necessario."""
    if not hasattr(_db_local, "conn") or _db_local.conn is None:
        conn = sqlite3.connect(HIST_DB, check_same_thread=False)
        conn.execute("PRAGMA journal_mode=WAL")   # leituras sem bloquear escritas
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS history (
                region_id  INTEGER NOT NULL,
                type_id    INTEGER NOT NULL,
                date       TEXT    NOT NULL,
                average    REAL    DEFAULT 0,
                highest    REAL    DEFAULT 0,
                lowest     REAL    DEFAULT 0,
                volume     INTEGER DEFAULT 0,
                PRIMARY KEY (region_id, type_id, date)
            )
        """)
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_hist_region_type "
            "ON history (region_id, type_id)")
        conn.commit()
        _db_local.conn = conn
    return _db_local.conn


def _db_max_date(region_id: int, type_id: int) -> str | None:
    """Retorna a data mais recente armazenada para (region_id, type_id), ou None."""
    try:
        conn = _get_db()
        row  = conn.execute(
            "SELECT MAX(date) FROM history WHERE region_id=? AND type_id=?",
            (region_id, type_id)).fetchone()
        return row[0] if row and row[0] else None
    except Exception as exc:
        log.warning("_db_max_date(%s,%s): %s", region_id, type_id, exc)
        return None


def _db_read(region_id: int, type_id: int) -> list:
    """Le todos os registros de historico do SQLite para um item."""
    try:
        conn = _get_db()
        rows = conn.execute(
            "SELECT date, average, highest, lowest, volume "
            "FROM history WHERE region_id=? AND type_id=? ORDER BY date",
            (region_id, type_id)).fetchall()
        return [{"date": r[0], "average": r[1], "highest": r[2],
                 "lowest": r[3], "volume": r[4]} for r in rows]
    except Exception as exc:
        log.warning("_db_read(%s,%s): %s", region_id, type_id, exc)
        return []


def _db_write(region_id: int, type_id: int, records: list):
    """Insere registros novos no SQLite (INSERT OR IGNORE = idempotente)."""
    if not records:
        return
    try:
        conn = _get_db()
        conn.executemany(
            "INSERT OR IGNORE INTO history "
            "(region_id, type_id, date, average, highest, lowest, volume) "
            "VALUES (?,?,?,?,?,?,?)",
            [(region_id, type_id,
              r.get("date",""),
              r.get("average", 0),
              r.get("highest", 0),
              r.get("lowest",  0),
              r.get("volume",  0)) for r in records])
        conn.commit()
    except Exception as exc:
        log.warning("_db_write(%s,%s): %s", region_id, type_id, exc)


def _db_item_count(region_id: int) -> int:
    """Retorna quantos type_ids distintos existem no banco para uma regiao."""
    try:
        conn = _get_db()
        row  = conn.execute(
            "SELECT COUNT(DISTINCT type_id) FROM history WHERE region_id=?",
            (region_id,)).fetchone()
        return row[0] if row else 0
    except Exception:
        return 0


def _db_bulk_load_region(region_id: int) -> dict:
    """
    Carrega TODO o historico de uma regiao em 1 unica query SQL.

    Em vez de 18747 queries individuais (padrao anterior do station_trader_scan),
    executa UMA query e organiza os dados em memoria como:
      {type_id: [{"date":..., "average":..., "highest":...,
                  "lowest":..., "volume":...}, ...]}

    Retorna dict vazio se o banco nao tiver dados para a regiao.
    Tipicamente carrega ~1-5 MB de dados em <500ms para regioes grandes.
    """
    result: dict = {}
    try:
        conn = _get_db()
        rows = conn.execute(
            "SELECT type_id, date, average, highest, lowest, volume "
            "FROM history WHERE region_id=? ORDER BY type_id, date",
            (region_id,)
        ).fetchall()
        for tid, date, avg, hi, lo, vol in rows:
            if tid not in result:
                result[tid] = []
            result[tid].append({
                "date":    date,
                "average": avg,
                "highest": hi,
                "lowest":  lo,
                "volume":  vol,
            })
    except Exception as exc:
        log.warning("_db_bulk_load_region(%s): %s", region_id, exc)
    return result


def _db_stats(region_id: int) -> dict:
    """Retorna estatisticas do banco para a regiao: n_items, min_date, max_date."""
    try:
        conn = _get_db()
        row  = conn.execute(
            "SELECT COUNT(DISTINCT type_id), MIN(date), MAX(date) "
            "FROM history WHERE region_id=?", (region_id,)).fetchone()
        if row:
            return {"n_items": row[0] or 0,
                    "min_date": row[1] or "--",
                    "max_date": row[2] or "--"}
    except Exception as _e:

        log.debug("silenced: %s", _e)
    return {"n_items": 0, "min_date": "--", "max_date": "--"}


def fetch_history_incremental(region_id: int, type_id: int) -> list:
    """
    Busca historico ESI apenas dos dias ausentes no banco local.

    FIX: compara com yesterday (D-1) em vez de today.
    A ESI so publica o historico do dia anterior apos ~02:00 UTC.
    Usar 'today' como cutoff fazia este metodo baixar todos os itens
    novamente a cada sync mesmo sem dados novos disponiveis.
    """
    max_local = _db_max_date(region_id, type_id)
    yesterday = (datetime.now(timezone.utc) - timedelta(days=1)).strftime("%Y-%m-%d")

    # Precisa buscar apenas se ainda não tem dados de ontem
    needs_fetch = (max_local is None) or (max_local < yesterday)

    if needs_fetch:
        remote = fetch_market_history(region_id, type_id)
        if remote is None:
            if max_local is None:           # falhou e nao ha nada local
                _hist_fetch_errors.add((region_id, type_id))
            else:                           # usa dado local (desatualizado)
                log.warning("historico %s/%s desatualizado (ESI falhou; local ate %s)",
                            region_id, type_id, max_local)
        else:
            _hist_fetch_errors.discard((region_id, type_id))
        if remote:
            if max_local:
                new_records = [r for r in remote if r.get("date","") > max_local]
            else:
                new_records = remote
            _db_write(region_id, type_id, new_records)

    return _db_read(region_id, type_id)


def _fetch_incremental_with_known(region_id: int, type_id: int,
                                   known_max: str | None,
                                   cutoff: str) -> None:
    """
    Versão otimizada de fetch_history_incremental que recebe a data maxima
    já conhecida pelo bulk load, evitando uma query SQL individual por item.

    known_max — data maxima do BD (já carregada pelo bulk load em sync_history_db)
    cutoff    — data alvo (yesterday): só baixa se known_max < cutoff
    """
    if known_max and known_max >= cutoff:
        return  # já atualizado até o limite disponível na ESI

    remote = fetch_market_history(region_id, type_id)
    if not remote:
        return

    if known_max:
        # Só insere registros mais novos que o último já armazenado
        new_records = [r for r in remote if r.get("date", "") > known_max]
    else:
        new_records = remote

    if new_records:
        _db_write(region_id, type_id, new_records)


def sync_history_db(region_id: int, type_ids: list,
                    progress_cb=None, stop_flag: list = None) -> dict:
    """
    Sincroniza o banco local para todos os type_ids de uma regiao.

    FIX: A ESI publica o historico de mercado com atraso de ~1 dia.
    O ultimo registro disponivel na ESI e sempre de ONTEM, nunca de hoje.
    Comparar com 'today' marcava TODOS os itens como desatualizados a cada
    execucao, tornando toda sync equivalente a uma primeira vez.
    Agora compara com 'yesterday' — se o BD ja tem dados de ontem, nada
    e baixado ate que o novo dia fique disponivel (~02:00 UTC).

    Otimizacao: carrega as datas maximas de TODOS os itens em 1 query SQL
    e passa o resultado para fetch_history_incremental, evitando N queries
    individuais de _db_max_date.

    stop_flag: lista de 1 elemento; se stop_flag[0]=True, aborta.
    Retorna dict com estatisticas {total, updated, skipped, errors}.
    """
    total     = len(type_ids)
    # ESI publica historico de D-1. Usar yesterday evita baixar tudo
    # novamente toda vez que a data local (hoje) ainda nao esta na ESI.
    yesterday = (datetime.now(timezone.utc) - timedelta(days=1)).strftime("%Y-%m-%d")

    # 1 query SQL carrega data maxima de todos os itens da região
    if progress_cb:
        progress_cb("Verificando banco de dados (1 query)...")
    try:
        conn = _get_db()
        rows = conn.execute(
            "SELECT type_id, MAX(date) FROM history "
            "WHERE region_id=? GROUP BY type_id",
            (region_id,)
        ).fetchall()
        max_dates = {tid: d for tid, d in rows}
    except Exception as exc:
        log.warning("sync_history_db bulk max_date: %s", exc)
        max_dates = {}

    # Item precisa de update se sua data maxima é anterior a ontem
    # (usa < em vez de != — mais robusto para datas muito antigas)
    needs_update = [tid for tid in type_ids
                    if (max_dates.get(tid) or "") < yesterday]
    already_ok   = total - len(needs_update)

    if progress_cb:
        progress_cb(f"{already_ok:,} já atualizados até {yesterday}, "
                    f"{len(needs_update):,} precisam de sync...")

    if not needs_update:
        return {"total": total, "updated": 0,
                "skipped": already_ok, "errors": 0}

    updated = [0]; errors_cnt = [0]
    done    = [0]
    lock    = threading.Lock()
    chunk   = max(1, len(needs_update) // 40)

    def _sync_one(tid):
        if stop_flag and stop_flag[0]:
            return
        try:
            # Passa max_date conhecido do bulk load — evita query individual
            known_max = max_dates.get(tid)
            _fetch_incremental_with_known(region_id, tid, known_max, yesterday)
            with lock:
                updated[0] += 1
        except Exception as exc:
            log.warning("sync_one(%s,%s): %s", region_id, tid, exc)
            with lock:
                errors_cnt[0] += 1
        with lock:
            done[0] += 1
            if progress_cb and done[0] % chunk == 0:
                progress_cb((done[0], len(needs_update),
                             f"Sincronizando... {done[0]:,}/{len(needs_update):,} "
                             f"(+{updated[0]} novos, {already_ok} ok)"))

    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as pool:
        list(pool.map(_sync_one, needs_update))

    return {"total": total,
            "updated":  updated[0],
            "skipped":  already_ok,
            "errors":   errors_cnt[0]}


# ── Cache RAM de 1h em cima do SQLite ─────────────────────────────────────────

def _get_hist_cached(region_id: int, type_id: int) -> list:
    """
    Retorna historico do item:
      1. Cache RAM (TTL 1h) — sem I/O.
      2. SQLite local       — sem request ESI se o dia ja foi baixado.
      3. ESI (fetch_history_incremental) — so se o banco esta desatualizado.
    """
    key   = (region_id, type_id)
    entry = _hist_cache.get(key)
    if entry is not None:
        ts, hist = entry
        if time.time() - ts < _HIST_TTL:
            return hist

    hist = fetch_history_incremental(region_id, type_id)
    _hist_cache[key] = (time.time(), hist)
    return hist


def _warm_hist_cache(region_id: int, progress_cb=None) -> int:
    """
    Pré-aquece o cache RAM com TODO o histórico da região em 1 query SQL.

    Chamado no início de analyze_price_movement, analyze_max_margin,
    analyze_supply_shortage, analyze_demand_supply, analyze_market_making,
    detect_cycle_region e analyze_budget — antes de lançar as threads.

    Após esta chamada, cada _get_hist_cached() encontra os dados no
    _hist_cache e retorna em O(1) sem tocar o SQLite ou a rede.

    Retorna o número de itens carregados.
    """
    # Verifica se já está quente (TTL ainda válido para pelo menos 1 item da região)
    now = time.time()
    sample = next(
        ((ts, h) for (rid, _), (ts, h) in _hist_cache.items()
         if rid == region_id), None)
    if sample and now - sample[0] < _HIST_TTL:
        # Cache já populado e fresco — conta quantos itens temos
        n = sum(1 for (rid, _) in _hist_cache if rid == region_id)
        if n > 0:
            if progress_cb:
                progress_cb(f"Cache RAM já aquecido: {n:,} itens da região.")
            return n

    if progress_cb:
        progress_cb("Carregando histórico da região (1 query SQL)...")

    bulk = _db_bulk_load_region(region_id)
    ts   = time.time()
    for tid, hist in bulk.items():
        _hist_cache[(region_id, tid)] = (ts, hist)

    n = len(bulk)
    if progress_cb:
        progress_cb(f"Cache aquecido: {n:,} itens em memória.")
    return n


# ══════════════════════════════════════════════════════════════════════════════
#  WATCHLIST — persistencia e alertas
# ══════════════════════════════════════════════════════════════════════════════

def load_watchlist() -> list:
    if not os.path.exists(WATCHLIST_FILE):
        return []
    try:
        with open(WATCHLIST_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def save_watchlist(items: list):
    try:
        with open(WATCHLIST_FILE, "w", encoding="utf-8") as f:
            json.dump(items, f, indent=2)
    except Exception as exc:
        log.warning("Erro ao salvar watchlist: %s", exc)

def check_watchlist_alerts(watchlist: list, region_id: int) -> list:
    """
    Busca preco de todos os itens da watchlist em paralelo (C2) e retorna
    lista de alertas disparados.  Salva watchlist apenas se houve mudanca (M6).
    """
    if not watchlist:
        return []

    # C2 — paralelo igual ao _live_worker
    results: dict = {}

    def _fetch_one(w):
        tid = w.get("type_id")
        if tid:
            d = fetch_live_price(region_id, tid)
            if d:
                results[tid] = d

    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
        list(pool.map(_fetch_one, watchlist))

    alerts  = []
    changed = False
    for w in watchlist:
        tid  = w.get("type_id")
        data = results.get(tid)
        if not data:
            continue
        price = data["sell_min"]

        # M6 — detectar mudanca antes de marcar changed
        if (w.get("last_sell") != price or
                w.get("last_buy")   != data["buy_max"] or
                w.get("last_net")   != data["net_pct"]):
            changed = True

        w["last_sell"]  = price
        w["last_buy"]   = data["buy_max"]
        w["last_net"]   = data["net_pct"]
        w["last_score"] = data["score"]

        if w.get("alert_below") and price <= w["alert_below"]:
            alerts.append({"name": w["name"], "type_id": tid,
                           "kind": "ABAIXO DE", "threshold": w["alert_below"],
                           "current_price": price})
        if w.get("alert_above") and price >= w["alert_above"]:
            alerts.append({"name": w["name"], "type_id": tid,
                           "kind": "ACIMA DE", "threshold": w["alert_above"],
                           "current_price": price})

    if changed:          # M6 — salva so se algo mudou
        save_watchlist(watchlist)
    return alerts


# ══════════════════════════════════════════════════════════════════════════════
#  PAPER TRADING — simulacao de ordens com impostos reais
# ══════════════════════════════════════════════════════════════════════════════

def load_paper_trades() -> dict:
    if not os.path.exists(PAPER_FILE):
        return {"positions": [], "history": [], "balance": 1_000_000_000.0,
                "next_id": 1}
    try:
        with open(PAPER_FILE, encoding="utf-8") as f:
            data = json.load(f)
        if "next_id" not in data:          # migrar arquivos antigos
            data["next_id"] = (
                max((p["id"] for p in data.get("positions", [])), default=0) +
                max((h["id"] for h in data.get("history",   [])), default=0) + 1)
        return data
    except Exception:
        return {"positions": [], "history": [], "balance": 1_000_000_000.0,
                "next_id": 1}

def save_paper_trades(data: dict):
    try:
        with open(PAPER_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, default=str)
    except Exception as exc:
        log.warning("Erro ao salvar paper trades: %s", exc)

def paper_open(data: dict, type_id: int, name: str, qty: int,
               buy_price: float, broker_rel: int, accounting: int) -> dict:
    """Abre posicao de compra no paper trading, debita saldo + broker fee."""
    rb, _  = _calc_fee_rates(broker_rel, accounting)
    fee    = buy_price * rb * qty
    total  = buy_price * qty + fee
    if total > data["balance"]:
        return {"ok": False, "msg": f"Saldo insuficiente ({data['balance']:,.0f} ISK)"}
    data["balance"] -= total
    pos_id = data.get("next_id", 1)        # A5 — ID monotonicamente crescente
    data["next_id"] = pos_id + 1
    pos = {
        "id":          pos_id,
        "type_id":     type_id,
        "name":        name,
        "qty":         qty,
        "buy_price":   round(buy_price, 2),
        "broker_fee":  round(fee, 2),
        "total_cost":  round(total, 2),
        "opened_at":   datetime.now(timezone.utc).isoformat(),
        "status":      "open",
    }
    data["positions"].append(pos)
    save_paper_trades(data)
    return {"ok": True, "pos": pos}

def paper_close(data: dict, pos_id: int, sell_price: float,
                broker_rel: int, accounting: int) -> dict:
    """Fecha posicao existente, calcula P&L liquido."""
    rb, rt = _calc_fee_rates(broker_rel, accounting)
    pos = next((p for p in data["positions"] if p["id"] == pos_id), None)
    if not pos:
        return {"ok": False, "msg": "Posicao nao encontrada"}
    qty    = pos["qty"]
    fee_b  = sell_price * rb * qty
    tax    = sell_price * rt * qty
    gross  = sell_price * qty - fee_b - tax
    pnl    = gross - pos["total_cost"]
    pnl_pct= pnl / pos["total_cost"] * 100 if pos["total_cost"] else 0
    data["balance"] += gross
    pos.update({
        "sell_price":  round(sell_price, 2),
        "sell_fee":    round(fee_b, 2),
        "sales_tax":   round(tax, 2),
        "pnl":         round(pnl, 2),
        "pnl_pct":     round(pnl_pct, 2),
        "closed_at":   datetime.now(timezone.utc).isoformat(),
        "status":      "closed",
    })
    data["positions"].remove(pos)
    data["history"].append(pos)
    save_paper_trades(data)
    return {"ok": True, "pnl": pnl, "pnl_pct": pnl_pct}


# ══════════════════════════════════════════════════════════════════════════════
#  CANDLESTICK + INDICADORES — calculo puro Python
# ══════════════════════════════════════════════════════════════════════════════

def compute_indicators(hist: list) -> dict:
    """
    Recebe lista de dicts ESI {date, average, volume, highest, lowest}.
    Retorna arrays: dates, opens, highs, lows, closes, volumes,
    sma7, sma26, bb_upper, bb_lower, bb_mid, rsi, macd_line, signal,
    macd_hist, fib_levels.

    Melhorias de eficiencia:
      A1 — sma() usa soma deslizante O(n) em vez de O(n*w)
      A2 — ema26 semente calculada direto (elimina sma26 extra)
      A3 — BB reutiliza bb_mid em vez de recalcular mu a cada iteracao
    """
    if len(hist) < 2:
        return {}
    hist   = sorted(hist, key=lambda d: d.get("date",""))
    dates  = [d["date"]           for d in hist]
    closes = [float(d["average"]) for d in hist]
    highs  = [float(d["highest"]) for d in hist]
    lows   = [float(d["lowest"])  for d in hist]
    vols   = [int(d["volume"])    for d in hist]
    opens  = [closes[0]] + closes[:-1]   # aproximacao: open = close anterior

    n = len(closes)

    # A1 — SMA com soma deslizante: O(n) em vez de O(n*w)
    def sma(data: list, w: int) -> list:
        out = [None] * n
        if w > n:
            return out
        s = sum(data[:w])
        out[w - 1] = s / w
        for i in range(w, n):
            s += data[i] - data[i - w]
            out[i] = s / w
        return out

    # EMA — seed no primeiro valor valido, propagacao exponencial
    def ema(data: list, w: int) -> list:
        out  = [None] * n
        k    = 2.0 / (w + 1)
        seed = next((i for i in range(n) if data[i] is not None), None)
        if seed is None:
            return out
        out[seed] = data[seed]
        for i in range(seed + 1, n):
            if data[i] is not None and out[i - 1] is not None:
                out[i] = data[i] * k + out[i - 1] * (1.0 - k)
        return out

    # SMA 7 e 26
    sma7  = sma(closes, 7)
    sma26 = sma(closes, 26)

    # A3 — Bollinger Bands (20, 2σ): reutiliza bb_mid em vez de recalcular mu
    bb_mid   = sma(closes, 20)
    bb_upper = [None] * n
    bb_lower = [None] * n
    for i in range(19, n):
        mu  = bb_mid[i]                              # A3: ja calculado
        slc = closes[i - 19: i + 1]
        sd  = (sum((x - mu) ** 2 for x in slc) / 20) ** 0.5
        bb_upper[i] = mu + 2.0 * sd
        bb_lower[i] = mu - 2.0 * sd

    # RSI (14) — Wilder smoothing
    rsi = [None] * n
    if n > 14:
        gains  = []
        losses = []
        for i in range(1, n):
            diff = closes[i] - closes[i - 1]
            gains.append(max(diff, 0.0))
            losses.append(max(-diff, 0.0))
        avg_g = sum(gains[:14])  / 14.0
        avg_l = sum(losses[:14]) / 14.0
        for i in range(14, n):
            avg_g = (avg_g * 13.0 + gains[i - 1])  / 14.0
            avg_l = (avg_l * 13.0 + losses[i - 1]) / 14.0
            rs    = avg_g / avg_l if avg_l > 0 else 9999.0
            rsi[i] = round(100.0 - 100.0 / (1.0 + rs), 2)

    # A2 — MACD (12, 26, 9): ema26 ja cobre a sma26 (mesma informacao)
    ema12     = ema(closes, 12)
    ema26     = ema(closes, 26)
    macd_line = [
        (ema12[i] - ema26[i])
        if (ema12[i] is not None and ema26[i] is not None) else None
        for i in range(n)]
    signal    = ema(macd_line, 9)
    macd_hist = [
        (macd_line[i] - signal[i])
        if (macd_line[i] is not None and signal[i] is not None) else None
        for i in range(n)]

    # Fibonacci — range das ultimas 90 entradas
    window    = closes[-90:] if len(closes) >= 90 else closes
    fib_hi    = max(window)
    fib_lo    = min(window)
    fib_diff  = fib_hi - fib_lo
    fib_levels = {
        "0.0":   round(fib_lo, 2),
        "23.6":  round(fib_lo + fib_diff * 0.236, 2),
        "38.2":  round(fib_lo + fib_diff * 0.382, 2),
        "50.0":  round(fib_lo + fib_diff * 0.500, 2),
        "61.8":  round(fib_lo + fib_diff * 0.618, 2),
        "78.6":  round(fib_lo + fib_diff * 0.786, 2),
        "100.0": round(fib_hi, 2),
    }

    return {
        "dates": dates, "opens": opens, "highs": highs,
        "lows":  lows,  "closes": closes, "volumes": vols,
        "sma7":  sma7,  "sma26": sma26,
        "bb_upper": bb_upper, "bb_lower": bb_lower, "bb_mid": bb_mid,
        "rsi":       rsi,
        "macd_line": macd_line, "signal": signal, "macd_hist": macd_hist,
        "fib_levels": fib_levels,
    }


def _ind_last(indicators: dict, key: str) -> float | None:
    """A6 — extrai o ultimo valor nao-None de um array de indicadores."""
    arr = indicators.get(key) or []
    for v in reversed(arr):
        if v is not None:
            return v
    return None


def ai_analyze(type_name: str, indicators: dict,
               sell_min: float, buy_max: float,
               net_margin: float, api_key: str,
               stream_cb=None) -> str:
    """
    Chama Anthropic API para analisar o item.

    A6 — usa _ind_last() centralizado em vez de filtrar listas duas vezes.
    O6 — suporta streaming via stream_cb(token: str) se fornecido;
         nesse caso retorna string vazia (o caller acumula via callback).
    """
    if not api_key or not api_key.strip():
        return "Configure sua API Key da Anthropic na aba AI ANALISTA."
    if not indicators:
        return "Dados insuficientes para analise."

    closes     = indicators.get("closes", [])
    last_close = closes[-1]        if closes else 0
    last_rsi   = _ind_last(indicators, "rsi")       or 0
    last_macd  = _ind_last(indicators, "macd_line") or 0
    last_sig   = _ind_last(indicators, "signal")    or 0
    fib        = indicators.get("fib_levels", {})

    chg7  = ((closes[-1] - closes[-7])  / closes[-7]  * 100) if len(closes) >= 7  and closes[-7]  else 0
    chg30 = ((closes[-1] - closes[-30]) / closes[-30] * 100) if len(closes) >= 30 and closes[-30] else 0

    prompt = (
        f'Voce e um especialista em trading no mercado do EVE Online.\n'
        f'Analise o item "{type_name}" e faca uma recomendacao objetiva.\n\n'
        f'DADOS DE MERCADO:\n'
        f'- Preco de venda (sell min): {sell_min:,.2f} ISK\n'
        f'- Melhor compra (buy max):   {buy_max:,.2f} ISK\n'
        f'- Margem liquida:            {net_margin:.2f}%\n'
        f'- Preco atual (close):       {last_close:,.2f} ISK\n'
        f'- Variacao 7 dias:           {chg7:+.2f}%\n'
        f'- Variacao 30 dias:          {chg30:+.2f}%\n\n'
        f'INDICADORES TECNICOS:\n'
        f'- RSI(14):        {last_rsi:.1f}\n'
        f'- MACD:           {last_macd:.4f}\n'
        f'- MACD Signal:    {last_sig:.4f}\n'
        f'- MACD Histograma:{(last_macd - last_sig):.4f}\n\n'
        f'NIVEIS DE FIBONACCI (90 dias):\n'
        + "\n".join(f'  {k}%: {v:,.2f} ISK' for k, v in fib.items()) +
        f'\n\nResponda SOMENTE em portugues com este formato exato:\n'
        f'RECOMENDACAO: [COMPRAR / VENDER / ESPERAR]\n'
        f'PRECO OBJETIVO: [valor em ISK ou "Nao aplica"]\n'
        f'STOP-LOSS: [valor em ISK ou "Nao aplica"]\n'
        f'RISCO: [BAIXO / MEDIO / ALTO], breve justificativa em 1 frase\n'
        f'ANALISE: [2-4 frases explicando a recomendacao com base nos indicadores]'
    )

    try:
        url     = "https://api.anthropic.com/v1/messages"
        headers = {
            "x-api-key":         api_key.strip(),
            "anthropic-version": "2023-06-01",
            "content-type":      "application/json",
        }
        # O6 — streaming quando stream_cb fornecido
        use_stream = stream_cb is not None
        payload = {
            "model":      "claude-sonnet-4-6",
            "max_tokens": 600,
            "stream":     use_stream,
            "messages":   [{"role": "user", "content": prompt}],
        }
        r = requests.post(url, headers=headers, json=payload,
                          timeout=60, stream=use_stream)

        if use_stream:
            if r.status_code != 200:
                return f"Erro API ({r.status_code}): {r.text[:200]}"
            full = []
            for raw_line in r.iter_lines():
                if not raw_line:
                    continue
                line = raw_line.decode("utf-8") if isinstance(raw_line, bytes) else raw_line
                if not line.startswith("data:"):
                    continue
                chunk = line[5:].strip()
                if chunk in ("", "[DONE]"):
                    continue
                try:
                    evt = json.loads(chunk)
                except Exception:
                    continue
                if evt.get("type") == "content_block_delta":
                    token = evt.get("delta", {}).get("text", "")
                    if token:
                        full.append(token)
                        stream_cb(token)
            return "".join(full)
        else:
            if r.status_code == 200:
                data = r.json()
                return data["content"][0]["text"].strip()
            return f"Erro API ({r.status_code}): {r.text[:200]}"
    except Exception as exc:
        return f"Erro ao chamar API: {exc}"


# ══════════════════════════════════════════════════════════════════════════════
#  BLACKLIST — persistencia (Tier 1)
# ══════════════════════════════════════════════════════════════════════════════

def load_blacklist() -> set:
    """Retorna conjunto de type_ids na blacklist."""
    if not os.path.exists(BLACKLIST_FILE):
        return set()
    try:
        with open(BLACKLIST_FILE, encoding="utf-8") as f:
            data = json.load(f)
        return set(int(x) for x in data.get("ids", []))
    except Exception:
        return set()

def save_blacklist(bl: set):
    try:
        with open(BLACKLIST_FILE, "w", encoding="utf-8") as f:
            json.dump({"ids": sorted(bl)}, f, indent=2)
    except Exception as exc:
        log.warning("Erro ao salvar blacklist: %s", exc)

def blacklist_add(bl: set, type_id: int) -> set:
    bl.add(type_id)
    save_blacklist(bl)
    return bl

def blacklist_remove(bl: set, type_id: int) -> set:
    bl.discard(type_id)
    save_blacklist(bl)
    return bl


# ══════════════════════════════════════════════════════════════════════════════
#  COMPARADOR DE HUBS — mesmo item em multiplos mercados (Tier 1)
# ══════════════════════════════════════════════════════════════════════════════

# Hubs principais exibidos por padrao
HUB_REGIONS = {
    "Jita":    10000002,
    "Amarr":   10000043,
    "Dodixie": 10000032,
    "Hek":     10000042,
    "Rens":    10000030,
}

def compare_hubs(type_id: int, broker_rel: int = 0, accounting: int = 0,
                 progress_cb=None) -> list:
    """
    Busca preco atual do type_id nos 5 hubs principais em paralelo.
    Retorna lista de dicts ordenada por net_pct desc.
    """
    rb, rt = _calc_fee_rates(broker_rel, accounting)

    def _fetch_hub(args):
        hub_name, rid = args
        try:
            url    = f"{ESI_BASE}/markets/{rid}/orders/"
            params = {"datasource": "tranquility",
                      "order_type": "all", "type_id": type_id}
            r      = _get_session().get(url, params=params, timeout=10)
            if r.status_code != 200:
                return None
            orders = r.json()
        except Exception:
            return None

        sells = sorted([float(o["price"]) for o in orders if not o["is_buy_order"]])
        buys  = sorted([float(o["price"]) for o in orders if  o["is_buy_order"]],
                       reverse=True)
        if not sells or not buys:
            return {"hub": hub_name, "sell_min": 0, "buy_max": 0,
                    "spread_pct": 0, "net_pct": 0,
                    "sell_orders": 0, "buy_orders": 0}

        sell_min  = sells[0]
        buy_max   = buys[0]
        spread    = (sell_min - buy_max) / sell_min * 100 if sell_min > 0 else 0
        cost      = buy_max + buy_max*rb + sell_min*rb + sell_min*rt
        net_pct   = (sell_min - cost) / cost * 100 if cost > 0 else 0
        return {
            "hub":         hub_name,
            "sell_min":    round(sell_min, 2),
            "buy_max":     round(buy_max,  2),
            "spread_pct":  round(spread,   2),
            "net_pct":     round(net_pct,  2),
            "sell_orders": len(sells),
            "buy_orders":  len(buys),
        }

    if progress_cb:
        progress_cb(f"Consultando {len(HUB_REGIONS)} hubs...")

    results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as pool:
        for res in pool.map(_fetch_hub, HUB_REGIONS.items()):
            if res:
                results.append(res)

    results.sort(key=lambda x: x["net_pct"], reverse=True)
    return results


# ══════════════════════════════════════════════════════════════════════════════
#  ARBITRAGEM MULTI-HOP A→B→C (Tier 2)
# ══════════════════════════════════════════════════════════════════════════════

def run_multihop(hop_names: list, orders_map: dict, item_cache: dict,
                 broker_rel: int = 0, accounting: int = 0,
                 min_net_pct: float = 3.0, top_n: int = 200,
                 progress_cb=None) -> list:
    """
    Calcula oportunidades de arbitragem encadeada A→B→C.

    hop_names : lista de 2 ou 3 nomes de regiao na ordem da jornada,
                e.g. ["The Forge (Jita)", "Domain (Amarr)", "Sinq Laison (Dodixie)"]
    orders_map: {region_name: [orders]} — ordens ja carregadas
    Retorna lista de dicts ordenada por net_pct_total desc.
    """
    if len(hop_names) < 2:
        return []

    rb, rt  = _calc_fee_rates(broker_rel, accounting)
    indexes = {r: build_order_index(orders_map[r])
               for r in hop_names if r in orders_map}
    if len(indexes) < len(hop_names):
        if progress_cb:
            progress_cb("Faltam ordens para alguns hubs. Baixe os CSVs primeiro.")
        return []

    # Universo de type_ids comuns a TODOS os hubs
    common = set.intersection(*[set(idx.keys()) for idx in indexes.values()])
    if progress_cb:
        progress_cb(f"{len(common)} itens comuns nos {len(hop_names)} hubs...")

    results = []
    for tid in common:
        prices = []   # [(buy_at_region, sell_at_next_region)]
        valid  = True
        for i, rname in enumerate(hop_names[:-1]):
            next_rname = hop_names[i + 1]
            sell_list  = indexes[rname][tid]["sell"]
            buy_list   = indexes[next_rname][tid]["buy"]
            if not sell_list or not buy_list:
                valid = False; break
            buy_p  = float(sell_list[0]["price"])
            sell_p = float(buy_list[0]["price"])
            prices.append((buy_p, sell_p))
        if not valid:
            continue

        # Simula a cadeia: investe buy_p[0], cada etapa gera receita que é
        # reinvestida na etapa seguinte
        capital = prices[0][0]   # custo inicial: compra no hub 0
        running = capital
        hop_details = []
        for i, (buy_p, sell_p) in enumerate(prices):
            cost_unit    = buy_p * (1 + rb)
            revenue_unit = sell_p * (1 - rb - rt)
            leg_net_pct  = (revenue_unit - cost_unit) / cost_unit * 100 if cost_unit > 0 else 0
            hop_details.append({
                "from":    hop_names[i],
                "to":      hop_names[i + 1],
                "buy_p":   round(buy_p,   2),
                "sell_p":  round(sell_p,  2),
                "net_pct": round(leg_net_pct, 2),
            })
            running = running / cost_unit * revenue_unit

        net_pct_total = (running - capital) / capital * 100 if capital > 0 else 0
        if net_pct_total < min_net_pct:
            continue

        info = item_cache.get(tid, {})
        results.append({
            "type_id":       tid,
            "item":          info.get("name", str(tid)),
            "capital_entry": round(capital, 2),
            "net_pct_total": round(net_pct_total, 2),
            "profit_unit":   round(running - capital, 2),
            "hops":          hop_details,
            "route":         " → ".join(h[:12] for h in hop_names),
            "n_hops":        len(hop_names) - 1,
        })

    results.sort(key=lambda x: x["net_pct_total"], reverse=True)
    results = results[:top_n]
    if progress_cb:
        progress_cb(f"{len(results)} oportunidades multi-hop encontradas.")
    return results


# ══════════════════════════════════════════════════════════════════════════════
#  CORRELACAO DE PRECOS ENTRE ITENS (Tier 2)
# ══════════════════════════════════════════════════════════════════════════════

def compute_price_correlation(hist_map: dict, min_days: int = 30) -> dict:
    """
    hist_map: {type_id: [ESI history dicts]}
    Retorna matriz {(tid_a, tid_b): pearson_r} para todos os pares.
    Usa apenas datas comuns a ambos os itens (inner join por data).
    """
    # Passo 1 — series de closes por item, indexadas por data
    series = {}
    for tid, hist in hist_map.items():
        if not hist:
            continue
        by_date = {d["date"]: float(d.get("average", 0)) for d in hist}
        series[tid] = by_date

    tids   = list(series.keys())
    matrix = {}

    for i in range(len(tids)):
        for j in range(i + 1, len(tids)):
            a, b  = tids[i], tids[j]
            dates = sorted(set(series[a]) & set(series[b]))
            if len(dates) < min_days:
                continue
            xs = [series[a][d] for d in dates]
            ys = [series[b][d] for d in dates]
            n  = len(xs)
            if n < 2:
                continue
            mx = sum(xs) / n
            my = sum(ys) / n
            num = sum((xi - mx) * (yi - my) for xi, yi in zip(xs, ys))
            sx  = (sum((xi - mx) ** 2 for xi in xs) / n) ** 0.5
            sy  = (sum((yi - my) ** 2 for yi in ys) / n) ** 0.5
            if sx == 0 or sy == 0:
                continue
            r   = num / (n * sx * sy)
            matrix[(a, b)] = round(r, 4)

    return matrix


# ══════════════════════════════════════════════════════════════════════════════
#  ALERTAS SONOROS / DESKTOP (Tier 2)
# ══════════════════════════════════════════════════════════════════════════════

def _desktop_notify(title: str, body: str):
    """
    Envia notificacao de desktop de forma nao-bloqueante.
    Tenta: plyer (cross-platform) → notify-send (Linux) → passa silenciosamente.
    """
    try:
        from plyer import notification
        notification.notify(title=title, message=body,
                            app_name="EVE Trade Finder", timeout=8)
        return
    except Exception as _e:

        log.debug("silenced: %s", _e)
    try:
        import subprocess
        subprocess.Popen(["notify-send", "-t", "8000", title, body],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as _e:

        log.debug("silenced: %s", _e)


# ══════════════════════════════════════════════════════════════════════════════
#  DETECCAO DE CICLOS DE MERCADO (Tier 3)
# ══════════════════════════════════════════════════════════════════════════════
#
#  Grupos de itens usados como "barometros" do ciclo de mercado EVE.
#  Cada grupo tem comportamento distinto em eventos e expansoes:
#   - Minerais:  sensíveis a mudancas de mineracao e guerras null-sec
#   - PLEX/Skill: sensíveis a campanhas CCP e promos
#   - Naves T1:   sensíveis a novos players e eventos PvP
#   - Módulos T2: sensíveis a guerras e industria avancada

CYCLE_ITEM_GROUPS = {
    # ── Minerais ──────────────────────────────────────────────────────────────
    "Minerais Basicos": {
        34: "Tritanium", 35: "Pyerite", 36: "Mexallon",
        37: "Isogen",    38: "Nocxium", 39: "Zydrine", 40: "Megacyte",
    },
    "Minerais Raros": {
        11399: "Morphite", 16634: "Arkonor",     16635: "Bistot",
        16636: "Crokite",  16637: "Dark Ochre",  16638: "Gneiss",
        16639: "Hedbergite",
    },
    "Minerais Moon": {
        45490: "Chromite",   45491: "Otavite",   45492: "Sperrylite",
        45493: "Vanadinite", 45494: "Cobaltite", 45495: "Euxenite",
        45496: "Scheelite",  45497: "Titanite",
    },
    # ── Combustivel e PI ──────────────────────────────────────────────────────
    "Combustivel Isotopes": {
        4051:  "Helium Isotopes",   17887: "Nitrogen Isotopes",
        17889: "Hydrogen Isotopes", 17888: "Oxygen Isotopes",
    },
    "Combustivel POS": {
        16273: "Strontium Clathrates", 4247: "Enriched Uranium",
        4312:  "Mechanical Parts",     4051: "Helium Isotopes",
    },
    "PI Tier 1 (P1)": {
        2267: "Bacteria",        2288: "Biofuels",
        9828: "Biomass",         2305: "Chiral Structures",
        2287: "Electrolytes",    2309: "Industrial Fibers",
        2393: "Oxidizing Compound", 2311: "Plasmoids",
        2312: "Precious Metals", 2306: "Proteins",
        2307: "Reactive Metals", 2392: "Silicon",
        9832: "Toxic Metals",    2313: "Water",
    },
    "PI Tier 2 (P2)": {
        2317: "Biocells",           2321: "Construction Blocks",
        9830: "Consumer Electronics",2324: "Coolant",
        2327: "Fertilizer",          2332: "Mechanical Parts",
        2346: "Microfiber Shielding",2348: "Miniature Electronics",
        2351: "Nanites",             2352: "Oxides",
        2354: "Polyaramids",         2358: "Polytextiles",
        9836: "Rocket Fuel",         2361: "Silicate Glass",
        2366: "Superconductors",     2367: "Supertensile Plastics",
        2370: "Synthetic Oil",       2374: "Test Cultures",
        2376: "Transmitter",         2378: "Viral Agent",
        2380: "Water-Cooled CPU",
    },
    "PI Tier 3 (P3)": {
        2382: "Biotech Research Reports",2384: "Camera Drones",
        2386: "Condensates",             2388: "Cryoprotectant Solution",
        2390: "Data Chips",              2395: "Guidance Systems",
        2397: "Hazmat Detection Systems",2399: "Hermetic Membranes",
        2401: "High-Tech Transmitters",  2403: "Industrial Explosives",
        2405: "Neocoms",                 2407: "Nuclear Reactors",
        2409: "Planet Vehicles",         2411: "Robotics",
        2413: "Smartfab Units",          2415: "Supercomputers",
        2417: "Synthetic Synapses",      2419: "Transcranial Microcontrollers",
        2421: "Ukomi Superconductors",   2423: "Vaccines",
    },
    # ── Commodities e Economia ────────────────────────────────────────────────
    "Commodities Premium": {
        29668: "PLEX",           40520: "Skill Injector",
        40519: "Skill Extractor",44992: "Daily Alpha Injector",
    },
    "Capsuleer Economy": {
        29668: "PLEX",  34: "Tritanium",
        35: "Pyerite",  40520: "Skill Injector",
    },
    # ── Naves ─────────────────────────────────────────────────────────────────
    "Fragatas T1": {
        587: "Rifter",   588: "Merlin",   582: "Kestrel",
        583: "Condor",   638: "Punisher", 648: "Executioner",
        596: "Incursus", 608: "Tristan",
    },
    "Fragatas T2 (Assault)": {
        11176: "Jaguar",    11178: "Wolf",
        11196: "Hawk",      11200: "Harpy",
        11172: "Crusader",  11174: "Malediction",
        11182: "Taranis",   11184: "Ares",
    },
    "Fragatas Especiais": {
        17703: "Dramiel",  17619: "Daredevil",
        33395: "Astero",   35779: "Pacifier",
        42135: "Sunesis",
    },
    "Destroyers T1": {
        16236: "Thrasher",  16238: "Catalyst",
        16240: "Cormorant", 16242: "Coercer",
        16243: "Algos",     16244: "Dragoon",
    },
    "Destroyers T2 (Interdictor)": {
        22444: "Sabre",    22440: "Flycatcher",
        22452: "Heretic",  22448: "Eris",
    },
    "Cruzadores T1": {
        621: "Rupture",   623: "Stabber",
        624: "Bellicose", 620: "Maller",
        633: "Arbitrator",634: "Omen",
    },
    "Cruzadores T2": {
        12003: "Vagabond", 12005: "Muninn",
        11993: "Zealot",   12021: "Sacrilege",
        11959: "Ishtar",   11987: "Ishkur",
    },
    "Cruzadores Estrategicos (T3)": {
        29986: "Tengu",   29988: "Loki",
        29984: "Proteus", 29990: "Legion",
    },
    "Battlecruisers T1": {
        16227: "Hurricane", 16229: "Cyclone",
        16231: "Drake",     16233: "Ferox",
        16225: "Brutix",    16235: "Prophecy",
    },
    "Battlecruisers T2 (Command)": {
        22442: "Sleipnir",  22446: "Claymore",
        22444: "Absolution",22448: "Astarte",
    },
    "Battleships T1": {
        639: "Tempest",  17919: "Typhoon",
        641: "Megathron",642: "Dominix",
        24692: "Apocalypse", 643: "Raven",
    },
    "Battleships T2 (Marauder)": {
        28661: "Vargur",  28659: "Paladin",
        28665: "Golem",   28663: "Kronos",
    },
    "Capitals (Carrier)": {
        34495: "Archon",    34517: "Chimera",
        34519: "Thanatos",  34521: "Nidhoggur",
        42241: "Hel",       42242: "Wyvern",
    },
    "Capitals (Dreadnought)": {
        34523: "Revelation", 34525: "Naglfar",
        34527: "Phoenix",    34529: "Moros",
    },
    "Supercapitais": {
        3514:  "Avatar",    11567: "Erebus",
        671:   "Leviathan", 45649: "Ragnarok",
    },
    # ── Modulos ───────────────────────────────────────────────────────────────
    "Modulos Escudo T2": {
        2488: "Invulnerability Field II",
        2456: "Large Shield Extender II",
        2281: "Adaptive Invulnerability Field II",
        2605: "Shield Boost Amplifier II",
    },
    "Modulos Armadura T2": {
        11370: "Armor Thermic Hardener II",
        11374: "Armor Kinetic Hardener II",
        11378: "Armor Explosive Hardener II",
        11380: "Armor EM Hardener II",
    },
    "Modulos Propulsao T2": {
        12058: "50MN Microwarpdrive II",
        12062: "100MN Microwarpdrive II",
        5973:  "10MN Afterburner II",
        5975:  "1MN Afterburner II",
    },
    "Armamento Hibrido T2": {
        2929: "250mm Railgun II",       2931: "425mm Railgun II",
        2933: "Neutron Blaster Cannon II",3001: "Ion Blaster Cannon II",
    },
    "Armamento Projectil T2": {
        2873: "220mm Vulcan AutoCannon II",  2875: "425mm AutoCannon II",
        3057: "1400mm Howitzer Artillery II",2877: "800mm Repeating Cannon II",
    },
    "Armamento Laser T2": {
        20448: "Mega Pulse Laser II",    20450: "Dual Heavy Pulse Laser II",
        20436: "Mega Beam Laser II",     20440: "Heavy Beam Laser II",
    },
    "Misseis T2": {
        209:  "Mjolnir Heavy Missile",  210: "Nova Heavy Missile",
        212:  "Inferno Heavy Missile",  211: "Scourge Heavy Missile",
        11489:"Caldari Navy Scourge Heavy Missile",
    },
    "Drones Combate T2": {
        2185: "Hobgoblin II",  2183: "Warrior II",
        2456: "Hammerhead II", 2454: "Ogre II",
        3244: "Valkyrie II",   2488: "Infiltrator II",
    },
    "Drones Miner": {
        22175: "Mining Drone II",  28209: "Harvester Mining Drone",
        23051: "Augmented Mining Drone",
    },
    "Rigs T1": {
        25589: "Small Trimark Armor Pump I",
        25591: "Medium Trimark Armor Pump I",
        25971: "Large Trimark Armor Pump I",
        26082: "Medium Core Defense Field Extender I",
    },
    "Rigs T2": {
        26450: "Small Trimark Armor Pump II",
        26452: "Medium Trimark Armor Pump II",
        26086: "Medium Core Defense Field Extender II",
    },
    # ── Implantes e Boosters ──────────────────────────────────────────────────
    "Implantes Padrão (Slot 1-5)": {
        10208: "Ocular Filter - Standard",
        10218: "Memory Augmentation - Standard",
        9942:  "Neural Boost - Standard",
        9943:  "Cybernetic Subprocessor - Standard",
        9947:  "Social Adaptation Chip - Standard",
    },
    "Implantes Enhanced/Improved": {
        10209: "Ocular Filter - Improved",
        10219: "Memory Augmentation - Improved",
        9944:  "Neural Boost - Improved",
        9945:  "Cybernetic Subprocessor - Improved",
    },
    "Implantes Strong/Elite": {
        10210: "Ocular Filter - Strong",
        10220: "Memory Augmentation - Strong",
        9946:  "Neural Boost - Strong",
        10248: "Cybernetic Subprocessor - Strong",
    },
    "Implantes Hardwiring": {
        10228: "Zainou Deadeye TA-706",
        10230: "Zainou Gypsy EE-603",
        13244: "Inherent Implants Squire EG-601",
        10212: "Inherent Implants Noble ZET-300",
    },
    "Boosters": {
        15464: "Blue Pill",    15470: "Exile",
        15468: "Drop",         15466: "Frentix",
        15472: "Mindflood",    15474: "X-Instinct",
        28672: "Agency Hardshell TB3",
    },
    # ── Industria e Blueprints ────────────────────────────────────────────────
    "Materiais Industria T2": {
        11399: "Morphite",
        3828:  "R.A.M.- Starship Tech",
        11478: "Fernite Carbide",
        11480: "Titanium Carbide",
        11482: "Crystalline Carbonide",
    },
    "Componentes Capital": {
        11530: "Capital Armor Plates",
        11532: "Capital Capacitor Battery",
        11534: "Capital Computer System",
        11536: "Capital Construction Parts",
        11538: "Capital Corporate Hangar Bay",
    },
    "Datacores": {
        20410: "Datacore - Caldari Starship Engineering",
        20411: "Datacore - Gallente Starship Engineering",
        20412: "Datacore - Minmatar Starship Engineering",
        20413: "Datacore - Amarr Starship Engineering",
        20414: "Datacore - Mechanical Engineering",
        20415: "Datacore - Electronic Engineering",
    },
    # ── Estruturas e Deployables ──────────────────────────────────────────────
    "Estruturas Upwell": {
        35825: "Astrahus",  35826: "Fortizar",
        35827: "Keepstar", 35835: "Raitaru",
        35836: "Azbel",    35832: "Sotiyo",
    },
    "Deployables": {
        33475: "Mobile Tractor Unit",
        33476: "Mobile Micro Jump Unit",
        33474: "Mobile Depot",
        22529: "Mobile Warp Disruptor",
        26872: "Mobile Cynosural Inhibitor",
    },
    # ── Salvage ───────────────────────────────────────────────────────────────
    "Salvage": {
        25590: "Alloyed Tritanium Bar",
        25595: "Burned Logic Circuit",
        25596: "Charred Micro Circuit",
        25600: "Damaged Artificial Neural Network",
        25601: "Defective Current Pump",
        25602: "Fried Interface Circuit",
    },
    # ── Asteroids (Minérios) ──────────────────────────────────────────────────
    # Categoria Asteroid (category_id=25) — minérios e versões compressed
    # Highsec ores
    "Asteroids Highsec": {
        1230:  "Veldspar",
        17470: "Concentrated Veldspar",
        17471: "Dense Veldspar",
        28432: "Compressed Veldspar",
        28430: "Compressed Concentrated Veldspar",
        28431: "Compressed Dense Veldspar",
        1228:  "Scordite",
        17463: "Condensed Scordite",
        17464: "Massive Scordite",
        28429: "Compressed Scordite",
        28427: "Compressed Condensed Scordite",
        28428: "Compressed Massive Scordite",
        1224:  "Pyroxeres",
        17459: "Solid Pyroxeres",
        17460: "Viscous Pyroxeres",
        28424: "Compressed Pyroxeres",
        28425: "Compressed Solid Pyroxeres",
        28426: "Compressed Viscous Pyroxeres",
        18:    "Plagioclase",
        17455: "Azure Plagioclase",
        17456: "Rich Plagioclase",
        28421: "Compressed Plagioclase",
        28422: "Compressed Azure Plagioclase",
        28423: "Compressed Rich Plagioclase",
        1227:  "Omber",
        17867: "Silvery Omber",
        17868: "Golden Omber",
        28416: "Compressed Omber",
        28414: "Compressed Silvery Omber",
        28415: "Compressed Golden Omber",
        20:    "Kernite",
        17452: "Luminous Kernite",
        17453: "Fiery Kernite",
        28418: "Compressed Kernite",
        28419: "Compressed Luminous Kernite",
        28420: "Compressed Fiery Kernite",
    },
    # Lowsec ores
    "Asteroids Lowsec": {
        1229:  "Jaspet",
        17448: "Pure Jaspet",
        17449: "Pristine Jaspet",
        28404: "Compressed Jaspet",
        28405: "Compressed Pure Jaspet",
        28406: "Compressed Pristine Jaspet",
        1232:  "Hemorphite",
        17444: "Vivid Hemorphite",
        17445: "Radiant Hemorphite",
        28410: "Compressed Hemorphite",
        28411: "Compressed Vivid Hemorphite",
        28412: "Compressed Radiant Hemorphite",
        21:    "Hedbergite",
        17440: "Vitric Hedbergite",
        17441: "Glazed Hedbergite",
        28407: "Compressed Hedbergite",
        28408: "Compressed Vitric Hedbergite",
        28409: "Compressed Glazed Hedbergite",
    },
    # Nullsec ores
    "Asteroids Nullsec": {
        17865: "Iridescent Gneiss",
        17866: "Prismatic Gneiss",
        1229:  "Gneiss",
        28401: "Compressed Gneiss",
        28402: "Compressed Iridescent Gneiss",
        28403: "Compressed Prismatic Gneiss",
        1236:  "Dark Ochre",
        17436: "Onyx Ochre",
        17437: "Obsidian Ochre",
        28398: "Compressed Dark Ochre",
        28399: "Compressed Onyx Ochre",
        28400: "Compressed Obsidian Ochre",
        19:    "Spodumain",
        17466: "Bright Spodumain",
        17467: "Gleaming Spodumain",
        28413: "Compressed Spodumain",
        28414: "Compressed Bright Spodumain",
        28415: "Compressed Gleaming Spodumain",
        22:    "Arkonor",
        17425: "Crimson Arkonor",
        17426: "Prime Arkonor",
        28367: "Compressed Arkonor",
        28365: "Compressed Crimson Arkonor",
        28366: "Compressed Prime Arkonor",
        1231:  "Bistot",
        17428: "Triclinic Bistot",
        17429: "Monoclinic Bistot",
        28395: "Compressed Bistot",
        28396: "Compressed Triclinic Bistot",
        28397: "Compressed Monoclinic Bistot",
        1225:  "Crokite",
        17432: "Sharp Crokite",
        17433: "Crystalline Crokite",
        28392: "Compressed Crokite",
        28393: "Compressed Sharp Crokite",
        28394: "Compressed Crystalline Crokite",
        11396: "Mercoxit",
        17869: "Magma Mercoxit",
        17870: "Vitreous Mercoxit",
        28389: "Compressed Mercoxit",
        28390: "Compressed Magma Mercoxit",
        28391: "Compressed Vitreous Mercoxit",
    },
    # Moon ores (também categoria Asteroid)
    "Asteroids Moon Ore R4": {
        45491: "Zeolites",         46676: "Compressed Zeolites",
        45490: "Sylvite",          46675: "Compressed Sylvite",
        45492: "Bitumens",         46677: "Compressed Bitumens",
        45493: "Coesite",          46678: "Compressed Coesite",
    },
    "Asteroids Moon Ore R8": {
        45494: "Cobaltite",        46679: "Compressed Cobaltite",
        45495: "Euxenite",         46680: "Compressed Euxenite",
        45496: "Scheelite",        46681: "Compressed Scheelite",
        45497: "Titanite",         46682: "Compressed Titanite",
    },
    "Asteroids Moon Ore R16": {
        45498: "Chromite",         46683: "Compressed Chromite",
        45499: "Otavite",          46684: "Compressed Otavite",
        45500: "Sperrylite",       46685: "Compressed Sperrylite",
        45501: "Vanadinite",       46686: "Compressed Vanadinite",
    },
    "Asteroids Moon Ore R32": {
        45502: "Carnotite",        46687: "Compressed Carnotite",
        45503: "Zircon",           46688: "Compressed Zircon",
        45504: "Loparite",         46689: "Compressed Loparite",
        45506: "Monazite",         46691: "Compressed Monazite",
    },
    "Asteroids Moon Ore R64": {
        45506: "Xenotime",         46691: "Compressed Xenotime",
        45507: "Ytterbite",        46692: "Compressed Ytterbite",
        45508: "Pollucite",        46693: "Compressed Pollucite",
        45509: "Cinnabar",         46694: "Compressed Cinnabar",
    },
    # Ice (também negociado no mercado de asteroids)
    "Asteroids Ice": {
        16262: "Blue Ice",         28443: "Compressed Blue Ice",
        16263: "Clear Icicle",     28444: "Compressed Clear Icicle",
        16264: "Glacial Mass",     28445: "Compressed Glacial Mass",
        16265: "White Glaze",      28446: "Compressed White Glaze",
        16266: "Glare Crust",      28447: "Compressed Glare Crust",
        16267: "Dark Glitter",     28448: "Compressed Dark Glitter",
        16268: "Gelidus",          28449: "Compressed Gelidus",
        16269: "Krystallos",       28450: "Compressed Krystallos",
    },
}

# Fases do ciclo com cor associada
CYCLE_PHASES = {
    "Acumulacao": {"color": "#00aaff",  "emoji": "◈",
                   "desc": "Preco baixo com volume crescendo. "
                           "Possivel oportunidade de compra antecipada."},
    "Alta":       {"color": "#00ff88",  "emoji": "▲",
                   "desc": "Preco e volume em alta. "
                           "Tendencia de alta ativa — momento de vender na forca."},
    "Distribuicao":{"color": "#ffd700", "emoji": "◉",
                   "desc": "Preco alto com volume caindo. "
                           "Possivel topo — considere realizar lucros."},
    "Queda":      {"color": "#ff4455",  "emoji": "▼",
                   "desc": "Preco e volume em queda. "
                           "Evite compras novas — aguarde estabilizacao."},
    "Lateral":    {"color": "#5a7a9a",  "emoji": "◇",
                   "desc": "Sem tendencia definida. "
                           "Mercado em equilibrio ou dados insuficientes."},
}


def _cycle_ema(data: list, w: int) -> list:
    """EMA simples O(n) sobre lista de floats. None onde nao ha dados."""
    out = [None] * len(data)
    k   = 2.0 / (w + 1)
    for i, v in enumerate(data):
        if v is None:
            out[i] = out[i-1] if i > 0 else None
            continue
        if out[i-1] is None:
            out[i] = v
        else:
            out[i] = v * k + out[i-1] * (1-k)
    return out


def detect_cycle_item(hist: list, days_back: int = 180) -> dict:
    """
    Analisa o historico de um item e retorna sinais de ciclo.

    Retorna dict com:
      phase         — fase detectada (str)
      phase_color   — cor hex da fase
      phase_emoji   — emoji da fase
      trend_score   — -1..+1 (negativo=queda, positivo=alta)
      vol_score     — 0..N (ratio vol_atual/vol_medio; >2 = spike)
      momentum_7d   — variacao % nos ultimos 7 dias
      momentum_30d  — variacao % nos ultimos 30 dias
      volatility    — coeficiente de variacao (std/mean) dos ultimos 90d
      rsi           — RSI(14) atual ou None
      anomaly       — bool, True se spike de volume ou preco fora do padrao
      anomaly_msg   — descricao da anomalia (se houver)
      dates         — lista de datas dos ultimos days_back dias
      prices        — precos medios diarios
      volumes       — volumes diarios
      ema_short     — EMA 14 dias
      ema_long      — EMA 50 dias
      n_days        — numero de dias de historico disponivel
    """
    if not hist or len(hist) < 14:
        return {"phase": "Lateral", "phase_color": CYCLE_PHASES["Lateral"]["color"],
                "phase_emoji": CYCLE_PHASES["Lateral"]["emoji"],
                "trend_score": 0, "vol_score": 1, "momentum_7d": 0,
                "momentum_30d": 0, "volatility": 0, "rsi": None,
                "anomaly": False, "anomaly_msg": "",
                "dates": [], "prices": [], "volumes": [],
                "ema_short": [], "ema_long": [], "n_days": len(hist)}

    hist_s  = sorted(hist, key=lambda d: d.get("date", ""))
    cutoff  = (datetime.now(timezone.utc) - timedelta(days=days_back)).strftime("%Y-%m-%d")
    hist_f  = [d for d in hist_s if d.get("date","") >= cutoff]
    if not hist_f:
        hist_f = hist_s[-days_back:] if len(hist_s) >= days_back else hist_s

    dates   = [d["date"]              for d in hist_f]
    prices  = [float(d.get("average", 0)) for d in hist_f]
    volumes = [int(d.get("volume", 0))    for d in hist_f]
    n       = len(prices)

    # ── EMAs ─────────────────────────────────────────────────────────────────
    ema14 = _cycle_ema(prices, 14)
    ema50 = _cycle_ema(prices, 50)

    # ── Trend score: EMA14 vs EMA50 (cruzamento) ─────────────────────────────
    last14 = next((v for v in reversed(ema14) if v is not None), None)
    last50 = next((v for v in reversed(ema50) if v is not None), None)
    if last14 and last50 and last50 > 0:
        trend_score = (last14 - last50) / last50
    else:
        trend_score = 0.0

    # ── Volume score: vol atual vs media 90d ─────────────────────────────────
    cutoff90 = (datetime.now(timezone.utc) - timedelta(days=90)).strftime("%Y-%m-%d")
    vols90   = [v for d, v in zip(dates, volumes) if d >= cutoff90]
    avg_vol  = sum(vols90) / len(vols90) if vols90 else 1
    cur_vol7 = sum(volumes[-7:]) / min(7, n)
    vol_score = cur_vol7 / avg_vol if avg_vol > 0 else 1.0

    # ── Momentum ──────────────────────────────────────────────────────────────
    def _chg(days):
        if n >= days and prices[-days] > 0:
            return (prices[-1] - prices[-days]) / prices[-days] * 100
        return 0.0
    mom7  = _chg(7)
    mom30 = _chg(min(30, n))

    # ── Volatilidade (CV = std/mean dos ultimos 90d) ──────────────────────────
    if len(vols90) >= 5 and avg_vol > 0:
        p90 = [float(h.get("average",0)) for h, d_ in zip(hist_f, dates)
               if d_ >= cutoff90]
        if p90:
            mu  = sum(p90) / len(p90)
            sd  = (sum((x-mu)**2 for x in p90) / len(p90)) ** 0.5
            volatility = sd / mu if mu > 0 else 0.0
        else:
            volatility = 0.0
    else:
        volatility = 0.0

    # ── RSI(14) simplificado ──────────────────────────────────────────────────
    rsi = None
    if n >= 15:
        gains  = [max(prices[i]-prices[i-1], 0) for i in range(1, n)]
        losses = [max(prices[i-1]-prices[i], 0) for i in range(1, n)]
        ag = sum(gains[:14])  / 14
        al = sum(losses[:14]) / 14
        for i in range(14, n-1):
            ag = (ag*13 + gains[i])  / 14
            al = (al*13 + losses[i]) / 14
        rs  = ag / al if al > 0 else 9999
        rsi = round(100 - 100 / (1 + rs), 1)

    # ── Anomalia ──────────────────────────────────────────────────────────────
    anomaly     = False
    anomaly_msg = ""
    if vol_score >= 3.0:
        anomaly     = True
        anomaly_msg = f"Spike de volume: {vol_score:.1f}x a media (possivel evento CCP)"
    elif abs(mom7) >= 15:
        anomaly     = True
        anomaly_msg = f"Variacao brusca de preco: {mom7:+.1f}% em 7 dias (possivel patch)"
    elif abs(mom7) >= 8 and vol_score >= 1.8:
        anomaly     = True
        anomaly_msg = f"Alta atividade: {mom7:+.1f}% / volume {vol_score:.1f}x"

    # ── Classificacao de fase ─────────────────────────────────────────────────
    # Regras simples e interpretáveis:
    #   trend_score > +0.03 e vol_score > 1.2 → Alta
    #   trend_score > +0.01 e vol_score < 0.8 → Distribuicao (subindo mas sem volume)
    #   trend_score < -0.03 e vol_score < 0.8 → Queda
    #   trend_score < -0.01 e vol_score > 1.2 → Acumulacao (caindo mas com volume)
    #   else                                   → Lateral
    if trend_score > 0.03 and vol_score >= 1.2:
        phase = "Alta"
    elif trend_score > 0.01 and vol_score < 0.8:
        phase = "Distribuicao"
    elif trend_score < -0.03 and vol_score < 0.8:
        phase = "Queda"
    elif trend_score < -0.01 and vol_score >= 1.2:
        phase = "Acumulacao"
    else:
        phase = "Lateral"

    return {
        "phase":        phase,
        "phase_color":  CYCLE_PHASES[phase]["color"],
        "phase_emoji":  CYCLE_PHASES[phase]["emoji"],
        "trend_score":  round(trend_score, 4),
        "vol_score":    round(vol_score,   2),
        "momentum_7d":  round(mom7,  2),
        "momentum_30d": round(mom30, 2),
        "volatility":   round(volatility, 4),
        "rsi":          rsi,
        "anomaly":      anomaly,
        "anomaly_msg":  anomaly_msg,
        "dates":        dates,
        "prices":       prices,
        "volumes":      volumes,
        "ema_short":    ema14,
        "ema_long":     ema50,
        "n_days":       n,
    }


def detect_cycle_region(region_id: int, region_name: str,
                         item_groups: dict = None,
                         progress_cb=None) -> dict:
    """
    Analisa multiplos grupos de itens numa regiao e retorna:
      - sinal por grupo (fase, anomalias)
      - sinal consolidado da regiao
      - lista de anomalias detectadas
      - dados por item para exibicao em tabela e graficos

    item_groups: {group_name: {type_id: name}}
                 Default: CYCLE_ITEM_GROUPS
    """
    if item_groups is None:
        item_groups = CYCLE_ITEM_GROUPS

    all_items   = {}   # {tid: {"name":..., "group":..., "signal":...}}
    group_sigs  = {}   # {group_name: [signals]}
    anomalies   = []   # lista de anomalias globais

    total_items = sum(len(v) for v in item_groups.values())
    done        = [0]
    _warm_hist_cache(region_id, progress_cb)  # pré-aquece cache RAM

    for group_name, items in item_groups.items():
        sigs = []
        for tid, iname in items.items():
            hist = _get_hist_cached(region_id, tid)
            sig  = detect_cycle_item(hist)
            sig["type_id"] = tid
            sig["name"]    = iname
            sig["group"]   = group_name
            all_items[tid] = sig
            sigs.append(sig)
            if sig["anomaly"]:
                anomalies.append({
                    "group": group_name,
                    "item":  iname,
                    "msg":   sig["anomaly_msg"],
                    "mom7":  sig["momentum_7d"],
                    "vol":   sig["vol_score"],
                })
            done[0] += 1
            if progress_cb and done[0] % 3 == 0:
                progress_cb((done[0], total_items,
                             f"Analisando {group_name}: {iname}..."))
        group_sigs[group_name] = sigs

    # ── Sinal consolidado da regiao ──────────────────────────────────────────
    # Votacao ponderada: cada item vota na sua fase
    phase_votes = {p: 0 for p in CYCLE_PHASES}
    total_trend  = 0.0
    total_vol    = 0.0
    n_items      = len(all_items)

    for sig in all_items.values():
        phase_votes[sig["phase"]] += 1
        total_trend  += sig["trend_score"]
        total_vol    += sig["vol_score"]

    region_phase  = max(phase_votes, key=phase_votes.get)
    avg_trend     = total_trend / n_items if n_items else 0
    avg_vol       = total_vol   / n_items if n_items else 1

    # Confianca: percentual do grupo dominante
    confidence    = phase_votes[region_phase] / n_items * 100 if n_items else 0

    # ── Deteccao de evento sistemico ─────────────────────────────────────────
    # Se >= 3 itens de grupos diferentes tiveram anomalia no mesmo periodo → evento
    n_anomaly_groups = len({a["group"] for a in anomalies})
    systemic_event   = n_anomaly_groups >= 2 and len(anomalies) >= 3
    systemic_msg     = ""
    if systemic_event:
        systemic_msg = (
            f"EVENTO SISTEMICO DETECTADO: {len(anomalies)} anomalias em "
            f"{n_anomaly_groups} grupos distintos. "
            "Provavel expansao, patch ou evento CCP ativo.")

    return {
        "region":          region_name,
        "region_phase":    region_phase,
        "region_color":    CYCLE_PHASES[region_phase]["color"],
        "region_emoji":    CYCLE_PHASES[region_phase]["emoji"],
        "confidence":      round(confidence, 1),
        "avg_trend":       round(avg_trend, 4),
        "avg_vol":         round(avg_vol, 2),
        "phase_votes":     phase_votes,
        "anomalies":       anomalies,
        "systemic_event":  systemic_event,
        "systemic_msg":    systemic_msg,
        "items":           all_items,
        "group_sigs":      group_sigs,
    }


# ══════════════════════════════════════════════════════════════════════════════
#  PREVISAO DE PRECOS — Ensemble de médias moveis (Tier 3)
# ══════════════════════════════════════════════════════════════════════════════
#
#  Modelos:
#   1. WMA  — média ponderada com pesos lineares (melhor janela por RMSE)
#   2. DES  — suavizacao exponencial dupla (Double Exponential Smoothing)
#             extrapola tendencia linear detectada
#   3. LR   — regressao linear sobre EMA (captura tendencias claras)
#   4. SAZ  — ajuste de sazonalidade semanal (efeito fim-de-semana EVE)
#
#  Ensemble: pesos = 1/RMSE (validacao walk-forward nos ultimos 30 dias)
#  Saida: previsao central + banda de confianca (±1σ da volatilidade historica)

def _wma(data: list, w: int) -> float | None:
    """Weighted Moving Average com pesos lineares (1,2,...,w)."""
    if len(data) < w:
        return None
    window = data[-w:]
    denom  = w * (w + 1) / 2
    return sum((i + 1) * v for i, v in enumerate(window)) / denom


def _des_forecast(data: list, alpha: float = 0.3, beta: float = 0.1,
                  horizon: int = 14) -> list:
    """
    Double Exponential Smoothing (Holt's linear trend method).
    Extrapola tendencia linear para 'horizon' passos a frente.
    Retorna lista de 'horizon' previsoes.
    """
    if len(data) < 3:
        return [data[-1] if data else 0.0] * horizon

    # Inicializacao: nivel = primeiro valor, tendencia = diferenca media inicial
    l = data[0]
    b = (data[min(3, len(data)-1)] - data[0]) / min(3, len(data)-1)

    for v in data[1:]:
        l_prev, b_prev = l, b
        l = alpha * v           + (1 - alpha) * (l_prev + b_prev)
        b = beta  * (l - l_prev) + (1 - beta)  * b_prev

    return [l + (h + 1) * b for h in range(horizon)]


def _linreg_forecast(data: list, w: int = 21, horizon: int = 14) -> list:
    """
    Regressao linear simples sobre os ultimos 'w' pontos.
    Extrapola a reta para 'horizon' passos a frente.
    """
    seg = data[-w:] if len(data) >= w else data
    n   = len(seg)
    if n < 3:
        return [seg[-1] if seg else 0.0] * horizon

    xs  = list(range(n))
    mx  = sum(xs) / n
    my  = sum(seg) / n
    num = sum((x - mx) * (y - my) for x, y in zip(xs, seg))
    den = sum((x - mx) ** 2 for x in xs)
    slope     = num / den if den > 0 else 0.0
    intercept = my - slope * mx

    base = n  # proximo ponto esta em posicao n
    return [intercept + slope * (base + h) for h in range(horizon)]


def _seasonal_factors(dates: list, values: list) -> dict:
    """
    Calcula fatores de sazonalidade por dia-da-semana (0=seg..6=dom).
    Retorna dict {0..6: fator_multiplicativo} onde 1.0 = media.
    """
    from datetime import datetime as _dt
    sums  = [0.0] * 7
    cnts  = [0]   * 7
    mean  = sum(values) / len(values) if values else 1.0

    for date_str, v in zip(dates, values):
        try:
            dow = _dt.strptime(date_str, "%Y-%m-%d").weekday()
            sums[dow] += v
            cnts[dow] += 1
        except Exception:
            continue

    factors = {}
    for dow in range(7):
        if cnts[dow] > 0:
            factors[dow] = (sums[dow] / cnts[dow]) / mean
        else:
            factors[dow] = 1.0
    return factors


def _rmse(actual: list, predicted: list) -> float:
    """RMSE entre dois arrays do mesmo tamanho."""
    n = min(len(actual), len(predicted))
    if n == 0:
        return 1e9
    return (sum((a - p) ** 2 for a, p in zip(actual[:n], predicted[:n])) / n) ** 0.5


def predict_prices(hist: list,
                   horizon: int = 14,
                   alpha: float = 0.3,
                   beta:  float = 0.1,
                   wma_windows: tuple = (7, 14, 21)) -> dict:
    """
    Gera previsao de precos para os proximos 'horizon' dias usando ensemble
    de WMA, DES e Regressao Linear, com ajuste de sazonalidade semanal e
    intervalos de confianca.

    Parametros
    ----------
    hist        : lista de dicts ESI (date, average, volume, ...)
    horizon     : numero de dias a prever (padrao 14)
    alpha, beta : suavizacao do DES
    wma_windows : janelas candidatas para WMA (melhor selecionada por RMSE)

    Retorna dict com:
      dates_hist    — datas historicas usadas
      prices_hist   — precos historicos (closes)
      dates_fcast   — datas das previsoes
      fcast_central — lista de previsoes centrais
      fcast_upper   — banda superior (+1σ volatilidade)
      fcast_lower   — banda inferior (-1σ volatilidade)
      model_weights — {nome: peso_no_ensemble}
      model_rmse    — {nome: RMSE_validacao}
      model_forecasts — {nome: [previsoes]}
      volatility    — desvio padrao relativo historico
      trend_pct     — tendencia linear nos ultimos 30d (% ao dia)
      confidence    — nivel de confianca estimado (0-100)
      n_days        — numero de dias de historico disponivel
    """
    if not hist or len(hist) < 10:
        return {}

    hist_s  = sorted(hist, key=lambda d: d.get("date", ""))
    dates   = [d["date"]              for d in hist_s]
    closes  = [float(d.get("average", 0)) for d in hist_s]
    n       = len(closes)

    # ── Datas futuras ─────────────────────────────────────────────────────────
    try:
        from datetime import datetime as _dt, timedelta as _td
        last_date = _dt.strptime(dates[-1], "%Y-%m-%d")
        dates_fcast = [(last_date + _td(days=i+1)).strftime("%Y-%m-%d")
                       for i in range(horizon)]
        fcast_dows  = [(last_date + _td(days=i+1)).weekday()
                       for i in range(horizon)]
    except Exception:
        dates_fcast = [f"+{i+1}d" for i in range(horizon)]
        fcast_dows  = [0] * horizon

    # ── Sazonalidade ──────────────────────────────────────────────────────────
    saz_factors = _seasonal_factors(dates, closes)

    # ── Validacao walk-forward (ultimos 30 dias como holdout) ─────────────────
    val_size  = min(30, n // 4)
    train_end = n - val_size
    train_c   = closes[:train_end]
    val_c     = closes[train_end:]

    if len(train_c) < 5 or len(val_c) < 3:
        # Dados insuficientes para validacao — usa pesos iguais
        train_c = closes
        val_c   = []

    # ── Gera previsoes de cada modelo ─────────────────────────────────────────
    def _gen(data, h):
        # WMA: escolhe melhor janela
        best_wma_w  = wma_windows[0]
        best_wma_rmse = 1e9
        for w in wma_windows:
            if len(data) < w + h:
                continue
            preds_wma = []
            for step in range(h):
                p = _wma(data[:len(data) - h + step] if step < h else data, w)
                preds_wma.append(p if p is not None else data[-1])
            r = _rmse(data[-h:], preds_wma) if len(data) >= h else 1e9
            if r < best_wma_rmse:
                best_wma_rmse = r
                best_wma_w    = w
        # Previsao WMA para horizonte futuro
        wma_fcast = []
        extended  = list(data)
        for _ in range(h):
            p = _wma(extended, best_wma_w) or extended[-1]
            wma_fcast.append(p)
            extended.append(p)
        return wma_fcast, best_wma_w

    wma_fcast, best_w = _gen(closes, horizon)
    des_fcast         = _des_forecast(closes, alpha, beta, horizon)
    lr_fcast          = _linreg_forecast(closes, 21, horizon)

    # ── Valida cada modelo ────────────────────────────────────────────────────
    if len(val_c) >= 3:
        wma_v, _ = _gen(train_c, len(val_c))
        des_v    = _des_forecast(train_c, alpha, beta, len(val_c))
        lr_v     = _linreg_forecast(train_c, 21, len(val_c))
        rmses    = {
            "WMA":  _rmse(val_c, wma_v),
            "DES":  _rmse(val_c, des_v),
            "LR":   _rmse(val_c, lr_v),
        }
    else:
        rmses = {"WMA": 1.0, "DES": 1.0, "LR": 1.0}

    # ── Pesos do ensemble: 1/RMSE (normalizado) ───────────────────────────────
    inv = {k: 1.0 / max(v, 1e-6) for k, v in rmses.items()}
    total_inv = sum(inv.values())
    weights   = {k: v / total_inv for k, v in inv.items()}

    # ── Previsao ensemble + ajuste sazonal ────────────────────────────────────
    fcast_central = []
    for i in range(horizon):
        base = (weights["WMA"] * wma_fcast[i] +
                weights["DES"] * des_fcast[i] +
                weights["LR"]  * lr_fcast[i])
        saz  = saz_factors.get(fcast_dows[i], 1.0) if fcast_dows else 1.0
        # Aplica sazonalidade suavizada (50% do efeito para nao exagerar)
        fcast_central.append(base * (1.0 + 0.5 * (saz - 1.0)))

    # ── Volatilidade historica (sigma das ultimas 90 variacoes diarias) ────────
    window_vol = closes[-90:] if n >= 90 else closes
    if len(window_vol) >= 5:
        diffs = [(window_vol[i] - window_vol[i-1]) / window_vol[i-1]
                 for i in range(1, len(window_vol)) if window_vol[i-1] > 0]
        vol   = (sum(d**2 for d in diffs) / len(diffs)) ** 0.5 if diffs else 0.01
    else:
        vol   = 0.02   # fallback: 2%

    fcast_upper = [v * (1 + vol * (i + 1) ** 0.5) for i, v in enumerate(fcast_central)]
    fcast_lower = [v * (1 - vol * (i + 1) ** 0.5) for i, v in enumerate(fcast_central)]

    # ── Tendencia linear (% ao dia, ultimos 30d) ──────────────────────────────
    seg30  = closes[-30:] if n >= 30 else closes
    n30    = len(seg30)
    trend_pct = 0.0
    if n30 >= 5:
        xs30  = list(range(n30))
        mx30  = sum(xs30) / n30
        my30  = sum(seg30) / n30
        num30 = sum((x - mx30) * (y - my30) for x, y in zip(xs30, seg30))
        den30 = sum((x - mx30) ** 2 for x in xs30)
        slope30    = num30 / den30 if den30 > 0 else 0.0
        trend_pct  = slope30 / my30 * 100 if my30 > 0 else 0.0  # % por dia

    # ── Confianca estimada: inversamente proporcional a volatilidade e erros ──
    avg_rmse   = sum(rmses.values()) / 3
    last_price = closes[-1] if closes else 1.0
    err_rel    = avg_rmse / last_price if last_price > 0 else 0.1
    confidence = max(0, min(100, round(100 * (1 - err_rel * 5 - vol * 10), 1)))

    return {
        "dates_hist":      dates,
        "prices_hist":     closes,
        "dates_fcast":     dates_fcast,
        "fcast_central":   [round(v, 2) for v in fcast_central],
        "fcast_upper":     [round(v, 2) for v in fcast_upper],
        "fcast_lower":     [round(v, 2) for v in fcast_lower],
        "model_weights":   {k: round(v, 4) for k, v in weights.items()},
        "model_rmse":      {k: round(v, 2) for k, v in rmses.items()},
        "model_forecasts": {
            "WMA": [round(v, 2) for v in wma_fcast],
            "DES": [round(v, 2) for v in des_fcast],
            "LR":  [round(v, 2) for v in lr_fcast],
        },
        "best_wma_window": best_w,
        "volatility":      round(vol, 6),
        "trend_pct":       round(trend_pct, 4),
        "confidence":      confidence,
        "n_days":          n,
        "alpha":           alpha,
        "beta":            beta,
        "saz_factors":     {str(k): round(v, 4) for k, v in saz_factors.items()},
    }


# ══════════════════════════════════════════════════════════════════════════════
#  MARKET MAKING — analise de spread otimo e scanner de oportunidades (Tier 3)
# ══════════════════════════════════════════════════════════════════════════════
#
#  O market maker coloca simultaneamente:
#   - Ordem de COMPRA: 0.01 ISK acima do melhor buy (frente da fila)
#   - Ordem de VENDA:  0.01 ISK abaixo do melhor sell (frente da fila)
#  Capturando o spread bruto menos 2x broker fee + sales tax.
#
#  Metricas calculadas por item:
#   spread_abs    — diferenca absoluta entre sell_min e buy_max (ISK)
#   spread_pct    — spread_abs / sell_min * 100 (%)
#   mm_buy_price  — preco otimo de compra (buy_max + 0.01)
#   mm_sell_price — preco otimo de venda (sell_min - 0.01)
#   cost_round_trip — custo total de ir e voltar (2*broker_fee + sales_tax)
#   net_mm_pct    — lucro liquido por ciclo completo (%)
#   fill_time_est — tempo estimado de preenchimento (dias) = capital/giro_dia
#   daily_profit  — lucro diario estimado se giro medio for mantido
#   queue_depth_buy  — ordens de compra na mesma faixa de preco (concorrencia)
#   queue_depth_sell — ordens de venda na mesma faixa de preco (concorrencia)
#   inventory_risk — volatilidade * capital = risco de inventario em ISK
#   mm_score      — score composite 0-100

def analyze_market_making(orders, region_name, item_cache,
                           broker_rel: int = 0, accounting: int = 0,
                           min_spread_pct: float = 1.0,
                           min_vol_dia: float = 5.0,
                           max_queue_depth: int = 50,
                           min_net_pct: float = 0.5,
                           top_n: int = 200,
                           blacklist: set = None,
                           progress_cb=None) -> list:
    """
    Escaneia todos os itens de uma regiao e retorna oportunidades de
    market making ordenadas por mm_score desc.

    Criterios de selecao:
      - spread_pct >= min_spread_pct   (espaco para cobrir custos)
      - avg_vol_90 >= min_vol_dia      (giro suficiente para preencher)
      - net_mm_pct >= min_net_pct      (lucro liquido positivo apos taxas)
      - queue_depth_buy <= max_queue_depth  (fila gerenciavel)
    """
    rb, rt   = _calc_fee_rates(broker_rel, accounting)
    # Custo de uma viagem de ida e volta: compra + venda
    cost_rtrip = rb + rb + rt   # broker na compra + broker na venda + sales tax na venda

    idx      = _filter_blacklist(build_order_index(orders), blacklist or set())
    rid      = _region_id(region_name)
    _ensure_names(idx, item_cache, progress_cb)
    _warm_hist_cache(rid, progress_cb)  # pré-aquece cache RAM

    candidates = []
    total = len(idx)
    chunk = max(1, total // 40)
    for i, (tid, entry) in enumerate(idx.items()):
        if progress_cb and i % chunk == 0:
            progress_cb((i, total, f"Analisando livro de ordens... {i:,}/{total:,}"))

        sells = entry["sell"]
        buys  = entry["buy"]
        if not sells or not buys:
            continue

        sell_min = float(sells[0]["price"])
        buy_max  = float(buys[0]["price"])
        if sell_min <= buy_max or sell_min <= 0:
            continue

        spread_abs = sell_min - buy_max
        spread_pct = spread_abs / sell_min * 100
        if spread_pct < min_spread_pct:
            continue

        # Precos otimos de MM (0.01 ISK a frente da fila)
        mm_buy  = round(buy_max  + 0.01, 2)
        mm_sell = round(sell_min - 0.01, 2)
        if mm_sell <= mm_buy:
            continue

        # Lucro liquido por ciclo: vender - comprar - taxas
        receita    = mm_sell * (1 - rb - rt)
        custo      = mm_buy  * (1 + rb)
        net_profit = receita - custo
        net_pct    = net_profit / custo * 100 if custo > 0 else 0
        if net_pct < min_net_pct:
            continue

        # Profundidade da fila (ordens dentro de 5% do melhor preco)
        thr_buy  = buy_max  * 0.95
        thr_sell = sell_min * 1.05
        q_buy  = sum(1 for o in buys  if float(o["price"]) >= thr_buy)
        q_sell = sum(1 for o in sells if float(o["price"]) <= thr_sell)
        if q_buy > max_queue_depth:
            continue

        # Volume de ordens na fila (ISK imobilizado pelos concorrentes)
        vol_buy_q  = sum(int(o.get("volume_remain",0)) * float(o["price"])
                         for o in buys[:5])
        vol_sell_q = sum(int(o.get("volume_remain",0)) * float(o["price"])
                         for o in sells[:5])

        info = item_cache.get(tid, {})
        candidates.append({
            "type_id":      tid,
            "item":         info.get("name", str(tid)),
            "sell_min":     round(sell_min, 2),
            "buy_max":      round(buy_max,  2),
            "spread_abs":   round(spread_abs, 2),
            "spread_pct":   round(spread_pct, 2),
            "mm_buy":       mm_buy,
            "mm_sell":      mm_sell,
            "net_pct":      round(net_pct, 3),
            "net_profit":   round(net_profit, 2),
            "cost_rtrip_pct": round(cost_rtrip * 100, 3),
            "q_buy":        q_buy,
            "q_sell":       q_sell,
            "vol_buy_q":    round(vol_buy_q, 0),
            "vol_sell_q":   round(vol_sell_q, 0),
            "region":       region_name,
            # Preenchidos apos historico:
            "avg_vol_90":   None,
            "fill_days":    None,
            "daily_profit": None,
            "volatility":   None,
            "inventory_risk": None,
            "mm_score":     None,
        })

    if not candidates:
        if progress_cb:
            progress_cb("Nenhum candidato encontrado com os filtros atuais.")
        return []

    # Ordena por spread para buscar historico dos top candidatos
    candidates.sort(key=lambda x: x["net_pct"], reverse=True)
    top_cands = candidates[:min(top_n * 3, len(candidates))]

    if rid and top_cands:
        if progress_cb:
            progress_cb(f"Buscando historico de {len(top_cands)} candidatos...")

        def _fetch(c):
            h  = _get_hist_cached(rid, c["type_id"])
            st = _history_stats(h)
            if not st or st["avg_vol_90"] < min_vol_dia:
                c["avg_vol_90"] = st["avg_vol_90"] if st else 0.0
                c["mm_score"]   = 0.0
                return

            vd  = st["avg_vol_90"]
            c["avg_vol_90"] = round(vd, 1)

            # Tempo de preenchimento: quantas unidades do capital seriam
            # negociadas por dia? Assume que capturamos ~20% do giro diario
            # (ja que compartilhamos a fila com outros market makers)
            share     = 0.20
            fill_qty  = vd * share
            # Lucro diario = fill_qty * net_profit por unidade
            c["daily_profit"] = round(fill_qty * c["net_profit"], 2)
            # Capital imobilizado na buy order (estimado: 10x o fill diario)
            capital_imob = c["mm_buy"] * fill_qty * 10
            # Dias para 1 ciclo completo
            c["fill_days"] = round(1 / share if share > 0 else 5, 1)

            # Volatilidade: desvio padrao relativo dos ultimos 30d
            hist_s  = sorted(h, key=lambda d: d.get("date",""))
            prices30= [float(d.get("average",0)) for d in hist_s[-30:]]
            if len(prices30) >= 5:
                mu30  = sum(prices30) / len(prices30)
                sd30  = (sum((p-mu30)**2 for p in prices30) / len(prices30))**0.5
                vol30 = sd30 / mu30 if mu30 > 0 else 0.0
            else:
                vol30 = 0.02
            c["volatility"]     = round(vol30, 4)
            c["inventory_risk"] = round(capital_imob * vol30, 2)

            # MM Score (0-100): combina net_pct, giro, fila e volatilidade
            #   net_pct   → ate 40 pts (0-10% → 0-40)
            #   giro/dia  → ate 30 pts (log10 escala)
            #   fila      → ate 20 pts (fila menor = mais pts)
            #   volatil.  → ate 10 pts (menor volatilidade = mais pts)
            score_net  = min(40, c["net_pct"]  * 4)
            score_giro = min(30, math.log10(vd + 1) / math.log10(10001) * 30)
            score_fila = max(0,  20 - c["q_buy"] * 0.4)
            score_vol  = max(0,  10 - vol30 * 100)
            c["mm_score"] = round(score_net + score_giro + score_fila + score_vol, 2)

        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as pool:
            list(pool.map(_fetch, top_cands))

    # Filtra por vol minimo e ordena por mm_score
    results = [c for c in top_cands
               if (c["avg_vol_90"] or 0) >= min_vol_dia
               and (c["mm_score"]  or 0) > 0]
    results.sort(key=lambda x: x["mm_score"] or 0, reverse=True)
    results = results[:top_n]

    if progress_cb:
        progress_cb(f"{len(results)} oportunidades de market making encontradas."
                    + _hist_err_note(rid, idx.keys()))
    return results


def simulate_mm_spread(orders_entry, type_id: int,
                        broker_rel: int = 0, accounting: int = 0,
                        capital: float = 100_000_000.0,
                        vol_dia: float = 100.0) -> dict:
    """
    Dado um book de ordens para um type_id especifico, simula diferentes
    configuracoes de spread e retorna a tabela de resultados.

    orders_entry: {"buy": [...], "sell": [...]} — entrada do build_order_index
    capital:      capital total disponivel para imobilizar na buy order
    vol_dia:      volume diario estimado do item

    Retorna dict com:
      book_levels   — top 5 niveis do book (buy e sell)
      scenarios     — lista de cenarios com diferentes spreads
      optimal       — cenario com melhor lucro diario
      breakeven_spread_pct — spread minimo para cobrir todas as taxas
    """
    rb, rt  = _calc_fee_rates(broker_rel, accounting)
    sells   = orders_entry.get("sell", [])
    buys    = orders_entry.get("buy",  [])

    if not sells or not buys:
        return {}

    sell_min = float(sells[0]["price"])
    buy_max  = float(buys[0]["price"])
    mid      = (sell_min + buy_max) / 2

    # Break-even: spread mínimo para que net_profit = 0
    # receita = sell * (1 - rb - rt)
    # custo   = buy  * (1 + rb)
    # net = 0 => sell * (1-rb-rt) = buy * (1+rb)
    # => sell/buy = (1+rb)/(1-rb-rt)
    # breakeven_spread_pct = (sell_be - buy_max) / sell_be * 100
    # onde sell_be = buy_max * (1+rb)/(1-rb-rt)
    if (1 - rb - rt) > 0:
        sell_be        = buy_max * (1 + rb) / (1 - rb - rt)
        breakeven_abs  = sell_be - buy_max
        breakeven_pct  = breakeven_abs / sell_be * 100
    else:
        breakeven_pct  = 100.0

    # Niveis do book (top 5 buy e sell)
    book_levels = {
        "sell": [{"price": float(o["price"]),
                  "vol":   int(o.get("volume_remain", 0))}
                 for o in sells[:5]],
        "buy":  [{"price": float(o["price"]),
                  "vol":   int(o.get("volume_remain", 0))}
                 for o in buys[:5]],
        "spread_abs":  round(sell_min - buy_max, 2),
        "spread_pct":  round((sell_min - buy_max) / sell_min * 100, 2),
        "mid":         round(mid, 2),
        "sell_min":    round(sell_min, 2),
        "buy_max":     round(buy_max,  2),
    }

    # Gera cenarios com diferentes posicionamentos dentro do spread
    # Posicao: 0.0 = frente da fila (buy_max+0.01 / sell_min-0.01)
    #          0.5 = meio do spread
    #          1.0 = spread maximo (longe da fila)
    scenarios = []
    positions = [0.0, 0.05, 0.1, 0.2, 0.3, 0.5, 0.75, 1.0]

    for pos in positions:
        mm_buy  = round(buy_max  + 0.01 + pos * (mid - buy_max  - 0.01), 2)
        mm_sell = round(sell_min - 0.01 - pos * (sell_min - 0.01 - mid), 2)
        if mm_sell <= mm_buy:
            continue

        receita     = mm_sell * (1 - rb - rt)
        custo       = mm_buy  * (1 + rb)
        net_unit    = receita - custo
        net_pct     = net_unit / custo * 100 if custo > 0 else 0
        sp_abs      = mm_sell - mm_buy
        sp_pct      = sp_abs / mm_sell * 100

        # Estimativa de fill: quanto mais longe da fila, menos trades capturamos
        # fill_rate = 1.0 na frente (pos=0) a 0.05 no fundo (pos=1.0)
        fill_rate    = max(0.05, 1.0 - pos * 0.95)
        qty_dia      = vol_dia * fill_rate * 0.2  # 20% do giro * fill_rate
        # Lucro diario (0 se net_unit negativo — posicao nao lucrativa)
        daily_profit = qty_dia * net_unit if net_unit > 0 else 0.0

        # Capital imobilizado: qty_dia * 10 dias * mm_buy
        capital_req  = max(qty_dia * 10 * mm_buy, 1.0)
        roi_capital  = (daily_profit * 30 / capital_req * 100
                        if capital_req > 0 and daily_profit > 0 else 0.0)

        scenarios.append({
            "pos":           pos,
            "pos_label":     f"{pos*100:.0f}%",
            "mm_buy":        mm_buy,
            "mm_sell":       mm_sell,
            "spread_abs":    round(sp_abs, 2),
            "spread_pct":    round(sp_pct, 2),
            "net_pct":       round(net_pct, 3),
            "net_unit":      round(net_unit, 2),
            "fill_rate_pct": round(fill_rate * 100, 1),
            "qty_dia":       round(qty_dia, 1),
            "daily_profit":  round(daily_profit, 2),
            "capital_req":   round(capital_req, 2),
            "roi_30d_pct":   round(roi_capital, 2),
            "profitable":    net_unit > 0,
        })

    if not scenarios:
        return {}

    optimal = max(scenarios, key=lambda s: s["daily_profit"])

    return {
        "type_id":           type_id,
        "broker_rel":        broker_rel,
        "accounting":        accounting,
        "rb":                round(rb, 6),
        "rt":                round(rt, 6),
        "book_levels":       book_levels,
        "scenarios":         scenarios,
        "optimal":           optimal,
        "breakeven_pct":     round(breakeven_pct, 3),
        "capital":           capital,
        "vol_dia":           vol_dia,
    }


# ══════════════════════════════════════════════════════════════════════════════
#  CALCULADORA DE INDUSTRIA (Alta prioridade)
# ══════════════════════════════════════════════════════════════════════════════

def fetch_industry_indices() -> dict:
    """
    Busca indices de custo de producao por sistema solar.
    Retorna {solar_system_id: {activity: cost_index}} onde activity pode ser:
    manufacturing, researching_material_efficiency, etc.
    """
    try:
        r = _get_session().get(f"{ESI_BASE}/industry/systems/",
                               timeout=20, params={"datasource": "tranquility"})
        if r.status_code == 200:
            return {item["solar_system_id"]: {
                a["activity"]: a["cost_index"]
                for a in item.get("cost_indices", [])}
                for item in r.json()}
    except Exception as exc:
        log.warning("fetch_industry_indices: %s", exc)
    return {}


def fetch_adjusted_prices() -> dict:
    """
    Busca precos medios ajustados da CCP (usados na formula de custo de instalacao).
    Retorna {type_id: adjusted_price}.
    """
    try:
        r = _get_session().get(f"{ESI_BASE}/markets/prices/",
                               timeout=20, params={"datasource": "tranquility"})
        if r.status_code == 200:
            return {item["type_id"]: item.get("adjusted_price", 0.0)
                    for item in r.json()}
    except Exception as exc:
        log.warning("fetch_adjusted_prices: %s", exc)
    return {}


def fetch_blueprint_info(type_id: int) -> dict:
    """
    Busca informacoes de blueprint via ESI /universe/types/ para nome/volume.
    Retorna dict com informacoes do tipo ou {} se nao encontrado.
    """
    try:
        r = _get_session().get(f"{ESI_BASE}/universe/types/{type_id}/",
                               timeout=10, params={"datasource": "tranquility",
                                                   "language": "en"})
        if r.status_code == 200:
            return r.json()
    except Exception as _e:
        log.debug("silenced: %s", _e)
    return {}


# ── Cache SQLite para materiais de blueprint ──────────────────────────────────
# Chave: (product_type_id, activity_id)  Valor: JSON dos materiais
# TTL: 24 horas (blueprints mudam raramente — apenas com patches CCP)
_BP_MATERIALS_TTL = 86_400   # 24h em segundos
_BP_DB_PATH       = os.path.join(DATA_DIR, "blueprint_cache.db")
_SDE_SQL_BZ2_PATH = os.path.join(DATA_DIR, "industryActivityMaterials.sql.bz2")
_bp_db_local      = threading.local()


def _get_bp_db():
    """
    Retorna conexão SQLite thread-local para cache de blueprints.
    Tabelas:
      bp_materials      — cache individual por lookup (TTL 24h)
      hoboleaks_idx     — índice SDE completo de materiais (TTL 7 dias)
      hoboleaks_products— mapeamento blueprint → produto fabricado
      hoboleaks_sync    — controle de quando foi feito o último sync
    """
    if not hasattr(_bp_db_local, "conn") or _bp_db_local.conn is None:
        conn = sqlite3.connect(_BP_DB_PATH, check_same_thread=False)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS bp_materials (
                lookup_tid  INTEGER NOT NULL,
                activity_id INTEGER NOT NULL DEFAULT 1,
                fetched_at  INTEGER NOT NULL,
                materials   TEXT    NOT NULL,
                qty_per_run INTEGER NOT NULL DEFAULT 1,
                bp_tid      INTEGER NOT NULL DEFAULT 0,
                PRIMARY KEY (lookup_tid, activity_id)
            )""")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS hoboleaks_idx (
                bp_tid      INTEGER NOT NULL,
                activity_id INTEGER NOT NULL DEFAULT 1,
                mat_tid     INTEGER NOT NULL,
                quantity    INTEGER NOT NULL DEFAULT 1,
                PRIMARY KEY (bp_tid, activity_id, mat_tid)
            )""")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS hoboleaks_products (
                bp_tid      INTEGER NOT NULL,
                activity_id INTEGER NOT NULL DEFAULT 1,
                prod_tid    INTEGER NOT NULL,
                qty_per_run INTEGER NOT NULL DEFAULT 1,
                PRIMARY KEY (bp_tid, activity_id)
            )""")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS hoboleaks_sync (
                key       TEXT PRIMARY KEY,
                synced_at INTEGER NOT NULL DEFAULT 0
            )""")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_hob_bp "
                     "ON hoboleaks_idx(bp_tid, activity_id)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_hob_prod "
                     "ON hoboleaks_products(prod_tid, activity_id)")
        conn.commit()
        _bp_db_local.conn = conn
    return _bp_db_local.conn


def _bp_cache_get(lookup_tid: int, activity_id: int = 1) -> dict | None:
    """Lê materiais do cache SQLite (TTL 24h)."""
    try:
        conn = _get_bp_db()
        row  = conn.execute(
            "SELECT fetched_at, materials, qty_per_run, bp_tid "
            "FROM bp_materials WHERE lookup_tid=? AND activity_id=?",
            (lookup_tid, activity_id)).fetchone()
        if row and (time.time() - row[0]) < _BP_MATERIALS_TTL:
            return {"materials": json.loads(row[1]),
                    "qty_per_run": row[2], "bp_tid": row[3]}
    except Exception as _e:
        log.debug("_bp_cache_get: %s", _e)
    return None


def _bp_cache_set(lookup_tid: int, activity_id: int,
                   materials: list, qty_per_run: int, bp_tid: int):
    """Armazena materiais no cache SQLite."""
    try:
        conn = _get_bp_db()
        conn.execute(
            "INSERT OR REPLACE INTO bp_materials "
            "(lookup_tid, activity_id, fetched_at, materials, qty_per_run, bp_tid)"
            " VALUES (?,?,?,?,?,?)",
            (lookup_tid, activity_id, int(time.time()),
             json.dumps(materials), qty_per_run, bp_tid))
        conn.commit()
    except Exception as _e:
        log.debug("_bp_cache_set: %s", _e)


# ── Índice Hoboleaks SDE local ─────────────────────────────────────────────────
# O Hoboleaks disponibiliza o SDE completo da CCP como JSON público.
# Baixamos uma vez (TTL 7 dias) e indexamos no SQLite local.
def _load_local_sql_bz2(progress_cb=None) -> bool:
    """
    Carrega e indexa o arquivo ``industryActivityMaterials.sql.bz2`` local.

    O arquivo fica em ``eve_data/industryActivityMaterials.sql.bz2``
    (mesma pasta do market_history.db). É um dump MySQL da tabela SDE:

        CREATE TABLE `industryActivityMaterials` (
          `typeID`         int  -- blueprint type_id
          `activityID`     int  -- 1=fabricação, 3=pesq.TE, 4=pesq.ME, 5=cópia
          `materialTypeID` int  -- material necessário
          `quantity`       int  -- quantidade necessária (sem ME)
        );

    Funcionamento:
      1. Abre o .bz2, lê o conteúdo SQL completo
      2. Extrai todas as tuplas com regex (um único INSERT INTO com 36k+ valores)
      3. Insere em lote na tabela ``hoboleaks_idx`` do SQLite local
      4. Registra o timestamp de carga em ``hoboleaks_sync``

    Retorna True se o índice foi criado com sucesso.
    """
    if not os.path.exists(_SDE_SQL_BZ2_PATH):
        log.debug("_load_local_sql_bz2: arquivo não encontrado em %s", _SDE_SQL_BZ2_PATH)
        return False

    if progress_cb:
        progress_cb(f"⟳ Carregando {os.path.basename(_SDE_SQL_BZ2_PATH)}...")
    log.info("Carregando SDE local de %s", _SDE_SQL_BZ2_PATH)

    try:
        import bz2 as _bz2
        with _bz2.open(_SDE_SQL_BZ2_PATH, "rb") as fh:
            raw = fh.read().decode("utf-8", errors="replace")
    except Exception as _e:
        log.warning("_load_local_sql_bz2 leitura: %s", _e)
        return False

    # Extrai tuplas: (typeID, activityID, materialTypeID, quantity)
    # O dump MySQL usa um único INSERT INTO com todos os valores concatenados
    tuples = re.findall(r'\((\d+),(\d+),(\d+),(\d+)\)', raw)
    if not tuples:
        log.warning("_load_local_sql_bz2: nenhuma tupla encontrada no SQL")
        return False

    if progress_cb:
        progress_cb(f"⟳ Indexando {len(tuples):,} registros no SQLite...")

    try:
        conn = _get_bp_db()
        conn.execute("DELETE FROM hoboleaks_idx")
        conn.execute("DELETE FROM hoboleaks_products")

        # Insere materiais em lote — eficiente para 36k+ registros
        mat_rows = [
            (int(bp_tid), int(act_id), int(mat_tid), int(qty))
            for bp_tid, act_id, mat_tid, qty in tuples
            if int(bp_tid) > 0 and int(mat_tid) > 0
        ]
        conn.executemany(
            "INSERT OR IGNORE INTO hoboleaks_idx VALUES (?,?,?,?)", mat_rows)

        # Marca como sincronizado com timestamp atual
        conn.execute(
            "INSERT OR REPLACE INTO hoboleaks_sync(key, synced_at) VALUES (?,?)",
            ("materials", int(time.time())))
        conn.commit()

        n_blueprints = len(set(r[0] for r in mat_rows))
        log.info("SDE local indexado: %d registros, %d blueprints únicos",
                 len(mat_rows), n_blueprints)
        if progress_cb:
            progress_cb(f"✓ SDE local: {len(mat_rows):,} materiais, "
                        f"{n_blueprints:,} blueprints")
        return True

    except Exception as _e:
        log.warning("_load_local_sql_bz2 indexação: %s", _e)
        return False


# ── Fontes remotas do SDE ─────────────────────────────────────────────────────

# Estratégia:
#   1. Fuzzwork typematerials.php     (HTTP, por blueprint)
#   2. Fuzzwork blueprint/api/        (HTTP, alternativo)
#   3. everef.net SDE JSON            (HTTP, espelho comunitário do SDE)
#   4. CCP SDE ZIP                    (HTTP, fonte oficial CCP)
#   5. Dados hardcoded                (offline, materiais dos blueprints comuns)
#
# O índice local SDE é construído a partir de qualquer das fontes acima
# e armazenado no SQLite (blueprint_cache.db), eliminando requests repetidos.

_SDE_TTL = 7 * 86_400   # 7 dias

# ── Fontes do SDE ──────────────────────────────────────────────────────────────
_SDE_SOURCES = {
    # everef.net — espelho comunitário do SDE, mantido por Squizz
    "everef_materials": "https://data.everef.net/industry/blueprints/latest/industry-activity-materials.json",
    "everef_products":  "https://data.everef.net/industry/blueprints/latest/industry-activity-products.json",
    # Hoboleaks — outro espelho, URLs alternativas a tentar
    "hoboleaks_mat_1":  "https://sde.hoboleaks.space/tq/industryActivityMaterials.json",
    "hoboleaks_mat_2":  "https://sde.hoboleaks.space/tq/industryactivitymaterials.json",
    "hoboleaks_mat_3":  "https://hoboleaks.space/tq/industryActivityMaterials.json",
    "hoboleaks_prod_1": "https://sde.hoboleaks.space/tq/industryActivityProducts.json",
    "hoboleaks_prod_2": "https://sde.hoboleaks.space/tq/industryactivityproducts.json",
    # CCP SDE oficial — ZIP com blueprints.yaml dentro
    "ccp_sde_zip":      "https://eve-static-data-export.s3-eu-west-1.amazonaws.com/tranquility/sde.zip",
}

# ── Materiais hardcoded para blueprints extremamente comuns ───────────────────
# Fallback offline de último recurso. Cobre ~90% dos blueprints de traders.
# Formato: {bp_tid: {"m": [(mat_tid, qty), ...], "q": qty_per_run, "prod": prod_tid}}
_HARDCODED_BP: dict = {
    # ── Munições T1 ──────────────────────────────────────────────────────────
    # prod_tid removido — type_ids de produtos variam por versão do EVE.
    # O mapeamento blueprint→produto é feito pelo nome no item_cache do usuário.
    807:  {"m":[(34,316),(35,108),(38,1)],       "q":100},   # Nova Heavy Missile Blueprint
    812:  {"m":[(34,316),(35,108),(38,1)],       "q":100},   # Mjolnir Heavy Missile Blueprint
    817:  {"m":[(34,316),(35,108),(38,1)],       "q":100},   # Scourge Heavy Missile Blueprint
    822:  {"m":[(34,316),(35,108),(38,1)],       "q":100},   # Inferno Heavy Missile Blueprint
    803:  {"m":[(34,77),(35,26)],                "q":100},   # Nova Light Missile Blueprint
    808:  {"m":[(34,77),(35,26)],                "q":100},   # Mjolnir Light Missile Blueprint
    813:  {"m":[(34,77),(35,26)],                "q":100},   # Scourge Light Missile Blueprint
    818:  {"m":[(34,77),(35,26)],                "q":100},   # Inferno Light Missile Blueprint
    # ── Ogivas (Rockets) T1 ────────────────────────────────────────────────
    801:  {"m":[(34,38),(35,13)],                "q":100},   # Nova Rocket Blueprint
    806:  {"m":[(34,38),(35,13)],                "q":100},   # Mjolnir Rocket Blueprint
    811:  {"m":[(34,38),(35,13)],                "q":100},   # Scourge Rocket Blueprint
    816:  {"m":[(34,38),(35,13)],                "q":100},   # Inferno Rocket Blueprint
    # ── Módulos T1 comuns ───────────────────────────────────────────────────
    3895: {"m":[(34,4000),(35,1500),(36,150),(37,60),(38,12)],      "q":1},  # Small Shield Extender I Blueprint
    11573:{"m":[(34,16000),(35,6000),(36,600),(37,240),(38,48)],    "q":1},  # Medium Shield Extender I Blueprint
    11575:{"m":[(34,64000),(35,24000),(36,2400),(37,960),(38,192)], "q":1},  # Large Shield Extender I Blueprint
    # ── Fragatas T1 ─────────────────────────────────────────────────────────
    596:  {"m":[(34,50000),(35,6000),(36,200),(37,200),(38,30)],    "q":1},  # Rifter Blueprint
    639:  {"m":[(34,28000),(35,3500),(36,110),(37,100),(38,15)],    "q":1},  # Merlin Blueprint
    643:  {"m":[(34,28000),(35,3500),(36,110),(37,100),(38,15)],    "q":1},  # Tristan Blueprint
}


def _sde_is_fresh() -> bool:
    """Verifica se o índice SDE local está recente."""
    try:
        row = _get_bp_db().execute(
            "SELECT synced_at FROM hoboleaks_sync WHERE key='materials'"
        ).fetchone()
        return bool(row and (time.time() - row[0]) < _SDE_TTL)
    except Exception:
        return False


def _sde_has_data() -> bool:
    """Verifica se o índice SDE tem entradas."""
    try:
        row = _get_bp_db().execute("SELECT COUNT(*) FROM hoboleaks_idx").fetchone()
        return bool(row and row[0] > 0)
    except Exception:
        return False


def _sde_try_download_flat_json(mat_urls: list, prod_urls: list,
                                 progress_cb=None) -> bool:
    """
    Tenta baixar o SDE no formato JSON plano (lista de dicts).
    Aceita múltiplas URLs de fallback para materiais e produtos.

    Formatos suportados:
      Everef/Hoboleaks: [{blueprintTypeID, activityID, materialTypeID, quantity}, ...]
    """
    materials_list = []
    for url in mat_urls:
        try:
            if progress_cb:
                progress_cb(f"⟳ Tentando: {url.split('//')[-1][:50]}...")
            r = _get_session().get(url, timeout=60)
            if r.status_code == 200:
                data = r.json()
                if isinstance(data, list) and data:
                    materials_list = data
                    log.info("SDE materials: %d entradas de %s", len(data), url)
                    break
                elif isinstance(data, dict):
                    # Formato alternativo: {bp_tid: [{activityID, materialTypeID, quantity}]}
                    flat = []
                    for bp_tid_str, entries in data.items():
                        for e in (entries if isinstance(entries, list) else []):
                            flat.append({
                                "blueprintTypeID": int(bp_tid_str),
                                "activityID":      e.get("activityID", 1),
                                "materialTypeID":  e.get("materialTypeID", 0),
                                "quantity":        e.get("quantity", 1),
                            })
                    if flat:
                        materials_list = flat
                        log.info("SDE materials (dict format): %d de %s", len(flat), url)
                        break
            else:
                log.debug("SDE %s HTTP %s", url, r.status_code)
        except Exception as _e:
            log.debug("SDE download %s: %s", url, _e)

    if not materials_list:
        return False

    products_list = []
    for url in prod_urls:
        try:
            r2 = _get_session().get(url, timeout=60)
            if r2.status_code == 200:
                data2 = r2.json()
                if isinstance(data2, list) and data2:
                    products_list = data2
                    break
                elif isinstance(data2, dict):
                    # Formato alternativo
                    for bp_tid_str, entries in data2.items():
                        for e in (entries if isinstance(entries, list) else []):
                            products_list.append({
                                "blueprintTypeID": int(bp_tid_str),
                                "activityID":      e.get("activityID", 1),
                                "productTypeID":   e.get("productTypeID", 0),
                                "quantity":        e.get("quantity", 1),
                            })
                    if products_list:
                        break
        except Exception as _e:
            log.debug("SDE products %s: %s", url, _e)

    if progress_cb:
        progress_cb(f"⟳ Indexando {len(materials_list):,} materiais no SQLite...")

    try:
        conn = _get_bp_db()
        conn.execute("DELETE FROM hoboleaks_idx")
        conn.execute("DELETE FROM hoboleaks_products")

        mat_rows = [
            (int(e.get("blueprintTypeID", 0)), int(e.get("activityID", 1)),
             int(e.get("materialTypeID", 0)),  int(e.get("quantity", 1)))
            for e in materials_list
            if int(e.get("blueprintTypeID", 0)) > 0 and int(e.get("materialTypeID", 0)) > 0
        ]
        conn.executemany("INSERT OR IGNORE INTO hoboleaks_idx VALUES (?,?,?,?)", mat_rows)

        if products_list:
            prod_rows = [
                (int(e.get("blueprintTypeID", 0)), int(e.get("activityID", 1)),
                 int(e.get("productTypeID", 0)),   int(e.get("quantity", 1)))
                for e in products_list
                if int(e.get("blueprintTypeID", 0)) > 0 and int(e.get("productTypeID", 0)) > 0
            ]
            conn.executemany("INSERT OR IGNORE INTO hoboleaks_products VALUES (?,?,?,?)", prod_rows)

        conn.execute(
            "INSERT OR REPLACE INTO hoboleaks_sync(key, synced_at) VALUES (?,?)",
            ("materials", int(time.time())))
        conn.commit()
        log.info("SDE indexado: %d materiais, %d produtos", len(mat_rows), len(products_list))
        if progress_cb:
            progress_cb(f"✓ SDE indexado: {len(mat_rows):,} materiais")
        return True
    except Exception as _e:
        log.warning("SDE indexing: %s", _e)
        return False


def _sde_try_download_ccp_zip(progress_cb=None) -> bool:
    """
    Baixa o SDE oficial da CCP (ZIP) e extrai blueprints.yaml.
    Requer PyYAML (pip install pyyaml).
    Download pesado (~300MB) — usado apenas como último recurso.
    """
    try:
        import yaml   # noqa
    except ImportError:
        log.debug("PyYAML não disponível para SDE CCP ZIP")
        return False

    import zipfile, io

    SDE_URL = ("https://eve-static-data-export.s3-eu-west-1.amazonaws.com"
               "/tranquility/sde.zip")
    if progress_cb:
        progress_cb("⟳ Baixando SDE oficial CCP (~300MB)... aguarde...")
    try:
        log.info("Baixando SDE CCP ZIP de %s", SDE_URL)
        r = _get_session().get(SDE_URL, timeout=300, stream=True)
        if r.status_code != 200:
            log.warning("SDE CCP ZIP HTTP %s", r.status_code)
            return False

        buf = io.BytesIO()
        total = 0
        for chunk in r.iter_content(chunk_size=1024 * 256):
            buf.write(chunk)
            total += len(chunk)
            if progress_cb and total % (10 * 1024 * 1024) == 0:
                progress_cb(f"⟳ SDE CCP: {total//1024//1024}MB baixados...")

        buf.seek(0)
        if progress_cb:
            progress_cb("⟳ Extraindo blueprints.yaml do ZIP...")

        with zipfile.ZipFile(buf) as zf:
            # O arquivo pode estar em sde/fsd/blueprints.yaml ou blueprints.yaml
            target = None
            for name in zf.namelist():
                if "blueprints.yaml" in name.lower():
                    target = name
                    break
            if not target:
                log.warning("blueprints.yaml não encontrado no SDE ZIP")
                return False

            if progress_cb:
                progress_cb("⟳ Parseando blueprints.yaml (PyYAML)...")
            with zf.open(target) as f:
                blueprints = yaml.safe_load(f)

        if not isinstance(blueprints, dict):
            return False

        if progress_cb:
            progress_cb(f"⟳ Indexando {len(blueprints):,} blueprints no SQLite...")

        conn = _get_bp_db()
        conn.execute("DELETE FROM hoboleaks_idx")
        conn.execute("DELETE FROM hoboleaks_products")

        mat_rows  = []
        prod_rows = []
        for bp_tid, bp_data in blueprints.items():
            if not isinstance(bp_data, dict):
                continue
            mfg = (bp_data.get("activities") or {}).get("manufacturing") or {}
            for mat in (mfg.get("materials") or []):
                mat_tid = mat.get("typeID", 0)
                qty     = mat.get("quantity", 1)
                if mat_tid:
                    mat_rows.append((int(bp_tid), 1, int(mat_tid), int(qty)))
            for prod in (mfg.get("products") or []):
                prod_tid = prod.get("typeID", 0)
                qty      = prod.get("quantity", 1)
                if prod_tid:
                    prod_rows.append((int(bp_tid), 1, int(prod_tid), int(qty)))

        conn.executemany("INSERT OR IGNORE INTO hoboleaks_idx VALUES (?,?,?,?)", mat_rows)
        conn.executemany("INSERT OR IGNORE INTO hoboleaks_products VALUES (?,?,?,?)", prod_rows)
        conn.execute("INSERT OR REPLACE INTO hoboleaks_sync(key, synced_at) VALUES (?,?)",
                     ("materials", int(time.time())))
        conn.commit()
        log.info("SDE CCP ZIP indexado: %d materiais, %d produtos", len(mat_rows), len(prod_rows))
        if progress_cb:
            progress_cb(f"✓ SDE CCP indexado: {len(mat_rows):,} materiais")
        return True
    except Exception as _e:
        log.warning("SDE CCP ZIP: %s", _e)
        return False


def _sde_download_and_index(progress_cb=None) -> bool:
    """
    Sincroniza o índice SDE local tentando as fontes em ordem de preferência:
      1. Arquivo local industryActivityMaterials.sql.bz2 (instantâneo, sem rede)
      2. everef.net JSON (espelho comunitário)
      3. Hoboleaks SDE JSON (múltiplas URLs variantes)
      4. CCP SDE ZIP oficial (download pesado, último recurso)
    """
    # 1. Arquivo local .sql.bz2 (mais rápido, sem rede)
    if _load_local_sql_bz2(progress_cb):
        return True

    # 2 e 3. Fontes JSON remotas
    mat_urls = [
        _SDE_SOURCES["everef_materials"],
        _SDE_SOURCES["hoboleaks_mat_1"],
        _SDE_SOURCES["hoboleaks_mat_2"],
        _SDE_SOURCES["hoboleaks_mat_3"],
    ]
    prod_urls = [
        _SDE_SOURCES["everef_products"],
        _SDE_SOURCES["hoboleaks_prod_1"],
        _SDE_SOURCES["hoboleaks_prod_2"],
    ]
    if _sde_try_download_flat_json(mat_urls, prod_urls, progress_cb):
        return True

    # 4. SDE oficial CCP (ZIP pesado)
    return _sde_try_download_ccp_zip(progress_cb)


def _sde_lookup(type_id: int, activity_id: int = 1) -> dict | None:
    """
    Busca materiais no índice SDE local.
    Aceita blueprint_tid OU product_tid (lookup cruzado automático).
    """
    try:
        conn = _get_bp_db()

        # Tenta como blueprint direto
        rows = conn.execute(
            "SELECT mat_tid, quantity FROM hoboleaks_idx "
            "WHERE bp_tid=? AND activity_id=?",
            (type_id, activity_id)).fetchall()
        if rows:
            p = conn.execute(
                "SELECT qty_per_run FROM hoboleaks_products "
                "WHERE bp_tid=? AND activity_id=?",
                (type_id, activity_id)).fetchone()
            return {
                "materials":   [{"type_id": r[0], "quantity": r[1], "name": str(r[0])}
                                for r in rows],
                "qty_per_run": p[0] if p else 1,
                "bp_tid":      type_id,
            }

        # Tenta como produto → acha o blueprint
        bp_row = conn.execute(
            "SELECT bp_tid, qty_per_run FROM hoboleaks_products "
            "WHERE prod_tid=? AND activity_id=?",
            (type_id, activity_id)).fetchone()
        if bp_row:
            bp_tid, qty_per_run = bp_row
            rows2 = conn.execute(
                "SELECT mat_tid, quantity FROM hoboleaks_idx "
                "WHERE bp_tid=? AND activity_id=?",
                (bp_tid, activity_id)).fetchall()
            if rows2:
                return {
                    "materials":   [{"type_id": r[0], "quantity": r[1], "name": str(r[0])}
                                   for r in rows2],
                    "qty_per_run": qty_per_run,
                    "bp_tid":      bp_tid,
                }
    except Exception as _e:
        log.debug("_sde_lookup: %s", _e)
    return None


def _hardcoded_lookup(type_id: int) -> dict | None:
    """
    Fallback de dados hardcoded para blueprints extremamente comuns.
    Funciona totalmente offline, sem rede.
    Resolve nomes via _item_cache mais tarde.
    """
    if type_id not in _HARDCODED_BP:
        return None
    entry = _HARDCODED_BP[type_id]
    return {
        "materials":   [{"type_id": m[0], "quantity": m[1], "name": str(m[0])}
                        for m in entry["m"]],
        "qty_per_run": entry["q"],
        "bp_tid":      type_id,
        "prod_tid":    entry.get("prod", 0),
    }


def _get_product_tid(bp_tid: int, bp_data: dict, item_cache: dict) -> int:
    """
    Retorna o type_id do PRODUTO fabricado por um blueprint.

    Tenta em ordem de confiabilidade:
      1. Dedução pelo nome no item_cache — mais confiável pois usa dados reais da ESI
         "Nova Heavy Missile Blueprint" → busca "Nova Heavy Missile" no cache
      2. bp_data["prod_tid"] — da Fuzzwork blueprint/api (quando disponível)
      3. hoboleaks_products SQLite — se produtos foram baixados
      4. Retorna 0 se nenhuma fonte funcionar

    NOTA: _HARDCODED_BP NÃO é usado para prod_tid porque os type_ids de produtos
    variam entre versões do EVE e podem estar incorretos. O nome é a fonte mais
    confiável pois é verificado diretamente no item_cache do usuário.
    """
    # 1. Dedução pelo nome no item_cache (mais confiável — dados reais da ESI)
    # "Nova Heavy Missile Blueprint" → remove " Blueprint" → busca "Nova Heavy Missile"
    bp_name = item_cache.get(bp_tid, {}).get("name", "")
    if bp_name.endswith(" Blueprint"):
        product_name = bp_name[:-len(" Blueprint")]
        for tid, info in item_cache.items():
            if info.get("name", "") == product_name and tid != bp_tid:
                return tid

    # 2. Da Fuzzwork blueprint/api (retorna prod_tid quando disponível)
    if bp_data.get("prod_tid", 0):
        return bp_data["prod_tid"]

    # 3. SQLite hoboleaks_products (preenchido se industryactivityproducts foi baixado)
    try:
        conn  = _get_bp_db()
        row_p = conn.execute(
            "SELECT prod_tid FROM hoboleaks_products WHERE bp_tid=? AND activity_id=1",
            (bp_tid,)).fetchone()
        if row_p and row_p[0]:
            return row_p[0]
    except Exception:
        pass

    return 0   # não encontrado


def fetch_blueprint_materials(type_id: int,
                               activity_id: int = 1,
                               progress_cb=None) -> dict:
    """
    Busca materiais de fabricação para um type_id (blueprint ou produto).

    Estratégia em 5 camadas — da mais rápida para a mais lenta:

      1. Cache SQLite individual (TTL 24h) — O(1), sem rede
      2. Fuzzwork typematerials.php
      3. Fuzzwork blueprint/api/blueprint/
      4. Índice SDE local (SQLite, TTL 7 dias)
         Se o índice estiver desatualizado ou vazio, tenta baixar de:
           a) everef.net (espelho comunitário)
           b) Hoboleaks SDE JSON (múltiplas URLs)
           c) CCP SDE ZIP oficial (fallback pesado, requer PyYAML)
      5. Dados hardcoded (offline, ~20 blueprints comuns)

    O parâmetro typeid aceita tanto blueprint quanto produto.
    Campo Fuzzwork: materialTypeID (não typeid).
    """
    # ── 1. Cache local ─────────────────────────────────────────────────────────
    cached = _bp_cache_get(type_id, activity_id)
    if cached:
        cached["source"] = "cache"
        return cached

    # ── 1b. Índice SDE local se já indexado (sem rede, O(log n)) ──────────────
    # Carrega automaticamente do arquivo .sql.bz2 se o índice estiver vazio
    if not _sde_has_data():
        if os.path.exists(_SDE_SQL_BZ2_PATH):
            log.info("Índice vazio — carregando arquivo SDE local...")
            _load_local_sql_bz2(progress_cb)
        # else: sem arquivo local, tenta fontes remotas após Fuzzwork

    if _sde_has_data():
        result = _sde_lookup(type_id, activity_id)
        if result:
            _bp_cache_set(type_id, activity_id,
                          result["materials"], result["qty_per_run"], result["bp_tid"])
            result["source"] = "sde_local"
            return result

    # ── Helpers de parsing ─────────────────────────────────────────────────────
    def _parse_typematerials(data) -> dict | None:
        if not isinstance(data, list) or not data:
            return None
        materials = [
            {"type_id":  int(item.get("materialTypeID", 0)),
             "quantity": int(item.get("quantity", 1)),
             "name":     str(item.get("typeName", item.get("materialTypeID", "")))}
            for item in data if int(item.get("materialTypeID", 0)) > 0
        ]
        return {"materials": materials,
                "qty_per_run": int(data[0].get("producedQty", 1)),
                "bp_tid":      int(data[0].get("blueprintTypeID", 0)),
                "prod_tid": 0} if materials else None

    def _parse_blueprint_api(data) -> dict | None:
        if not isinstance(data, dict):
            return None
        mfg = (data.get("activities") or {}).get("manufacturing") or {}
        if not mfg:
            return None
        raw_mats = mfg.get("materials") or []
        products = mfg.get("products") or []
        materials = [
            {"type_id": int(m.get("typeID", 0)),
             "quantity": int(m.get("quantity", 1)),
             "name": str(m.get("typeID", 0))}
            for m in raw_mats if int(m.get("typeID", 0)) > 0
        ]
        return {"materials": materials,
                "qty_per_run": int(products[0].get("quantity", 1)) if products else 1,
                "bp_tid":      int(data.get("blueprintTypeID", type_id)),
                "prod_tid":    int(products[0].get("typeID", 0)) if products else 0
                } if materials else None

    # ── 2. Fuzzwork typematerials.php ──────────────────────────────────────────
    try:
        r = _get_session().get(
            "https://www.fuzzwork.co.uk/api/typematerials.php",
            params={"typeid": type_id, "activityid": activity_id},
            timeout=15)
        if r.status_code == 200:
            result = _parse_typematerials(r.json())
            if result:
                _bp_cache_set(type_id, activity_id,
                              result["materials"], result["qty_per_run"], result["bp_tid"])
                result["source"] = "fuzzwork"
                return result
    except Exception as _e:
        log.debug("fuzzwork typematerials(%s): %s", type_id, _e)

    # ── 3. Fuzzwork blueprint/api/blueprint/ ───────────────────────────────────
    try:
        r2 = _get_session().get(
            "https://www.fuzzwork.co.uk/blueprint/api/blueprint/",
            params={"typeid": type_id},
            timeout=15)
        if r2.status_code == 200:
            result = _parse_blueprint_api(r2.json())
            if result:
                _bp_cache_set(type_id, activity_id,
                              result["materials"], result["qty_per_run"], result["bp_tid"])
                result["source"] = "fuzzwork_bp"
                return result
    except Exception as _e:
        log.debug("fuzzwork blueprint_api(%s): %s", type_id, _e)

    # ── 4. Índice SDE via download remoto (se 1b não produziu resultado) ────────
    if not _sde_has_data():
        log.info("Sem arquivo SDE local — tentando fontes remotas...")
        _sde_download_and_index(progress_cb)
    elif not _sde_is_fresh():
        log.debug("Índice SDE desatualizado — usando dados existentes")

    result = _sde_lookup(type_id, activity_id)
    if result:
        _bp_cache_set(type_id, activity_id,
                      result["materials"], result["qty_per_run"], result["bp_tid"])
        result["source"] = "sde_local"
        return result

    # ── 5. Dados hardcoded (offline, blueprints mais comuns) ──────────────────
    result = _hardcoded_lookup(type_id)
    if result:
        _bp_cache_set(type_id, activity_id,
                      result["materials"], result["qty_per_run"], result["bp_tid"])
        result["source"] = "hardcoded"
        return result

    return {"materials": [], "qty_per_run": 1, "bp_tid": 0,
            "prod_tid": 0, "source": "error"}


def fetch_material_prices_bulk(type_ids: list, region_id: int = 10000002) -> dict:
    """
    Busca preços de mercado para múltiplos materiais em paralelo.

    Retorna {type_id: {"sell_min": float, "buy_max": float}} para cada item.
    Usa ThreadPoolExecutor com max 8 workers para respeitar rate limits da ESI.
    Aproveita o cache ETag: se o preço não mudou desde a última busca,
    a ESI retorna 304 Not Modified sem trafegar dados.

    Parâmetros:
      type_ids  — lista de type_ids dos materiais
      region_id — ID da região de mercado (padrão: The Forge / Jita = 10000002)
    """
    results: dict = {}
    lock = threading.Lock()

    def _fetch_one(tid: int):
        data = fetch_live_price(region_id, tid)
        if data:
            with lock:
                results[tid] = {
                    "sell_min": data.get("sell_min", 0),
                    "buy_max":  data.get("buy_max",  0),
                }

    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as pool:
        list(pool.map(_fetch_one, type_ids))

    return results







# ══════════════════════════════════════════════════════════════════════════════
#  BACKTESTING DE ESTRATEGIAS (Alta prioridade)
# ══════════════════════════════════════════════════════════════════════════════

def backtest_strategy(
    hist: list,
    strategy: str = "mm",      # "mm" | "route" | "cycle_buy"
    capital: float = 1_000_000_000.0,
    days_back: int = 90,
    # MM params
    spread_threshold_pct: float = 5.0,
    target_net_pct: float = 1.0,
    broker_rel: int = 0,
    accounting: int = 0,
    # Route params
    buy_region_factor: float = 0.98,  # compra a 98% do preco do dia
    sell_region_factor: float = 1.02,
    # Cycle params
    phase_buy: str = "Acumulacao",
    phase_sell: str = "Distribuicao",
) -> dict:
    """
    Simula uma estrategia de trading sobre historico real do SQLite.

    Estrategias:
      'mm'         — market making: entra quando spread >= threshold, sai quando cai
      'route'      — compra num dia, vende N dias depois com markup fixo
      'cycle_buy'  — compra na fase de Acumulacao, vende na fase de Distribuicao

    Retorna dict com:
      trades        — lista de trades simulados
      equity_curve  — curva de capital ao longo do tempo
      metrics       — PnL total, win_rate, max_drawdown, sharpe
    """
    if not hist or len(hist) < 14:
        return {}

    hist_s = sorted(hist, key=lambda d: d.get("date", ""))
    # Usa os ultimos days_back dias do historico disponivel (nao data absoluta)
    hist_f = hist_s[-days_back:] if len(hist_s) >= days_back else hist_s
    if len(hist_f) < 7:
        return {}

    rb, rt = _calc_fee_rates(broker_rel, accounting)
    trades = []
    equity = capital
    equity_curve = []  # [(date, equity)]
    position = None    # {buy_date, buy_price, qty}

    prices  = [float(d.get("average", 0)) for d in hist_f]
    highs   = [float(d.get("highest", 0)) for d in hist_f]
    lows    = [float(d.get("lowest",  0)) for d in hist_f]
    dates   = [d["date"] for d in hist_f]
    n       = len(hist_f)

    # Calcula EMA14 e EMA50 para ciclo
    ema14_series = _cycle_ema(prices, 14)
    ema50_series = _cycle_ema(prices, 50)

    for i in range(1, n):
        p        = prices[i]
        hi       = highs[i]
        lo       = lows[i]
        date     = dates[i]
        ema14    = ema14_series[i]
        ema50    = ema50_series[i]

        if p <= 0:
            equity_curve.append((date, round(equity, 0)))
            continue

        # ── Market Making ─────────────────────────────────────────────────────
        if strategy == "mm":
            spread_pct = (hi - lo) / hi * 100 if hi > 0 else 0
            net_pct    = spread_pct - (rb + rb + rt) * 100

            if net_pct >= target_net_pct and position is None:
                # Entra: aloca 10% do capital por trade
                qty      = (equity * 0.10) / (p * (1 + rb))
                cost     = qty * p * (1 + rb)
                if cost <= equity:
                    position = {"buy_date": date, "buy_price": p,
                                "qty": qty, "cost": cost,
                                "strategy": "mm", "spread_pct": round(spread_pct, 2)}
                    equity  -= cost

            elif position is not None and net_pct < target_net_pct * 0.5:
                # Saída: spread caiu muito — realiza posicao ao preco médio
                revenue = position["qty"] * p * (1 - rb - rt)
                pnl     = revenue - position["cost"]
                equity += revenue
                trades.append({
                    "buy_date":    position["buy_date"],
                    "sell_date":   date,
                    "buy_price":   round(position["buy_price"], 2),
                    "sell_price":  round(p, 2),
                    "qty":         round(position["qty"], 4),
                    "pnl":         round(pnl, 2),
                    "pnl_pct":     round(pnl / position["cost"] * 100, 3),
                    "strategy":    "mm",
                    "spread_pct":  position["spread_pct"],
                })
                position = None

        # ── Route (compra/venda com intervalo) ────────────────────────────────
        elif strategy == "route":
            hold_days = 7
            if position is None and i % hold_days == 0:
                buy_p = p * buy_region_factor
                qty   = (equity * 0.20) / (buy_p * (1 + rb))
                cost  = qty * buy_p * (1 + rb)
                if cost <= equity:
                    position = {"buy_date": date, "buy_price": buy_p,
                                "qty": qty, "cost": cost, "strategy": "route",
                                "sell_at_day": i + hold_days}
                    equity  -= cost

            elif position is not None and i >= position.get("sell_at_day", 9999):
                sell_p  = p * sell_region_factor
                revenue = position["qty"] * sell_p * (1 - rb - rt)
                pnl     = revenue - position["cost"]
                equity += revenue
                trades.append({
                    "buy_date":   position["buy_date"],
                    "sell_date":  date,
                    "buy_price":  round(position["buy_price"], 2),
                    "sell_price": round(sell_p, 2),
                    "qty":        round(position["qty"], 4),
                    "pnl":        round(pnl, 2),
                    "pnl_pct":    round(pnl / position["cost"] * 100, 3),
                    "strategy":   "route",
                })
                position = None

        # ── Cycle Buy/Sell ────────────────────────────────────────────────────
        elif strategy == "cycle_buy":
            trend = ((ema14 or p) - (ema50 or p)) / (ema50 or p) if ema50 else 0
            vols  = [int(d.get("volume", 0)) for d in hist_f[max(0,i-7):i+1]]
            avg_v = sum(vols) / len(vols) if vols else 1
            cur_v = vols[-1] if vols else 0
            vol_sc = cur_v / avg_v if avg_v > 0 else 1

            # Fase aproximada
            if trend < -0.01 and vol_sc >= 1.2:
                cur_phase = phase_buy      # Acumulacao
            elif trend > 0.03 and vol_sc >= 1.2:
                cur_phase = "Alta"
            elif trend > 0.01 and vol_sc < 0.8:
                cur_phase = phase_sell     # Distribuicao
            else:
                cur_phase = "Lateral"

            if position is None and cur_phase == phase_buy:
                qty  = (equity * 0.25) / (p * (1 + rb))
                cost = qty * p * (1 + rb)
                if cost <= equity:
                    position = {"buy_date": date, "buy_price": p,
                                "qty": qty, "cost": cost,
                                "strategy": "cycle", "phase": cur_phase}
                    equity  -= cost

            elif position is not None and cur_phase == phase_sell:
                revenue = position["qty"] * p * (1 - rb - rt)
                pnl     = revenue - position["cost"]
                equity += revenue
                trades.append({
                    "buy_date":   position["buy_date"],
                    "sell_date":  date,
                    "buy_price":  round(position["buy_price"], 2),
                    "sell_price": round(p, 2),
                    "qty":        round(position["qty"], 4),
                    "pnl":        round(pnl, 2),
                    "pnl_pct":    round(pnl / position["cost"] * 100, 3),
                    "strategy":   "cycle",
                    "phase_buy":  phase_buy,
                    "phase_sell": phase_sell,
                })
                position = None

        # Liquidar posicao aberta no final
        if i == n - 1 and position is not None:
            revenue = position["qty"] * p * (1 - rb - rt)
            pnl     = revenue - position["cost"]
            equity += revenue
            trades.append({
                "buy_date":   position["buy_date"],
                "sell_date":  date + " (fechado)",
                "buy_price":  round(position["buy_price"], 2),
                "sell_price": round(p, 2),
                "qty":        round(position["qty"], 4),
                "pnl":        round(pnl, 2),
                "pnl_pct":    round(pnl / position["cost"] * 100, 3),
                "strategy":   strategy,
                "forced":     True,
            })
            position = None

        equity_curve.append((date, round(equity + (
            position["qty"] * p * (1 - rb - rt) - position["cost"]
            if position else 0), 0)))

    # ── Metricas ──────────────────────────────────────────────────────────────
    pnls       = [t["pnl"] for t in trades]
    total_pnl  = sum(pnls)
    wins       = [p for p in pnls if p > 0]
    losses     = [p for p in pnls if p <= 0]
    win_rate   = len(wins) / len(pnls) * 100 if pnls else 0
    avg_win    = sum(wins)   / len(wins)   if wins   else 0
    avg_loss   = sum(losses) / len(losses) if losses else 0
    profit_factor = abs(sum(wins) / sum(losses)) if losses and sum(losses) != 0 else 999

    # Max drawdown
    peak = capital
    max_dd = 0.0
    for _, eq in equity_curve:
        if eq > peak: peak = eq
        dd = (peak - eq) / peak * 100
        if dd > max_dd: max_dd = dd

    # Sharpe (anualizado, taxa livre de risco = 0)
    if len(pnls) >= 2:
        mu_r  = sum(p / capital for p in pnls) / len(pnls)
        sd_r  = (sum((p/capital - mu_r)**2 for p in pnls) / len(pnls))**0.5
        sharpe = (mu_r / sd_r * (365/days_back)**0.5) if sd_r > 0 else 0
    else:
        sharpe = 0.0

    return {
        "strategy":      strategy,
        "capital_start": capital,
        "capital_end":   round(equity, 2),
        "total_pnl":     round(total_pnl, 2),
        "total_pnl_pct": round(total_pnl / capital * 100, 3),
        "n_trades":      len(trades),
        "win_rate":      round(win_rate, 1),
        "avg_win":       round(avg_win, 2),
        "avg_loss":      round(avg_loss, 2),
        "profit_factor": round(profit_factor, 3),
        "max_drawdown":  round(max_dd, 2),
        "sharpe":        round(sharpe, 3),
        "days_back":     days_back,
        "trades":        trades,
        "equity_curve":  equity_curve,
    }


# ══════════════════════════════════════════════════════════════════════════════
#  HEATMAP DE OPORTUNIDADES (Alta prioridade)
# ══════════════════════════════════════════════════════════════════════════════

# Grupos de itens para o heatmap (type_ids representativos por grupo)
HEATMAP_ITEM_GROUPS = {
    # ── Minerais e Minérios ───────────────────────────────────────────────────
    "Minerais":       [34, 35, 36, 37, 38, 39, 40],          # Trit→Mega
    "Minerais Raros": [11399, 22, 1231, 1225, 16636, 16638,  # Morphite + Nullsec ores
                       28367, 28395, 28392, 28391],           # Compressed variants
    "Ice":            [16262, 16263, 16264, 16265,            # Blue Ice, Clear Icicle,
                       16266, 16267, 16268, 16269],           # Glacial Mass → Krystallos
    # ── PLEX e Economy ───────────────────────────────────────────────────────
    "PLEX/Injector":  [29668, 40520, 40519, 44992],           # PLEX, Injector, Extractor
    # ── Combustível ──────────────────────────────────────────────────────────
    "Isotopes":       [4051, 17887, 17889, 17888,             # He/N/H/O Isotopes
                       16273, 4247],                          # Strontium, Enriched Uranium
    # ── PI Materials ─────────────────────────────────────────────────────────
    "PI Tier 1 (P1)": [2267, 2288, 9828, 2305, 2287,         # Bacteria, Biofuels...
                       2309, 2393, 2311, 2312, 2307],
    "PI Tier 2 (P2)": [2317, 2321, 9830, 2324, 2327,         # Biocells, Constr. Blocks...
                       2332, 2346, 2351, 2366, 9836],
    "PI Tier 3 (P3)": [2411, 2395, 2403, 2413, 2415,         # Robotics, Guidance Sys...
                       2417, 2419, 2421, 2382, 2390],
    # ── Moon Ores ────────────────────────────────────────────────────────────
    "Moon Ore R4/R8": [45491, 45490, 45492, 45493,           # Zeolites, Sylvite...
                       45494, 45495, 45496, 45497],           # Cobaltite → Titanite
    "Moon Ore R16+":  [45498, 45499, 45500, 45501,           # Chromite → Vanadinite
                       45502, 45503, 45504, 45506],           # Carnotite → Xenotime
    # ── Naves ────────────────────────────────────────────────────────────────
    "Fragatas T1":    [587, 588, 582, 583, 596, 608,         # Rifter, Merlin...
                       638, 648],
    "Cruzadores T1":  [621, 623, 624, 620, 633, 634,         # Rupture, Stabber...
                       625, 626],
    "Fragatas T2":    [11176, 11178, 11196, 11200,           # Jaguar, Wolf, Hawk...
                       11172, 11174, 11182, 11184],
    # ── Módulos ──────────────────────────────────────────────────────────────
    "Módulos Escudo": [2488, 2456, 2281, 2605,               # Shield hardeners/extenders
                       2455, 2457, 2363],
    "Módulos Propuls":[12058, 12062, 5973, 5975,             # MWDs and ABs T2
                       3244, 4405, 4405],
    # ── Munição e Drones ─────────────────────────────────────────────────────
    "Munições T2":    [209, 210, 211, 212,                   # Heavy missiles
                       223, 224, 225, 226],                  # Light missiles
    "Drones T2":      [2185, 2183, 2456, 2454,               # Hobgoblin II, Ogre II...
                       3244, 23721, 23727, 23731],
    # ── Estruturas e Industria ────────────────────────────────────────────────
    "Estruturas":     [35825, 35826, 35835, 35836,           # Astrahus, Fortizar, Raitaru...
                       33474, 33475, 33476],                  # Deployables
}

# Hubs para o eixo X do heatmap
HEATMAP_HUBS = {
    "Jita":    {"region": "The Forge (Jita)",     "rid": 10000002},
    "Amarr":   {"region": "Domain (Amarr)",        "rid": 10000043},
    "Dodixie": {"region": "Sinq Laison (Dodixie)", "rid": 10000032},
    "Hek":     {"region": "Metropolis (Hek)",      "rid": 10000042},
    "Rens":    {"region": "Heimatar (Rens)",        "rid": 10000030},
}


def compute_heatmap_cell(rid: int, tids: list, broker_rel: int = 0,
                          accounting: int = 0) -> dict:
    """
    Calcula o score de oportunidade para uma celula do heatmap.
    Usa cache RAM populado por _warm_hist_cache em compute_full_heatmap.
    Sem I/O — apenas leituras do _hist_cache em memoria.
    """
    rb, rt = _calc_fee_rates(broker_rel, accounting)
    scores, margins, vols, trends = [], [], [], []
    best_score = -1
    best_tid   = None

    for tid in tids:
        hist = _get_hist_cached(rid, tid)          # cache RAM — O(1)
        if not hist or len(hist) < 14:
            continue

        # hist já ordenado por date (ORDER BY na query _db_bulk_load_region)
        hist_r  = hist[-30:]
        prices  = [float(d.get("average",  0)) for d in hist_r]
        highs   = [float(d.get("highest",  0)) for d in hist_r]
        lows    = [float(d.get("lowest",   0)) for d in hist_r]
        volumes = [int(d.get("volume",     0)) for d in hist_r]

        if not prices or prices[-1] <= 0:
            continue

        # Spread medio dos últimos 30 dias
        spreads = [(hi - lo) / hi * 100 if hi > 0 else 0
                   for hi, lo in zip(highs, lows)]
        avg_spread = sum(spreads) / len(spreads)

        # Margem liquida
        net_m = avg_spread - (rb + rb + rt) * 100

        # Volume score
        avg_vol = sum(volumes) / len(volumes) if volumes else 1
        recent  = sum(volumes[-7:]) / 7 if len(volumes) >= 7 else avg_vol
        vol_sc  = recent / avg_vol if avg_vol > 0 else 1

        # Tendencia (EMA14 vs EMA50)
        ema14_s = _cycle_ema(prices, min(14, len(prices)))
        ema50_s = _cycle_ema(prices, min(50, len(prices)))
        e14 = next((v for v in reversed(ema14_s) if v), prices[-1])
        e50 = next((v for v in reversed(ema50_s) if v), prices[-1])
        trend = (e14 - e50) / e50 if e50 > 0 else 0

        # Score individual (0-100)
        s = (min(40, max(0, net_m * 4)) +
             min(30, vol_sc * 15) +
             min(20, max(0, trend * 200 + 10)) +
             10)
        s = min(100, max(0, s))

        scores.append(s)
        margins.append(net_m)
        vols.append(vol_sc)
        trends.append(trend)

        if s > best_score:
            best_score = s
            best_tid   = tid

    if not scores:
        return {"score": 0, "net_margin": 0, "vol_score": 0,
                "trend": 0, "n_items": 0, "best_item_tid": None}

    return {
        "score":        round(sum(scores) / len(scores), 1),
        "net_margin":   round(sum(margins) / len(margins), 2),
        "vol_score":    round(sum(vols) / len(vols), 2),
        "trend":        round(sum(trends) / len(trends), 4),
        "n_items":      len(scores),
        "best_item_tid":best_tid,
    }


def compute_full_heatmap(broker_rel: int = 0, accounting: int = 0,
                          progress_cb=None) -> dict:
    """
    Calcula todas as celulas do heatmap (hub x grupo).
    Retorna {hub_name: {group_name: cell_dict}}.

    Otimizacao: pré-aquece o cache RAM para cada hub com 1 query SQL por hub
    antes de lançar as threads de calculo — as celulas fazem apenas lookups
    em memoria, sem I/O.
    """
    # Aquece o cache de todos os hubs antes de processar as células
    n_hubs = len(HEATMAP_HUBS)
    for idx_hub, (hub_name, hub_info) in enumerate(HEATMAP_HUBS.items(), 1):
        if progress_cb:
            progress_cb(f"Carregando histórico: {hub_name} ({idx_hub}/{n_hubs})...")
        _warm_hist_cache(hub_info["rid"])

    total   = len(HEATMAP_HUBS) * len(HEATMAP_ITEM_GROUPS)
    done    = [0]
    results = {hub: {} for hub in HEATMAP_HUBS}

    def _calc_cell(args):
        hub_name, group_name, rid, tids = args
        cell = compute_heatmap_cell(rid, tids, broker_rel, accounting)
        cell["hub"]   = hub_name
        cell["group"] = group_name
        return hub_name, group_name, cell

    work = [(hub_name, group_name, info["rid"], tids)
            for hub_name, info in HEATMAP_HUBS.items()
            for group_name, tids in HEATMAP_ITEM_GROUPS.items()]

    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as pool:
        for hub_name, group_name, cell in pool.map(_calc_cell, work):
            results[hub_name][group_name] = cell
            done[0] += 1
            if progress_cb:
                progress_cb((done[0], total,
                             f"Heatmap: {hub_name} / {group_name}..."))

    return results


# ══════════════════════════════════════════════════════════════════════════════
#  CATALOGO DE ITENS — enriquecimento com grupo e categoria (ESI)
# ══════════════════════════════════════════════════════════════════════════════

def fetch_item_details(type_id: int) -> dict:
    """
    Busca detalhes completos de um item via ESI:
    nome, group_id, group_name, category_id, category_name,
    volume, packaged_volume, description.
    Retorna dict ou {} se falhar.
    """
    try:
        data = esi_get(f"/universe/types/{type_id}/")
        if not data:
            return {}
        group_id   = data.get("group_id")
        group_name = ""
        cat_id     = None
        cat_name   = ""
        if group_id:
            gdata = esi_get(f"/universe/groups/{group_id}/")
            if gdata:
                group_name = gdata.get("name", "")
                cat_id     = gdata.get("category_id")
                if cat_id:
                    cdata = esi_get(f"/universe/categories/{cat_id}/")
                    if cdata:
                        cat_name = cdata.get("name", "")
        return {
            "name":          data.get("name", ""),
            "group_id":      group_id,
            "group_name":    group_name,
            "category_id":   cat_id,
            "category_name": cat_name,
            "volume":        float(data.get("packaged_volume")
                                   or data.get("volume") or 0.0),
            "description":   data.get("description", "")[:300],
        }
    except Exception as exc:
        log.debug("fetch_item_details(%s): %s", type_id, exc)
        return {}


def enrich_item_cache_groups(item_cache: dict, tids: list,
                              progress_cb=None) -> int:
    """
    Enriquece o item_cache com group_id/group_name/category_id/category_name
    para os type_ids da lista que ainda nao tem esses campos.
    Retorna o numero de itens atualizados.
    """
    sem_grupo = [tid for tid in tids
                 if tid in item_cache
                 and not item_cache[tid].get("group_id")]
    # Tambem itens ainda desconhecidos
    sem_nome  = [tid for tid in tids if tid not in item_cache]
    pendentes = list(set(sem_grupo + sem_nome))

    if not pendentes:
        return 0

    total = len(pendentes)
    done  = [0]

    def _enrich(tid):
        details = fetch_item_details(tid)
        if details:
            item_cache[tid] = details
        done[0] += 1
        if progress_cb and done[0] % 10 == 0:
            progress_cb((done[0], total,
                         f"Enriquecendo catalogo: {done[0]}/{total}..."))

    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as pool:
        list(pool.map(_enrich, pendentes))

    save_item_cache(item_cache)
    return len(pendentes)


def search_item_catalog(item_cache: dict, query: str = "",
                         filter_group: str = "",
                         filter_category: str = "") -> list:
    """
    Filtra o item_cache com base em query (busca no nome),
    grupo e categoria. Retorna lista de dicts ordenada por nome.
    """
    q   = query.strip().lower()
    grp = filter_group.strip().lower()
    cat = filter_category.strip().lower()

    results = []
    for tid, info in item_cache.items():
        name     = info.get("name", "")
        grp_name = info.get("group_name", "")
        cat_name = info.get("category_name", "")

        if q   and q   not in name.lower():
            continue
        if grp and grp not in grp_name.lower():
            continue
        if cat and cat not in cat_name.lower():
            continue

        results.append({
            "type_id":       tid,
            "name":          name,
            "group_id":      info.get("group_id", ""),
            "group_name":    grp_name,
            "category_id":   info.get("category_id", ""),
            "category_name": cat_name,
            "volume":        info.get("volume", 0.0),
        })

    results.sort(key=lambda x: x["name"].lower())
    return results


# ══════════════════════════════════════════════════════════════════════════════
#  STATION TRADER — scanner inspirado no EVE Trade Guru (módulo Station Trader)
# ══════════════════════════════════════════════════════════════════════════════
#
#  Colunas calculadas (equivalentes ao EVE Guru):
#   purchase_units    — unidades que cabem no capital × days_of_trade
#   daily_tradable    — média do volume diário negociado
#   buy_units_day     — volume do lado buy por dia
#   sell_units_day    — volume do lado sell por dia
#   bvs_ratio         — buy_vol / sell_vol (equilíbrio de mercado)
#   supply_units      — total à venda no mercado
#   dos               — days of supply = supply / sell_units_day
#   region_traded     — volume total negociado na região no período
#   share_region      — % do volume da região nesta estação
#   now_buy_price     — melhor buy order atual
#   now_sell_price    — melhor sell order atual
#   period_avg_price  — preço médio do período configurado
#   now_profit        — lucro líquido por unidade com preços atuais
#   period_profit     — lucro líquido por unidade com preço médio do período
#   capital_required  — capital necessário para comprar purchase_units
#   now_roi           — ROI com preços atuais (%)
#   period_roi        — ROI com preço médio do período (%)
#   now_order_margin  — margem bruta atual (%)
#   period_order_margin — margem bruta do período (%)
#   active_buy_orders — quantidade de ordens de compra abertas
#   active_sell_orders — quantidade de ordens de venda abertas
#   n_sellers         — número de vendedores distintos
#   price_vs_low      — preço atual vs mínimo do período (%)
#   price_vs_high     — preço atual vs máximo do período (%)

def station_trader_scan(
    orders: list,
    region_id: int,
    item_cache: dict,
    broker_rel: int = 0,
    accounting: int = 0,
    # Purchase Volume
    days_of_trade: float = 3.0,
    max_capital: float = 0.0,
    skip_expired: bool = True,
    # Profit Filters
    min_item_profit: float = 0.0,
    min_order_margin: float = 0.0,
    min_daily_tradable: float = 0.0,
    use_ranged_buy: bool = True,
    # Risk Profile
    avg_price_period: int = 14,
    min_period_roi: float = 0.0,
    min_bvs: float = 0.0,
    max_bvs: float = 9999.0,
    limit_buy_plow: float = 0.0,
    limit_sell_phigh: float = 0.0,
    # Blacklist
    blacklist: set = None,
    progress_cb=None,
) -> list:
    """
    Scanner de station trading — equivalente ao EVE Trade Guru Station Trader.

    OTIMIZAÇÃO CENTRAL (v14.1):
      Antes: 18747 queries SQL individuais (2 por item × N itens)
      Agora:  1 query SQL bulk (_db_bulk_load_region) + processamento em RAM

    Fluxo:
      1. Indexa ordens por type_id  (O(N_orders))
      2. Carrega TODO o histórico da região em 1 query SQL  (O(1) I/O)
      3. Processa cada item em paralelo (10 threads) — sem I/O nas threads
      4. Filtra e ordena os resultados por period_roi
    """
    rb, rt  = _calc_fee_rates(broker_rel, accounting)
    bl      = blacklist or set()
    lock    = threading.Lock()

    # ── Etapa 1: Indexa ordens por type_id ───────────────────────────────────
    if progress_cb:
        progress_cb("Indexando ordens...")

    by_type: dict = {}
    for o in orders:
        tid = int(o["type_id"])
        if tid in bl:
            continue
        is_buy = o["is_buy_order"]
        if isinstance(is_buy, str):
            is_buy = is_buy.strip().lower() == "true"
        by_type.setdefault(tid, {"sell": [], "buy": []})
        if is_buy:
            by_type[tid]["buy"].append(o)
        else:
            by_type[tid]["sell"].append(o)

    for entry in by_type.values():
        entry["sell"].sort(key=lambda o: float(o["price"]))
        entry["buy"].sort(key=lambda o: float(o["price"]), reverse=True)

    # ── Etapa 2: Carrega HISTÓRICO COMPLETO em 1 query SQL ────────────────────
    # Esta é a otimização central: elimina 18747 queries individuais
    if progress_cb:
        progress_cb("Carregando histórico do banco (1 query)...")

    bulk_hist = _db_bulk_load_region(region_id)
    n_hist    = len(bulk_hist)

    if progress_cb:
        progress_cb(f"Histórico: {n_hist:,} itens em memória. Analisando...")

    # ── Etapa 3: Processa cada item em paralelo ───────────────────────────────
    total   = len(by_type)
    done    = [0]
    results = []

    def _process(tid):
        entry = by_type[tid]
        sells = entry["sell"]
        buys  = entry["buy"]
        if not sells or not buys:
            return None

        sell_min = float(sells[0]["price"])
        buy_max  = float(buys[0]["price"])
        if sell_min <= buy_max or sell_min <= 0:
            return None

        # Histórico da RAM (sem I/O) — O(1) lookup
        hist = bulk_hist.get(tid)
        if not hist:
            return None

        # Últimos N dias do histórico (já ordenado por date pelo ORDER BY da query)
        hist_p = hist[-avg_price_period:] if len(hist) >= avg_price_period else hist
        if not hist_p:
            return None

        prices_p = [d["average"] for d in hist_p]
        vols_p   = [d["volume"]  for d in hist_p]
        highs_p  = [d["highest"] for d in hist_p]
        lows_p   = [d["lowest"]  for d in hist_p]

        n_p          = len(hist_p)
        avg_price_p  = sum(prices_p) / n_p
        p_high       = max(highs_p)
        p_low        = min(lows_p)
        daily_vol    = sum(vols_p)   / n_p
        region_traded= sum(vols_p)

        if daily_vol < min_daily_tradable:
            return None

        # Métricas de ordens
        supply_units  = sum(int(o.get("volume_remain", 0)) for o in sells)
        n_buy_orders  = len(buys)
        n_sell_orders = len(sells)
        n_sellers     = len(set(str(o.get("location_id","")) for o in sells))

        bvs = n_buy_orders / n_sell_orders if n_sell_orders > 0 else 1.0
        if not (min_bvs <= bvs <= max_bvs):
            return None

        sell_units_day = daily_vol
        dos            = supply_units / sell_units_day if sell_units_day > 0 else 999

        # Limites de preço
        buy_limit  = p_low  * (1 + limit_buy_plow   / 100) if limit_buy_plow  != 0 else buy_max
        sell_limit = p_high * (1 + limit_sell_phigh  / 100) if limit_sell_phigh!= 0 else sell_min
        effective_buy  = min(buy_max,  buy_limit)
        effective_sell = max(sell_min, sell_limit)

        # Margem bruta (filtro de seleção)
        now_margin = (sell_min - buy_max) / sell_min * 100
        if now_margin < min_order_margin:
            return None

        # Lucros
        revenue_now     = sell_min      * (1 - rb - rt)
        cost_now        = effective_buy * (1 + rb)
        now_profit_u    = revenue_now - cost_now
        now_roi         = now_profit_u / cost_now * 100 if cost_now > 0 else 0

        revenue_per     = avg_price_p   * (1 - rb - rt)
        cost_per        = effective_buy * (1 + rb)
        period_profit_u = revenue_per - cost_per
        period_roi      = period_profit_u / cost_per * 100 if cost_per > 0 else 0
        period_margin   = (avg_price_p - buy_max) / avg_price_p * 100 if avg_price_p > 0 else 0

        if period_roi < min_period_roi:
            return None

        # Purchase units
        cap_per_day    = (max_capital / days_of_trade) if max_capital > 0 else 1e18
        purchase_units = int(min(
            cap_per_day / (effective_buy * (1 + rb)) if effective_buy > 0 else 0,
            supply_units,
            daily_vol * days_of_trade,
        ))
        if purchase_units <= 0:
            return None

        capital_req = purchase_units * effective_buy * (1 + rb)
        item_profit = purchase_units * now_profit_u
        if item_profit < min_item_profit:
            return None

        # Preço vs histórico
        price_vs_low  = (sell_min - p_low)  / p_low  * 100 if p_low  > 0 else 0
        price_vs_high = (sell_min - p_high) / p_high * 100 if p_high > 0 else 0

        # Alertas de risco (prioridade: preço > volume > desequilíbrio)
        risk_flag = ""
        if sell_min > p_high * 1.20:
            risk_flag = "⚠ Preco acima do teto historico"
        elif buy_max < p_low * 0.80:
            risk_flag = "⚠ Buy abaixo do piso historico"
        elif dos > 30:
            risk_flag = "⚠ D.O.S. alto — giro lento"
        elif bvs < 0.5:
            risk_flag = "⚠ Muito mais sell do que buy"
        elif (now_profit_u > 0 and period_profit_u > 0
              and now_profit_u > period_profit_u * 3.0):
            risk_flag = "⚠ Now Profit muito acima da media"

        info = item_cache.get(tid, {})
        return {
            "type_id":          tid,
            "name":             info.get("name", str(tid)),
            "group_name":       info.get("group_name", ""),
            "category_name":    info.get("category_name", ""),
            "purchase_units":   purchase_units,
            "daily_tradable":   round(daily_vol, 1),
            "buy_units_day":    round(daily_vol, 1),
            "sell_units_day":   round(daily_vol, 1),
            "bvs_ratio":        round(bvs, 2),
            "supply_units":     supply_units,
            "dos":              round(dos, 1),
            "region_traded":    int(region_traded),
            "share_region":     100.0,
            "now_buy_price":    round(buy_max, 2),
            "now_sell_price":   round(sell_min, 2),
            "period_avg_price": round(avg_price_p, 2),
            "now_profit":       round(now_profit_u, 2),
            "period_profit":    round(period_profit_u, 2),
            "item_profit_total":round(item_profit, 0),
            "capital_required": round(capital_req, 0),
            "now_roi":          round(now_roi, 2),
            "period_roi":       round(period_roi, 2),
            "now_margin":       round(now_margin, 2),
            "period_margin":    round(period_margin, 2),
            "n_buy_orders":     n_buy_orders,
            "n_sell_orders":    n_sell_orders,
            "n_sellers":        n_sellers,
            "p_low":            round(p_low, 2),
            "p_high":           round(p_high, 2),
            "price_vs_low":     round(price_vs_low, 1),
            "price_vs_high":    round(price_vs_high, 1),
            "risk_flag":        risk_flag,
            "avg_price_period": avg_price_period,
        }

    # Processamento paralelo — threads fazem apenas cálculos em RAM (sem I/O)
    chunk = max(1, total // 40)   # progresso ~40 atualizações ao longo do scan
    with concurrent.futures.ThreadPoolExecutor(max_workers=20) as pool:
        futures = {pool.submit(_process, tid): tid for tid in by_type}
        for fut in concurrent.futures.as_completed(futures):
            with lock:
                done[0] += 1
                n_done = done[0]
            if progress_cb and n_done % chunk == 0:
                progress_cb((n_done, total,
                             f"Station Trader: {n_done:,}/{total:,} itens..."))
            r = fut.result()
            if r:
                results.append(r)

    # ── Etapa 4: Ordena e retorna ──────────────────────────────────────────────
    results.sort(key=lambda x: x["period_roi"], reverse=True)
    if progress_cb:
        progress_cb(f"{len(results):,} oportunidades encontradas.")
    return results


# ── Tabela de tipos do SDE (nome + volume em lote) ────────────────────────────
# Evita 1 GET /universe/types/{id}/ por item: um unico download (~3 MB gzip)
# cobre todos os tipos. Itens ausentes da tabela (recem-lancados) ainda caem
# no caminho ESI de fetch_type_ids_bulk.
SDE_TYPES_URL     = "https://www.fuzzwork.co.uk/dump/latest/csv/invTypes.csv"
SDE_TYPES_PATH    = os.path.join(DATA_DIR, "invTypes.csv")
SDE_TYPES_MAX_AGE = 30 * 86400   # re-baixa se o arquivo tiver mais de 30 dias
_sde_types: dict | None = None
_sde_types_lock = threading.Lock()


def _download_sde_types() -> bool:
    """Baixa invTypes.csv para SDE_TYPES_PATH (escrita atomica). True se ok."""
    tmp = SDE_TYPES_PATH + ".part"
    try:
        with requests.get(SDE_TYPES_URL, stream=True, timeout=(10, 300)) as r:
            r.raise_for_status()
            with open(tmp, "wb") as f:
                for block in r.iter_content(chunk_size=1 << 16):
                    f.write(block)
        os.replace(tmp, SDE_TYPES_PATH)
        return True
    except (requests.RequestException, OSError) as e:
        log.warning("Falha ao baixar tabela de tipos do SDE: %s", e)
        if os.path.exists(tmp):
            os.remove(tmp)
        return False


def load_sde_types() -> dict:
    """Retorna {type_id: {"name", "volume"}} do SDE; {} se indisponivel.

    volume = packagedVolume quando existir (igual a preferencia da ESI).
    Resultado fica em memoria; o arquivo e baixado no primeiro uso.
    """
    global _sde_types
    with _sde_types_lock:
        if _sde_types is not None:
            return _sde_types
        stale = (not os.path.exists(SDE_TYPES_PATH) or
                 time.time() - os.path.getmtime(SDE_TYPES_PATH) > SDE_TYPES_MAX_AGE)
        if stale:
            _download_sde_types()   # se falhar, usa o arquivo antigo (se houver)
        types = {}
        try:
            with open(SDE_TYPES_PATH, newline="", encoding="utf-8-sig") as f:
                for row in csv.DictReader(f):
                    try:
                        vol = float(row.get("packagedVolume") or row.get("volume") or 0.0)
                        types[int(row["typeID"])] = {"name": row["typeName"], "volume": vol}
                    except (KeyError, ValueError):
                        continue
        except OSError as e:
            log.warning("Tabela de tipos do SDE indisponivel: %s", e)
        _sde_types = types
        return types


def fetch_type_ids_bulk(ids):
    """Resolve {type_id: {"name", "volume"}} — SDE primeiro, ESI so p/ ausentes."""
    if not ids:
        return {}
    sde    = load_sde_types()
    found  = {t: dict(sde[t]) for t in set(ids) if t in sde}
    unique = [t for t in set(ids) if t not in sde]
    if not unique:
        return found
    log.info("fetch_type_ids_bulk: %d via SDE, %d via ESI", len(found), len(unique))
    names_map = {}
    chunks = [unique[i:i+1000] for i in range(0, len(unique), 1000)]

    def _fetch_names_chunk(chunk):
        items = esi_post("/universe/names/", chunk)
        return {item["id"]: item["name"] for item in items
                if item.get("category") == "inventory_type"}

    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as pool:
        for partial in pool.map(_fetch_names_chunk, chunks):
            names_map.update(partial)

    if not names_map:
        return found

    def _fetch_volume(tid):
        data = esi_get(f"/universe/types/{tid}/")
        if data:
            vol = data.get("packaged_volume") or data.get("volume") or 0.0
            return tid, float(vol)
        return tid, 0.0

    result = found
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as pool:
        for tid, vol in pool.map(_fetch_volume, list(names_map.keys())):
            result[tid] = {"name": names_map[tid], "volume": vol}
    return result


# ── Cache de pulos (ESI route) ────────────────────────────────────────────────
_sys_id_cache = {}
_route_cache  = {}

def get_system_id_cached(location_name):
    sys_guess = str(location_name).split(" - ")[0].split(" ")[0]
    if sys_guess in _sys_id_cache:
        return _sys_id_cache[sys_guess]
    try:
        r = requests.post(f"{ESI_BASE}/universe/ids/", json=[sys_guess], timeout=10)
        if r.status_code == 200:
            data = r.json()
            if "systems" in data:
                sid = data["systems"][0]["id"]
                _sys_id_cache[sys_guess] = sid
                return sid
    except Exception as _e:

        log.debug("silenced: %s", _e)
    return None


def get_jumps_cached(origin, dest):
    if origin == dest:
        return 0
    id_orig = get_system_id_cached(origin)
    id_dest = get_system_id_cached(dest)
    if not id_orig or not id_dest:
        return 999
    pair = tuple(sorted((id_orig, id_dest)))
    if pair in _route_cache:
        return _route_cache[pair]
    try:
        r = requests.get(
            f"{ESI_BASE}/route/{id_orig}/{id_dest}/?flag=shortest", timeout=10)
        if r.status_code == 200:
            jumps = len(r.json())
            _route_cache[pair] = jumps
            return jumps
    except Exception as _e:

        log.debug("silenced: %s", _e)
    return 999


# ── Cache persistente ─────────────────────────────────────────────────────────

def load_item_cache():
    if not os.path.exists(ITEM_CACHE_FILE):
        return {}
    try:
        with open(ITEM_CACHE_FILE, encoding="utf-8") as f:
            raw = json.load(f)
        return {int(k): v for k, v in raw.items()}
    except Exception as exc:
        log.warning("Erro ao carregar cache de itens: %s", exc)
        return {}


def save_item_cache(cache):
    try:
        with open(ITEM_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump({str(k): v for k, v in cache.items()}, f)
    except Exception as exc:
        log.warning("Erro ao salvar cache de itens: %s", exc)


# ── Config persistente (M6) ───────────────────────────────────────────────────

_CONFIG_DEFAULTS = {
    # Aba Rotas
    "r_roi":   "10",  "r_ltmin": "1000000", "r_cmin": "0",
    "r_cmax":  "0",   "r_vol":   "0",
    # Aba Arbitragem
    "a_bbuy":  "2.50","a_bsell": "3.50","a_stax":  "4.00",
    "a_fret":  "100", "a_mmin":  "0",   "a_ltmin": "1000000",
    # Otimizador
    "o_cap":   "1600000","o_lucro": "10000","o_marg":  "1.5",
    "o_ispulo":"10000",  "o_pulos": "15",
    # Skills globais (Tier 1 — Configuracoes)
    "g_brel":  "0",   "g_acc":   "0",
    # AI
    "ai_key":  "",
    "experimental": "0",   # Previsão / AI Analista / Conviction (sem backtest)
}

EXPERIMENTAL_NOTICE = ("EXPERIMENTAL — sem validação estatística (backtest pendente). "
                       "Use como triagem, não como base de decisão.")


def experimental_enabled(cfg: dict | None = None) -> bool:
    """Recursos sem backtest (Previsão, AI Analista, Conviction) estão ligados?"""
    cfg = cfg if cfg is not None else load_config()
    return str(cfg.get("experimental", "0")).strip().lower() in ("1", "true", "yes")


def load_config() -> dict:
    if not os.path.exists(CONFIG_FILE):
        return dict(_CONFIG_DEFAULTS)
    try:
        with open(CONFIG_FILE, encoding="utf-8") as f:
            data = json.load(f)
        cfg = dict(_CONFIG_DEFAULTS)
        cfg.update({k: str(v) for k, v in data.items() if k in cfg})
        return cfg
    except Exception as exc:
        log.warning("Erro ao carregar config: %s", exc)
        return dict(_CONFIG_DEFAULTS)

def save_config(cfg: dict):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception as exc:
        log.warning("Erro ao salvar config: %s", exc)


def load_station_cache():
    if not os.path.exists(STATION_CACHE_FILE):
        return {}
    result = {}
    try:
        with open(STATION_CACHE_FILE, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                try:
                    result[int(row["location_id"])] = row["station_name"]
                except Exception as _e:

                    log.debug("silenced: %s", _e)
    except Exception as exc:
        log.error("Erro ao carregar cache de estacoes: %s", exc)
    return result


# ── Download de ordens ────────────────────────────────────────────────────────

def _fetch_page(base_url, page):
    """
    Baixa uma página de ordens ESI com suporte a ETag (If-None-Match).
    Se a página não mudou desde o último download, retorna o cache sem
    trafegar os dados novamente — economia de ~60-80% em syncs repetidas.
    """
    url       = f"{base_url}?datasource=tranquility&order_type=all&page={page}"
    cache_key = url
    headers   = dict(_ESI_HEADERS)
    if cache_key in _etag_cache:
        etag, _ = _etag_cache[cache_key]
        headers["If-None-Match"] = etag
    try:
        r = _get_session().get(url, headers=headers, timeout=10)
        if r.status_code == 304:
            # Not Modified — retorna dado em cache sem baixar novamente
            return _etag_cache[cache_key][1]
        if r.status_code == 200:
            data = r.json()
            etag = r.headers.get("ETag")
            if etag:
                _etag_cache[cache_key] = (etag, data)
            return data
    except Exception as _e:
        log.debug("silenced: %s", _e)
    return []


def download_region_orders(region_id, region_name, progress_cb=None):
    """
    Baixa todas as ordens de mercado de uma região via ESI.
    Cada página usa ETag (If-None-Match): páginas que não mudaram desde
    o último download retornam 304 Not Modified e são servidas do cache
    sem trafegar os dados — redução de 60-80% do tráfego em syncs repetidas.
    """
    base_url = f"{ESI_BASE}/markets/{region_id}/orders/"
    try:
        init_req = _get_session().get(
            f"{base_url}?datasource=tranquility&order_type=all&page=1",
            headers=_ESI_HEADERS, timeout=10)
        pages = int(init_req.headers.get("X-Pages", 1))
        if progress_cb:
            progress_cb(f"[{region_name}] {pages} paginas detectadas.")
    except Exception as exc:
        msg = f"Erro ao conectar em {region_name}: {exc}"
        log.error(msg)
        if progress_cb:
            progress_cb(f"  {msg}")
        return []

    all_orders = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        futures = {executor.submit(_fetch_page, base_url, p): p
                   for p in range(1, pages + 1)}
        for future in concurrent.futures.as_completed(futures):
            data = future.result()
            if data:
                all_orders.extend(data)

    if all_orders:
        fname = _orders_csv_path(region_name)
        try:
            if HAS_PANDAS:
                pd.DataFrame(all_orders).to_csv(fname, index=False)
            else:
                keys = list(all_orders[0].keys())
                with open(fname, "w", newline="", encoding="utf-8") as f:
                    w = csv.DictWriter(f, fieldnames=keys)
                    w.writeheader()
                    w.writerows(all_orders)
            if progress_cb:
                progress_cb(f"[{region_name}] {len(all_orders):,} ordens salvas.")
        except Exception as exc:
            log.error("Erro ao salvar CSV de %s: %s", region_name, exc)
    return all_orders


def _orders_csv_path(region_name):
    return os.path.join(DATA_DIR,
                        f"market_orders_{region_name.replace(' ', '_')}.csv")


def load_orders_from_csv(region_name, force_reload=False):
    """
    Carrega ordens do CSV com cache RAM de 30 minutos (TTL = _ORDERS_TTL).
    Sem force_reload: retorna cache se ainda fresco, relê o disco se expirado.
    force_reload=True: ignora o TTL e sempre relê do disco.
    """
    if not force_reload:
        entry = _orders_mem_cache.get(region_name)
        if entry is not None:
            ts, cached = entry
            if time.time() - ts < _ORDERS_TTL:
                return cached
    fname = _orders_csv_path(region_name)
    if not os.path.exists(fname):
        return []
    orders = []
    skipped = 0
    try:
        with open(fname, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                try:
                    row["type_id"]       = int(row["type_id"])
                    row["price"]         = float(row["price"])
                    row["volume_remain"] = int(row["volume_remain"])
                    row["is_buy_order"]  = str(row["is_buy_order"]).strip().lower() == "true"
                    row["location_id"]   = int(row["location_id"])
                    orders.append(row)
                except (KeyError, TypeError, ValueError):
                    skipped += 1
    except (OSError, csv.Error) as exc:
        log.error("Erro ao ler CSV de %s: %s", region_name, exc)
    if skipped:
        log.warning("CSV de %s: %d linhas invalidas ignoradas (%d validas)",
                    region_name, skipped, len(orders))
    _orders_mem_cache[region_name] = (time.time(), orders)
    return orders


def cache_file_info(region_name):
    fname = _orders_csv_path(region_name)
    if not os.path.exists(fname):
        return "--", COLORS["subtext"]
    mtime = datetime.fromtimestamp(os.path.getmtime(fname))  # uma única chamada
    age   = datetime.now() - mtime
    label = mtime.strftime("%d/%m %H:%M")
    if age < timedelta(hours=1):
        return label, COLORS["green"]
    if age < timedelta(hours=24):
        return label, COLORS["yellow"]
    return label, COLORS["red"]


# ── Cache de estacoes ─────────────────────────────────────────────────────────

def build_station_cache(region_names, progress_cb=None):
    all_ids = set()
    for rname in region_names:
        fname = _orders_csv_path(rname)
        if not os.path.exists(fname):
            continue
        try:
            with open(fname, newline="", encoding="utf-8") as f:
                for row in csv.DictReader(f):
                    try:
                        all_ids.add(int(row["location_id"]))
                    except Exception as _e:

                        log.debug("silenced: %s", _e)
        except Exception as exc:
            log.error("Erro ao ler IDs de %s: %s", rname, exc)

    if not all_ids:
        if progress_cb:
            progress_cb("Nenhum ID encontrado -- baixe os dados primeiro.")
        return {}

    ids_npc       = [i for i in all_ids if i <= MAX_INT32]
    ids_structure = [i for i in all_ids if i >  MAX_INT32]

    if progress_cb:
        progress_cb(f"IDs: {len(all_ids)} total | {len(ids_npc)} NPC | "
                    f"{len(ids_structure)} estruturas")

    final_map = {}
    for i in range(0, len(ids_npc), 500):
        chunk = ids_npc[i:i+500]
        try:
            items = esi_post("/universe/names/", chunk)
            final_map.update({item["id"]: item["name"] for item in items})
        except Exception as exc:
            log.error("Batch de nomes falhou: %s", exc)
        if progress_cb:
            progress_cb(f"Nomes NPC: {len(final_map)}/{len(ids_npc)}...")
        time.sleep(0.1)

    for sid in ids_structure:
        final_map[sid] = f"Estrutura Privada {sid}"

    try:
        with open(STATION_CACHE_FILE, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["location_id", "station_name"])
            w.writeheader()
            w.writerows({"location_id": k, "station_name": v}
                         for k, v in final_map.items())
    except Exception as exc:
        log.error("Erro ao salvar cache de estacoes: %s", exc)

    if progress_cb:
        progress_cb(f"Cache salvo: {len(final_map)} estacoes.")
    return final_map


# ══════════════════════════════════════════════════════════════════════════════
#  LOGICA DE ROTAS — ABA 1
# ══════════════════════════════════════════════════════════════════════════════







def resolve_item_names(routes, item_cache, progress_cb=None):
    ids_needed = [r["type_id"] for r in routes
                  if str(r["type_id"]) == r["name"]
                  and r["type_id"] not in item_cache]
    if not ids_needed:
        return item_cache
    if progress_cb:
        progress_cb(f"Resolvendo nomes de {len(ids_needed)} itens...")
    new_info = fetch_type_ids_bulk(ids_needed)
    item_cache.update(new_info)
    for r in routes:
        tid = r["type_id"]
        if tid in item_cache:
            r["name"]         = item_cache[tid].get("name", str(tid))
            r["item_volume"]  = float(item_cache[tid].get("volume", 0))
            r["total_volume"] = r["item_volume"] * r.get("vol_realizavel", r["volume_avail"])
    save_item_cache(item_cache)
    return item_cache


def enrich_routes_lucro_dia(routes, dst_region_id, progress_cb=None):
    if not routes:
        return
    if progress_cb:
        progress_cb(f"Buscando historico de {len(routes)} itens para Lucro/Dia...")
    _warm_hist_cache(dst_region_id)   # 1 query SQL antes das threads

    def _fetch(r):
        h = _get_hist_cached(dst_region_id, r["type_id"])   # cache RAM — O(1)
        ld, gd = calc_lucro_dia(h, r["profit_unit"])
        r["lucro_dia"] = ld
        r["giro_dia"]  = gd
        r["hist_erro"] = hist_failed(dst_region_id, r["type_id"])

    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as pool:
        list(pool.map(_fetch, routes))
    if progress_cb:
        progress_cb("Lucro/Dia calculado."
                    + _hist_err_note(dst_region_id, [r["type_id"] for r in routes]))


# ══════════════════════════════════════════════════════════════════════════════
#  LOGICA DE ARBITRAGEM — ABA 2
# ══════════════════════════════════════════════════════════════════════════════



def run_arbitrage(src_orders, dst_orders, src_name, dst_name,
                  tax_broker_buy, tax_broker_sell, tax_sales,
                  custo_frete_m3, station_cache, item_cache,
                  min_margem=0.0, min_lucro_total=0.0,
                  progress_cb=None):
    """Arbitragem (resolve nomes aqui; calculo em evetrade.core.routes)."""
    tids = ({int(o["type_id"]) for o in src_orders}
            & {int(o["type_id"]) for o in dst_orders})
    _ensure_names(list(tids), item_cache, progress_cb)  # 2B
    return arbitrage_opportunities(
        src_orders, dst_orders, src_name, dst_name,
        tax_broker_buy, tax_broker_sell, tax_sales, custo_frete_m3,
        station_cache, item_cache, min_margem, min_lucro_total, progress_cb)


def enrich_arb_lucro_dia(opportunities, dst_region_id, progress_cb=None):
    if not opportunities:
        return
    if progress_cb:
        progress_cb(f"Buscando historico de {len(opportunities)} itens para Lucro/Dia...")
    _warm_hist_cache(dst_region_id)   # 1 query SQL antes das threads

    tids = list({o["type_id"] for o in opportunities})

    def _fetch(tid):
        return tid, _get_hist_cached(dst_region_id, tid)   # cache RAM — O(1)

    hist_map = {}
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as pool:
        for tid, h in pool.map(_fetch, tids):
            hist_map[tid] = h

    for opp in opportunities:
        h = hist_map.get(opp["type_id"], [])
        ld, gd = calc_lucro_dia(h, opp["lucro_liq_unit"])
        opp["lucro_dia"] = ld
        opp["giro_dia"]  = gd
        opp["hist_erro"] = hist_failed(dst_region_id, opp["type_id"])

    if progress_cb:
        progress_cb("Lucro/Dia calculado." + _hist_err_note(dst_region_id, tids))


# ══════════════════════════════════════════════════════════════════════════════
#  LOGICA DO OTIMIZADOR — ABA 3
# ══════════════════════════════════════════════════════════════════════════════

def solve_cargo(opportunities, capacity_m3, min_lucro, min_margem,
                min_isk_pulo, max_pulos, filtro_origem, filtro_destino,
                progress_cb=None):
    if not HAS_ORTOOLS:
        return [], {"erro": "ortools nao instalado. Execute: pip install ortools"}

    items = list(opportunities)

    if filtro_origem.strip():
        items = [x for x in items if filtro_origem.lower() in x["loc_origem"].lower()]
    if filtro_destino.strip():
        items = [x for x in items if filtro_destino.lower() in x["loc_destino"].lower()]
    items = [x for x in items if x["roi_total"] >= min_lucro]
    items = [x for x in items if x["margem_pct"] >= min_margem]

    if not items:
        return [], {"erro": "Nenhum item apos filtros."}

    if progress_cb:
        pairs_count = len({(x["loc_origem"], x["loc_destino"]) for x in items})
        progress_cb(f"Calculando pulos para {pairs_count} pares de estacoes (paralelo)...")

    pairs_set = list({(x["loc_origem"], x["loc_destino"]) for x in items})

    # Pulos em paralelo -- cada par e uma chamada ESI independente
    def _calc_pair(pair):
        o, d = pair
        return pair, get_jumps_cached(o, d)

    jumps_map = {}
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as pool:
        for pair, j in pool.map(_calc_pair, pairs_set):
            jumps_map[pair] = j

    for x in items:
        x["pulos"] = jumps_map.get((x["loc_origem"], x["loc_destino"]), 999)

    items = [x for x in items if x["pulos"] <= max_pulos]
    for x in items:
        p = x["pulos"]
        x["isk_pulo"] = (x["roi_total"] / p) if p > 0 else x["roi_total"]
    items = [x for x in items if x["isk_pulo"] >= min_isk_pulo]

    if not items:
        return [], {"erro": "Nenhum item apos filtro de pulos/ISK por pulo."}

    if progress_cb:
        progress_cb(f"Resolvendo knapsack com {len(items)} itens candidatos...")

    items_s = sorted(items,
                     key=lambda x: x["lucro_liq_unit"] / max(x["volume_m3"], 0.001),
                     reverse=True)[:500]

    solver = pywraplp.Solver.CreateSolver("SCIP")
    if not solver:
        return [], {"erro": "Solver SCIP nao disponivel."}

    # Estrutura unica: lista de (var, vol, lucro, row) -- sem 4 listas paralelas dessincronizadas
    solver_items = []
    for i, row in enumerate(items_s):
        max_q = int(row["qtd_origem"])
        if max_q > 0:
            var = solver.IntVar(0, max_q, f"x_{i}")
            solver_items.append((var, row["volume_m3"], row["lucro_liq_unit"], row))

    solver.Add(
        solver.Sum([var * vol   for var, vol, _,   _ in solver_items]) <= capacity_m3)
    solver.Maximize(
        solver.Sum([var * lucro for var, _,   lucro, _ in solver_items]))

    status = solver.Solve()
    plan   = []

    if status == pywraplp.Solver.OPTIMAL:
        for var, _, _, row in solver_items:
            qtd = int(var.solution_value())
            if qtd > 0:
                lucro_tot = qtd * row["lucro_liq_unit"]
                p         = row["pulos"]
                plan.append({
                    "Item":           row["item"],
                    "Origem":         row["loc_origem"],
                    "Destino":        row["loc_destino"],
                    "Pulos":          int(p),
                    "Qtd":            qtd,
                    "Vol Total (m3)": round(qtd * row["volume_m3"], 2),
                    "Lucro Total":    round(lucro_tot, 2),
                    "ISK/Pulo":       round(lucro_tot / p if p > 0 else lucro_tot, 2),
                    "Margem (%)":     round(row["margem_pct"], 2),
                })
    elif progress_cb:
        progress_cb("Solver nao encontrou solucao otima com os filtros atuais.")

    vol_total   = sum(r["Vol Total (m3)"] for r in plan)
    lucro_total = sum(r["Lucro Total"]    for r in plan)
    stats = {
        "n_items":      len(plan),
        "vol_total":    vol_total,
        "capacity":     capacity_m3,
        "lucro_total":  lucro_total,
        "margem_media": round(sum(r["Margem (%)"] for r in plan) / len(plan), 2) if plan else 0,
    }
    return plan, stats


# ══════════════════════════════════════════════════════════════════════════════
#  INTERFACE
# ══════════════════════════════════════════════════════════════════════════════

def _apply_styles():
    s = ttk.Style()
    s.theme_use("default")
    s.configure("EVE.Treeview",
                 background=COLORS["row_odd"], foreground=COLORS["text"],
                 fieldbackground=COLORS["row_odd"], rowheight=22,
                 font=("Courier", 11), borderwidth=0)
    s.configure("EVE.Treeview.Heading",
                 background=COLORS["border"], foreground=COLORS["header"],
                 font=("Courier", 11, "bold"), relief="flat")
    s.map("EVE.Treeview",
          background=[("selected", COLORS["border"])],
          foreground=[("selected", COLORS["accent"])])
    s.configure("TCombobox",
                 fieldbackground=COLORS["border"], background=COLORS["border"],
                 foreground=COLORS["text"], arrowcolor=COLORS["accent"],
                 selectbackground=COLORS["border"], selectforeground=COLORS["accent"])
    s.configure("TNotebook", background=COLORS["bg"], borderwidth=0)
    s.configure("TNotebook.Tab",
                 background=COLORS["border"], foreground=COLORS["subtext"],
                 font=("Courier", 11, "bold"), padding=(12, 4))
    s.map("TNotebook.Tab",
          background=[("selected", COLORS["tab_sel"])],
          foreground=[("selected", COLORS["accent"])])




# ══════════════════════════════════════════════════════════════════════════════
#  TOOLTIP — exibe dica ao passar o mouse sobre qualquer widget
# ══════════════════════════════════════════════════════════════════════════════

class ToolTip:
    """Tooltip leve para qualquer widget tkinter.
    Uso: ToolTip(widget, "Texto explicativo aqui")
    """
    _BG   = "#1e2433"
    _FG   = "#c8d0e0"
    _FONT = ("Courier", 10)
    _DELAY = 600   # ms antes de exibir

    def __init__(self, widget, text: str):
        self.widget = widget
        self.text   = text
        self._id    = None
        self._tw    = None
        widget.bind("<Enter>",   self._schedule, add="+")
        widget.bind("<Leave>",   self._cancel,   add="+")
        widget.bind("<Button>",  self._cancel,   add="+")

    def _schedule(self, _e=None):
        self._cancel()
        self._id = self.widget.after(self._DELAY, self._show)

    def _cancel(self, _e=None):
        if self._id:
            self.widget.after_cancel(self._id)
            self._id = None
        if self._tw:
            self._tw.destroy()
            self._tw = None

    def _show(self):
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self._tw = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.configure(bg=self._BG)
        tk.Label(tw, text=self.text, bg=self._BG, fg=self._FG,
                 font=self._FONT, relief="flat", bd=0,
                 padx=8, pady=4, wraplength=300,
                 justify="left").pack()
        # Auto-fecha após 4 segundos
        tw.after(4000, self._cancel)


class EVETradeApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("EVE Trade Finder  v14")
        self.configure(bg=COLORS["bg"])
        self.geometry("1600x960")
        self.resizable(True, True)

        self._cfg           = load_config()
        self._experimental  = tk.BooleanVar(value=experimental_enabled(self._cfg))
        self._routes        = []
        self._arb_opps      = []
        self._opt_pool      = []
        self._opt_plan      = []
        self._item_cache    = load_item_cache()
        self._station_cache = load_station_cache()
        self._blacklist     = load_blacklist()          # Tier 1
        self._loading_rotas  = False
        self._loading_arb   = False
        self._loading_opt   = False
        # abas de analise
        self._mm_results    = []
        self._sc_results    = []
        self._ds_results    = []
        self._pm_results    = []
        self._bud_results   = []
        self._bud_stats     = {}
        self._loading_mm    = False
        self._loading_sc    = False
        self._loading_ds    = False
        self._loading_pm    = False
        self._loading_bud   = False
        # mercado live
        self._live_data     = {}
        self._live_running  = False
        self._live_region   = "The Forge (Jita)"
        self._live_after_id = None
        self._live_items    = list(MARKET_ITEMS)
        self._live_cur_tid  = None  # tipo selecionado (robusto para Linux/GTK)
        self._wmap: dict    = {}    # {id(StringVar): tk.Widget} — tooltip lookup
        self._live_sort_col = "score"
        self._live_sort_rev = True
        # watchlist
        self._watchlist     = load_watchlist()
        self._wl_running    = False
        self._wl_after_id   = None
        # paper trading
        self._paper         = load_paper_trades()
        self._paper_pos_sel = None
        # AI analista
        # API Key da Anthropic: env > keyring seguro > config.json legado.
        # Chave legada em config.json é migrada p/ o keyring e apagada do arquivo
        # (sem keyring seguro, fica só em memória). Nunca volta a ser escrita lá.
        purge_plaintext_keyring_entry()
        _legacy_key = self._cfg.get("ai_key", "")
        if _legacy_key and HAS_KEYRING and not keyring_load_key():
            keyring_save_key(_legacy_key)
        _kr_key = os.environ.get(API_KEY_ENV) or keyring_load_key() or _legacy_key
        self._ai_api_key    = tk.StringVar(value=_kr_key)
        self._ai_last       = {}
        self._loading_ai    = False
        # comparador de hubs
        self._hub_results   = []
        self._loading_hub   = False
        # skills globais (Tier 1)
        self._g_brel        = tk.StringVar(value=self._cfg.get("g_brel","0"))
        self._g_acc         = tk.StringVar(value=self._cfg.get("g_acc", "0"))
        # Tier 2 — Dashboard
        self._dash_after_id = None
        # Tier 2 — Arbitragem multi-hop
        self._mh_results    = []
        self._loading_mh    = False
        self._mh_sort_state = {}
        # Tier 2 — Correlacao
        self._corr_results  = {}   # {tid: {"name":..., "closes":[...]}}
        self._loading_corr  = False
        # Tier 2 — Spread historico
        self._sph_tid       = None
        self._loading_sph   = False
        # Tier 2 — Alertas sonoros / desktop
        self._alert_history = []   # [(ts, msg), ...]
        # Tier 3 — Ciclos de mercado
        self._cycle_result     = None
        self._loading_cycle    = False
        self._cycle_sort_state = {}
        # Tier 3 — Previsao de precos
        self._fcast_result     = None   # ultimo resultado de predict_prices
        self._fcast_tid        = None
        self._fcast_name       = ""
        self._loading_fcast    = False
        # Tier 3 — Market Making
        self._mm2_results      = []    # scanner de oportunidades MM
        self._mm2_sim          = None  # simulacao de spread para item especifico
        self._loading_mm2      = False
        self._mm2_sort_state   = {}
        # Alta prioridade — Calculadora de Industria
        self._ind_materials    = []    # lista de materiais adicionados
        self._ind_result       = None
        self._loading_ind      = False
        self._ind_want_qty     = None  # StringVar — widget criado em _build_tab_ind
        self._ind_reg          = None  # StringVar — widget criado em _build_tab_ind
        self._ind_auto_btn     = None  # Button AUTO-PREENCHER
        self._ind_status_lbl   = None  # Label de status inline
        self._ind_str_mat      = None  # StringVar — Estrutura Mat%
        self._ind_str_job      = None  # StringVar — Estrutura Job%
        # Alta prioridade — Backtesting
        self._bt_result        = None
        self._loading_bt       = False
        # Alta prioridade — Heatmap
        self._hm_result        = {}
        self._loading_hm       = False
        # Watchlist — item atualmente selecionado (armazenado via <<TreeviewSelect>>)
        self._wl_cur_tid       = None  # type_id do último item clicado na WL tree
        # Tabs: _cur_tid para seleção robusta (Linux/GTK)
        self._mm_cur_tid       = None
        self._sc_cur_tid       = None
        self._ds_cur_tid       = None
        self._pm_cur_tid       = None
        self._bud_cur_tid      = None
        # Tabs restantes — cur_tid (type_id disponível) ou cur_idx (índice)
        self._r_cur_tid        = None   # ROTAS
        self._a_cur_tid        = None   # ARBITRAGEM
        self._a_cur_idx        = None   # índice de linha (iid=str(i))
        self._mh_cur_tid       = None   # MULTI-HOP
        self._cyc_cur_tid      = None   # CICLOS
        self._corr_cur_idx     = 0      # CORRELAÇÃO (sem type_id único)
        self._opt_cur_idx      = 0      # OTIMIZADOR (sem type_id único)
        self._ind_cur_tid      = None   # INDÚSTRIA
        # Catalogo de itens
        self._cat_rows         = []   # lista filtrada atual
        self._cat_sort_state   = {}
        self._loading_cat      = False
        # Station Trader
        self._st_results       = []
        self._st_sort_state    = {}
        self._loading_st       = False
        # misc
        self._chart_win     = None
        self._chk_labels    = {}
        self._r_sort_state  = {}
        self._a_sort_state  = {}
        self._mm_sort_state = {}
        self._sc_sort_state = {}
        self._ds_sort_state = {}
        self._pm_sort_state = {}
        self._bud_sort_state= {}
        self._hub_sort_state= {}

        _apply_styles()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        n  = len(self._station_cache)
        ni = len(self._item_cache)
        self._set_status(
            f"Cache: {n} estacoes | {ni} itens" if n
            else "Cache vazio -- use DADOS & CACHE")

    # ── layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        top = tk.Frame(self, bg=COLORS["bg"])
        top.pack(fill="x", padx=12, pady=(8, 0))
        tk.Label(top, text="EVE TRADE FINDER  v14",
                 font=("Courier", 17, "bold"),
                 fg=COLORS["accent"], bg=COLORS["bg"]).pack(side="left")
        self._status_lbl = tk.Label(top, text="...", font=("Courier", 10),
                                    fg=COLORS["subtext"], bg=COLORS["bg"])
        self._status_lbl.pack(side="right")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=6)
        self._nb = nb   # saved for programmatic tab switching

        self._tab_rotas  = tk.Frame(nb, bg=COLORS["bg"])
        self._tab_arb    = tk.Frame(nb, bg=COLORS["bg"])
        self._tab_opt    = tk.Frame(nb, bg=COLORS["bg"])
        self._tab_mm     = tk.Frame(nb, bg=COLORS["bg"])
        self._tab_sc     = tk.Frame(nb, bg=COLORS["bg"])
        self._tab_ds     = tk.Frame(nb, bg=COLORS["bg"])
        self._tab_pm     = tk.Frame(nb, bg=COLORS["bg"])
        self._tab_bud    = tk.Frame(nb, bg=COLORS["bg"])
        self._tab_live   = tk.Frame(nb, bg=COLORS["bg"])
        self._tab_wl     = tk.Frame(nb, bg=COLORS["bg"])
        self._tab_paper  = tk.Frame(nb, bg=COLORS["bg"])
        self._tab_ai     = tk.Frame(nb, bg=COLORS["bg"])
        self._tab_hub    = tk.Frame(nb, bg=COLORS["bg"])   # Tier 1
        self._tab_cfg    = tk.Frame(nb, bg=COLORS["bg"])   # Tier 1
        self._tab_dash   = tk.Frame(nb, bg=COLORS["bg"])   # Tier 2 Dashboard
        self._tab_mh     = tk.Frame(nb, bg=COLORS["bg"])   # Tier 2 Multi-hop
        self._tab_corr   = tk.Frame(nb, bg=COLORS["bg"])   # Tier 2 Correlacao
        self._tab_sph    = tk.Frame(nb, bg=COLORS["bg"])   # Tier 2 Spread historico
        self._tab_cycle  = tk.Frame(nb, bg=COLORS["bg"])   # Tier 3 Ciclos
        self._tab_fcast  = tk.Frame(nb, bg=COLORS["bg"])   # Tier 3 Previsao
        self._tab_mm2    = tk.Frame(nb, bg=COLORS["bg"])   # Tier 3 Market Making
        self._tab_ind    = tk.Frame(nb, bg=COLORS["bg"])   # Alta — Industria
        self._tab_bt     = tk.Frame(nb, bg=COLORS["bg"])   # Alta — Backtesting
        self._tab_hm     = tk.Frame(nb, bg=COLORS["bg"])   # Alta — Heatmap
        self._tab_catalog= tk.Frame(nb, bg=COLORS["bg"])   # Catalogo de itens
        self._tab_st     = tk.Frame(nb, bg=COLORS["bg"])   # Station Trader
        self._tab_dados  = tk.Frame(nb, bg=COLORS["bg"])
        self._tab_ind_scan = tk.Frame(nb, bg=COLORS["bg"])  # Scanner de Industria

        nb.add(self._tab_rotas,  text="  ROTAS  ")
        nb.add(self._tab_arb,    text="  ARBITRAGEM  ")
        nb.add(self._tab_opt,    text="  OTIMIZADOR  ")
        nb.add(self._tab_mm,     text="  MAX MARGEM  ")
        nb.add(self._tab_sc,     text="  ESCASSEZ  ")
        nb.add(self._tab_ds,     text="  DEMANDA/OFERTA  ")
        nb.add(self._tab_pm,     text="  MOV. PRECOS  ")
        nb.add(self._tab_bud,    text="  ORCAMENTO  ")
        nb.add(self._tab_live,   text="  MERCADO LIVE  ")
        nb.add(self._tab_wl,     text="  WATCHLIST  ")
        nb.add(self._tab_paper,  text="  PAPER TRADING  ")
        nb.add(self._tab_ai,     text="  AI ANALISTA  ")
        nb.add(self._tab_hub,    text="  COMPARADOR  ")
        nb.add(self._tab_cfg,    text="  CONFIGURACOES  ")
        nb.add(self._tab_dash,   text="  DASHBOARD  ")
        nb.add(self._tab_mh,     text="  MULTI-HOP  ")
        nb.add(self._tab_corr,   text="  CORRELACAO  ")
        nb.add(self._tab_sph,    text="  SPREAD HIST.  ")
        nb.add(self._tab_cycle,  text="  CICLOS  ")
        nb.add(self._tab_fcast,  text="  PREVISAO  ")
        nb.add(self._tab_mm2,    text="  MARKET MAKING  ")
        nb.add(self._tab_ind,    text="  INDUSTRIA  ")
        nb.add(self._tab_bt,     text="  BACKTESTING  ")
        nb.add(self._tab_hm,     text="  HEATMAP  ")
        nb.add(self._tab_catalog,text="  CATALOGO  ")
        nb.add(self._tab_st,     text="  STATION TRADER  ")
        nb.add(self._tab_dados,  text="  DADOS & CACHE  ")
        nb.add(self._tab_ind_scan, text="  IND. SCANNER  ")

        self._build_tab_rotas()
        self._build_tab_arb()
        self._build_tab_opt()
        self._build_tab_mm()
        self._build_tab_sc()
        self._build_tab_ds()
        self._build_tab_pm()
        self._build_tab_bud()
        self._build_tab_live()
        self._build_tab_wl()
        self._build_tab_paper()
        self._build_tab_ai()
        self._build_tab_hub()
        self._build_tab_cfg()
        self._build_tab_dash()
        self._build_tab_mh()
        self._build_tab_corr()
        self._build_tab_sph()
        self._build_tab_cycle()
        self._build_tab_fcast()
        self._build_tab_mm2()
        self._build_tab_ind()
        self._build_tab_bt()
        self._build_tab_hm()
        self._build_tab_catalog()
        self._build_tab_st()
        self._build_tab_dados()
        self._build_tab_ind_scan()
        self._setup_experimental_gate()
        # Tooltips e atalhos de teclado (após todos os widgets criados)
        self._setup_tooltips()
        self._setup_keybindings()
        # Auto-save periódico — watchlist + paper trades a cada 5 minutos
        self._auto_save_loop()

    # ── helpers ───────────────────────────────────────────────────────────────

    def _setup_tooltips(self):
        """Registra tooltips em todos os campos de filtro do aplicativo.
        Funciona via self._wmap, populado em _entry() e _combo() durante o build.
        Para widgets criados com tk.Entry direto, usa W(widget) em vez de T(sv)."""

        def T(sv, text):
            """Tooltip via StringVar → _wmap."""
            w = self._wmap.get(id(sv))
            if w:
                ToolTip(w, text)

        def W(widget, text):
            """Tooltip direto num widget (para campos não criados via _entry)."""
            if widget:
                ToolTip(widget, text)

        # ── STATION TRADER ────────────────────────────────────────────────────
        T(self._st_days,      "Days of Trade: dias de estoque a comprar.\nEx: 3 = compra para 3 dias de giro.")
        T(self._st_max_cap,   "Capital máximo por item (ISK).\n0 = sem limite.")
        T(self._st_min_profit,"Min. Item Profit: lucro mínimo do lote (ISK).")
        T(self._st_min_margin,"Min. Order Margin: margem bruta mínima (%).\n(sell−buy)/sell × 100.")
        T(self._st_min_daily, "Min. Daily Tradable: volume mínimo por dia.")
        T(self._st_avg_period,"Avg Price Period: dias para preço médio do período.")
        T(self._st_min_proi,  "Min. Period ROI: ROI mínimo no período (%).")
        T(self._st_min_bvs,   "Min. B/S Ratio: proporção mínima compra/venda.\n<0.5 = mais vendedores.")
        T(self._st_max_bvs,   "Max. B/S Ratio: proporção máxima compra/venda.\n>1.5 = mais compradores.")
        T(self._st_buy_plow,  "Limit Buy to P.Low (%): limite de compra ao mínimo histórico.\nEx: -10 = até 10% abaixo do mínimo.")
        T(self._st_sell_phigh,"Limit Sell to P.High (%): limite de venda ao máximo histórico.")

        # ── PREVISÃO ──────────────────────────────────────────────────────────
        T(self._fcast_tid_e,  "Type ID do item a prever.\nCole aqui após copiar com botão direito.")
        T(self._fcast_name_e, "Nome do item (preenchido automaticamente).")
        T(self._fcast_reg,    "Região de mercado para buscar o histórico.")
        T(self._fcast_horizon,"Horizonte de previsão em dias (7, 14 ou 30).")
        T(self._fcast_alpha,  "Alpha DES: peso das observações recentes.\n0.01 = suave · 0.99 = reativo. Padrão: 0.30")
        T(self._fcast_beta,   "Beta DES: peso da tendência.\n0.01 = estável · 0.99 = reativo. Padrão: 0.10")

        # ── BACKTESTING ───────────────────────────────────────────────────────
        T(self._bt_tid,       "Type ID do item para backtesting.")
        T(self._bt_name,      "Nome do item (preenchido automaticamente).")
        T(self._bt_reg,       "Região de mercado para o backtest.")
        T(self._bt_days,      "Dias de histórico. Mínimo: 14. Recomendado: 90–365.")
        T(self._bt_cap,       "Capital inicial (ISK). Ex: 1000000000 = 1B ISK.")

        # ── AI ANALISTA ───────────────────────────────────────────────────────
        T(self._ai_tid,       "Type ID do item a analisar com IA.")
        T(self._ai_name_e,    "Nome do item (preenchido automaticamente).")
        T(self._ai_reg,       "Região de mercado para buscar dados.")

        # ── MARKET MAKING (Simulador) ─────────────────────────────────────────
        T(self._mm2_sim_tid,  "Type ID do item para simular spread de Market Making.")
        T(self._mm2_sim_name, "Nome do item (preenchido automaticamente).")

        # ── SPREAD HISTÓRICO ──────────────────────────────────────────────────
        T(self._sph_tid,      "Type ID do item para análise de spread histórico.")
        T(self._sph_name,     "Nome do item (preenchido automaticamente).")

        # ── COMPARADOR DE HUBS ────────────────────────────────────────────────
        T(self._hub_tid,      "Type ID do item para comparar nos 5 hubs principais.")
        T(self._hub_name,     "Nome do item (preenchido automaticamente).")

        # ── INDÚSTRIA ─────────────────────────────────────────────────────────
        T(self._ind_prod_tid, "Type ID do produto fabricado.\nEx: 587 = Rifter.")
        T(self._ind_prod_name,"Nome do produto (preenchido automaticamente).")

        # ── WATCHLIST (widgets diretos — não via _entry) ──────────────────────
        T(self._wl_interval,  "Intervalo de atualização (segundos).\nMínimo recomendado: 60s.")
        W(self._wl_ab,        "Alerta ↓: dispara quando preço sell cair abaixo deste valor (ISK).")
        W(self._wl_aa,        "Alerta ↑: dispara quando preço sell subir acima deste valor (ISK).")

        # ── PAPER TRADING ─────────────────────────────────────────────────────
        T(self._pt_tid,       "Type ID do item para Paper Trading.")
        T(self._pt_name,      "Nome do item (preenchido automaticamente).")

        # ── STATION TRADER ────────────────────────────────────────────────────

        # ── PREVISÃO ──────────────────────────────────────────────────────────

        # ── BACKTESTING ───────────────────────────────────────────────────────

        # ── AI ANALISTA ───────────────────────────────────────────────────────

        # ── MARKET MAKING (Simulador) ─────────────────────────────────────────

        # ── SPREAD HISTÓRICO ──────────────────────────────────────────────────

        # ── COMPARADOR DE HUBS ────────────────────────────────────────────────

        # ── INDÚSTRIA ─────────────────────────────────────────────────────────

        # ── WATCHLIST ─────────────────────────────────────────────────────────

        # ── PAPER TRADING ─────────────────────────────────────────────────────

    def _find_entry_widget(self, sv: tk.StringVar):
        """Encontra o widget associado a um StringVar via _wmap (O(1))."""
        return self._wmap.get(id(sv))
    def _setup_keybindings(self):
        """Registra atalhos de teclado globais na janela principal."""
        nb = self._nb

        # F5 — roda o cálculo/scanner da aba atual
        def _f5(_e=None):
            try:
                tab = nb.select()
                tab_map = {
                    str(self._tab_rotas):  self._start_rotas,
                    str(self._tab_arb):    self._start_arb,
                    str(self._tab_mm):     self._start_mm,
                    str(self._tab_sc):     self._start_sc,
                    str(self._tab_ds):     self._start_ds,
                    str(self._tab_pm):     self._start_pm,
                    str(self._tab_bud):    self._start_bud,
                    str(self._tab_fcast):  self._start_fcast,
                    str(self._tab_bt):     self._start_bt,
                    str(self._tab_sph):    self._start_sph,
                    str(self._tab_st):     self._start_st,
                    str(self._tab_hm):     self._start_hm,
                    str(self._tab_mh):     self._start_mh,
                    str(self._tab_corr):   self._start_corr,
                    str(self._tab_cycle):  self._start_cycle,
                    str(self._tab_mm2):    self._start_mm2,
                }
                fn = tab_map.get(tab)
                if fn:
                    fn()
            except Exception as _e:

                log.debug("silenced: %s", _e)

        # Ctrl+número — navega para aba
        tab_shortcuts = {
            "1": self._tab_rotas,  "2": self._tab_arb,
            "3": self._tab_opt,    "4": self._tab_mm,
            "5": self._tab_sc,     "6": self._tab_ds,
            "7": self._tab_pm,     "8": self._tab_bud,
            "9": self._tab_live,   "0": self._tab_wl,
        }

        self.bind("<F5>", _f5)
        self.bind("<F9>", lambda e: self._start_wl_monitor() if hasattr(self, '_start_wl_monitor') else None)

        for key, tab in tab_shortcuts.items():
            self.bind(f"<Control-{key}>",
                      lambda e, t=tab: self._nb.select(t))

        # Atalhos nomeados para abas mais usadas
        named = {
            "<Control-w>": self._tab_wl,        # Watchlist
            "<Control-s>": self._tab_st,         # Station Trader
            "<Control-p>": self._tab_pm,         # Mov. Preços
            "<Control-f>": self._tab_fcast,      # Previsão (Forecast)
            "<Control-b>": self._tab_bt,         # Backtesting
            "<Control-h>": self._tab_hm,         # Heatmap
            "<Control-d>": self._tab_dash,       # Dashboard
            "<Control-i>": self._tab_ind,        # Indústria
            "<Control-a>": self._tab_ai,         # AI Analista
        }
        for key, tab in named.items():
            self.bind(key, lambda e, t=tab: self._nb.select(t))

        # Escape — cancela loading se possível (fecha messagebox)
        self.bind("<Escape>", lambda e: None)

        self._set_status(
            "F5=executar  Ctrl+1-9=abas  Ctrl+W=Watchlist  "
            "Ctrl+S=StationTrader  Ctrl+F=Previsão  Ctrl+B=Backtest  "
            "Ctrl+H=Heatmap  Ctrl+D=Dashboard  Ctrl+A=AI")

    def _export_table(self, title: str, headers: list,
                       rows: list, default_filename: str = "export"):
        """
        Mostra diálogo de exportação e salva como CSV ou Excel (.xlsx).
        headers — lista de (key, label) ou strings.
        rows    — lista de dicts.
        """
        if not rows:
            messagebox.showinfo("Vazio", "Sem dados para exportar."); return

        filetypes = [("CSV", "*.csv")]
        if HAS_OPENPYXL:
            filetypes.insert(0, ("Excel (.xlsx)", "*.xlsx"))

        path = filedialog.asksaveasfilename(
            title=f"Exportar — {title}",
            defaultextension=".xlsx" if HAS_OPENPYXL else ".csv",
            filetypes=filetypes,
            initialfile=default_filename)
        if not path: return

        try:
            if path.endswith(".xlsx"):
                ok = export_xlsx(path, title[:31], headers, rows)
                if ok:
                    messagebox.showinfo("Exportado", f"Excel salvo em:\n{path}")
                else:
                    messagebox.showerror("Erro",
                        "Falha ao salvar Excel. Tente exportar como CSV.")
            else:
                keys = [k if isinstance(k, str) else k[0] for k in headers]
                lbls = [k if isinstance(k, str) else k[1] for k in headers]
                with open(path, "w", newline="", encoding="utf-8") as f:
                    w = csv.writer(f)
                    w.writerow(lbls)
                    for row in rows:
                        w.writerow([row.get(k, "") if isinstance(row, dict)
                                    else "" for k in keys])
                messagebox.showinfo("Exportado", f"CSV salvo em:\n{path}")
        except Exception as exc:
            messagebox.showerror("Erro ao exportar", str(exc))

    def _auto_save_loop(self):
        """Auto-save silencioso da watchlist e paper trades a cada 5 minutos.
        Protege contra perda de dados em crashes ou fechamentos inesperados."""
        try:
            save_watchlist(self._watchlist)
        except Exception as _e:
            log.debug("auto_save watchlist: %s", _e)
        try:
            save_paper_trades(self._paper)
        except Exception as _e:
            log.debug("auto_save paper: %s", _e)
        # Reagenda para daqui a 5 minutos (300 000 ms)
        self._auto_save_id = self.after(300_000, self._auto_save_loop)

    def _lbl(self, parent, text, row, col):
        tk.Label(parent, text=text, bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 10)
                 ).grid(row=row, column=col, sticky="w", padx=4, pady=2)

    def _entry(self, parent, default, row, col, width=12):
        v = tk.StringVar(value=default)
        w = tk.Entry(parent, textvariable=v, bg=COLORS["border"],
                     fg=COLORS["text"], insertbackground=COLORS["text"],
                     font=("Courier", 11), width=width, relief="flat", bd=2)
        w.grid(row=row, column=col, padx=4, pady=2, sticky="w")
        self._wmap[id(v)] = w   # regista widget para tooltip lookup
        return v

    # ── Auto-lookup de nome pelo Type ID ──────────────────────────────────────
    # Conecta qualquer campo tid_var → name_var.
    # Quando o usuário cola ou digita um Type ID válido, o nome é buscado
    # automaticamente no _item_cache (cache local) e, se ausente,
    # via ESI em background — sem bloquear a UI.

    def _bind_tid_lookup(self, tid_var: tk.StringVar,
                          name_var: tk.StringVar,
                          reg_var:  tk.StringVar = None):
        """
        Liga o campo tid_var ao name_var para preenchimento automático.
        Dispara em qualquer escrita (digitação, paste, preenchimento via código).

        tid_var  — StringVar do campo Type ID
        name_var — StringVar do campo Nome
        reg_var  — StringVar do combo Região (opcional, para busca histórico)
        """
        def _on_tid_change(*_):
            raw = tid_var.get().strip()
            # Só age quando o valor for um inteiro válido
            if not raw.isdigit():
                return
            tid = int(raw)
            if tid <= 0:
                return
            # 1. Verifica cache local imediatamente (O(1), sem I/O)
            cached = self._item_cache.get(tid, {})
            name   = cached.get("name", "")
            if name and name != str(tid):
                # Nome já conhecido — preenche imediatamente
                current = name_var.get().strip()
                if not current or current == str(tid):
                    name_var.set(name)
                return
            # 2. Nome não está no cache → busca ESI em background
            def _fetch():
                try:
                    _ensure_names([tid], self._item_cache)
                    found = self._item_cache.get(tid, {}).get("name", "")
                    if found and found != str(tid):
                        current = name_var.get().strip()
                        if not current or current == str(tid):
                            self.after(0, lambda n=found: name_var.set(n))
                except Exception as _e:

                    log.debug("silenced: %s", _e)
            threading.Thread(target=_fetch, daemon=True).start()

        tid_var.trace_add("write", _on_tid_change)

    def _combo(self, parent, values, row, col, width=22):
        v = tk.StringVar(value=values[0])
        w = ttk.Combobox(parent, textvariable=v, values=values,
                         state="readonly", width=width, font=("Courier", 11))
        w.grid(row=row, column=col, padx=4, pady=2, sticky="w")
        self._wmap[id(v)] = w   # regista widget para tooltip lookup
        return v

    def _csv_age_label(self, parent, mode_var, region_vars, row, col):
        """Mostra sob o seletor 'Fonte' a data do CSV usado no calculo.

        Com varias regioes (rotas, arbitragem, multi-hop) mostra a MAIS ANTIGA.
        Atualiza ao trocar regiao/fonte, a cada 60 s e apos downloads.
        Cores iguais a aba Dados & Cache: verde < 1h, amarelo < 24h, vermelho.
        """
        lbl = tk.Label(parent, text="", bg=parent.cget("bg"),
                       font=("Courier", 9), anchor="w")
        lbl.grid(row=row, column=col, padx=4, sticky="w")

        def _update(*_):
            if not lbl.winfo_exists():
                return
            if not mode_var.get().startswith("CSV"):
                lbl.config(text="ao vivo (ESI)", fg=COLORS["subtext"])
                return
            regs = [v.get() for v in region_vars if v.get() in REGIONS]
            missing = [r for r in regs if not os.path.exists(_orders_csv_path(r))]
            if missing:
                lbl.config(text=f"sem CSV: {missing[0].split(' (')[0]}", fg=COLORS["red"])
                return
            if not regs:
                lbl.config(text="")
                return
            oldest = min(regs, key=lambda r: os.path.getmtime(_orders_csv_path(r)))
            label_txt, color = cache_file_info(oldest)
            suffix = f" ({oldest.split(' (')[0]})" if len(regs) > 1 else ""
            lbl.config(text=f"CSV {label_txt}{suffix}", fg=color)

        for v in (mode_var, *region_vars):
            v.trace_add("write", _update)
        if not hasattr(self, "_csv_age_updaters"):
            self._csv_age_updaters = []
            self.after(60_000, self._tick_csv_age)
        self._csv_age_updaters.append(_update)
        _update()

    def _tick_csv_age(self):
        """Reavalia todos os rotulos de idade do CSV (cor muda com o tempo)."""
        for upd in getattr(self, "_csv_age_updaters", []):
            try:
                upd()
            except tk.TclError as e:
                log.debug("csv_age_label: %s", e)
        self.after(60_000, self._tick_csv_age)

    def _make_tree(self, parent, cols, height=None):
        outer = tk.Frame(parent, bg=COLORS["bg"])
        outer.pack(fill="both", expand=True, padx=6, pady=(0, 2))
        kw = {"style": "EVE.Treeview",
              "columns": [c[0] for c in cols],
              "show": "headings", "selectmode": "browse"}
        if height is not None:
            kw["height"] = height
        tree = ttk.Treeview(outer, **kw)
        _text = {"name","item","station","dst_station","loc_origem",
                 "loc_destino","Origem","Destino","Item"}
        for cid, cname, cw in cols:
            tree.heading(cid, text=cname)
            tree.column(cid, width=cw, minwidth=40,
                        anchor="w" if cid in _text else "e")
        vsb = ttk.Scrollbar(outer, orient="vertical",   command=tree.yview)
        hsb = ttk.Scrollbar(outer, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)
        tree.tag_configure("odd",  background=COLORS["row_odd"])
        tree.tag_configure("even", background=COLORS["row_even"])
        tree.tag_configure("hot",  foreground=COLORS["green"])
        return tree

    def _set_heading_arrow(self, tree, col_names: dict, active_col: str, reverse: bool):
        """M7: atualiza cabecalhos mostrando seta apenas na coluna ativa."""
        arrow = " ▼" if reverse else " ▲"
        for cid, base_text in col_names.items():
            tree.heading(cid, text=base_text + (arrow if cid == active_col else ""))

    def _fmt(self, v, kind="isk"):
        try:
            if kind == "isk": return f"{float(v):,.2f}"
            if kind == "pct": return f"{float(v):.1f}%"
            if kind == "int": return f"{int(v):,}"
            if kind == "vol": return f"{float(v):,.2f}"
        except Exception as _e:

            log.debug("silenced: %s", _e)
        return str(v)

    def _set_status(self, msg):
        # M9: aceita string ou tupla (done, total, msg)
        if isinstance(msg, tuple):
            msg = msg[2]
        self.after(0, lambda: self._status_lbl.config(text=str(msg)[:180]))

    def _on_close(self):
        """Salva configuracoes, para timers e fecha."""
        self._live_running = False
        self._wl_running   = False
        if self._live_after_id:
            try: self.after_cancel(self._live_after_id)
            except Exception: pass
        if self._wl_after_id:
            try: self.after_cancel(self._wl_after_id)
            except Exception: pass
        if self._dash_after_id:
            try: self.after_cancel(self._dash_after_id)
            except Exception: pass
        self._persist_config()
        self.destroy()

    def _persist_config(self):
        """Le todos os campos de filtro e salva em config.json."""
        try:
            cfg = {
                "r_roi":   self._r_roi.get(),   "r_ltmin": self._r_ltmin.get(),
                "r_cmin":  self._r_cmin.get(),  "r_cmax":  self._r_cmax.get(),
                "r_vol":   self._r_vol.get(),
                "a_bbuy":  self._a_bbuy.get(),  "a_bsell": self._a_bsell.get(),
                "a_stax":  self._a_stax.get(),  "a_fret":  self._a_fret.get(),
                "a_mmin":  self._a_mmin.get(),  "a_ltmin": self._a_ltmin.get(),
                "ai_key":  "",   # nunca em disco — ver keyring_save_key
                # Skills globais
                "g_brel":  self._g_brel.get(),
                "g_acc":   self._g_acc.get(),
                "experimental": "1" if self._experimental.get() else "0",
            }
            if HAS_ORTOOLS:
                cfg.update({
                    "o_cap":    self._o_cap.get(),
                    "o_lucro":  self._o_lucro.get(),
                    "o_marg":   self._o_marg.get(),
                    "o_ispulo": self._o_ispulo.get(),
                    "o_pulos":  self._o_pulos.get(),
                })
            save_config(cfg)
            # Salva API Key só no keyring seguro (nunca se veio do ambiente)
            api_key = self._ai_api_key.get().strip()
            if api_key and api_key != os.environ.get(API_KEY_ENV):
                keyring_save_key(api_key)
        except Exception as _e:

            log.debug("silenced: %s", _e)

    def _fread(self, var, default):
        """Le float de StringVar. Retorna default se vazio ou invalido."""
        try:
            return float(var.get().strip() or str(default))
        except ValueError:
            return float(default)

    def _fread_limit(self, var, zero_means=1e18):
        """Le um limite numerico: 0 (ou vazio) significa 'sem limite' (zero_means)."""
        try:
            v = float(var.get().strip() or "0")
            return zero_means if v == 0 else v
        except ValueError:
            return zero_means

    def _make_prog(self, prog_bar=None, pct_lbl=None):
        """
        C1 — Fabrica de callbacks de progresso (elimina 10 closures identicas).
        Retorna funcao prog(msg) que:
          - se msg e tuple (done, total, txt): atualiza barra + label + status
          - se msg e str: atualiza so o status
        Ambos os args sao opcionais; funciona mesmo sem barra.
        """
        def prog(msg):
            if isinstance(msg, tuple):
                done, total, txt = msg
                pct = int(done / total * 100) if total else 0
                def _upd(p=pct, t=txt):
                    if prog_bar is not None:
                        prog_bar.config(value=p)
                    if pct_lbl is not None:
                        pct_lbl.config(text=f"{p}%")
                    self._status_lbl.config(text=t[:180])
                self.after(0, _upd)
            else:
                def _set(t=str(msg)):
                    fg = COLORS["red"] if HIST_ERROR_MARK in t else COLORS["subtext"]
                    self._status_lbl.config(text=t[:220], fg=fg)
                self.after(0, _set)
        return prog

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA 1 — ROTAS
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_rotas(self):
        frame = tk.Frame(self._tab_rotas, bg=COLORS["panel"],
                         highlightbackground=COLORS["border"], highlightthickness=1)
        frame.pack(fill="x", padx=6, pady=(6, 2))

        c = self._cfg   # M6
        regions = list(REGIONS.keys())
        self._lbl(frame, "Origem",           0, 0)
        self._r_src   = self._combo(frame, regions, 1, 0)
        self._lbl(frame, "Destino",          0, 2)
        self._r_dst   = self._combo(frame, regions[1:]+[regions[0]], 1, 2)
        self._lbl(frame, "Fonte",            0, 4)
        self._r_mode  = self._combo(frame, ["API (live)", "CSV (cache)"], 1, 4, 13)
        self._csv_age_label(frame, self._r_mode, [self._r_src, self._r_dst], 2, 4)
        self._lbl(frame, "ROI min (%)",      0, 6)
        self._r_roi   = self._entry(frame, c["r_roi"],   1, 6, 7)
        self._lbl(frame, "Lucro Total min",  0, 8)
        self._r_ltmin = self._entry(frame, c["r_ltmin"], 1, 8, 12)
        self._lbl(frame, "Capital min",      0, 10)
        self._r_cmin  = self._entry(frame, c["r_cmin"],  1, 10, 12)
        self._lbl(frame, "Capital max",      0, 12)
        self._r_cmax  = self._entry(frame, c["r_cmax"],  1, 12, 12)
        self._lbl(frame, "Vol max (m3)",     0, 14)
        self._r_vol   = self._entry(frame, c["r_vol"],   1, 14, 10)

        self._r_btn = tk.Button(frame, text="  BUSCAR ROTAS  ",
                                command=self._start_rotas,
                                bg=COLORS["accent"], fg="#000",
                                font=("Courier", 11, "bold"),
                                relief="flat", cursor="hand2")
        self._r_btn.grid(row=1, column=16, padx=8)
        tk.Button(frame, text=" EXPORTAR CSV ", command=self._export_rotas,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=17, padx=4)
        # M9: barra deterministica
        self._r_prog = ttk.Progressbar(frame, mode="determinate", length=120,
                                        maximum=100, value=0)
        self._r_prog.grid(row=1, column=18, padx=6)
        self._r_pct = tk.Label(frame, text="", bg=COLORS["panel"],
                               fg=COLORS["subtext"], font=("Courier", 10), width=5)
        self._r_pct.grid(row=1, column=19, padx=2)

        self._r_tree = self._make_tree(self._tab_rotas, [
            ("name",           "Item",             210),
            ("station",        "Estacao (compra)", 180),
            ("dst_station",    "Estacao (venda)",  180),
            ("dst_vol_demand", "Qtd destino",        75),
            ("vol_realizavel", "Qtd realizavel",      75),
            ("roi",            "ROI %",              65),
            ("buy_price",      "Compra (ISK)",      115),
            ("sell_price",     "Venda (ISK)",       115),
            ("profit_unit",    "Lucro/un",          108),
            ("lucro_total",    "Lucro Total",       128),
            ("lucro_dia",      "Lucro/Dia",         118),
            ("capital",        "Capital (ISK)",     118),
            ("competition",    "Concorr.",            60),
            ("total_volume",   "Vol m3",              85),
        ])
        for cid in self._R_COL_NAMES:
            self._r_tree.heading(cid, text=self._R_COL_NAMES[cid],
                                  command=lambda c=cid: self._sort_rotas(c))
        self._r_tree.bind("<Double-1>", self._on_dbl_rotas)

        bar = tk.Frame(self._tab_rotas, bg=COLORS["bg"])
        bar.pack(fill="x", padx=6, pady=(0, 4))
        self._r_count = tk.Label(bar, text="0 rotas", fg=COLORS["subtext"],
                                  bg=COLORS["bg"], font=("Courier", 10))
        self._r_count.pack(side="left")
        self._r_tree.bind("<<TreeviewSelect>>", self._r_on_select)
        self._r_tree.bind("<Double-1>",         self._r_on_dbl)
        tk.Button(bar, text=" GRÁFICO ",     command=self._r_open_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" PREVER ",      command=self._r_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" BACKTEST ",    command=self._r_open_bt,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" ANALISAR IA ", command=self._r_open_ai,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Label(bar, text="Double-click -> historico 90d  |  Lucro/Dia = giro medio x lucro/un",
                 fg=COLORS["subtext"], bg=COLORS["bg"],
                 font=("Courier", 10)).pack(side="right")

    def _start_rotas(self):
        if self._loading_rotas: return
        self._loading_rotas = True
        self._r_btn.config(state="disabled")
        self._r_prog.config(value=0)
        self._r_pct.config(text="")
        # C2 — lê widgets no thread principal antes de lançar o worker
        params = dict(
            src       = self._r_src.get(),
            dst       = self._r_dst.get(),
            use_csv   = self._r_mode.get().startswith("CSV"),
            min_roi   = self._fread(self._r_roi,   10),
            ltmin     = self._fread(self._r_ltmin,  0),
            cmin      = self._fread(self._r_cmin,   0),
            cmax      = self._fread_limit(self._r_cmax),
            vol       = self._fread(self._r_vol,    0),
        )
        threading.Thread(target=self._worker_rotas,
                         args=(params,), daemon=True).start()

    def _worker_rotas(self, params: dict):
        prog = self._make_prog(self._r_prog, self._r_pct)
        try:
            src     = params["src"]
            dst     = params["dst"]
            if params["use_csv"]:
                prog("Carregando CSVs...")
                so = load_orders_from_csv(src)
                do = load_orders_from_csv(dst)
                if not so or not do:
                    self.after(0, lambda: messagebox.showerror(
                        "CSV nao encontrado",
                        "Baixe os dados primeiro na aba DADOS & CACHE."))
                    return
            else:
                prog("Baixando ordens via API...")
                with concurrent.futures.ThreadPoolExecutor(max_workers=2) as p:
                    fs = p.submit(download_region_orders, REGIONS[src], src, prog)
                    fd = p.submit(download_region_orders, REGIONS[dst], dst, prog)
                    so, do = fs.result(), fd.result()
                _orders_mem_cache[src] = (time.time(), so)
                _orders_mem_cache[dst] = (time.time(), do)

            prog("Calculando rotas...")
            routes = find_routes(
                so, do, src, dst,
                min_roi         = params["min_roi"],
                min_lucro_total = params["ltmin"],
                min_budget      = params["cmin"],
                max_budget      = params["cmax"],
                max_volume      = params["vol"],
                station_cache=self._station_cache,
                item_cache=self._item_cache,
            )
            self._item_cache = resolve_item_names(routes, self._item_cache, prog)
            self._routes = routes
            self.after(0, lambda: (self._r_prog.config(value=100),
                                   self._r_pct.config(text="100%")))
            self.after(0, self._pop_rotas)

            dst_region_id = REGIONS.get(dst)
            if dst_region_id and routes:
                prog("Buscando historico para Lucro/Dia...")
                enrich_routes_lucro_dia(routes, dst_region_id, prog)
                self.after(0, self._pop_rotas)

        except Exception as exc:
            log.exception("worker_rotas")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro", e))
        finally:
            self._loading_rotas = False
            self.after(0, lambda: (self._r_btn.config(state="normal"),
                                   self._set_status(f"Pronto — {len(self._routes)} rotas")))

    def _pop_rotas(self, reorder_only=False):
        """reorder_only=True: reordena iids in-place sem delete/reinsert (sem flickering)."""
        if reorder_only:
            for i in range(len(self._routes)):
                self._r_tree.move(str(i), "", i)
            return
        for row in self._r_tree.get_children(): self._r_tree.delete(row)
        for i, r in enumerate(self._routes):
            tag = ("hot",) if r["roi"] >= 50 else ("even" if i%2==0 else "odd",)
            ld  = (self._fmt(r["lucro_dia"], "isk") if r.get("lucro_dia") is not None
                   else "ERRO" if r.get("hist_erro") else "--")
            self._r_tree.insert("", "end", iid=str(r["type_id"]), tags=tag, values=(
                r["name"],
                r["station"],
                r.get("dst_station","--"),
                self._fmt(r.get("dst_vol_demand",0),"int"),
                self._fmt(r.get("vol_realizavel",0),"int"),
                self._fmt(r["roi"],                  "pct"),
                self._fmt(r["buy_price"],            "isk"),
                self._fmt(r["sell_price"],           "isk"),
                self._fmt(r["profit_unit"],          "isk"),
                self._fmt(r.get("lucro_total",0),    "isk"),
                ld,
                self._fmt(r["capital"],              "isk"),
                self._fmt(r["competition"],          "int"),
                self._fmt(r["total_volume"],         "vol"),
            ))
        self._r_count.config(text=f"{len(self._routes)} rotas encontradas")

    _R_COL_NAMES = {
        "name":"Item","station":"Estacao (compra)","dst_station":"Estacao (venda)",
        "dst_vol_demand":"Qtd destino","vol_realizavel":"Qtd realizavel","roi":"ROI %",
        "buy_price":"Compra (ISK)","sell_price":"Venda (ISK)","profit_unit":"Lucro/un",
        "lucro_total":"Lucro Total","lucro_dia":"Lucro/Dia","capital":"Capital (ISK)",
        "competition":"Concorr.","total_volume":"Vol m3",
    }

    def _sort_rotas(self, col):
        # 3A — padrao uniforme com set de colunas numericas (igual _sort_arb etc.)
        num = {"roi","buy_price","sell_price","profit_unit","lucro_total",
               "lucro_dia","vol_realizavel","dst_vol_demand",
               "capital","competition","total_volume"}
        rev = not self._r_sort_state.get(col, False)
        self._r_sort_state = {col: rev}
        self._routes.sort(
            key=lambda r: (r.get(col) or 0) if col in num else r.get(col, ""),
            reverse=rev)
        self._set_heading_arrow(self._r_tree, self._R_COL_NAMES, col, rev)
        self._pop_rotas(reorder_only=True)

    def _on_dbl_rotas(self, _):
        sel = self._r_tree.selection()
        if not sel: return
        idx = int(sel[0])
        if idx >= len(self._routes): return
        if not HAS_MATPLOTLIB:
            messagebox.showinfo("matplotlib ausente", "pip install matplotlib")
            return
        self._open_chart(self._routes[idx])

    def _open_candle_chart(self, tid: int, name: str, rid: int):
        """
        Abre janela de gráfico técnico para um item numa região.
        Mostra: OHLC bars, SMA7/26, Bollinger Bands, RSI(14), Volume.
        Reutiliza self._chart_win — fecha a janela anterior se existir.
        """
        if not HAS_MATPLOTLIB:
            messagebox.showinfo("Falta matplotlib", "pip install matplotlib"); return

        # Reutiliza ou cria janela
        if self._chart_win and self._chart_win.winfo_exists():
            self._chart_win.destroy()
        win = tk.Toplevel(self)
        win.title(f"Gráfico — {name}")
        win.configure(bg=COLORS["bg"])
        win.geometry("1100x640")
        self._chart_win = win

        lbl_load = tk.Label(win, text=f"Carregando histórico de {name}...",
                            fg=COLORS["subtext"], bg=COLORS["bg"],
                            font=("Courier", 11))
        lbl_load.pack(pady=20)

        def _worker():
            hist = _get_hist_cached(rid, tid)
            ind  = compute_indicators(hist) if hist else {}
            win.after(0, lambda: _draw(hist, ind))

        def _draw(hist, ind):
            lbl_load.destroy()

            if not hist or not ind:
                tk.Label(win,
                         text=f"Sem histórico disponível para {name}\n"
                              f"na região selecionada.\n\n"
                              f"Baixe o histórico na aba DADOS & CACHE.",
                         fg=COLORS["subtext"], bg=COLORS["bg"],
                         font=("Courier", 12), justify="center").pack(expand=True)
                return

            # Limita a 180 dias para performance
            n_max  = 180
            dates  = ind["dates"][-n_max:]
            closes = ind["closes"][-n_max:]
            highs  = ind["highs"][-n_max:]
            lows   = ind["lows"][-n_max:]
            vols   = ind["volumes"][-n_max:]
            sma7   = ind["sma7"][-n_max:]
            sma26  = ind["sma26"][-n_max:]
            bb_u   = ind["bb_upper"][-n_max:]
            bb_l   = ind["bb_lower"][-n_max:]
            bb_m   = ind["bb_mid"][-n_max:]
            rsi    = ind["rsi"][-n_max:]
            xs     = list(range(len(dates)))

            fig = Figure(figsize=(13, 7), facecolor=COLORS["bg"])
            fig.subplots_adjust(left=0.06, right=0.97, top=0.93,
                                bottom=0.10, hspace=0.12)

            # ── Layout: 3 linhas — Preço (3), RSI (1), Volume (1) ───────────
            ax_price = fig.add_subplot(3, 1, 1)
            ax_rsi   = fig.add_subplot(3, 1, 2, sharex=ax_price)
            ax_vol   = fig.add_subplot(3, 1, 3, sharex=ax_price)

            C = COLORS
            for ax in (ax_price, ax_rsi, ax_vol):
                ax.set_facecolor(C["panel"])
                ax.tick_params(colors=C["subtext"], labelsize=7)
                for sp in ax.spines.values():
                    sp.set_edgecolor(C["border"])
                ax.grid(color=C["border"], linewidth=0.4, alpha=0.5)

            # ── Painel de Preço: OHLC bars + SMA + BB ───────────────────────
            ax_price.set_title(f"{name}  |  {dates[0]} → {dates[-1]}  "
                               f"|  {len(dates)} dias",
                               color=C["header"], fontsize=9, pad=4)

            # OHLC bars (linha vertical hi-lo + tick horizontal close)
            for i, x in enumerate(xs):
                color = C["green"] if closes[i] >= (closes[i-1] if i > 0 else closes[i]) else C["red"]
                ax_price.plot([x, x], [lows[i], highs[i]],
                              color=color, linewidth=0.8, alpha=0.7)
                ax_price.plot([x, x+0.3], [closes[i], closes[i]],
                              color=color, linewidth=1.5)

            # SMA 7 e 26
            xs_arr = xs
            valid7  = [(x, v) for x, v in zip(xs_arr, sma7)  if v is not None]
            valid26 = [(x, v) for x, v in zip(xs_arr, sma26) if v is not None]
            if valid7:
                ax_price.plot(*zip(*valid7),  color=C["accent"],  linewidth=1.2,
                              label="SMA7", alpha=0.9)
            if valid26:
                ax_price.plot(*zip(*valid26), color=C["yellow"],  linewidth=1.2,
                              label="SMA26", alpha=0.9)

            # Bollinger Bands
            valid_bb = [(x, u, l, m) for x, u, l, m
                        in zip(xs_arr, bb_u, bb_l, bb_m)
                        if u is not None]
            if valid_bb:
                xb, ub, lb, mb = zip(*valid_bb)
                ax_price.fill_between(xb, lb, ub,
                                      color=C["accent"], alpha=0.07)
                ax_price.plot(xb, ub, color=C["accent"],
                              linewidth=0.6, linestyle="--", alpha=0.5)
                ax_price.plot(xb, lb, color=C["accent"],
                              linewidth=0.6, linestyle="--", alpha=0.5)

            ax_price.legend(fontsize=7, facecolor=C["panel"],
                            labelcolor=C["text"], loc="upper left",
                            framealpha=0.7)
            ax_price.yaxis.set_major_formatter(
                lambda v, _: f"{v/1e6:.1f}M" if v >= 1e6
                else (f"{v/1e3:.0f}K" if v >= 1e3 else f"{v:.0f}"))
            ax_price.tick_params(labelbottom=False)

            # ── Painel RSI ───────────────────────────────────────────────────
            valid_rsi = [(x, v) for x, v in zip(xs_arr, rsi) if v is not None]
            if valid_rsi:
                xr, yr = zip(*valid_rsi)
                ax_rsi.plot(xr, yr, color=C["orange"], linewidth=1.1)
                ax_rsi.axhline(70, color=C["red"],    linewidth=0.7,
                               linestyle="--", alpha=0.6)
                ax_rsi.axhline(30, color=C["green"],  linewidth=0.7,
                               linestyle="--", alpha=0.6)
                ax_rsi.fill_between(xr, 70, yr,
                                    where=[v >= 70 for v in yr],
                                    color=C["red"],   alpha=0.2)
                ax_rsi.fill_between(xr, 30, yr,
                                    where=[v <= 30 for v in yr],
                                    color=C["green"], alpha=0.2)
            ax_rsi.set_ylim(0, 100)
            ax_rsi.set_ylabel("RSI(14)", color=C["subtext"], fontsize=7)
            ax_rsi.tick_params(labelbottom=False)

            # ── Painel Volume ────────────────────────────────────────────────
            bar_colors = [C["green"] if closes[i] >= (closes[i-1] if i > 0 else closes[i])
                          else C["red"] for i in range(len(closes))]
            ax_vol.bar(xs_arr, vols, color=bar_colors, alpha=0.7)
            ax_vol.set_ylabel("Volume", color=C["subtext"], fontsize=7)
            ax_vol.yaxis.set_major_formatter(
                lambda v, _: f"{v/1e6:.1f}M" if v >= 1e6
                else (f"{v/1e3:.0f}K" if v >= 1e3 else f"{v:.0f}"))

            # Eixo X — datas (amostra espaçada)
            step = max(1, len(dates) // 10)
            tick_xs = xs_arr[::step]
            tick_lb = dates[::step]
            ax_vol.set_xticks(tick_xs)
            ax_vol.set_xticklabels(tick_lb, rotation=35, ha="right",
                                    fontsize=7, color=C["subtext"])

            canvas = FigureCanvasTkAgg(fig, master=win)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)

            # Legenda de preço atual
            last = closes[-1]
            prev = closes[-2] if len(closes) > 1 else last
            chg  = (last - prev) / prev * 100 if prev else 0
            color_chg = C["green"] if chg >= 0 else C["red"]
            tk.Label(win,
                     text=f"Último:  {last:,.2f} ISK    "
                          f"Variação: {chg:+.2f}%    "
                          f"Máx 180d: {max(highs):,.0f}    "
                          f"Mín 180d: {min(lows):,.0f}",
                     fg=color_chg, bg=C["bg"],
                     font=("Courier", 11)).pack(pady=3)

        threading.Thread(target=_worker, daemon=True).start()

    def _open_chart(self, route):
        if self._chart_win and self._chart_win.winfo_exists():
            self._chart_win.destroy()
        tid = route["type_id"]
        src_name, dst_name = route["src"], route["dst"]
        src_id, dst_id = REGIONS[src_name], REGIONS[dst_name]
        win = tk.Toplevel(self)
        win.title(f"Historico 90d -- {route['name']}")
        win.configure(bg=COLORS["bg"])
        win.geometry("960x560")
        self._chart_win = win
        tk.Label(win, text=f"Carregando {route['name']}...",
                 fg=COLORS["subtext"], bg=COLORS["bg"],
                 font=("Courier", 11)).pack(pady=8)

        def worker():
            # O1 — cache; busca src e dst em paralelo
            with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
                f_src = pool.submit(_get_hist_cached, src_id, tid)
                f_dst = pool.submit(_get_hist_cached, dst_id, tid)
                h_src = f_src.result()
                h_dst = f_dst.result()
            win.after(0, lambda: _draw(h_src, h_dst))

        def _draw(h_src, h_dst):
            for w in win.winfo_children(): w.destroy()
            cutoff = (datetime.now(timezone.utc) - timedelta(days=90)).strftime("%Y-%m-%d")
            def flt(h):
                return sorted([d for d in h if d.get("date","") >= cutoff],
                               key=lambda x: x["date"])
            fig = Figure(figsize=(11.5, 5.5), facecolor=COLORS["bg"])
            for ai, (data, label, col) in enumerate([
                    (flt(h_src), src_name, COLORS["accent"]),
                    (flt(h_dst), dst_name, COLORS["green"])]):
                ax = fig.add_subplot(1, 2, ai+1)
                ax.set_facecolor(COLORS["panel"])
                ax.tick_params(colors=COLORS["subtext"], labelsize=7)
                for sp in ax.spines.values(): sp.set_edgecolor(COLORS["border"])
                ax.set_title(label, color=COLORS["header"], fontsize=9)
                if not data:
                    ax.text(0.5, 0.5, "Sem dados", ha="center", va="center",
                            color=COLORS["subtext"], transform=ax.transAxes)
                    continue
                dates  = [d["date"]    for d in data]
                prices = [d["average"] for d in data]
                vols   = [d["volume"]  for d in data]
                xs = range(len(dates))
                ax.plot(xs, prices, color=col, linewidth=1.5)
                ax2 = ax.twinx()
                ax2.bar(xs, vols, color=COLORS["yellow"], alpha=0.2)
                ax2.tick_params(colors=COLORS["subtext"], labelsize=7)
                step = max(1, len(dates)//8)
                ax.set_xticks(list(xs)[::step])
                ax.set_xticklabels(dates[::step], rotation=38, ha="right",
                                   fontsize=6, color=COLORS["subtext"])
            fig.tight_layout(pad=1.5)
            canvas = FigureCanvasTkAgg(fig, master=win)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)

        threading.Thread(target=worker, daemon=True).start()

    def _export_rotas(self):
        headers = [("name","Item"),("src","Origem"),("dst","Destino"),
                   ("buy_price","Compra (ISK)"),("sell_price","Venda (ISK)"),
                   ("profit_unit","Lucro/un"),("roi","ROI %"),
                   ("lucro_total","Lucro Total"),("lucro_dia","Lucro/Dia"),
                   ("vol_realizavel","Vol Real"),("capital","Capital")]
        self._export_table("Rotas", headers, self._routes or [], "rotas")

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA 2 — ARBITRAGEM
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_arb(self):
        frame = tk.Frame(self._tab_arb, bg=COLORS["panel"],
                         highlightbackground=COLORS["border"], highlightthickness=1)
        frame.pack(fill="x", padx=6, pady=(6, 2))

        c = self._cfg   # M6
        regions = list(REGIONS.keys())
        self._lbl(frame, "Origem",           0, 0)
        self._a_src   = self._combo(frame, regions, 1, 0)
        self._lbl(frame, "Destino",          0, 2)
        self._a_dst   = self._combo(frame, regions[1:]+[regions[0]], 1, 2)
        self._lbl(frame, "Fonte",            0, 4)
        self._a_mode  = self._combo(frame, ["CSV (cache)", "API (live)"], 1, 4, 13)
        self._csv_age_label(frame, self._a_mode, [self._a_src, self._a_dst], 2, 4)
        self._lbl(frame, "Broker Buy (%)",   0, 6)
        self._a_bbuy  = self._entry(frame, c["a_bbuy"],  1, 6, 7)
        self._lbl(frame, "Broker Sell (%)",  0, 8)
        self._a_bsell = self._entry(frame, c["a_bsell"], 1, 8, 7)
        self._lbl(frame, "Sales Tax (%)",    0, 10)
        self._a_stax  = self._entry(frame, c["a_stax"],  1, 10, 7)
        self._lbl(frame, "Frete (ISK/m3)",   0, 12)
        self._a_fret  = self._entry(frame, c["a_fret"],  1, 12, 8)
        self._lbl(frame, "Margem min (%)",   0, 14)
        self._a_mmin  = self._entry(frame, c["a_mmin"],  1, 14, 7)
        self._lbl(frame, "Lucro Total min",  0, 16)
        self._a_ltmin = self._entry(frame, c["a_ltmin"], 1, 16, 12)

        self._a_btn = tk.Button(frame, text="  CALCULAR ARBITRAGEM  ",
                                command=self._start_arb,
                                bg=COLORS["green"], fg="#000",
                                font=("Courier", 11, "bold"),
                                relief="flat", cursor="hand2")
        self._a_btn.grid(row=1, column=18, padx=8)
        tk.Button(frame, text=" EXPORTAR CSV ", command=self._export_arb,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=19, padx=4)
        # M9: barra deterministica
        self._a_prog = ttk.Progressbar(frame, mode="determinate", length=120,
                                        maximum=100, value=0)
        self._a_prog.grid(row=1, column=20, padx=6)
        self._a_pct = tk.Label(frame, text="", bg=COLORS["panel"],
                               fg=COLORS["subtext"], font=("Courier", 10), width=5)
        self._a_pct.grid(row=1, column=21, padx=2)

        self._a_tree = self._make_tree(self._tab_arb, [
            ("item",          "Item",               210),
            ("loc_origem",    "Estacao Origem",      185),
            ("loc_destino",   "Estacao Destino",     185),
            ("qtd_origem",    "Qtd",                  68),
            ("spread_pct",    "Spread %",             72),
            ("margem_pct",    "Margem Liq %",         88),
            ("preco_origem",  "Compra (ISK)",        113),
            ("preco_destino", "Venda (ISK)",         113),
            ("lucro_liq_unit","Lucro/un",            108),
            ("roi_total",     "ROI Total (ISK)",     138),
            ("lucro_dia",     "Lucro/Dia",           118),
            ("custo_total",   "Custo Total",         128),
            ("vol_total_m3",  "Vol Total m3",         92),
            ("concorr_dst",   "Concorr.",             68),
        ])
        for cid in self._A_COL_NAMES:
            self._a_tree.heading(cid, text=self._A_COL_NAMES[cid],
                                  command=lambda c=cid: self._sort_arb(c))
        self._a_tree.bind("<Double-1>", self._on_dbl_arb)

        bar = tk.Frame(self._tab_arb, bg=COLORS["bg"])
        bar.pack(fill="x", padx=6, pady=(0, 4))
        self._a_count = tk.Label(bar, text="0 oportunidades",
                                  fg=COLORS["subtext"], bg=COLORS["bg"],
                                  font=("Courier", 10))
        self._a_count.pack(side="left")
        self._a_tree.bind("<<TreeviewSelect>>", self._a_on_select)
        self._a_tree.bind("<Double-1>",         self._a_on_dbl)
        tk.Button(bar, text=" GRÁFICO ",     command=self._a_open_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" PREVER ",      command=self._a_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" BACKTEST ",    command=self._a_open_bt,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" ANALISAR IA ", command=self._a_open_ai,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" 🗺 MAPA DA ROTA ",
                  command=self._arb_open_map,
                  bg="#1a3a1a", fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=6)
        tk.Label(bar, text="Double-click -> historico 90d  |  Lucro/Dia = giro medio x lucro/un",
                 fg=COLORS["subtext"], bg=COLORS["bg"],
                 font=("Courier", 10)).pack(side="right")

    _A_COL_NAMES = {
        "item":"Item","loc_origem":"Estacao Origem","loc_destino":"Estacao Destino",
        "qtd_origem":"Qtd","spread_pct":"Spread %","margem_pct":"Margem Liq %",
        "preco_origem":"Compra (ISK)","preco_destino":"Venda (ISK)",
        "lucro_liq_unit":"Lucro/un","roi_total":"ROI Total (ISK)",
        "lucro_dia":"Lucro/Dia","custo_total":"Custo Total",
        "vol_total_m3":"Vol Total m3","concorr_dst":"Concorr.",
    }

    def _start_arb(self):
        if self._loading_arb: return
        self._loading_arb = True
        self._a_btn.config(state="disabled")
        self._a_prog.config(value=0)
        self._a_pct.config(text="")
        params = dict(
            src    = self._a_src.get(),
            dst    = self._a_dst.get(),
            use_csv= self._a_mode.get().startswith("CSV"),
            tax_bb = self._fread(self._a_bbuy,  2.5),
            tax_bs = self._fread(self._a_bsell, 3.5),
            tax_st = self._fread(self._a_stax,  4.0),
            frete  = self._fread(self._a_fret,  100.0),
            mmin   = self._fread(self._a_mmin,  0.0),
            ltmin  = self._fread(self._a_ltmin, 0.0),
        )
        threading.Thread(target=self._worker_arb, args=(params,), daemon=True).start()

    def _worker_arb(self, params: dict):
        prog = self._make_prog(self._a_prog, self._a_pct)
        try:
            src    = params["src"]
            dst    = params["dst"]
            tax_bb = params["tax_bb"]
            tax_bs = params["tax_bs"]
            tax_st = params["tax_st"]
            frete  = params["frete"]
            mmin   = params["mmin"]
            ltmin  = params["ltmin"]

            if params["use_csv"]:
                prog("Carregando CSVs...")
                so = load_orders_from_csv(src)   # M5
                do = load_orders_from_csv(dst)
                if not so or not do:
                    self.after(0, lambda: messagebox.showerror(
                        "CSV nao encontrado",
                        "Baixe os dados primeiro na aba DADOS & CACHE."))
                    return
            else:
                prog("Baixando ordens via API...")
                with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
                    fs = pool.submit(download_region_orders, REGIONS[src], src, prog)
                    fd = pool.submit(download_region_orders, REGIONS[dst], dst, prog)
                    so, do = fs.result(), fd.result()
                _orders_mem_cache[src] = (time.time(), so)   # M5
                _orders_mem_cache[dst] = (time.time(), do)

            opps = run_arbitrage(
                so, do, src, dst,
                tax_bb, tax_bs, tax_st, frete,
                self._station_cache, self._item_cache,
                min_margem=mmin, min_lucro_total=ltmin,
                progress_cb=prog)
            self._arb_opps = opps
            self.after(0, lambda: (self._a_prog.config(value=100),
                                   self._a_pct.config(text="100%")))
            self.after(0, self._pop_arb)
            self.after(0, self._dash_refresh)   # Tier 2 — atualiza Dashboard

            dst_region_id = REGIONS.get(dst)
            if dst_region_id and opps:
                prog("Buscando historico para Lucro/Dia...")
                enrich_arb_lucro_dia(opps, dst_region_id, prog)
                self.after(0, self._pop_arb)

        except Exception as exc:
            log.exception("worker_arb")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro", e))
        finally:
            self._loading_arb = False
            self.after(0, lambda: self._a_btn.config(state="normal"))

    def _pop_arb(self, reorder_only=False):
        if reorder_only:
            for i in range(len(self._arb_opps)):
                self._a_tree.move(str(i), "", i)
            return
        for row in self._a_tree.get_children(): self._a_tree.delete(row)
        for i, r in enumerate(self._arb_opps):
            tag = ("hot",) if r["margem_pct"] >= 20 else ("even" if i%2==0 else "odd",)
            ld  = (self._fmt(r["lucro_dia"], "isk") if r.get("lucro_dia") is not None
                   else "ERRO" if r.get("hist_erro") else "--")
            cc  = r.get("concorr_dst", 0)
            # iid = str(i) para garantir unicidade mesmo que o mesmo type_id
            # apareça em múltiplos pares de estações (origem/destino distintos).
            # O type_id real é armazenado via _a_on_select ao clicar.
            self._a_tree.insert("", "end", iid=str(i), tags=tag, values=(
                r["item"],
                r["loc_origem"],
                r["loc_destino"],
                self._fmt(r["qtd_origem"],    "int"),
                self._fmt(r["spread_pct"],    "pct"),
                self._fmt(r["margem_pct"],    "pct"),
                self._fmt(r["preco_origem"],  "isk"),
                self._fmt(r["preco_destino"], "isk"),
                self._fmt(r["lucro_liq_unit"],"isk"),
                self._fmt(r["roi_total"],     "isk"),
                ld,
                self._fmt(r["custo_total"],   "isk"),
                self._fmt(r["vol_total_m3"],  "vol"),
                f"{cc}",
            ))
        total_roi = sum(x["roi_total"] for x in self._arb_opps)
        self._a_count.config(
            text=f"{len(self._arb_opps)} oportunidades  |  "
                 f"ROI total: {self._fmt(total_roi, 'isk')} ISK  |  "
                 f"Concorr. = vendedores ja posicionados abaixo do preco planejado")

    def _sort_arb(self, col):
        num_cols = {"qtd_origem","spread_pct","margem_pct","preco_origem",
                    "preco_destino","lucro_liq_unit","roi_total","lucro_dia",
                    "custo_total","vol_total_m3","concorr_dst"}
        default_rev = col in num_cols
        rev = not self._a_sort_state.get(col, not default_rev)
        self._a_sort_state = {col: rev}
        self._arb_opps.sort(
            key=lambda r: (r.get(col) or 0) if col in num_cols else r.get(col, ""),
            reverse=rev)
        self._set_heading_arrow(self._a_tree, self._A_COL_NAMES, col, rev)
        self._pop_arb(reorder_only=True)

    def _on_dbl_arb(self, _):
        """Double-click na aba Arbitragem: abre historico 90d igual ao da aba Rotas."""
        sel = self._a_tree.selection()
        if not sel: return
        idx = int(sel[0])
        if idx >= len(self._arb_opps): return
        if not HAS_MATPLOTLIB:
            messagebox.showinfo("matplotlib ausente", "pip install matplotlib")
            return
        opp = self._arb_opps[idx]
        # Reutiliza _open_chart passando um dict compativel com o formato de rota
        route_compat = {
            "type_id":  opp["type_id"],
            "name":     opp["item"],
            "src":      opp["src"],
            "dst":      opp["dst"],
        }
        self._open_chart(route_compat)

    def _export_arb(self):
        headers = [("item","Item"),("loc_origem","Origem"),("loc_destino","Destino"),
                   ("spread_pct","Spread %"),("margem_pct","Margem %"),
                   ("lucro_liq_unit","Lucro/un"),("lucro_dia","Lucro/Dia"),
                   ("giro_dia","Giro/Dia"),("concorr_dst","Concorrência")]
        self._export_table("Arbitragem", headers, self._arb_opps or [], "arbitragem")
    def _build_tab_opt(self):
        if not HAS_ORTOOLS:
            tk.Label(self._tab_opt,
                     text="ortools nao instalado.\n\npip install ortools",
                     bg=COLORS["bg"], fg=COLORS["yellow"],
                     font=("Courier", 14, "bold")).pack(expand=True)
            return

        c = self._cfg   # M6
        frame = tk.Frame(self._tab_opt, bg=COLORS["panel"],
                         highlightbackground=COLORS["border"], highlightthickness=1)
        frame.pack(fill="x", padx=6, pady=(6, 2))

        self._lbl(frame, "Capacidade (m3)",  0, 0)
        self._o_cap    = self._entry(frame, c["o_cap"],    1, 0, 12)
        self._lbl(frame, "Lucro min (ISK)",  0, 2)
        self._o_lucro  = self._entry(frame, c["o_lucro"],  1, 2, 12)
        self._lbl(frame, "Margem min (%)",   0, 4)
        self._o_marg   = self._entry(frame, c["o_marg"],   1, 4, 8)
        self._lbl(frame, "ISK/Pulo min",     0, 6)
        self._o_ispulo = self._entry(frame, c["o_ispulo"], 1, 6, 12)
        self._lbl(frame, "Max pulos",        0, 8)
        self._o_pulos  = self._entry(frame, c["o_pulos"],  1, 8, 6)
        self._lbl(frame, "Filtro origem",    0, 10)
        self._o_forig  = self._entry(frame, "",            1, 10, 12)
        self._lbl(frame, "Filtro destino",   0, 12)
        self._o_fdest  = self._entry(frame, "",            1, 12, 12)

        self._o_btn = tk.Button(frame, text="  OTIMIZAR CARGA  ",
                                command=self._start_opt,
                                bg=COLORS["yellow"], fg="#000",
                                font=("Courier", 11, "bold"),
                                relief="flat", cursor="hand2")
        self._o_btn.grid(row=1, column=14, padx=8)
        tk.Button(frame, text=" EXPORTAR CSV ", command=self._export_opt,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=15, padx=4)
        tk.Button(frame, text=" LISTA COMPRAS ", command=self._export_lista_compras,
                  bg=COLORS["green"], fg="#000",
                  font=("Courier", 11, "bold"), relief="flat", cursor="hand2"
                  ).grid(row=1, column=16, padx=4)
        self._o_prog = ttk.Progressbar(frame, mode="indeterminate", length=90)
        self._o_prog.grid(row=1, column=17, padx=6)

        # M8 — pool multi-rota
        pool_bar = tk.Frame(self._tab_opt, bg=COLORS["bg"])
        pool_bar.pack(fill="x", padx=8, pady=(4, 0))
        tk.Label(pool_bar,
                 text="  Pool multi-rota: acumule oportunidades de varios pares "
                      "antes de otimizar.",
                 bg=COLORS["bg"], fg=COLORS["subtext"],
                 font=("Courier", 10)).pack(side="left")
        tk.Button(pool_bar, text=" + ACUMULAR ARB ",
                  command=self._pool_add,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=(12, 4))
        tk.Button(pool_bar, text=" LIMPAR POOL ",
                  command=self._pool_clear,
                  bg=COLORS["border"], fg=COLORS["red"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)
        self._o_pool_lbl = tk.Label(pool_bar, text="Pool: 0 oportunidades",
                                     bg=COLORS["bg"], fg=COLORS["subtext"],
                                     font=("Courier", 10))
        self._o_pool_lbl.pack(side="left", padx=12)

        self._o_tree = self._make_tree(self._tab_opt, [
            ("Item",           "Item",           225),
            ("Origem",         "Origem",         205),
            ("Destino",        "Destino",        205),
            ("Pulos",          "Pulos",           52),
            ("Qtd",            "Qtd",             58),
            ("Vol Total (m3)", "Vol Total (m3)", 108),
            ("Lucro Total",    "Lucro Total",    128),
            ("ISK/Pulo",       "ISK/Pulo",       108),
            ("Margem (%)",     "Margem (%)",      82),
        ])

        self._o_stats = tk.Label(self._tab_opt, text="--",
                                  bg=COLORS["panel"], fg=COLORS["green"],
                                  font=("Courier", 11, "bold"))
        self._o_stats.pack(fill="x", padx=6, pady=(2, 4))

        # Barra de ações do Otimizador
        opt_act = tk.Frame(self._tab_opt, bg=COLORS["bg"])
        opt_act.pack(fill="x", padx=6, pady=(0, 4))
        tk.Button(opt_act, text=" GRÁFICO ",     command=self._opt_open_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(opt_act, text=" PREVER ",      command=self._opt_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(opt_act, text=" BACKTEST ",    command=self._opt_open_bt,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(opt_act, text=" ANALISAR IA ", command=self._opt_open_ai,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(opt_act, text=" 🗺 MAPA DA ROTA ",
                  command=self._opt_open_map,
                  bg="#1a3a1a", fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=6)
        self._o_tree.bind("<<TreeviewSelect>>", self._opt_on_select)
        self._o_tree.bind("<Double-1>",         self._opt_on_dbl)

    def _pool_add(self):
        """M8: adiciona oportunidades da aba Arbitragem ao pool do Otimizador."""
        if not self._arb_opps:
            messagebox.showinfo("Sem dados",
                "Calcule a Arbitragem primeiro (aba ARBITRAGEM)."); return
        # Deduplicacao por (type_id, loc_origem, loc_destino)
        existing = {(o["type_id"], o["loc_origem"], o["loc_destino"])
                    for o in self._opt_pool}
        added = 0
        for opp in self._arb_opps:
            key = (opp["type_id"], opp["loc_origem"], opp["loc_destino"])
            if key not in existing:
                self._opt_pool.append(opp)
                existing.add(key)
                added += 1
        self._o_pool_lbl.config(
            text=f"Pool: {len(self._opt_pool)} oportunidades  (+{added} adicionadas)",
            fg=COLORS["green"])

    def _pool_clear(self):
        """M8: limpa o pool multi-rota."""
        self._opt_pool.clear()
        self._o_pool_lbl.config(text="Pool: 0 oportunidades", fg=COLORS["subtext"])

    def _start_opt(self):
        if self._loading_opt: return
        source = self._opt_pool if self._opt_pool else self._arb_opps
        if not source:
            messagebox.showinfo("Sem dados",
                "Calcule Arbitragem ou acumule pares no pool."); return
        self._loading_opt = True
        self._o_btn.config(state="disabled")
        self._o_prog.start(12)
        params = dict(
            capacity_m3    = self._fread(self._o_cap,    1_600_000),
            min_lucro      = self._fread(self._o_lucro,  10_000),
            min_margem     = self._fread(self._o_marg,   1.5),
            min_isk_pulo   = self._fread(self._o_ispulo, 10_000),
            max_pulos      = int(self._fread(self._o_pulos, 15)),
            filtro_origem  = self._o_forig.get().strip(),
            filtro_destino = self._o_fdest.get().strip(),
        )
        threading.Thread(target=self._worker_opt, args=(source, params),
                         daemon=True).start()

    def _worker_opt(self, source, params: dict):
        try:
            plan, stats = solve_cargo(
                source,
                capacity_m3    = params["capacity_m3"],
                min_lucro      = params["min_lucro"],
                min_margem     = params["min_margem"],
                min_isk_pulo   = params["min_isk_pulo"],
                max_pulos      = params["max_pulos"],
                filtro_origem  = params["filtro_origem"],
                filtro_destino = params["filtro_destino"],
                progress_cb    = self._set_status,
            )
            self._opt_plan = plan
            self.after(0, lambda: self._pop_opt(plan, stats))
        except Exception as exc:
            log.exception("worker_opt")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro", e))
        finally:
            self._loading_opt = False
            self.after(0, lambda: (self._o_prog.stop(),
                                   self._o_btn.config(state="normal")))

    def _pop_opt(self, plan, stats):
        for row in self._o_tree.get_children(): self._o_tree.delete(row)

        if "erro" in stats:
            self._o_stats.config(text=f"  {stats['erro']}", fg=COLORS["red"])
            return

        for i, r in enumerate(plan):
            tag = ("hot",) if r["Margem (%)"] >= 10 else ("even" if i%2==0 else "odd",)
            self._o_tree.insert("", "end", iid=str(i), tags=tag, values=(
                r["Item"], r["Origem"], r["Destino"],
                r["Pulos"],
                self._fmt(r["Qtd"],            "int"),
                self._fmt(r["Vol Total (m3)"], "vol"),
                self._fmt(r["Lucro Total"],    "isk"),
                self._fmt(r["ISK/Pulo"],       "isk"),
                self._fmt(r["Margem (%)"],     "pct"),
            ))

        pct = (stats["vol_total"] / stats["capacity"] * 100) if stats["capacity"] else 0
        self._o_stats.config(
            fg=COLORS["green"],
            text=(f"  {stats['n_items']} itens  |  "
                  f"Vol: {stats['vol_total']:,.0f} / {stats['capacity']:,.0f} m3 ({pct:.1f}%)  |  "
                  f"Lucro estimado: {stats['lucro_total']:,.2f} ISK  |  "
                  f"Margem media: {stats['margem_media']:.1f}%"))

    def _export_opt(self):
        headers = [("Item","Item"),("Origem","Origem"),("Destino","Destino"),
                   ("Pulos","Pulos"),("Qtd","Qtd"),("Vol Total (m3)","Vol (m³)"),
                   ("Lucro Total","Lucro Total"),("ISK/Pulo","ISK/Pulo"),("Margem (%)","Margem %")]
        self._export_table("Otimizador", headers, self._opt_plan or [], "otimizador")
    def _export_lista_compras(self):
        """
        Exporta a Lista de Compras do plano do Otimizador:
        Item | Estacao de Origem | Qtd | Capital Necessario | Lucro Estimado | Margem %
        Acompanhada de um resumo de capital total necessario.
        """
        if not self._opt_plan:
            messagebox.showinfo("Vazio", "Execute o Otimizador primeiro."); return

        # Precisa das oportunidades originais para ter o preco de compra
        source = self._opt_pool if self._opt_pool else self._arb_opps
        price_map = {}   # (item, origem) -> preco_origem
        if source:
            for opp in source:
                key = (opp.get("item",""), opp.get("loc_origem",""))
                price_map[key] = opp.get("preco_origem", 0)

        rows = []
        total_capital = 0.0
        total_lucro   = 0.0
        for r in self._opt_plan:
            key          = (r["Item"], r["Origem"])
            preco_compra = price_map.get(key, 0)
            capital      = preco_compra * r["Qtd"]
            total_capital += capital
            total_lucro   += r["Lucro Total"]
            rows.append({
                "Item":              r["Item"],
                "Estacao Origem":    r["Origem"],
                "Estacao Destino":   r["Destino"],
                "Pulos":             r["Pulos"],
                "Qtd":               r["Qtd"],
                "Preco Compra (ISK)":round(preco_compra, 2),
                "Capital (ISK)":     round(capital, 2),
                "Lucro Estimado":    round(r["Lucro Total"], 2),
                "Margem (%)":        r["Margem (%)"],
                "Vol Total (m3)":    r["Vol Total (m3)"],
            })

        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=([("Excel (.xlsx)","*.xlsx"),("CSV","*.csv")]
                       if HAS_OPENPYXL else [("CSV","*.csv")]),
            initialfile="lista_compras.csv")
        if not path: return

        # Adiciona linha de totais como última entrada
        rows.append({
            "Item":              "─── TOTAIS ───",
            "Estacao Origem":    "",
            "Estacao Destino":   "",
            "Pulos":             "",
            "Qtd":               "",
            "Preco Compra (ISK)":"",
            "Capital (ISK)":     round(total_capital, 2),
            "Lucro Estimado":    round(total_lucro, 2),
            "Margem (%)":        round(total_lucro/total_capital*100
                                       if total_capital else 0, 1),
            "Vol Total (m3)":    "",
        })
        headers = [("Item","Item"),("Estacao Origem","Origem"),
                   ("Estacao Destino","Destino"),("Pulos","Pulos"),
                   ("Qtd","Qtd"),("Preco Compra (ISK)","Preço Compra"),
                   ("Capital (ISK)","Capital (ISK)"),
                   ("Lucro Estimado","Lucro Estimado"),
                   ("Margem (%)","Margem %"),("Vol Total (m3)","Vol (m³)")]
        self._export_table("Lista de Compras", headers, rows, "lista_compras")
    # ══════════════════════════════════════════════════════════════════════════
    #  ABA CICLOS — Deteccao de ciclos de mercado (Tier 3)
    # ══════════════════════════════════════════════════════════════════════════

    _CYC_COLS = {
        "name":        "Item",
        "group":       "Grupo",
        "phase":       "Fase",
        "trend_score": "Trend",
        "vol_score":   "Vol Score",
        "momentum_7d": "Var 7d %",
        "momentum_30d":"Var 30d %",
        "volatility":  "Volatil.",
        "rsi":         "RSI(14)",
        "anomaly_msg": "Anomalia",
    }

    def _build_tab_cycle(self):
        p = self._tab_cycle

        # ── Painel de controle ────────────────────────────────────────────────
        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(6, 2))

        self._lbl(fr, "Regiao",        0, 0)
        self._cyc_reg = self._combo(fr, list(REGIONS.keys()), 1, 0)

        self._lbl(fr, "Dias de hist.", 0, 2)
        self._cyc_days = self._entry(fr, "180", 1, 2, 6)

        self._lbl(fr, "Grupos de itens", 0, 4)
        self._cyc_groups = self._combo(
            fr, ["Todos"] + list(CYCLE_ITEM_GROUPS.keys()), 1, 4, 30)

        self._cyc_btn = tk.Button(fr, text="  ANALISAR CICLOS  ",
                                   command=self._start_cycle,
                                   bg="#7f4adf", fg="#fff",
                                   font=("Courier", 11, "bold"),
                                   relief="flat", cursor="hand2")
        self._cyc_btn.grid(row=1, column=6, padx=8)

        tk.Button(fr, text=" GRAFICO ITEM ",
                  command=self._cycle_chart_item,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=7, padx=4)

        tk.Button(fr, text=" GRAFICO REGIAO ",
                  command=self._cycle_chart_region,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=8, padx=4)

        tk.Button(fr, text=" + WATCHLIST ",
                  command=self._cycle_add_wl,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=9, padx=4)

        tk.Button(fr, text=" SPREAD HIST. ",
                  command=self._cycle_open_spread,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=10, padx=4)

        tk.Button(fr, text=" EXPORTAR ",
                  command=self._export_cycle,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=11, padx=4)

        self._cyc_prog = ttk.Progressbar(fr, mode="determinate",
                                          length=120, maximum=100)
        self._cyc_prog.grid(row=1, column=12, padx=6)
        self._cyc_pct = tk.Label(fr, text="", bg=COLORS["panel"],
                                  fg=COLORS["subtext"], font=("Courier", 10), width=5)
        self._cyc_pct.grid(row=1, column=13, padx=2)

        # Linha de dica
        tk.Label(fr,
                 text="Double-click = grafico detalhado do item  |  "
                      "'Todos' analisa todos os grupos predefinidos",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 9)
                 ).grid(row=0, column=6, columnspan=8, padx=4, sticky="w")

        # Campo de Type IDs adicionais (coluna 0-5, linha 2)
        tk.Label(fr, text="Type IDs extras:", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 10)
                 ).grid(row=2, column=0, padx=4, pady=(0, 4), sticky="w")
        self._cyc_custom = tk.Entry(fr, width=50,
                                     bg=COLORS["border"], fg=COLORS["text"],
                                     insertbackground=COLORS["text"],
                                     font=("Courier", 11), relief="flat", bd=2)
        self._cyc_custom.grid(row=2, column=1, columnspan=5,
                               padx=4, pady=(0, 4), sticky="w")
        tk.Label(fr, text="(type_ids separados por virgula, opcional)",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 9)
                 ).grid(row=2, column=6, columnspan=4, padx=4, sticky="w")

        # ── Painel de resumo da regiao ────────────────────────────────────────
        sum_fr = tk.Frame(p, bg=COLORS["panel"],
                          highlightbackground=COLORS["border"], highlightthickness=1)
        sum_fr.pack(fill="x", padx=6, pady=(2, 2))

        self._cyc_summary_lbl = tk.Label(
            sum_fr,
            text="Execute ANALISAR CICLOS para ver o resumo da regiao.",
            bg=COLORS["panel"], fg=COLORS["subtext"],
            font=("Courier", 11), justify="left", anchor="w")
        self._cyc_summary_lbl.pack(fill="x", padx=10, pady=6)

        self._cyc_systemic_lbl = tk.Label(
            sum_fr, text="",
            bg=COLORS["panel"], fg=COLORS["red"],
            font=("Courier", 11, "bold"), justify="left", anchor="w")
        self._cyc_systemic_lbl.pack(fill="x", padx=10, pady=(0, 4))

        # ── Tabela de itens ───────────────────────────────────────────────────
        self._cyc_tree = self._make_tree(p, [
            ("name",        "Item",        200),
            ("group",       "Grupo",       160),
            ("phase",       "Fase",         96),
            ("trend_score", "Trend",         72),
            ("vol_score",   "Vol Score",     80),
            ("momentum_7d", "Var 7d %",      80),
            ("momentum_30d","Var 30d %",     80),
            ("volatility",  "Volatil.",      72),
            ("rsi",         "RSI(14)",       68),
            ("anomaly_msg", "Anomalia",     280),
        ])
        for cid in self._CYC_COLS:
            self._cyc_tree.heading(cid, text=self._CYC_COLS[cid],
                                    command=lambda c=cid: self._sort_cycle(c))
        self._cyc_tree.bind("<Double-1>", self._on_dbl_cycle)

        # Tags de cor por fase
        for phase, info in CYCLE_PHASES.items():
            self._cyc_tree.tag_configure(phase, foreground=info["color"])
        self._cyc_tree.tag_configure("anomaly",
                                      foreground=COLORS["red"],
                                      font=("Courier", 11, "bold"))

        bar = tk.Frame(p, bg=COLORS["bg"])
        bar.pack(fill="x", padx=6, pady=(0, 4))
        self._cyc_count = tk.Label(bar, text="0 itens",
                                    fg=COLORS["subtext"], bg=COLORS["bg"],
                                    font=("Courier", 10))
        self._cyc_count.pack(side="left")
        self._cyc_tree.bind("<<TreeviewSelect>>", self._cyc_on_select)
        self._cyc_tree.bind("<Double-1>",         self._cyc_on_dbl)
        tk.Button(bar, text=" GRÁFICO ",     command=self._cyc_open_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" PREVER ",      command=self._cyc_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" BACKTEST ",    command=self._cyc_open_bt,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" ANALISAR IA ", command=self._cyc_open_ai,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        # Legenda de fases
        legend = tk.Frame(bar, bg=COLORS["bg"])
        legend.pack(side="right")
        for phase, info in CYCLE_PHASES.items():
            tk.Label(legend, text=f"{info['emoji']} {phase}  ",
                     fg=info["color"], bg=COLORS["bg"],
                     font=("Courier", 10)).pack(side="left")

    # ── Workers e handlers ────────────────────────────────────────────────────

    def _start_cycle(self):
        if self._loading_cycle: return
        self._loading_cycle = True
        self._cyc_btn.config(state="disabled")
        self._cyc_prog.config(value=0); self._cyc_pct.config(text="")
        params = dict(
            rname   = self._cyc_reg.get(),
            days    = max(30, int(self._fread(self._cyc_days, 180))),
            sel_grp = self._cyc_groups.get(),
            custom  = self._cyc_custom.get().strip(),
        )
        threading.Thread(target=self._worker_cycle, args=(params,), daemon=True).start()

    def _worker_cycle(self, params: dict):
        prog = self._make_prog(self._cyc_prog, self._cyc_pct)
        try:
            rname    = params["rname"]
            rid      = REGIONS.get(rname, 10000002)
            days     = params["days"]
            sel_grp  = params["sel_grp"]

            # Filtra grupos se usuario selecionou um especifico
            if sel_grp == "Todos":
                groups = dict(CYCLE_ITEM_GROUPS)
            else:
                groups = {sel_grp: CYCLE_ITEM_GROUPS.get(sel_grp, {})}

            # Adiciona type_ids personalizados como grupo "Custom"
            custom_raw = params["custom"]
            if custom_raw:
                try:
                    custom_tids = [int(x.strip()) for x in custom_raw.split(",")
                                   if x.strip().isdigit()]
                    if custom_tids:
                        groups["Custom"] = {tid: str(tid) for tid in custom_tids}
                except Exception as _e:

                    log.debug("silenced: %s", _e)

            # Resolve nomes de TODOS os itens em 1 única chamada (era 2 chamadas)
            all_tids = [tid for g in groups.values() for tid in g]
            _ensure_names(all_tids, self._item_cache)
            # Preenche nomes no grupo Custom agora que estão no cache
            if "Custom" in groups:
                groups["Custom"] = {
                    tid: self._item_cache.get(tid, {}).get("name", str(tid))
                    for tid in groups["Custom"]}

            result = detect_cycle_region(rid, rname,
                                          item_groups=groups,
                                          progress_cb=prog)
            # Re-analisa com janela de dias personalizada se diferente do default
            if days != 180:
                for tid, sig in result["items"].items():
                    hist  = _get_hist_cached(rid, tid)
                    new_s = detect_cycle_item(hist, days_back=days)
                    new_s["type_id"] = tid
                    new_s["name"]    = sig["name"]
                    new_s["group"]   = sig["group"]
                    result["items"][tid] = new_s

            self._cycle_result = result
            self.after(0, lambda: (self._cyc_prog.config(value=100),
                                   self._cyc_pct.config(text="100%")))
            self.after(0, self._pop_cycle)
            self.after(0, self._dash_refresh)
        except Exception as exc:
            log.exception("worker_cycle")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro Ciclo", e))
        finally:
            self._loading_cycle = False
            self.after(0, lambda: self._cyc_btn.config(state="normal"))

    def _cycle_add_wl(self):
        """Adiciona o item selecionado na tabela de ciclos à Watchlist."""
        if not self._cycle_result:
            messagebox.showinfo("Sem dados", "Execute ANALISAR CICLOS primeiro."); return
        sel = self._cyc_tree.selection()
        if not sel:
            messagebox.showinfo("Selecione", "Clique num item da tabela."); return
        try:
            tid = int(sel[0])
        except ValueError:
            return
        sig = self._cycle_result["items"].get(tid)
        if not sig:
            return
        if any(w["type_id"] == tid for w in self._watchlist):
            messagebox.showinfo("Já existe", f"{sig['name']} já está na Watchlist.")
            return
        self._watchlist.append({
            "type_id": tid, "name": sig["name"],
            "alert_below": None, "alert_above": None,
            "last_sell": None, "last_buy": None,
            "last_net": None, "last_score": None,
        })
        save_watchlist(self._watchlist)
        self._pop_wl()
        messagebox.showinfo("Adicionado",
            f"{sig['name']} adicionado à Watchlist.\n"
            f"Fase atual: {sig['phase']}")

    def _cycle_open_spread(self):
        """Abre Spread Hist. para o item selecionado na tabela de ciclos."""
        if not self._cycle_result:
            messagebox.showinfo("Sem dados", "Execute ANALISAR CICLOS primeiro."); return
        sel = self._cyc_tree.selection()
        if not sel:
            messagebox.showinfo("Selecione", "Clique num item da tabela."); return
        try:
            tid = int(sel[0])
        except ValueError:
            return
        sig = self._cycle_result["items"].get(tid)
        if not sig:
            return
        self._sph_tid.set(str(tid))
        self._sph_name.set(sig["name"])
        self._sph_reg.set(self._cyc_reg.get())
        self._nb.select(self._tab_sph)
        self.after(50, self._start_sph)

    def _pop_cycle(self):
        r = self._cycle_result
        if not r:
            return

        # ── Resumo da regiao ──────────────────────────────────────────────────
        phase    = r["region_phase"]
        emoji    = CYCLE_PHASES[phase]["emoji"]
        color    = r["region_color"]
        votes    = r["phase_votes"]
        vote_str = "  ".join(
            f"{CYCLE_PHASES[p]['emoji']}{p[:4]}:{votes[p]}"
            for p in CYCLE_PHASES if votes[p] > 0)

        summary = (
            f"{emoji}  FASE DA REGIAO: {phase.upper()}  "
            f"|  Confianca: {r['confidence']:.0f}%  "
            f"|  Trend medio: {r['avg_trend']:+.3f}  "
            f"|  Vol medio: {r['avg_vol']:.2f}x  "
            f"|  Anomalias: {len(r['anomalies'])}  "
            f"|  [{vote_str}]")
        self._cyc_summary_lbl.config(text=summary, fg=color)

        # Evento sistemico
        if r["systemic_event"]:
            self._cyc_systemic_lbl.config(
                text=f"  ⚠  {r['systemic_msg']}", fg=COLORS["red"])
            # Registra no historico de alertas
            ts = datetime.now().strftime("%H:%M")
            self._alert_history.append((ts, f"[CICLO] {r['systemic_msg'][:80]}"))
        else:
            self._cyc_systemic_lbl.config(text="")

        # ── Tabela de itens ───────────────────────────────────────────────────
        for row in self._cyc_tree.get_children():
            self._cyc_tree.delete(row)

        items = sorted(r["items"].values(),
                       key=lambda x: abs(x["trend_score"]), reverse=True)

        for i, sig in enumerate(items):
            phase_i = sig["phase"]
            tag     = ("anomaly",) if sig["anomaly"] else (phase_i,)
            rsi_txt = f"{sig['rsi']:.1f}" if sig["rsi"] is not None else "--"
            anom    = sig.get("anomaly_msg", "")
            self._cyc_tree.insert("", "end", iid=str(sig["type_id"]),
                                   tags=tag, values=(
                sig["name"],
                sig["group"],
                f"{CYCLE_PHASES[phase_i]['emoji']} {phase_i}",
                f"{sig['trend_score']:+.4f}",
                f"{sig['vol_score']:.2f}x",
                f"{sig['momentum_7d']:+.2f}%",
                f"{sig['momentum_30d']:+.2f}%",
                f"{sig['volatility']:.4f}",
                rsi_txt,
                anom,
            ))

        n_anom = len(r["anomalies"])
        self._cyc_count.config(
            text=(f"{len(items)} itens analisados  |  "
                  f"{n_anom} anomalia(s) detectada(s)  |  "
                  f"Fase: {phase}"))

    def _sort_cycle(self, col):
        if not self._cycle_result:
            return
        num = {"trend_score","vol_score","momentum_7d","momentum_30d",
               "volatility","rsi"}
        rev = not self._cycle_sort_state.get(col, col in num)
        self._cycle_sort_state = {col: rev}
        items = list(self._cycle_result["items"].values())
        items.sort(
            key=lambda x: (x.get(col) or 0) if col in num
            else x.get(col, ""),
            reverse=rev)
        # P5 — move() em vez de delete+reinsert (sem flickering, O(n) em vez de O(n*insert))
        for i, sig in enumerate(items):
            self._cyc_tree.move(str(sig["type_id"]), "", i)
        self._set_heading_arrow(self._cyc_tree, self._CYC_COLS, col, rev)

    def _on_dbl_cycle(self, _):
        """Double-click: abre grafico de ciclo do item selecionado."""
        self._cycle_chart_item()

    def _cycle_chart_item(self):
        """Grafico detalhado de ciclo para o item selecionado na tabela."""
        if not HAS_MATPLOTLIB:
            messagebox.showinfo("Falta matplotlib", "pip install matplotlib"); return
        if not self._cycle_result:
            messagebox.showinfo("Sem dados", "Execute ANALISAR CICLOS primeiro."); return
        sel = self._cyc_tree.selection()
        if not sel:
            messagebox.showinfo("Selecione", "Clique num item da tabela."); return
        try:
            tid = int(sel[0])
        except ValueError:
            return
        sig = self._cycle_result["items"].get(tid)
        if not sig:
            return

        if self._chart_win and self._chart_win.winfo_exists():
            self._chart_win.destroy()
        win = tk.Toplevel(self)
        win.title(f"Ciclo de Mercado — {sig['name']}")
        win.configure(bg=COLORS["bg"])
        win.geometry("1100x700")
        self._chart_win = win

        dates  = sig["dates"]
        prices = sig["prices"]
        vols   = sig["volumes"]
        ema_s  = sig["ema_short"]
        ema_l  = sig["ema_long"]
        n      = len(dates)
        if n < 2:
            tk.Label(win, text="Dados insuficientes.",
                     bg=COLORS["bg"], fg=COLORS["red"],
                     font=("Courier", 12)).pack(pady=20)
            return

        xs     = list(range(n))
        bg     = COLORS["bg"]
        pan    = COLORS["panel"]
        phase  = sig["phase"]
        pcolor = sig["phase_color"]

        fig = Figure(figsize=(13, 8), facecolor=bg)
        fig.subplots_adjust(left=0.07, right=0.96, hspace=0.32,
                            top=0.91, bottom=0.09)

        def _style(ax):
            ax.set_facecolor(pan)
            ax.tick_params(colors=COLORS["subtext"], labelsize=7)
            for sp in ax.spines.values():
                sp.set_edgecolor(COLORS["border"])

        # ── Subplot 1: Preco + EMAs ───────────────────────────────────────────
        ax1 = fig.add_subplot(3, 1, 1)
        _style(ax1)
        ax1.plot(xs, prices, color=COLORS["text"], lw=1.2, alpha=0.7,
                 label="Preco")

        # EMA curta e longa
        xs_s = [x for x, v in zip(xs, ema_s) if v is not None]
        ys_s = [v for v in ema_s if v is not None]
        xs_l = [x for x, v in zip(xs, ema_l) if v is not None]
        ys_l = [v for v in ema_l if v is not None]
        if xs_s: ax1.plot(xs_s, ys_s, color=COLORS["accent"],  lw=1.5, label="EMA14")
        if xs_l: ax1.plot(xs_l, ys_l, color=COLORS["yellow"],  lw=1.5, label="EMA50")

        # Colorir fundo pela fase detectada
        ax1.set_facecolor(pcolor + "18")   # cor da fase com alpha ~10%

        # Marca anomalia com linha vermelha vertical
        if sig["anomaly"]:
            ax1.axvline(xs[-7] if n > 7 else xs[-1],
                        color=COLORS["red"], lw=1.0, ls="--", alpha=0.7)
            ax1.text(xs[-7] if n > 7 else xs[-1], prices[-1],
                     "  ⚠ anomalia", color=COLORS["red"],
                     fontsize=7, va="bottom")

        ax1.set_title(
            f"{sig['name']}  |  Fase: {CYCLE_PHASES[phase]['emoji']} {phase}  "
            f"|  Trend: {sig['trend_score']:+.4f}  "
            f"|  Momentum 7d: {sig['momentum_7d']:+.2f}%  "
            f"|  RSI: {sig['rsi'] if sig['rsi'] is not None else '--'}",
            color=pcolor, fontsize=9, pad=5)
        ax1.set_ylabel("ISK", color=COLORS["subtext"], fontsize=7)
        ax1.legend(loc="upper left", fontsize=6,
                   facecolor=pan, labelcolor=COLORS["text"], framealpha=0.8)
        ax1.tick_params(labelbottom=False)

        # ── Subplot 2: Volume + score ─────────────────────────────────────────
        ax2 = fig.add_subplot(3, 1, 2)
        _style(ax2)
        # Media movel do volume (14d)
        vol_ema = _cycle_ema([float(v) for v in vols], 14)
        avg_v   = sum(vols) / n if n else 1
        vcols   = [COLORS["red"]    if v >= avg_v * 3.0
                   else (COLORS["green"] if v >= avg_v * 1.5
                   else COLORS["subtext"])
                   for v in vols]
        ax2.bar(xs, vols, color=vcols, alpha=0.6, width=0.8)
        xs_ve = [x for x, v in zip(xs, vol_ema) if v is not None]
        ys_ve = [v for v in vol_ema if v is not None]
        if xs_ve:
            ax2.plot(xs_ve, ys_ve, color=COLORS["orange"], lw=1.2,
                     label="EMA14 Volume")
        ax2.axhline(avg_v,       color=COLORS["yellow"], lw=0.8, ls="--",
                    label=f"Media {avg_v:,.0f}")
        ax2.axhline(avg_v * 3.0, color=COLORS["red"],    lw=0.6, ls=":",
                    label="Spike 3x")
        ax2.set_ylabel("Volume", color=COLORS["subtext"], fontsize=7)
        ax2.legend(loc="upper left", fontsize=6,
                   facecolor=pan, labelcolor=COLORS["text"], framealpha=0.8)
        ax2.tick_params(labelbottom=False)

        # ── Subplot 3: Momentum rolling 7d ───────────────────────────────────
        ax3 = fig.add_subplot(3, 1, 3)
        _style(ax3)
        mom_roll = []
        for i in range(n):
            if i >= 7 and prices[i-7] > 0:
                mom_roll.append((prices[i] - prices[i-7]) / prices[i-7] * 100)
            else:
                mom_roll.append(0.0)
        mcols = [COLORS["green"] if m >= 0 else COLORS["red"] for m in mom_roll]
        ax3.bar(xs, mom_roll, color=mcols, alpha=0.7, width=0.8)
        ax3.axhline(0,    color=COLORS["subtext"], lw=0.5)
        ax3.axhline(15,   color=COLORS["red"],  lw=0.6, ls="--", alpha=0.6)
        ax3.axhline(-15,  color=COLORS["red"],  lw=0.6, ls="--", alpha=0.6)
        ax3.set_ylabel("Momentum 7d %", color=COLORS["subtext"], fontsize=7)

        # Eixo X
        step = max(1, n // 12)
        ax3.set_xticks(xs[::step])
        ax3.set_xticklabels(dates[::step], rotation=30, ha="right",
                             fontsize=6, color=COLORS["subtext"])

        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def _cycle_chart_region(self):
        """Grafico de pizza + barras mostrando distribuicao de fases da regiao."""
        if not HAS_MATPLOTLIB:
            messagebox.showinfo("Falta matplotlib", "pip install matplotlib"); return
        if not self._cycle_result:
            messagebox.showinfo("Sem dados", "Execute ANALISAR CICLOS primeiro."); return

        r = self._cycle_result
        if self._chart_win and self._chart_win.winfo_exists():
            self._chart_win.destroy()
        win = tk.Toplevel(self)
        win.title(f"Resumo de Ciclo — {r['region']}")
        win.configure(bg=COLORS["bg"])
        win.geometry("1100x700")
        self._chart_win = win

        bg  = COLORS["bg"]
        pan = COLORS["panel"]

        fig = Figure(figsize=(13, 8), facecolor=bg)
        fig.subplots_adjust(left=0.06, right=0.96, top=0.90,
                            bottom=0.10, wspace=0.35, hspace=0.4)

        def _style(ax):
            ax.set_facecolor(pan)
            ax.tick_params(colors=COLORS["subtext"], labelsize=7)
            for sp in ax.spines.values():
                sp.set_edgecolor(COLORS["border"])

        # ── Subplot 1: Pizza de fases ─────────────────────────────────────────
        ax1 = fig.add_subplot(2, 2, 1)
        ax1.set_facecolor(pan)
        votes  = r["phase_votes"]
        labels = [p for p in CYCLE_PHASES if votes[p] > 0]
        sizes  = [votes[p] for p in labels]
        colors = [CYCLE_PHASES[p]["color"] for p in labels]
        wedge_props = {"edgecolor": bg, "linewidth": 2}
        ax1.pie(sizes, labels=labels, colors=colors,
                autopct="%1.0f%%", pctdistance=0.75,
                wedgeprops=wedge_props,
                textprops={"color": COLORS["text"], "fontsize": 8})
        ax1.set_title(f"Distribuicao de Fases\n{r['region']}",
                      color=COLORS["header"], fontsize=9, pad=4)

        # ── Subplot 2: Barras de trend_score por item ─────────────────────────
        ax2 = fig.add_subplot(2, 2, 2)
        _style(ax2)
        items_s = sorted(r["items"].values(),
                         key=lambda x: x["trend_score"], reverse=True)
        names_s = [sig["name"][:12] for sig in items_s]
        trends  = [sig["trend_score"] for sig in items_s]
        bcolors = [CYCLE_PHASES[sig["phase"]]["color"] for sig in items_s]
        yi = list(range(len(names_s)))
        ax2.barh(yi, trends, color=bcolors, alpha=0.85, height=0.6)
        ax2.axvline(0, color=COLORS["subtext"], lw=0.8)
        ax2.set_yticks(yi)
        ax2.set_yticklabels(names_s, fontsize=7, color=COLORS["text"])
        ax2.set_xlabel("Trend Score (EMA14/EMA50 - 1)",
                       color=COLORS["subtext"], fontsize=7)
        ax2.set_title("Trend Score por Item",
                      color=COLORS["header"], fontsize=9, pad=4)

        # ── Subplot 3: Vol score por item ─────────────────────────────────────
        ax3 = fig.add_subplot(2, 2, 3)
        _style(ax3)
        items_v = sorted(r["items"].values(),
                         key=lambda x: x["vol_score"], reverse=True)
        names_v = [sig["name"][:12] for sig in items_v]
        vscores = [sig["vol_score"] for sig in items_v]
        vcols   = [COLORS["red"]    if v >= 3.0
                   else (COLORS["green"] if v >= 1.5
                   else COLORS["subtext"])
                   for v in vscores]
        yj = list(range(len(names_v)))
        ax3.barh(yj, vscores, color=vcols, alpha=0.85, height=0.6)
        ax3.axvline(1.0, color=COLORS["yellow"], lw=0.8, ls="--",
                    label="Normal (1x)")
        ax3.axvline(3.0, color=COLORS["red"],    lw=0.8, ls="--",
                    label="Spike (3x)")
        ax3.set_yticks(yj)
        ax3.set_yticklabels(names_v, fontsize=7, color=COLORS["text"])
        ax3.set_xlabel("Volume Score (vol_atual / vol_medio)",
                       color=COLORS["subtext"], fontsize=7)
        ax3.set_title("Volume Score por Item",
                      color=COLORS["header"], fontsize=9, pad=4)
        ax3.legend(fontsize=6, facecolor=pan,
                   labelcolor=COLORS["text"], framealpha=0.8)

        # ── Subplot 4: Anomalias + texto de resumo ────────────────────────────
        ax4 = fig.add_subplot(2, 2, 4)
        ax4.set_facecolor(pan)
        ax4.axis("off")

        phase    = r["region_phase"]
        pinfo    = CYCLE_PHASES[phase]
        n_anom   = len(r["anomalies"])
        systemic = r["systemic_event"]

        lines = [
            f"Fase da regiao:  {pinfo['emoji']} {phase}",
            f"Confianca:       {r['confidence']:.0f}%",
            f"Trend medio:     {r['avg_trend']:+.4f}",
            f"Vol medio:       {r['avg_vol']:.2f}x",
            f"Anomalias:       {n_anom}",
            "",
            f"Descricao:",
            *[f"  {l}" for l in pinfo["desc"].split(". ") if l],
        ]
        if systemic:
            lines += ["", "EVENTO SISTEMICO DETECTADO", r["systemic_msg"][:60]]
        if r["anomalies"]:
            lines += ["", "Anomalias detectadas:"]
            for a in r["anomalies"][:5]:
                lines.append(f"  {a['item'][:14]}: {a['msg'][:40]}")

        for i, line in enumerate(lines[:18]):
            col = (COLORS["red"]    if "EVENTO" in line
                   else (pinfo["color"] if i == 0
                   else (COLORS["yellow"] if line.startswith("Anomalias")
                   else COLORS["text"])))
            ax4.text(0.02, 1.0 - i * 0.056, line,
                     transform=ax4.transAxes, color=col,
                     fontsize=8, va="top",
                     fontfamily="monospace")

        ax4.set_title("Resumo do Ciclo",
                      color=COLORS["header"], fontsize=9,
                      loc="left", pad=4)

        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def _export_cycle(self):
        """Exporta resultados da analise de ciclo para CSV."""
        if not self._cycle_result:
            messagebox.showinfo("Sem dados", "Execute ANALISAR CICLOS primeiro.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=([("Excel (.xlsx)","*.xlsx"),("CSV","*.csv")] if HAS_OPENPYXL else [("CSV","*.csv")]),
            initialfile="ciclos_mercado.csv")
        if not path:
            return
        try:
            r = self._cycle_result
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["item","grupo","fase","trend_score","vol_score",
                             "momentum_7d","momentum_30d","volatility","rsi",
                             "anomalia","msg_anomalia"])
                for sig in r["items"].values():
                    w.writerow([
                        sig["name"], sig["group"], sig["phase"],
                        sig["trend_score"], sig["vol_score"],
                        sig["momentum_7d"], sig["momentum_30d"],
                        sig["volatility"],
                        sig["rsi"] if sig["rsi"] is not None else "",
                        int(sig["anomaly"]), sig.get("anomaly_msg",""),
                    ])
                w.writerow([])
                w.writerow(["RESUMO DA REGIAO"])
                w.writerow(["regiao", r["region"]])
                w.writerow(["fase",   r["region_phase"]])
                w.writerow(["confianca_%", r["confidence"]])
                w.writerow(["trend_medio", r["avg_trend"]])
                w.writerow(["vol_medio",   r["avg_vol"]])
                w.writerow(["n_anomalias", len(r["anomalies"])])
                w.writerow(["evento_sistemico", int(r["systemic_event"])])
                if r["systemic_msg"]:
                    w.writerow(["msg_evento", r["systemic_msg"]])
            messagebox.showinfo("Exportado", f"Salvo em:\n{path}")
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))
    # ══════════════════════════════════════════════════════════════════════════
    #  ABA PREVISAO — Previsao de precos com ensemble de medias moveis (Tier 3)
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_fcast(self):
        p = self._tab_fcast

        # ── Painel de controle ────────────────────────────────────────────────
        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(6, 2))

        self._lbl(fr, "Type ID",        0, 0)
        self._fcast_tid_e  = self._entry(fr, "",    1, 0, 12)
        self._lbl(fr, "Nome",           0, 2)
        self._fcast_name_e = self._entry(fr, "",    1, 2, 22)
        self._lbl(fr, "Regiao",         0, 4)
        self._fcast_reg    = self._combo(fr, list(REGIONS.keys()), 1, 4)
        self._lbl(fr, "Horizonte (dias)", 0, 6)
        self._fcast_horizon = self._combo(fr, ["7", "14", "30"], 1, 6, 6)
        self._fcast_horizon.set("14")
        self._lbl(fr, "Alpha (DES)",    0, 8)
        self._fcast_alpha  = self._entry(fr, "0.30", 1, 8, 7)
        self._lbl(fr, "Beta (DES)",     0, 10)
        self._fcast_beta   = self._entry(fr, "0.10", 1, 10, 7)

        self._fcast_btn = tk.Button(fr, text="  PREVER PRECOS  ",
                                     command=self._start_fcast,
                                     bg=COLORS["green"], fg="#000",
                                     font=("Courier", 11, "bold"),
                                     relief="flat", cursor="hand2")
        self._fcast_btn.grid(row=1, column=12, padx=8)

        tk.Button(fr, text=" + WATCHLIST ", command=self._fcast_add_wl,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=13, padx=4)
        tk.Button(fr, text=" EXPORTAR ", command=self._export_fcast,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=14, padx=4)

        self._fcast_spin = ttk.Progressbar(fr, mode="indeterminate", length=80)
        self._fcast_spin.grid(row=1, column=15, padx=6)

        tk.Label(fr,
                 text="Ensemble: WMA + DES (Holt) + Regressao Linear | "
                      "Pesos auto-ajustados por RMSE | Banda = ±1σ volatilidade",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 9)
                 ).grid(row=0, column=6, columnspan=10, padx=4, sticky="w")

        # ── Painel de metricas ────────────────────────────────────────────────
        met_fr = tk.Frame(p, bg=COLORS["panel"],
                          highlightbackground=COLORS["border"], highlightthickness=1)
        met_fr.pack(fill="x", padx=6, pady=(2, 2))

        self._fcast_meta_lbl = tk.Label(
            met_fr,
            text="Execute PREVER PRECOS para ver as metricas do modelo.",
            bg=COLORS["panel"], fg=COLORS["subtext"],
            font=("Courier", 11), justify="left", anchor="w")
        self._fcast_meta_lbl.pack(fill="x", padx=10, pady=5)

        # ── Area do grafico ───────────────────────────────────────────────────
        self._fcast_chart_frame = tk.Frame(p, bg=COLORS["bg"])
        self._fcast_chart_frame.pack(fill="both", expand=True, padx=4, pady=4)

        # ── Tabela de previsoes (dias futuros) ────────────────────────────────
        tbl_fr = tk.Frame(p, bg=COLORS["bg"])
        tbl_fr.pack(fill="x", padx=6, pady=(0, 6))

        self._fcast_tree = self._make_tree(tbl_fr, [
            ("date",    "Data",          100),
            ("central", "Previsao (ISK)", 130),
            ("lower",   "Minimo (-1σ)",  120),
            ("upper",   "Maximo (+1σ)",  120),
            ("wma",     "WMA",           110),
            ("des",     "DES (Holt)",    110),
            ("lr",      "Reg. Linear",   110),
            ("chg_pct", "Var vs hoje %",  100),
        ], height=8)
        for cid, label in [
            ("date","Data"),("central","Previsao (ISK)"),
            ("lower","Minimo (-1σ)"),("upper","Maximo (+1σ)"),
            ("wma","WMA"),("des","DES (Holt)"),("lr","Reg. Linear"),
            ("chg_pct","Var vs hoje %")]:
            self._fcast_tree.heading(cid, text=label)
        self._fcast_tree.tag_configure("up",   foreground=COLORS["green"])
        self._fcast_tree.tag_configure("down", foreground=COLORS["red"])
        self._fcast_tree.tag_configure("flat", foreground=COLORS["subtext"])

    # ── Workers e handlers ────────────────────────────────────────────────────

        # Auto-lookup nome pelo Type ID (busca no cache local + ESI)
        self._bind_tid_lookup(self._fcast_tid_e, self._fcast_name_e, self._fcast_reg)

    def _start_fcast(self):
        if not self._experimental_ok(): return
        if self._loading_fcast: return
        try:
            tid = int(self._fcast_tid_e.get().strip())
        except ValueError:
            messagebox.showerror("Erro", "Type ID invalido."); return
        self._loading_fcast = True
        self._fcast_btn.config(state="disabled")
        self._fcast_spin.start(10)
        params = dict(
            rname   = self._fcast_reg.get(),
            horizon = int(self._fcast_horizon.get()),
            alpha   = max(0.01, min(0.99, self._fread(self._fcast_alpha, 0.3))),
            beta    = max(0.01, min(0.99, self._fread(self._fcast_beta,  0.1))),
            name    = self._fcast_name_e.get().strip() or str(tid),
        )
        threading.Thread(target=self._worker_fcast, args=(tid, params), daemon=True).start()

    def _worker_fcast(self, tid: int, params: dict):
        try:
            rname   = params["rname"]
            rid     = REGIONS.get(rname, 10000002)
            horizon = params["horizon"]
            alpha   = params["alpha"]
            beta    = params["beta"]
            name    = params["name"]

            self.after(0, lambda: self._set_status(
                f"Buscando historico para {name}..."))
            _warm_hist_cache(rid)  # pré-aquece cache RAM: 1 query SQL → O(1) nos _get_hist_cached
            hist = _get_hist_cached(rid, tid)
            if not hist:
                self.after(0, lambda: messagebox.showinfo(
                    "Sem dados", "Sem historico disponivel para este item/regiao."))
                return

            self.after(0, lambda: self._set_status(
                f"Calculando previsao ({len(hist)} dias de historico)..."))
            result = predict_prices(hist, horizon=horizon,
                                    alpha=alpha, beta=beta)
            if not result:
                self.after(0, lambda: messagebox.showinfo(
                    "Dados insuficientes",
                    f"Minimo 10 dias de historico necessario. "
                    f"Disponivel: {len(hist)}."))
                return

            self._fcast_result = result
            self._fcast_tid    = tid
            self._fcast_name   = name
            self.after(0, self._pop_fcast)
            self.after(0, self._dash_refresh)
        except Exception as exc:
            log.exception("worker_fcast")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro Previsao", e))
        finally:
            self._loading_fcast = False
            self.after(0, lambda: (self._fcast_btn.config(state="normal"),
                                   self._fcast_spin.stop()))

    def _pop_fcast(self):
        r = self._fcast_result
        if not r:
            return

        last_price = r["prices_hist"][-1] if r["prices_hist"] else 0
        horizon    = len(r["dates_fcast"])
        trend_dir  = "▲ Alta" if r["trend_pct"] > 0.02 else (
                     "▼ Queda" if r["trend_pct"] < -0.02 else "◇ Lateral")
        trend_col  = (COLORS["green"] if r["trend_pct"] > 0.02 else
                      COLORS["red"]   if r["trend_pct"] < -0.02 else
                      COLORS["subtext"])

        # ── Atualiza metricas ─────────────────────────────────────────────────
        w = r["model_weights"]
        rmse_str = "  |  ".join(
            f"{m}: RMSE={r['model_rmse'][m]:,.0f} peso={w[m]:.2f}"
            for m in ["WMA","DES","LR"])
        saz_max  = max(r["saz_factors"].values(), default=1.0)
        saz_min  = min(r["saz_factors"].values(), default=1.0)

        meta = (
            f"Item: {self._fcast_name}  |  "
            f"Hist: {r['n_days']}d  |  "
            f"Tendencia: {r['trend_pct']:+.4f}%/dia  {trend_dir}  |  "
            f"Volatilidade: {r['volatility']*100:.2f}%  |  "
            f"Confianca: {r['confidence']:.0f}%  |  "
            f"Sazonalidade: {saz_min:.3f}–{saz_max:.3f}\n"
            f"Modelos: {rmse_str}  |  "
            f"Melhor WMA janela: {r['best_wma_window']}d  |  "
            f"Alpha: {r['alpha']}  Beta: {r['beta']}")
        self._fcast_meta_lbl.config(text=meta, fg=trend_col)

        # ── Tabela de dias futuros ────────────────────────────────────────────
        for row in self._fcast_tree.get_children():
            self._fcast_tree.delete(row)

        mf = r["model_forecasts"]
        for i, (date, central, lower, upper) in enumerate(zip(
                r["dates_fcast"],
                r["fcast_central"],
                r["fcast_lower"],
                r["fcast_upper"])):
            chg = (central - last_price) / last_price * 100 if last_price > 0 else 0
            tag = ("up",) if chg > 0.5 else ("down",) if chg < -0.5 else ("flat",)
            self._fcast_tree.insert("", "end", iid=str(i), tags=tag, values=(
                date,
                f"{central:>14,.2f}",
                f"{lower:>14,.2f}",
                f"{upper:>14,.2f}",
                f"{mf['WMA'][i]:>14,.2f}" if i < len(mf['WMA']) else "--",
                f"{mf['DES'][i]:>14,.2f}" if i < len(mf['DES']) else "--",
                f"{mf['LR'][i]:>14,.2f}"  if i < len(mf['LR'])  else "--",
                f"{chg:>+.2f}%",
            ))

        # ── Grafico ───────────────────────────────────────────────────────────
        self._draw_fcast_chart(r, self._fcast_name)

    def _draw_fcast_chart(self, r: dict, name: str):
        if not HAS_MATPLOTLIB:
            return
        for w in self._fcast_chart_frame.winfo_children():
            w.destroy()

        prices   = r["prices_hist"]
        dates    = r["dates_hist"]
        fc       = r["fcast_central"]
        fu       = r["fcast_upper"]
        fl       = r["fcast_lower"]
        fd       = r["dates_fcast"]
        n        = len(prices)
        nh       = len(fc)
        mf       = r["model_forecasts"]

        # Exibe apenas ultimos 90 dias de historico
        win      = min(90, n)
        prices_w = prices[-win:]
        dates_w  = dates[-win:]
        xs_h     = list(range(win))
        xs_f     = list(range(win, win + nh))

        bg   = COLORS["bg"]
        pan  = COLORS["panel"]

        fig  = Figure(figsize=(13, 6.5), facecolor=bg)
        fig.subplots_adjust(left=0.07, right=0.96, top=0.90,
                            bottom=0.12, hspace=0.35)

        def _style(ax):
            ax.set_facecolor(pan)
            ax.tick_params(colors=COLORS["subtext"], labelsize=7)
            for sp in ax.spines.values():
                sp.set_edgecolor(COLORS["border"])

        # ── Subplot principal: historico + previsao ───────────────────────────
        ax1 = fig.add_subplot(2, 1, 1)
        _style(ax1)

        # Historico
        ax1.plot(xs_h, prices_w, color=COLORS["text"], lw=1.5,
                 alpha=0.8, label="Historico")

        # Linha de separacao historico / previsao
        ax1.axvline(win - 1, color=COLORS["subtext"], lw=0.8,
                    ls="--", alpha=0.6, label="Hoje")

        # Banda de confianca
        ax1.fill_between(xs_f, fl, fu,
                         alpha=0.18, color=COLORS["accent"],
                         label=f"Banda ±1σ (vol {r['volatility']*100:.1f}%)")

        # Previsoes individuais dos modelos
        wma_col = COLORS["yellow"]
        des_col = COLORS["orange"]
        lr_col  = "#ff88cc"
        if len(mf["WMA"]) == nh:
            ax1.plot(xs_f, mf["WMA"], color=wma_col, lw=1.0,
                     ls="--", alpha=0.7,
                     label=f"WMA{r['best_wma_window']} (w={r['model_weights']['WMA']:.2f})")
        if len(mf["DES"]) == nh:
            ax1.plot(xs_f, mf["DES"], color=des_col, lw=1.0,
                     ls=":",  alpha=0.7,
                     label=f"DES  (w={r['model_weights']['DES']:.2f})")
        if len(mf["LR"]) == nh:
            ax1.plot(xs_f, mf["LR"],  color=lr_col,  lw=1.0,
                     ls="-.", alpha=0.7,
                     label=f"LR   (w={r['model_weights']['LR']:.2f})")

        # Ensemble (linha principal de previsao)
        ax1.plot(xs_f, fc, color=COLORS["accent"], lw=2.5,
                 label="Ensemble")

        # Marcador no ultimo ponto historico
        ax1.scatter([win - 1], [prices_w[-1]],
                    color=COLORS["text"], s=40, zorder=5)

        trend_dir = ("▲" if r["trend_pct"] > 0.02 else
                     "▼" if r["trend_pct"] < -0.02 else "◇")
        ax1.set_title(
            f"{name}  |  Previsao {nh} dias  |  "
            f"Tendencia: {r['trend_pct']:+.4f}%/dia {trend_dir}  |  "
            f"Confianca: {r['confidence']:.0f}%",
            color=COLORS["header"], fontsize=9, pad=5)
        ax1.set_ylabel("ISK", color=COLORS["subtext"], fontsize=7)
        ax1.legend(loc="upper left", fontsize=6,
                   facecolor=pan, labelcolor=COLORS["text"],
                   framealpha=0.8, ncol=3)
        ax1.tick_params(labelbottom=False)

        # ── Subplot 2: Variacao % projetada ──────────────────────────────────
        ax2 = fig.add_subplot(2, 1, 2)
        _style(ax2)

        last_p = prices_w[-1] if prices_w else 1.0
        chg_fc = [(v - last_p) / last_p * 100 for v in fc]
        chg_up = [(v - last_p) / last_p * 100 for v in fu]
        chg_lo = [(v - last_p) / last_p * 100 for v in fl]

        xcols = [COLORS["green"] if c >= 0 else COLORS["red"] for c in chg_fc]
        ax2.bar(range(nh), chg_fc, color=xcols, alpha=0.75, width=0.7)
        ax2.fill_between(range(nh), chg_lo, chg_up,
                         alpha=0.12, color=COLORS["accent"])
        ax2.axhline(0, color=COLORS["subtext"], lw=0.5)
        ax2.set_ylabel("Variacao vs hoje %", color=COLORS["subtext"], fontsize=7)
        ax2.set_xticks(range(nh))
        ax2.set_xticklabels(
            [fd[i] if i < len(fd) else "" for i in range(nh)],
            rotation=35, ha="right", fontsize=6, color=COLORS["subtext"])

        canvas = FigureCanvasTkAgg(fig, master=self._fcast_chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def _fcast_add_wl(self):
        """Adiciona o item atual da previsao à Watchlist."""
        tid  = self._fcast_tid
        name = self._fcast_name
        if not tid:
            messagebox.showinfo("Sem item", "Execute PREVER PRECOS primeiro.")
            return
        if any(w["type_id"] == tid for w in self._watchlist):
            messagebox.showinfo("Já existe", f"{name} já está na Watchlist.")
            return
        self._watchlist.append({
            "type_id": tid, "name": name,
            "alert_below": None, "alert_above": None,
            "last_sell": None, "last_buy": None,
            "last_net": None, "last_score": None,
        })
        save_watchlist(self._watchlist)
        self._pop_wl()
        messagebox.showinfo("Adicionado",
            f"{name} adicionado à Watchlist.\n"
            f"Previsao em {len(self._fcast_result.get('dates_fcast',[]))} dias calculada.")

    def _export_fcast(self):
        """Exporta previsao completa para CSV."""
        if not self._fcast_result:
            messagebox.showinfo("Sem dados", "Execute PREVER PRECOS primeiro.")
            return
        r    = self._fcast_result
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=([("Excel (.xlsx)","*.xlsx"),("CSV","*.csv")] if HAS_OPENPYXL else [("CSV","*.csv")]),
            initialfile=f"previsao_{self._fcast_name.replace(' ','_')}.csv")
        if not path:
            return
        try:
            last_p = r["prices_hist"][-1] if r["prices_hist"] else 1.0
            mf     = r["model_forecasts"]
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                # Cabecalho com metadados
                w.writerow(["item", self._fcast_name])
                w.writerow(["regiao", self._fcast_name])
                w.writerow(["dias_historico", r["n_days"]])
                w.writerow(["tendencia_%_por_dia", r["trend_pct"]])
                w.writerow(["volatilidade", r["volatility"]])
                w.writerow(["confianca_%", r["confidence"]])
                w.writerow(["alpha_des", r["alpha"]])
                w.writerow(["beta_des",  r["beta"]])
                w.writerow([])
                # Pesos do ensemble
                w.writerow(["modelo","rmse","peso"])
                for m in ["WMA","DES","LR"]:
                    w.writerow([m, r["model_rmse"][m], r["model_weights"][m]])
                w.writerow([])
                # Previsoes
                w.writerow(["data","central_isk","minimo_isk","maximo_isk",
                             "wma","des","lr","var_pct_vs_hoje"])
                for i, (date, c, lo, hi) in enumerate(zip(
                        r["dates_fcast"], r["fcast_central"],
                        r["fcast_lower"],  r["fcast_upper"])):
                    chg = (c - last_p) / last_p * 100 if last_p else 0
                    w.writerow([date, c, lo, hi,
                                 mf["WMA"][i] if i < len(mf["WMA"]) else "",
                                 mf["DES"][i] if i < len(mf["DES"]) else "",
                                 mf["LR"][i]  if i < len(mf["LR"])  else "",
                                 round(chg, 3)])
                w.writerow([])
                # Sazonalidade
                days_name = ["Seg","Ter","Qua","Qui","Sex","Sab","Dom"]
                w.writerow(["dia_semana","fator_sazonal"])
                for dow_s, factor in r["saz_factors"].items():
                    w.writerow([days_name[int(dow_s)], factor])
            messagebox.showinfo("Exportado", f"Salvo em:\n{path}")
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))
    def _sph_open_backtest(self):
        """Abre aba BACKTESTING pre-preenchida com o item atual do Spread Hist."""
        try:
            tid = int(self._sph_tid.get().strip())
        except (ValueError, AttributeError):
            messagebox.showerror("Erro", "Digite um Type ID valido."); return
        name = self._sph_name.get().strip() or str(tid)
        self._bt_tid.set(str(tid))
        self._bt_name.set(name)
        self._bt_reg.set(self._sph_reg.get())
        self._nb.select(self._tab_bt)
        self.after(50, self._start_bt)
    # ══════════════════════════════════════════════════════════════════════════

    _MM2_COLS = {
        "item":        "Item",
        "spread_pct":  "Spread %",
        "net_pct":     "Margem Líq %",
        "mm_buy":      "Buy Ótimo (ISK)",
        "mm_sell":     "Sell Ótimo (ISK)",
        "avg_vol_90":  "Vol/Dia (90d)",
        "daily_profit":"Lucro/Dia (est.)",
        "fill_days":   "Fill (dias)",
        "q_buy":       "Fila Compra",
        "volatility":  "Volatil.",
        "mm_score":    "Score MM",
    }

    def _build_tab_mm2(self):
        p = self._tab_mm2

        # ── Sub-abas: Scanner | Simulador ─────────────────────────────────────
        nb2 = ttk.Notebook(p)
        nb2.pack(fill="both", expand=True, padx=4, pady=4)
        self._mm2_nb = nb2

        scan_fr = tk.Frame(nb2, bg=COLORS["bg"])
        sim_fr  = tk.Frame(nb2, bg=COLORS["bg"])
        nb2.add(scan_fr, text="  SCANNER  ")
        nb2.add(sim_fr,  text="  SIMULADOR DE SPREAD  ")

        self._build_mm2_scanner(scan_fr)
        self._build_mm2_simulator(sim_fr)

    # ── Scanner ───────────────────────────────────────────────────────────────

        # Auto-lookup nome pelo Type ID (busca no cache local + ESI)
        self._bind_tid_lookup(self._mm2_sim_tid, self._mm2_sim_name, None)

    def _build_mm2_scanner(self, p):
        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(6, 2))

        self._lbl(fr, "Regiao",           0, 0)
        self._mm2_reg   = self._combo(fr, list(REGIONS.keys()), 1, 0)
        self._lbl(fr, "Fonte",            0, 2)
        self._mm2_mode  = self._combo(fr, ["CSV (cache)", "API (live)"], 1, 2, 13)
        self._csv_age_label(fr, self._mm2_mode, [self._mm2_reg], 2, 2)
        self._lbl(fr, "Spread min %",     0, 4)
        self._mm2_smin  = self._entry(fr, "1.0",  1, 4, 7)
        self._lbl(fr, "Net MM min %",     0, 6)
        self._mm2_nmin  = self._entry(fr, "0.5",  1, 6, 7)
        self._lbl(fr, "Vol/dia min",      0, 8)
        self._mm2_vmin  = self._entry(fr, "10",   1, 8, 7)
        self._lbl(fr, "Fila max",         0, 10)
        self._mm2_qmax  = self._entry(fr, "30",   1, 10, 6)
        self._lbl(fr, "Top N",            0, 12)
        self._mm2_topn  = self._entry(fr, "200",  1, 12, 6)

        self._mm2_btn = tk.Button(fr, text="  ESCANEAR MM  ",
                                   command=self._start_mm2,
                                   bg=COLORS["accent"], fg="#000",
                                   font=("Courier", 11, "bold"),
                                   relief="flat", cursor="hand2")
        self._mm2_btn.grid(row=1, column=14, padx=8)

        tk.Button(fr, text=" SIMULAR ITEM ",
                  command=self._mm2_open_sim,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=15, padx=4)

        tk.Button(fr, text=" + WATCHLIST ",
                  command=self._mm2_add_wl,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=16, padx=4)

        tk.Button(fr, text=" PREVER ",
                  command=self._mm2_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=17, padx=4)

        tk.Button(fr, text=" EXPORTAR ",
                  command=self._export_mm2,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=18, padx=4)

        self._mm2_prog = ttk.Progressbar(fr, mode="determinate",
                                          length=120, maximum=100)
        self._mm2_prog.grid(row=1, column=19, padx=6)
        self._mm2_pct = tk.Label(fr, text="", bg=COLORS["panel"],
                                  fg=COLORS["subtext"], font=("Courier", 10), width=5)
        self._mm2_pct.grid(row=1, column=20, padx=2)

        tk.Label(fr,
                 text="Score MM = lucro%(40) + giro(30) + fila(20) + baixa volatil.(10)  |  "
                      "Skills globais via aba CONFIGURACOES",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 9)
                 ).grid(row=0, column=4, columnspan=16, padx=4, sticky="w")

        # Tabela de resultados
        self._mm2_tree = self._make_tree(p, [
            ("item",         "Item",           220),
            ("spread_pct",   "Spread %",         80),
            ("net_pct",      "Net MM %",          80),
            ("mm_buy",       "Buy Ótimo",        130),
            ("mm_sell",      "Sell Ótimo",       130),
            ("avg_vol_90",   "Vol/Dia",           80),
            ("daily_profit", "Lucro/Dia",        120),
            ("fill_days",    "Fill (d)",          68),
            ("q_buy",        "Fila",              55),
            ("volatility",   "Volatil.",          72),
            ("mm_score",     "Score",             65),
        ])
        for cid in self._MM2_COLS:
            self._mm2_tree.heading(cid, text=self._MM2_COLS[cid],
                                    command=lambda c=cid: self._sort_mm2(c))
        self._mm2_tree.bind("<Double-1>", self._on_dbl_mm2)
        self._mm2_tree.tag_configure("hot",  foreground=COLORS["green"])
        self._mm2_tree.tag_configure("warn", foreground=COLORS["yellow"])

        bar = tk.Frame(p, bg=COLORS["bg"])
        bar.pack(fill="x", padx=6, pady=(0, 4))
        self._mm2_count = tk.Label(bar, text="0 oportunidades",
                                    fg=COLORS["subtext"], bg=COLORS["bg"],
                                    font=("Courier", 10))
        self._mm2_count.pack(side="left")
        tk.Label(bar,
                 text="Verde = Score >= 60  |  Double-click = Simulador de spread",
                 fg=COLORS["subtext"], bg=COLORS["bg"],
                 font=("Courier", 10)).pack(side="right")

    # ── Simulador de spread ───────────────────────────────────────────────────

    def _build_mm2_simulator(self, p):
        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(6, 2))

        self._lbl(fr, "Type ID",        0, 0)
        self._mm2_sim_tid  = self._entry(fr, "", 1, 0, 12)
        self._lbl(fr, "Nome",           0, 2)
        self._mm2_sim_name = self._entry(fr, "", 1, 2, 20)
        self._lbl(fr, "Regiao",         0, 4)
        self._mm2_sim_reg  = self._combo(fr, list(REGIONS.keys()), 1, 4)
        self._lbl(fr, "Capital (ISK)",  0, 6)
        self._mm2_sim_cap  = self._entry(fr, "500000000", 1, 6, 14)
        self._lbl(fr, "Vol/dia est.",   0, 8)
        self._mm2_sim_vol  = self._entry(fr, "100", 1, 8, 8)

        self._mm2_sim_btn = tk.Button(fr, text="  SIMULAR SPREAD  ",
                                       command=self._start_mm2_sim,
                                       bg=COLORS["yellow"], fg="#000",
                                       font=("Courier", 11, "bold"),
                                       relief="flat", cursor="hand2")
        self._mm2_sim_btn.grid(row=1, column=10, padx=8)
        tk.Button(fr, text=" GRAFICO BOOK ",
                  command=self._draw_mm2_book_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=11, padx=4)
        self._mm2_sim_spin = ttk.Progressbar(fr, mode="indeterminate", length=80)
        self._mm2_sim_spin.grid(row=1, column=12, padx=6)

        tk.Label(fr,
                 text="Simula posicionamento em diferentes pontos do spread e "
                      "calcula lucro diario estimado para cada cenario.",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 9)
                 ).grid(row=0, column=6, columnspan=6, padx=4, sticky="w")

        # Painel de resumo do book
        book_fr = tk.Frame(p, bg=COLORS["panel"],
                           highlightbackground=COLORS["border"], highlightthickness=1)
        book_fr.pack(fill="x", padx=6, pady=(2, 2))
        self._mm2_book_lbl = tk.Label(
            book_fr,
            text="Execute SIMULAR SPREAD para ver o book de ordens e cenarios.",
            bg=COLORS["panel"], fg=COLORS["subtext"],
            font=("Courier", 11), justify="left", anchor="w")
        self._mm2_book_lbl.pack(fill="x", padx=10, pady=6)

        # Tabela de cenarios
        self._mm2_scen_tree = self._make_tree(p, [
            ("pos_label",    "Posicao",         80),
            ("mm_buy",       "Buy (ISK)",       130),
            ("mm_sell",      "Sell (ISK)",      130),
            ("spread_pct",   "Spread %",         80),
            ("net_pct",      "Net %",            75),
            ("fill_rate_pct","Fill Rate %",       85),
            ("qty_dia",      "Qtd/Dia",           75),
            ("daily_profit", "Lucro/Dia",        115),
            ("capital_req",  "Capital Req.",     120),
            ("roi_30d_pct",  "ROI 30d %",         85),
        ], height=10)
        for cid, lbl in [
            ("pos_label","Posicao"),("mm_buy","Buy (ISK)"),("mm_sell","Sell (ISK)"),
            ("spread_pct","Spread %"),("net_pct","Net %"),
            ("fill_rate_pct","Fill Rate %"),("qty_dia","Qtd/Dia"),
            ("daily_profit","Lucro/Dia"),("capital_req","Capital Req."),
            ("roi_30d_pct","ROI 30d %")]:
            self._mm2_scen_tree.heading(cid, text=lbl)
        self._mm2_scen_tree.tag_configure("optimal", foreground=COLORS["green"],
                                           font=("Courier", 11, "bold"))
        self._mm2_scen_tree.tag_configure("loss",    foreground=COLORS["red"])

        bar2 = tk.Frame(p, bg=COLORS["bg"])
        bar2.pack(fill="x", padx=6, pady=(0, 4))
        self._mm2_sim_info = tk.Label(bar2, text="",
                                       fg=COLORS["subtext"], bg=COLORS["bg"],
                                       font=("Courier", 10))
        self._mm2_sim_info.pack(side="left")

    # ── Workers Scanner ───────────────────────────────────────────────────────

    def _start_mm2(self):
        if self._loading_mm2: return
        self._loading_mm2 = True
        self._mm2_btn.config(state="disabled")
        self._mm2_prog.config(value=0); self._mm2_pct.config(text="")
        br, ac = self._gskills()
        params = dict(
            rname   = self._mm2_reg.get(),
            use_csv = self._mm2_mode.get().startswith("CSV"),
            br=br, ac=ac,
            smin    = self._fread(self._mm2_smin, 1.0),
            vmin    = self._fread(self._mm2_vmin, 10.0),
            qmax    = int(self._fread(self._mm2_qmax, 30)),
            nmin    = self._fread(self._mm2_nmin, 0.5),
            topn    = int(self._fread(self._mm2_topn, 200)),
        )
        threading.Thread(target=self._worker_mm2, args=(params,), daemon=True).start()

    def _worker_mm2(self, params: dict):
        prog = self._make_prog(self._mm2_prog, self._mm2_pct)
        try:
            rname  = params["rname"]
            br, ac = params["br"], params["ac"]
            if params["use_csv"]:
                orders = load_orders_from_csv(rname)
                if not orders:
                    self.after(0, lambda: messagebox.showerror(
                        "CSV nao encontrado",
                        "Baixe os dados na aba DADOS & CACHE.")); return
            else:
                orders = download_region_orders(REGIONS[rname], rname, prog)
                _orders_mem_cache[rname] = (time.time(), orders)

            results = analyze_market_making(
                orders, rname, self._item_cache,
                broker_rel      = br,
                accounting      = ac,
                min_spread_pct  = params["smin"],
                min_vol_dia     = params["vmin"],
                max_queue_depth = params["qmax"],
                min_net_pct     = params["nmin"],
                top_n           = params["topn"],
                blacklist       = self._blacklist,
                progress_cb     = prog,
            )
            self._mm2_results = results
            self.after(0, lambda: (self._mm2_prog.config(value=100),
                                   self._mm2_pct.config(text="100%")))
            self.after(0, self._pop_mm2)
            self.after(0, self._dash_refresh)
        except Exception as exc:
            log.exception("worker_mm2")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro MM", e))
        finally:
            self._loading_mm2 = False
            self.after(0, lambda: self._mm2_btn.config(state="normal"))

    def _pop_mm2(self, reorder_only=False):
        if reorder_only:
            for i in range(len(self._mm2_results)):
                self._mm2_tree.move(str(i), "", i)
            return
        for row in self._mm2_tree.get_children():
            self._mm2_tree.delete(row)
        for i, r in enumerate(self._mm2_results):
            score = r.get("mm_score") or 0
            tag   = ("hot",)  if score >= 60 else \
                    ("warn",) if score >= 35 else \
                    ("even" if i % 2 == 0 else "odd",)
            dp = r.get("daily_profit")
            vd = r.get("avg_vol_90")
            fd = r.get("fill_days")
            vo = r.get("volatility")
            self._mm2_tree.insert("", "end", iid=str(i), tags=tag, values=(
                r["item"],
                f"{r['spread_pct']:.2f}%",
                f"{r['net_pct']:.3f}%",
                f"{r['mm_buy']:>16,.2f}",
                f"{r['mm_sell']:>16,.2f}",
                f"{vd:>8,.1f}" if vd is not None else "--",
                f"{dp:>14,.0f}" if dp is not None else "--",
                f"{fd:.1f}d"    if fd is not None else "--",
                str(r["q_buy"]),
                f"{vo*100:.2f}%" if vo is not None else "--",
                f"{score:.1f}",
            ))
        total_daily = sum(r.get("daily_profit") or 0 for r in self._mm2_results)
        self._mm2_count.config(
            text=(f"{len(self._mm2_results)} oportunidades  |  "
                  f"Lucro/dia total (est.): {total_daily:,.0f} ISK"))

    def _sort_mm2(self, col):
        num = {"spread_pct","net_pct","mm_buy","mm_sell","avg_vol_90",
               "daily_profit","fill_days","q_buy","volatility","mm_score"}
        rev = not self._mm2_sort_state.get(col, col in num)
        self._mm2_sort_state = {col: rev}
        self._mm2_results.sort(
            key=lambda r: (r.get(col) or 0) if col in num else r.get(col, ""),
            reverse=rev)
        self._set_heading_arrow(self._mm2_tree, self._MM2_COLS, col, rev)
        self._pop_mm2(reorder_only=True)

    def _on_dbl_mm2(self, _):
        """Double-click no scanner: abre Simulador para o item."""
        self._mm2_open_sim()

    def _mm2_open_sim(self):
        """Pre-preenche o Simulador com o item selecionado no Scanner."""
        sel = self._mm2_tree.selection()
        if not sel:
            messagebox.showinfo("Selecione", "Clique num item no Scanner."); return
        try:
            idx = int(sel[0])
        except ValueError:
            return
        r = self._mm2_results[idx]
        self._mm2_sim_tid.set(str(r["type_id"]))
        self._mm2_sim_name.set(r["item"])
        self._mm2_sim_reg.set(self._mm2_reg.get())
        vd = r.get("avg_vol_90") or 100
        self._mm2_sim_vol.set(str(int(vd)))
        # Muda para sub-aba Simulador e navega para a aba Market Making
        self._nb.select(self._tab_mm2)
        self._mm2_nb.select(1)
        self._start_mm2_sim()

    def _mm2_add_wl(self):
        """Adiciona item selecionado no Scanner à Watchlist."""
        sel = self._mm2_tree.selection()
        if not sel:
            messagebox.showinfo("Selecione", "Clique num item."); return
        try:
            idx = int(sel[0])
        except ValueError:
            return
        r = self._mm2_results[idx]
        tid, name = r["type_id"], r["item"]
        if any(w["type_id"] == tid for w in self._watchlist):
            messagebox.showinfo("Já existe", f"{name} já está na Watchlist."); return
        self._watchlist.append({
            "type_id": tid, "name": name,
            "alert_below": None, "alert_above": None,
            "last_sell": None, "last_buy": None,
            "last_net": None, "last_score": None,
        })
        save_watchlist(self._watchlist)
        self._pop_wl()
        messagebox.showinfo("Adicionado",
            f"{name} adicionado à Watchlist.\n"
            f"Score MM: {r.get('mm_score',0):.1f}  |  "
            f"Net: {r['net_pct']:.3f}%")

    def _export_mm2(self):
        headers = [("item","Item"),("type_id","Type ID"),("spread_pct","Spread %"),
                   ("net_pct","Net %"),("mm_buy","MM Buy"),("mm_sell","MM Sell"),
                   ("avg_vol_90","Vol/Dia"),("daily_profit","Lucro/Dia"),
                   ("fill_days","Fill Days"),("q_buy","Q Buy"),("q_sell","Q Sell"),
                   ("volatility","Volatilidade"),("inventory_risk","Risco Inv."),
                   ("mm_score","Score"),("region","Região")]
        self._export_table("Market Making", headers,
                           self._mm2_results or [], "market_making")
    def _mm2_open_fcast(self):
        """Abre aba PREVISAO pre-preenchida com o item selecionado no scanner MM."""
        sel = self._mm2_tree.selection()
        if not sel:
            messagebox.showinfo("Selecione", "Clique num item no Scanner."); return
        try:
            idx = int(sel[0])
        except ValueError:
            return
        r = self._mm2_results[idx]
        self._fcast_tid_e.set(str(r["type_id"]))
        self._fcast_name_e.set(r["item"])
        self._fcast_reg.set(self._mm2_reg.get())
        self._nb.select(self._tab_fcast)
        self.after(50, self._start_fcast)

    def _draw_mm2_book_chart(self):
        """Gráfico do livro de ordens (depth chart) do item simulado."""
        if not HAS_MATPLOTLIB:
            messagebox.showinfo("Falta matplotlib", "pip install matplotlib"); return
        sim = self._mm2_sim
        if not sim:
            messagebox.showinfo("Sem dados", "Execute SIMULAR SPREAD primeiro."); return

        book  = sim["book_levels"]
        sells = book["sell"]
        buys  = book["buy"]
        if not sells or not buys:
            messagebox.showinfo("Sem dados", "Book insuficiente."); return

        if self._chart_win and self._chart_win.winfo_exists():
            self._chart_win.destroy()
        name = getattr(self, "_mm2_sim_name_cur", "Item")
        win  = tk.Toplevel(self)
        win.title(f"Book de Ordens — {name}")
        win.configure(bg=COLORS["bg"])
        win.geometry("1000x600")
        self._chart_win = win

        bg  = COLORS["bg"]
        pan = COLORS["panel"]
        fig = Figure(figsize=(12, 6), facecolor=bg)
        fig.subplots_adjust(left=0.08, right=0.96, top=0.88, bottom=0.12)

        ax = fig.add_subplot(1, 1, 1)
        ax.set_facecolor(pan)
        ax.tick_params(colors=COLORS["subtext"], labelsize=8)
        for sp in ax.spines.values():
            sp.set_edgecolor(COLORS["border"])

        # Constrói curvas de profundidade acumulada
        sell_prices  = sorted([s["price"]  for s in sells])
        sell_vols    = [next(s["vol"] for s in sells if s["price"] == p)
                        for p in sell_prices]
        buy_prices   = sorted([b["price"]  for b in buys], reverse=True)
        buy_vols     = [next(b["vol"] for b in buys if b["price"] == p)
                        for p in buy_prices]

        # Acumula
        sell_cum = []
        acc = 0
        for v in sell_vols:
            acc += v; sell_cum.append(acc)
        buy_cum = []
        acc = 0
        for v in buy_vols:
            acc += v; buy_cum.append(acc)

        # Step chart
        ax.step(sell_prices, sell_cum, where="post",
                color=COLORS["red"], lw=2, label="Ordens de venda")
        ax.fill_between(sell_prices, sell_cum,
                        step="post", alpha=0.2, color=COLORS["red"])
        ax.step(buy_prices, buy_cum, where="post",
                color=COLORS["green"], lw=2, label="Ordens de compra")
        ax.fill_between(buy_prices, buy_cum,
                        step="post", alpha=0.2, color=COLORS["green"])

        # Linhas de referência
        ax.axvline(book["sell_min"], color=COLORS["red"],
                   lw=1, ls="--", alpha=0.7, label=f"Sell min {book['sell_min']:,.2f}")
        ax.axvline(book["buy_max"], color=COLORS["green"],
                   lw=1, ls="--", alpha=0.7, label=f"Buy max {book['buy_max']:,.2f}")
        ax.axvline(book["mid"], color=COLORS["yellow"],
                   lw=0.8, ls=":", alpha=0.6, label=f"Mid {book['mid']:,.2f}")

        # Linhas do cenário ótimo
        opt = sim["optimal"]
        ax.axvline(opt["mm_buy"],  color=COLORS["accent"],
                   lw=1.5, ls="-",  alpha=0.8,
                   label=f"MM Buy ★ {opt['mm_buy']:,.2f}")
        ax.axvline(opt["mm_sell"], color=COLORS["orange"],
                   lw=1.5, ls="-",  alpha=0.8,
                   label=f"MM Sell ★ {opt['mm_sell']:,.2f}")

        ax.set_xlabel("Preço (ISK)", color=COLORS["subtext"], fontsize=8)
        ax.set_ylabel("Volume acumulado", color=COLORS["subtext"], fontsize=8)
        be_pct = sim["breakeven_pct"]
        ax.set_title(
            f"Book de Ordens — {name}  |  "
            f"Spread: {book['spread_pct']:.2f}%  |  "
            f"Break-even: {be_pct:.3f}%  |  "
            f"Net MM ótimo: {opt['net_pct']:.3f}%",
            color=COLORS["header"], fontsize=9, pad=6)
        ax.legend(loc="upper center", fontsize=7, ncol=4,
                  facecolor=pan, labelcolor=COLORS["text"], framealpha=0.8)

        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    # ── Workers Simulador ─────────────────────────────────────────────────────

    def _start_mm2_sim(self):
        if getattr(self, '_loading_mm2_sim', False): return
        try:
            tid = int(self._mm2_sim_tid.get().strip())
        except ValueError:
            messagebox.showerror("Erro", "Type ID invalido."); return
        self._loading_mm2_sim = True
        self._mm2_sim_btn.config(state="disabled")
        self._mm2_sim_spin.start(10)
        br, ac = self._gskills()
        params = dict(
            rname   = self._mm2_sim_reg.get(),
            capital = self._fread(self._mm2_sim_cap, 500_000_000),
            vol_dia = self._fread(self._mm2_sim_vol, 100.0),
            br=br, ac=ac,
            name    = self._mm2_sim_name.get().strip() or str(tid),
        )
        threading.Thread(target=self._worker_mm2_sim, args=(tid, params),
                         daemon=True).start()

    def _worker_mm2_sim(self, tid: int, params: dict):
        try:
            rname   = params["rname"]
            rid     = REGIONS.get(rname, 10000002)
            capital = params["capital"]
            vol_dia = params["vol_dia"]
            br, ac  = params["br"], params["ac"]
            name    = params["name"]

            # Busca ordens ao vivo para o item especifico
            self.after(0, lambda: self._set_status(
                f"Buscando book de ordens para {name}..."))
            url    = f"{ESI_BASE}/markets/{rid}/orders/"
            params = {"datasource": "tranquility",
                      "order_type": "all", "type_id": tid}
            r_esi  = _get_session().get(url, params=params, timeout=12)
            if r_esi.status_code != 200:
                self.after(0, lambda: messagebox.showerror(
                    "Erro ESI", f"HTTP {r_esi.status_code}")); return
            orders_raw = r_esi.json()

            # Monta entry no formato do build_order_index
            sells = sorted([o for o in orders_raw if not o["is_buy_order"]],
                           key=lambda o: float(o["price"]))
            buys  = sorted([o for o in orders_raw if  o["is_buy_order"]],
                           key=lambda o: float(o["price"]), reverse=True)
            entry  = {"sell": sells, "buy": buys}

            # Se vol_dia nao foi pre-preenchido, usa historico
            if vol_dia <= 0:
                hist   = _get_hist_cached(rid, tid)
                st     = _history_stats(hist)
                vol_dia= st["avg_vol_90"] if st else 100.0

            sim = simulate_mm_spread(entry, tid, br, ac, capital, vol_dia)
            if not sim:
                self.after(0, lambda: messagebox.showinfo(
                    "Sem dados", "Book de ordens insuficiente.")); return

            self._mm2_sim    = sim
            self._mm2_sim_tid_cur  = tid
            self._mm2_sim_name_cur = name
            self.after(0, self._pop_mm2_sim)
        except Exception as exc:
            log.exception("worker_mm2_sim")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro Sim", e))
        finally:
            self._loading_mm2_sim = False
            self.after(0, lambda: (self._mm2_sim_btn.config(state="normal"),
                                   self._mm2_sim_spin.stop()))

    def _pop_mm2_sim(self):
        sim = self._mm2_sim
        if not sim:
            return
        book = sim["book_levels"]
        opt  = sim["optimal"]
        br_pct = sim["rb"] * 100
        rt_pct = sim["rt"] * 100
        be_pct = sim["breakeven_pct"]

        # Indica se o spread atual é suficiente para lucrar
        spread_ok = book["spread_pct"] > be_pct
        spread_indicator = "✓ VIAVEL" if spread_ok else f"✗ INSUFICIENTE (precisa >{be_pct:.2f}%)"
        spread_col = COLORS["green"] if spread_ok else COLORS["red"]

        # Resumo do book
        book_txt = (
            f"Item: {self._mm2_sim_name_cur}  |  "
            f"Sell min: {book['sell_min']:,.2f}  |  "
            f"Buy max: {book['buy_max']:,.2f}  |  "
            f"Spread: {book['spread_abs']:,.2f} ISK ({book['spread_pct']:.2f}%)  |  "
            f"Break-even: {be_pct:.3f}%  [{spread_indicator}]\n"
            f"Broker: {br_pct:.2f}%  Sales tax: {rt_pct:.3f}%  "
            f"Custo round-trip: {(sim['rb']*2 + sim['rt'])*100:.3f}%"
            + (f"\n  CENARIO OTIMO: Buy {opt['mm_buy']:,.2f} / Sell {opt['mm_sell']:,.2f}  |  "
               f"Net {opt['net_pct']:.3f}%  |  Fill {opt['fill_rate_pct']:.0f}%  |  "
               f"Lucro/dia {opt['daily_profit']:,.0f} ISK  |  ROI 30d {opt['roi_30d_pct']:.1f}%"
               if spread_ok else ""))
        self._mm2_book_lbl.config(text=book_txt, fg=spread_col)

        # Tabela de cenários — inclui todos, marca negativos em vermelho
        for row in self._mm2_scen_tree.get_children():
            self._mm2_scen_tree.delete(row)
        for i, sc in enumerate(sim["scenarios"]):
            is_opt  = (sc is opt)
            profitbl = sc.get("profitable", True)
            if is_opt and profitbl:
                tag = ("optimal",)
            elif not profitbl:
                tag = ("loss",)
            else:
                tag = ("even" if i % 2 == 0 else "odd",)
            dp_str = f"{sc['daily_profit']:>12,.0f}" if profitbl else "N/A (spread<be)"
            self._mm2_scen_tree.insert("", "end", iid=str(i), tags=tag, values=(
                sc["pos_label"] + (" ★" if is_opt else ""),
                f"{sc['mm_buy']:>16,.2f}",
                f"{sc['mm_sell']:>16,.2f}",
                f"{sc['spread_pct']:.2f}%",
                f"{sc['net_pct']:.3f}%",
                f"{sc['fill_rate_pct']:.0f}%",
                f"{sc['qty_dia']:.1f}",
                dp_str,
                f"{sc['capital_req']:>14,.0f}",
                f"{sc['roi_30d_pct']:.1f}%" if profitbl else "--",
            ))

        self._mm2_sim_info.config(
            fg=COLORS["green"] if spread_ok else COLORS["red"],
            text=(f"Vol/dia: {sim['vol_dia']:.0f}  |  "
                  f"Capital: {sim['capital']:,.0f} ISK  |  "
                  f"Break-even: {be_pct:.3f}%  |  "
                  f"Spread atual: {book['spread_pct']:.2f}%  [{spread_indicator}]  |  "
                  f"Skills: Broker Rel={sim['broker_rel']} Accounting={sim['accounting']}"))

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA INDUSTRIA — Calculadora de custo de fabricacao vs compra
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_ind(self):
        p = self._tab_ind

        # ── Painel superior: produto e parâmetros ─────────────────────────────
        top = tk.Frame(p, bg=COLORS["panel"],
                       highlightbackground=COLORS["border"], highlightthickness=1)
        top.pack(fill="x", padx=6, pady=(6, 2))

        # Linha 1: identificação do produto
        self._lbl(top, "Blueprint ou Produto ID", 0, 0)
        self._ind_prod_tid  = self._entry(top, "", 1, 0, 12)
        self._lbl(top, "Nome",              0, 2)
        self._ind_prod_name = self._entry(top, "", 1, 2, 24)
        self._lbl(top, "Qtd desejada",      0, 4)
        self._ind_want_qty  = self._entry(top, "100", 1, 4, 8)
        self._lbl(top, "ME (0-10)",         0, 6)
        self._ind_me        = self._entry(top, "0", 1, 6, 5)
        self._lbl(top, "Região",            0, 8)
        self._ind_reg       = self._combo(top, list(REGIONS.keys()), 1, 8)

        # Linha 2: parâmetros financeiros
        self._lbl(top, "Preço mercado (ISK)", 2, 0)
        self._ind_mkt_price = self._entry(top, "", 3, 0, 16)
        self._lbl(top, "Estrutura Mat%",  2, 2)
        self._ind_str_mat   = self._entry(top, "0", 3, 2, 5)
        self._lbl(top, "Estrutura Job%",  2, 4)
        self._ind_str_job   = self._entry(top, "0", 3, 4, 5)
        self._lbl(top, "Taxa estrutura%", 2, 6)
        self._ind_fac_tax   = self._entry(top, "5", 3, 6, 5)
        self._lbl(top, "Cost Index %",    2, 8)
        self._ind_ci        = self._entry(top, "5", 3, 8, 5)
        self._lbl(top, "Qtd/run",         2, 10)
        self._ind_qty       = self._entry(top, "1", 3, 10, 6)
        self._lbl(top, "Runs",            2, 12)
        self._ind_runs      = self._entry(top, "1", 3, 12, 5)

        # ── Botões de ação principais ─────────────────────────────────────────
        act_row = tk.Frame(p, bg=COLORS["bg"])
        act_row.pack(fill="x", padx=6, pady=(3, 1))

        self._ind_auto_btn = tk.Button(
            act_row,
            text="  ⟳ AUTO-PREENCHER MATERIAIS E PREÇOS  ",
            command=self._ind_auto_load,
            bg=COLORS["accent"], fg="#000",
            font=("Courier", 12, "bold"), relief="flat", cursor="hand2")
        self._ind_auto_btn.pack(side="left", padx=4)

        self._ind_btn = tk.Button(
            act_row, text="  CALCULAR  ",
            command=self._start_ind,
            bg=COLORS["green"], fg="#000",
            font=("Courier", 11, "bold"), relief="flat", cursor="hand2")
        self._ind_btn.pack(side="left", padx=4)

        tk.Button(act_row, text=" BUSCAR PREÇO MERCADO ",
                  command=self._ind_fetch_price,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)

        tk.Button(act_row, text=" EXPORTAR ",
                  command=self._export_ind,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)

        self._ind_spin = ttk.Progressbar(act_row, mode="indeterminate", length=80)
        self._ind_spin.pack(side="left", padx=8)

        self._ind_status_lbl = tk.Label(
            act_row, text="",
            bg=COLORS["bg"], fg=COLORS["subtext"],
            font=("Courier", 10))
        self._ind_status_lbl.pack(side="left", padx=4)

        # ── Painel de materiais ────────────────────────────────────────────────
        mat_fr = tk.LabelFrame(p, text="  Materiais  ",
                               bg=COLORS["panel"], fg=COLORS["header"],
                               font=("Courier", 11, "bold"),
                               highlightbackground=COLORS["border"], highlightthickness=1)
        mat_fr.pack(fill="x", padx=6, pady=(2, 2))

        add_row = tk.Frame(mat_fr, bg=COLORS["panel"])
        add_row.pack(fill="x", padx=6, pady=4)
        tk.Label(add_row, text="Type ID:", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 10)).pack(side="left")
        self._ind_mat_tid = tk.Entry(add_row, width=10,
                                     bg=COLORS["border"], fg=COLORS["text"],
                                     insertbackground=COLORS["text"],
                                     font=("Courier", 11), relief="flat", bd=2)
        self._ind_mat_tid.pack(side="left", padx=4)
        tk.Label(add_row, text="Qtd:", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 10)).pack(side="left")
        self._ind_mat_qty = tk.Entry(add_row, width=8,
                                     bg=COLORS["border"], fg=COLORS["text"],
                                     insertbackground=COLORS["text"],
                                     font=("Courier", 11), relief="flat", bd=2)
        self._ind_mat_qty.pack(side="left", padx=4)
        tk.Label(add_row, text="Preço unit (ISK):", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 10)).pack(side="left")
        self._ind_mat_price = tk.Entry(add_row, width=14,
                                       bg=COLORS["border"], fg=COLORS["text"],
                                       insertbackground=COLORS["text"],
                                       font=("Courier", 11), relief="flat", bd=2)
        self._ind_mat_price.pack(side="left", padx=4)
        tk.Button(add_row, text=" + ADICIONAR ",
                  command=self._ind_add_material,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)
        tk.Button(add_row, text=" REMOVER ",
                  command=self._ind_remove_material,
                  bg=COLORS["border"], fg=COLORS["red"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)
        tk.Button(add_row, text=" LIMPAR TUDO ",
                  command=self._ind_clear_materials,
                  bg=COLORS["border"], fg=COLORS["subtext"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)

        self._ind_mat_tree = self._make_tree(mat_fr, [
            ("tid",      "Type ID",        90),
            ("name",     "Material",      200),
            ("qty_base", "Qtd/run",        80),
            ("qty_me",   "Qtd c/ME",       80),
            ("qty_total","Qtd total",       90),
            ("price",    "Preço unit",     130),
            ("total",    "Subtotal",       140),
        ], height=6)

        # ── Painel de resultado ────────────────────────────────────────────────
        res_fr = tk.LabelFrame(p, text="  Resultado  ",
                               bg=COLORS["panel"], fg=COLORS["header"],
                               font=("Courier", 11, "bold"),
                               highlightbackground=COLORS["border"], highlightthickness=1)
        res_fr.pack(fill="both", expand=True, padx=6, pady=(2, 6))

        self._ind_result_lbl = tk.Label(
            res_fr,
            text="Digite um Type ID e clique  ⟳ AUTO-PREENCHER MATERIAIS E PREÇOS.",
            bg=COLORS["panel"], fg=COLORS["subtext"],
            font=("Courier", 11), justify="left", anchor="nw")
        self._ind_result_lbl.pack(fill="both", expand=True, padx=10, pady=8)

        # Auto-lookup nome pelo Type ID (busca no cache local + ESI)
        self._bind_tid_lookup(self._ind_prod_tid, self._ind_prod_name, None)

        # Botões de navegação para outras abas
        ind_act = tk.Frame(res_fr, bg=COLORS["panel"])
        ind_act.pack(fill="x", padx=4, pady=(2, 4))
        tk.Button(ind_act, text=" GRÁFICO ",     command=self._ind_open_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(ind_act, text=" PREVER ",      command=self._ind_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(ind_act, text=" BACKTEST ",    command=self._ind_open_bt,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(ind_act, text=" ANALISAR IA ", command=self._ind_open_ai,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)

        # ── Painel de materiais ────────────────────────────────────────────────
        mat_fr = tk.LabelFrame(p, text="  Materiais  ",
                               bg=COLORS["panel"], fg=COLORS["header"],
                               font=("Courier", 11, "bold"),
                               highlightbackground=COLORS["border"], highlightthickness=1)
        mat_fr.pack(fill="x", padx=6, pady=(2, 2))

        add_row = tk.Frame(mat_fr, bg=COLORS["panel"])
        add_row.pack(fill="x", padx=6, pady=4)
        tk.Label(add_row, text="Type ID:", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 10)).pack(side="left")
        self._ind_mat_tid = tk.Entry(add_row, width=10,
                                     bg=COLORS["border"], fg=COLORS["text"],
                                     insertbackground=COLORS["text"],
                                     font=("Courier", 11), relief="flat", bd=2)
        self._ind_mat_tid.pack(side="left", padx=4)
        tk.Label(add_row, text="Qtd:", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 10)).pack(side="left")
        self._ind_mat_qty = tk.Entry(add_row, width=8,
                                     bg=COLORS["border"], fg=COLORS["text"],
                                     insertbackground=COLORS["text"],
                                     font=("Courier", 11), relief="flat", bd=2)
        self._ind_mat_qty.pack(side="left", padx=4)
        tk.Label(add_row, text="Preco unit (ISK):", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 10)).pack(side="left")
        self._ind_mat_price = tk.Entry(add_row, width=14,
                                       bg=COLORS["border"], fg=COLORS["text"],
                                       insertbackground=COLORS["text"],
                                       font=("Courier", 11), relief="flat", bd=2)
        self._ind_mat_price.pack(side="left", padx=4)
        tk.Button(add_row, text=" + ADICIONAR ",
                  command=self._ind_add_material,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)
        tk.Button(add_row, text=" REMOVER ",
                  command=self._ind_remove_material,
                  bg=COLORS["border"], fg=COLORS["red"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)
        tk.Button(add_row, text=" LIMPAR TUDO ",
                  command=self._ind_clear_materials,
                  bg=COLORS["border"], fg=COLORS["subtext"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)

        self._ind_mat_tree = self._make_tree(mat_fr, [
            ("tid",   "Type ID",        90),
            ("name",  "Material",      200),
            ("qty",   "Qtd/run",        80),
            ("qty_me","Qtd c/ ME",      80),
            ("price", "Preco unit",    120),
            ("total", "Subtotal",      130),
        ], height=5)

        # ── Botoes de acao ─────────────────────────────────────────────────────
        btn_row = tk.Frame(p, bg=COLORS["bg"])
        btn_row.pack(fill="x", padx=6, pady=(2, 2))

        self._ind_btn = tk.Button(btn_row, text="  CALCULAR  ",
                                   command=self._start_ind,
                                   bg=COLORS["green"], fg="#000",
                                   font=("Courier", 11, "bold"),
                                   relief="flat", cursor="hand2")
        self._ind_btn.pack(side="left", padx=4)
        tk.Button(btn_row, text=" BUSCAR PRECO MERCADO ",
                  command=self._ind_fetch_price,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)
        tk.Button(btn_row, text=" EXPORTAR ",
                  command=self._export_ind,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)
        self._ind_spin = ttk.Progressbar(btn_row, mode="indeterminate", length=80)
        self._ind_spin.pack(side="left", padx=8)

        # ── Painel de resultado ────────────────────────────────────────────────
        res_fr = tk.LabelFrame(p, text="  Resultado  ",
                               bg=COLORS["panel"], fg=COLORS["header"],
                               font=("Courier", 11, "bold"),
                               highlightbackground=COLORS["border"], highlightthickness=1)
        res_fr.pack(fill="both", expand=True, padx=6, pady=(2, 6))

        self._ind_result_lbl = tk.Label(
            res_fr,
            text="Configure os materiais e clique CALCULAR.",
            bg=COLORS["panel"], fg=COLORS["subtext"],
            font=("Courier", 11), justify="left", anchor="nw")
        self._ind_result_lbl.pack(fill="both", expand=True, padx=10, pady=8)

        # Auto-lookup nome pelo Type ID (busca no cache local + ESI)
        self._bind_tid_lookup(self._ind_prod_tid, self._ind_prod_name, None)

        # Botões de ação para o produto fabricado
        ind_act = tk.Frame(res_fr, bg=COLORS["panel"])
        ind_act.pack(fill="x", padx=4, pady=(2, 4))
        tk.Button(ind_act, text=" GRÁFICO ",     command=self._ind_open_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(ind_act, text=" PREVER ",      command=self._ind_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(ind_act, text=" BACKTEST ",    command=self._ind_open_bt,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(ind_act, text=" ANALISAR IA ", command=self._ind_open_ai,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)

    def _ind_add_material(self):
        try:
            tid   = int(self._ind_mat_tid.get().strip())
            qty   = int(self._ind_mat_qty.get().strip())
            price = float(self._ind_mat_price.get().strip().replace(",", ""))
        except ValueError:
            messagebox.showerror("Erro", "Type ID, Qtd e Preco devem ser numericos.")
            return
        name = self._item_cache.get(tid, {}).get("name", str(tid))
        # Resolve nome se necessario
        if name == str(tid):
            threading.Thread(target=lambda: (
                _ensure_names([tid], self._item_cache),
                self.after(0, self._ind_refresh_mat_tree)
            ), daemon=True).start()
        # Atualiza ou adiciona
        for m in self._ind_materials:
            if m["type_id"] == tid:
                m["quantity"]   = qty
                m["unit_price"] = price
                self._ind_refresh_mat_tree(); return
        self._ind_materials.append({"type_id": tid, "quantity": qty,
                                     "unit_price": price})
        self._ind_refresh_mat_tree()
        self._ind_mat_tid.delete(0, "end")
        self._ind_mat_qty.delete(0, "end")
        self._ind_mat_price.delete(0, "end")

    def _ind_remove_material(self):
        sel = self._ind_mat_tree.selection()
        if not sel: return
        try:
            tid = int(self._ind_mat_tree.item(sel[0], "values")[0])
        except (ValueError, IndexError):
            return
        self._ind_materials = [m for m in self._ind_materials if m["type_id"] != tid]
        self._ind_refresh_mat_tree()

    def _ind_clear_materials(self):
        self._ind_materials = []
        self._ind_refresh_mat_tree()

    def _ind_refresh_mat_tree(self):
        """
        Atualiza a tabela de materiais com as colunas:
          Qtd/run → Qtd c/ME → Qtd total (× runs necessários para qtd desejada)
        """
        me = max(0, min(10, int(self._fread(self._ind_me, 0))))
        me_factor = 1.0 - me / 100.0
        # Calcula runs necessários para atingir a quantidade desejada
        try:
            want      = max(1, int(self._fread(self._ind_want_qty, 1)))
            qty_p_run = max(1, int(self._fread(self._ind_qty, 1)))
            runs      = max(1, int(self._fread(self._ind_runs, 1)))
        except Exception:
            want = qty_p_run = runs = 1
        # Se runs não foi preenchido manualmente, calcula automaticamente
        total_produced = qty_p_run * runs
        if total_produced < want:
            runs = math.ceil(want / max(1, qty_p_run))

        for row in self._ind_mat_tree.get_children():
            self._ind_mat_tree.delete(row)
        for m in self._ind_materials:
            tid       = m["type_id"]
            qty       = m["quantity"]           # por run
            price     = m.get("unit_price", 0)
            name      = self._item_cache.get(tid, {}).get("name", str(tid))
            qty_me    = max(1, math.ceil(qty * me_factor))  # com ME aplicado
            qty_total = qty_me * runs                        # total para todos os runs
            tag = ("hot",) if price == 0 else ()
            self._ind_mat_tree.insert("", "end", iid=str(tid), tags=tag, values=(
                tid, name,
                f"{qty:,}",
                f"{qty_me:,}",
                f"{qty_total:,}",
                f"{price:,.2f}",
                f"{qty_total * price:,.2f}"))

    def _ind_auto_load(self):
        """
        Passo a passo automático:
          1. Lê o Type ID do produto digitado
          2. Busca materiais via Fuzzwork API (cache SQLite 24h)
          3. Calcula runs necessários para a quantidade desejada
          4. Busca preços de mercado de todos os materiais em paralelo
          5. Preenche _ind_materials, atualiza tabela e dispara o cálculo
        """
        try:
            tid = int(self._ind_prod_tid.get().strip())
        except ValueError:
            messagebox.showerror("Erro", "Digite um Type ID válido."); return
        if self._loading_ind:
            return
        self._loading_ind = True
        self._ind_auto_btn.config(state="disabled")
        self._ind_btn.config(state="disabled")
        self._ind_spin.start(10)
        self._ind_status_lbl.config(text="Buscando blueprint...", fg=COLORS["yellow"])

        try:
            want_qty  = max(1, int(self._fread(self._ind_want_qty, 100)))
        except Exception:
            want_qty  = 100
        rname = self._ind_reg.get()
        rid   = REGIONS.get(rname, 10000002)

        threading.Thread(
            target=self._worker_ind_auto,
            args=(tid, want_qty, rid, rname),
            daemon=True).start()

    def _worker_ind_auto(self, input_tid: int, want_qty: int,
                          region_id: int, region_name: str):
        """
        Worker thread do AUTO-PREENCHER.
        Aceita TANTO o type_id do blueprint (807) quanto do produto (209).
        A Fuzzwork API resolve internamente qual é o blueprint correto.
        """
        try:
            # ── 1. Resolve nome via cache ou ESI ──────────────────────────────
            self.after(0, lambda: self._ind_status_lbl.config(
                text=f"Buscando informações do ID {input_tid}...",
                fg=COLORS["yellow"]))

            _ensure_names([input_tid], self._item_cache)
            input_name = self._item_cache.get(input_tid, {}).get("name", "")

            # Se nome vazio, tenta ESI /universe/types/ como fallback
            if not input_name:
                try:
                    info = fetch_blueprint_info(input_tid)
                    input_name = info.get("name", str(input_tid))
                    if input_name and input_name != str(input_tid):
                        self._item_cache[input_tid] = {"name": input_name}
                except Exception:
                    input_name = str(input_tid)

            display_name = input_name or str(input_tid)
            self.after(0, lambda n=display_name: self._ind_prod_name.set(n))

            # ── 2. Busca materiais (Fuzzwork aceita tanto produto quanto BP) ──
            self.after(0, lambda: self._ind_status_lbl.config(
                text=f"Consultando blueprint: {display_name[:40]}...",
                fg=COLORS["yellow"]))

            bp_data       = fetch_blueprint_materials(input_tid, activity_id=1)
            materials_raw = bp_data.get("materials", [])
            qty_per_run   = bp_data.get("qty_per_run", 1)
            source        = bp_data.get("source", "?")

            if not materials_raw:
                is_bp = "Blueprint" in display_name or "blueprint" in display_name
                def _show_not_found(tid=input_tid, name=display_name, bp=is_bp):
                    if bp:
                        hint = (
                            f"Blueprint reconhecido: {name}\n"
                            f"Mas a API não retornou materiais.\n\n"
                            "Causas prováveis:\n"
                            "• API Fuzzwork temporariamente indisponível\n"
                            "• Blueprint não-fabricável (ex: T3, único, NPC drop)\n\n"
                            "Adicione os materiais manualmente ou tente novamente.")
                    else:
                        hint = (
                            f"Sem materiais para: ID {tid} = '{name}'\n\n"
                            "Dicas:\n"
                            "• Use o Type ID do BLUEPRINT, não do produto\n"
                            "  Ex: 'Nova Heavy Missile Blueprint' = 807\n"
                            "  (não 209 que é o produto)\n"
                            "• No EVE: clique direito no item → Mostrar Informações\n"
                            "  → aba Industry → o Type ID que aparece é do blueprint\n"
                            "• Minerals e combustível não têm blueprint\n"
                            "• API Fuzzwork pode estar temporariamente fora do ar")
                    messagebox.showinfo("Materiais não encontrados", hint)
                    self._ind_status_lbl.config(
                        text=f"Sem materiais para ID {tid}. Verifique se é blueprint.",
                        fg=COLORS["red"])
                self.after(0, _show_not_found)
                return

            # ── 3. Calcula runs ───────────────────────────────────────────────
            runs_needed = max(1, math.ceil(want_qty / max(1, qty_per_run)))
            total_units = runs_needed * qty_per_run

            self.after(0, lambda r=runs_needed, q=qty_per_run: (
                self._ind_runs.set(str(r)),
                self._ind_qty.set(str(q))))

            self.after(0, lambda: self._ind_status_lbl.config(
                text=(f"{len(materials_raw)} materiais ({source}) | "
                      f"{runs_needed} runs → {total_units:,} un | buscando preços..."),
                fg=COLORS["yellow"]))

            # ── 4. Resolve nomes dos materiais ────────────────────────────────
            mat_tids = [m["type_id"] for m in materials_raw]
            _ensure_names(mat_tids, self._item_cache)
            for m in materials_raw:
                cached = self._item_cache.get(m["type_id"], {}).get("name")
                if cached and cached != str(m["type_id"]):
                    m["name"] = cached

            # ── 5. Preços em paralelo ─────────────────────────────────────────
            prices = fetch_material_prices_bulk(mat_tids, region_id)

            # ── 6. Monta materiais com preços ─────────────────────────────────
            materials_final = []
            missing_prices  = []
            for m in materials_raw:
                tid_m = m["type_id"]
                sell  = prices.get(tid_m, {}).get("sell_min", 0)
                if sell <= 0:
                    missing_prices.append(m.get("name", str(tid_m)))
                materials_final.append({
                    "type_id":    tid_m,
                    "quantity":   m["quantity"],
                    "unit_price": sell,
                })

            # ── 7. Atualiza tabela ────────────────────────────────────────────
            def _update_ui():
                self._ind_materials = materials_final
                self._ind_refresh_mat_tree()
                n_miss = len(missing_prices)
                msg = (f"✓ {len(materials_final)} materiais | "
                       f"{len(materials_final)-n_miss} com preço | "
                       f"{runs_needed} runs → {total_units:,} un | {source}")
                if n_miss:
                    msg += f" | ⚠ {n_miss} sem preço: {', '.join(missing_prices[:3])}"
                self._ind_status_lbl.config(
                    text=msg,
                    fg=COLORS["green"] if not n_miss else COLORS["yellow"])

            self.after(0, _update_ui)

            # ── 8. Maior buy order (buy_max) do PRODUTO fabricado ────────────
            # Usa _get_product_tid para resolver blueprint_id → product_id.
            # Ex: 807 (blueprint) → 209 (Nova Heavy Missile)
            price_tid = _get_product_tid(input_tid, bp_data, self._item_cache)
            if not price_tid:
                price_tid = input_tid  # fallback: usa o próprio input_tid

            self.after(0, lambda: self._ind_status_lbl.config(
                text=f"Buscando maior buy order (produto ID {price_tid})...",
                fg=COLORS["yellow"]))

            prod_price = fetch_live_price(region_id, price_tid)
            if prod_price and prod_price.get("buy_max", 0) > 0:
                _buy_max   = prod_price["buy_max"]
                _prod_name = self._item_cache.get(price_tid, {}).get("name", "")
                self.after(0, lambda p=_buy_max, n=_prod_name: (
                    self._ind_mkt_price.set(f"{p:.4f}"),
                    self._ind_status_lbl.config(
                        text=f"Buy max {n or price_tid}: {p:,.4f} ISK/un",
                        fg=COLORS["green"])))

            # ── 9. Dispara cálculo ────────────────────────────────────────────
            self.after(200, self._start_ind)

        except Exception as exc:
            log.exception("worker_ind_auto")
            _err = str(exc)
            self.after(0, lambda e=_err: (
                messagebox.showerror("Erro AUTO-PREENCHER", e),
                self._ind_status_lbl.config(
                    text=f"Erro: {e[:60]}", fg=COLORS["red"])))
        finally:
            self._loading_ind = False
            self.after(0, lambda: (
                self._ind_auto_btn.config(state="normal"),
                self._ind_btn.config(state="normal"),
                self._ind_spin.stop()))

    def _ind_fetch_price(self):
        """Busca o maior buy order do PRODUTO (não do blueprint) na região selecionada."""
        try:
            input_tid = int(self._ind_prod_tid.get().strip())
        except ValueError:
            messagebox.showerror("Erro", "Digite um Type ID válido."); return
        rname = self._ind_reg.get()
        rid   = REGIONS.get(rname, 10000002)
        self._ind_btn.config(state="disabled")
        self._ind_spin.start(10)
        def worker():
            try:
                # Resolve o type_id do produto (não do blueprint)
                bp_data  = _bp_cache_get(input_tid, 1) or {"prod_tid": 0, "bp_tid": input_tid}
                price_tid= _get_product_tid(input_tid, bp_data, self._item_cache)
                if not price_tid:
                    price_tid = input_tid
                data = fetch_live_price(rid, price_tid)
                if data:
                    price = data.get("buy_max") or data.get("sell_min") or 0
                    name  = self._item_cache.get(price_tid, {}).get("name", str(price_tid))
                    self.after(0, lambda: (
                        self._ind_mkt_price.set(f"{price:.4f}"),
                        self._set_status(f"Buy max de {name}: {price:,.4f} ISK/un")))
                else:
                    self.after(0, lambda: messagebox.showinfo(
                        "Sem dados", "Não foi possível obter o preço."))
            except Exception as exc:
                self.after(0, lambda e=str(exc): messagebox.showerror("Erro", e))
            finally:
                self.after(0, lambda: (self._ind_btn.config(state="normal"),
                                       self._ind_spin.stop()))
        threading.Thread(target=worker, daemon=True).start()

    def _start_ind(self):
        if self._loading_ind: return
        if not self._ind_materials:
            messagebox.showinfo("Sem materiais",
                "Adicione materiais ou use AUTO-PREENCHER."); return
        try:
            want_qty   = max(1, int(self._fread(self._ind_want_qty, 1)))
            qty_p_run  = max(1, int(self._fread(self._ind_qty, 1)))
            me         = max(0, min(10, int(self._fread(self._ind_me, 0))))
            fac_tax    = self._fread(self._ind_fac_tax, 5) / 100
            ci         = self._fread(self._ind_ci, 5) / 100
            str_mat    = max(0.0, self._fread(self._ind_str_mat, 0) / 100)
            str_job    = max(0.0, self._fread(self._ind_str_job, 0) / 100)
            mkt_price  = float(self._ind_mkt_price.get().strip().replace(",", "") or "0")
            # Runs calculados automaticamente
            runs = max(1, math.ceil(want_qty / qty_p_run))
            self._ind_runs.set(str(runs))
        except ValueError:
            messagebox.showerror("Erro", "Parâmetros inválidos."); return
        rname = self._ind_reg.get()
        params = dict(qty=qty_p_run, runs=runs, me=me,
                      fac_tax=fac_tax, ci=ci,
                      str_mat=str_mat, str_job=str_job,
                      mkt_price=mkt_price, want_qty=want_qty,
                      region_name=rname)
        self._loading_ind = True
        self._ind_btn.config(state="disabled")
        self._ind_spin.start(10)
        threading.Thread(target=self._worker_ind,
                         args=(list(self._ind_materials), params),
                         daemon=True).start()

    def _worker_ind(self, materials: list, params: dict):
        """
        Calcula o custo total de fabricação e compara com o preço de mercado.
        Inclui:
          - Custo de materiais (usando qtd com ME × preço de mercado atual)
          - Job cost (fórmula CCP: EIV × system_cost_index)
          - Taxa de facilidade (facility_tax × job_cost)
          - Breakdown por unidade produzida
          - Comparação direta: fabricar vs comprar no mercado
        """
        try:
            self.after(0, lambda: self._set_status("Buscando preço ajustado CCP..."))
            adj_prices = fetch_adjusted_prices()
            result = calculate_industry_cost(
                materials            = materials,
                product_qty          = params["qty"],
                facility_tax         = params["fac_tax"],
                system_cost_index    = params["ci"],
                adjusted_prices      = adj_prices,
                me_level             = params["me"],
                runs                 = params["runs"],
                structure_mat_bonus  = params.get("str_mat", 0.0),
                structure_job_bonus  = params.get("str_job", 0.0),
            )
            mkt_price   = params["mkt_price"]
            want_qty    = params.get("want_qty", params["qty"] * params["runs"])
            total_units = result["product_qty_total"]

            result["market_price"]      = mkt_price
            result["want_qty"]          = want_qty
            result["profit_unit"]       = round(mkt_price - result["cost_per_unit"], 2)
            result["profit_pct"]        = (round(result["profit_unit"] /
                                                  result["cost_per_unit"] * 100, 2)
                                            if result["cost_per_unit"] > 0 else 0)
            # Receita total = buy_max × total_units_produzidas
            # (quanto o fabricante recebe ao vender tudo aos buy orders)
            result["total_market_value"]= round(mkt_price * total_units, 2)
            result["total_profit"]      = round(result["total_market_value"] -
                                                  result["total_cost"], 2)
            result["roi"]               = (round(result["total_profit"] /
                                                  result["total_cost"] * 100, 2)
                                            if result["total_cost"] > 0 else 0)
            # buy_cost_total = buy_max × want_qty = receita de vender toda a produção
            result["buy_cost_total"]    = round(mkt_price * want_qty, 2)
            result["fabricate_saves"]   = round(result["buy_cost_total"] -
                                                  result["total_cost"], 2)
            # Veredito com análise qualitativa
            if mkt_price <= 0:
                result["verdict"]       = "⚠ SEM PREÇO DE MERCADO"
                result["verdict_color"] = "yellow"
                result["advice"]        = "Configure o maior buy order do produto para análise completa."
            elif result["profit_unit"] > 0:
                roi_str = f"{result['roi']:+.1f}%"
                result["verdict"]       = "✓ FABRICAR — LUCRATIVO"
                result["verdict_color"] = "green"
                result["advice"]        = (
                    f"Receita: {result['buy_cost_total']:,.0f} ISK "
                    f"({want_qty:,} un × {mkt_price:.4f} buy max) "
                    f"− Custo: {result['total_cost']:,.0f} ISK "
                    f"= Lucro: {result['total_profit']:+,.0f} ISK ({roi_str} ROI)")
            elif result["profit_unit"] > -result["cost_per_unit"] * 0.05:
                result["verdict"]       = "≈ MARGEM MÍNIMA — AVALIAR"
                result["verdict_color"] = "yellow"
                result["advice"]        = (
                    f"Margem de {result['profit_pct']:+.1f}% por unidade. "
                    "Considere negociar taxas ou aguardar alta do buy order.")
            else:
                result["verdict"]       = "✗ NÃO LUCRATIVO"
                result["verdict_color"] = "red"
                result["advice"]        = (
                    f"Custo/un ({result['cost_per_unit']:,.4f} ISK) > "
                    f"buy max ({mkt_price:,.4f} ISK). "
                    f"Prejuízo de {abs(result['profit_unit']):,.4f} ISK/un. "
                    "Reduza o Cost Index ou aguarde valorização.")

            self._ind_result = result
            _ensure_names([m["type_id"] for m in materials], self._item_cache)
            self.after(0, self._pop_ind)
            self.after(0, self._dash_refresh)
        except Exception as exc:
            log.exception("worker_ind")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro Indústria", e))
        finally:
            self._loading_ind = False
            self.after(0, lambda: (self._ind_btn.config(state="normal"),
                                   self._ind_spin.stop()))

    def _pop_ind(self):
        """
        Exibe o relatório completo de fabricação, estruturado em 4 blocos:
          1. Veredicto e conselho estratégico
          2. Breakdown de custos (materiais, job, taxa)
          3. Análise de lucratividade por UNIDADE (dividido por qty_per_run)
          4. Comparação fabricar vs comprar
        """
        r = self._ind_result
        if not r: return
        self._ind_refresh_mat_tree()

        vc = r.get("verdict_color", "green")
        color = COLORS.get(vc, COLORS["green"])

        saves    = r.get("fabricate_saves", 0)
        saves_str= f"{saves:+,.0f} ISK"
        saves_lbl= "Lucro total (receita-custo):"

        buy_total  = r.get("buy_cost_total", 0)
        tot_profit = r.get("total_profit", 0)
        roi        = r.get("roi", 0)
        qty_total  = r.get("product_qty_total", 1)
        runs_n     = r.get("runs", 1)
        # qty_per_run = total produzido / runs
        qty_per_run = qty_total // runs_n if runs_n > 0 else qty_total

        text = (
            f"{'═'*64}\n"
            f"  {r.get('verdict', '?')}\n"
            f"{'═'*64}\n"
            f"  {r.get('advice', '')}\n"
            f"{'─'*64}\n"
            f"\n"
            f"  ── MATERIAIS  (ME{r['me_level']}"
            + (f" + Estrut.{r['structure_mat_bonus']:.0f}%" if r.get('structure_mat_bonus', 0) else "")
            + f")  ×{runs_n} run — {qty_per_run}/run ──\n"
            f"  Custo dos materiais:  {r['total_mat_cost']:>18,.2f} ISK\n"
            f"\n"
            f"  ── JOB COST ──────────────────────────────────────────────\n"
            f"  EIV (Adjusted Prices): {r['eiv']:>17,.2f} ISK\n"
            f"  Job gross (CI {r['system_cost_index']*100:.2f}%):  {r['job_gross_before']:>17,.2f} ISK\n"
            + (f"  Bônus estrutura (-{r['structure_job_bonus']:.0f}%):  {-r['job_struct_saving']:>17,.2f} ISK\n"
               if r.get('structure_job_bonus', 0) else "")
            + f"  Job cost final:       {r['job_cost']:>18,.2f} ISK\n"
            f"  SCC Surcharge (4%EIV):{r['scc_surcharge']:>18,.2f} ISK\n"
            + (f"  Taxa estrutura ({r['facility_tax_pct']:.0f}%):  {r['facility_cost']:>18,.2f} ISK\n"
               if r.get('facility_cost', 0) else "")
            + f"  {'─'*46}\n"
            f"  TOTAL JOB COST:       {r['total_job_cost']:>18,.2f} ISK\n"
            f"\n"
            f"  ── TOTAL ─────────────────────────────────────────────────\n"
            f"  Mat + Job cost:       {r['total_cost']:>18,.2f} ISK\n"
            f"  Unidades produzidas:  {qty_total:>15,}  ({runs_n}×{qty_per_run}/run)\n"
            f"  Custo por unidade:    {r['cost_per_unit']:>18,.4f} ISK\n"
            f"\n"
            f"  ── ANÁLISE DE MERCADO (por unidade) ──────────────────────\n"
            f"  Maior buy order:      {r['market_price']:>18,.4f} ISK/un\n"
            f"  Lucro por unidade:    {r['profit_unit']:>+18,.4f} ISK\n"
            f"  Margem líquida:       {r['profit_pct']:>+17.2f}%\n"
            f"  ROI total:            {roi:>+17.2f}%\n"
            f"\n"
            f"  ── RECEITA E LUCRO ({r.get('want_qty',0):,} unidades produzidas) ──────\n"
            f"  Custo total fabricação:{r['total_cost']:>17,.2f} ISK\n"
            f"  Receita (buy_max×qtd):{r.get('buy_cost_total',0):>18,.2f} ISK\n"
            f"  {saves_lbl:<24}{saves_str:>16}\n"
            f"  Lucro total estimado: {tot_profit:>+18,.0f} ISK\n"
        )
        self._ind_result_lbl.config(text=text, fg=color)



    def _export_ind(self):
        if not self._ind_result:
            messagebox.showinfo("Sem dados", "Calcule primeiro."); return
        r = self._ind_result
        # Monta linhas de materiais enriquecidas com nome
        rows = []
        for m in r.get("materials", []):
            name = self._item_cache.get(m["type_id"], {}).get("name", str(m["type_id"]))
            rows.append({**m, "nome": name})
        # Linha de resumo no topo
        rows = [{"nome": "RESUMO", "type_id": "", "quantity": "",
                 "unit_price": "", "subtotal": "",
                 "custo_total": r.get("total_cost",""),
                 "custo_unit": r.get("cost_per_unit",""),
                 "preco_mercado": r.get("market_price",""),
                 "lucro_unit": r.get("profit_unit",""),
                 "margem_pct": r.get("profit_pct",""),
                 "veredicto": r.get("verdict","")}] + rows
        headers = [("nome","Item/Resumo"),("type_id","Type ID"),
                   ("quantity","Qtd ME"),("unit_price","Preço/un"),("subtotal","Subtotal"),
                   ("custo_total","Custo Total"),("custo_unit","Custo/un"),
                   ("preco_mercado","Preço Mercado"),("lucro_unit","Lucro/un"),
                   ("margem_pct","Margem %"),("veredicto","Veredicto")]
        self._export_table("Indústria", headers, rows, "industria")
    # ══════════════════════════════════════════════════════════════════════════
    #  ABA BACKTESTING — simula estrategias sobre historico SQLite
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_bt(self):
        p = self._tab_bt

        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(6, 2))

        self._lbl(fr, "Type ID",        0, 0)
        self._bt_tid   = self._entry(fr, "", 1, 0, 12)
        self._lbl(fr, "Nome",           0, 2)
        self._bt_name  = self._entry(fr, "", 1, 2, 20)
        self._lbl(fr, "Regiao",         0, 4)
        self._bt_reg   = self._combo(fr, list(REGIONS.keys()), 1, 4)
        self._lbl(fr, "Dias",           0, 6)
        self._bt_days  = self._entry(fr, "90", 1, 6, 6)
        self._lbl(fr, "Capital (ISK)",  0, 8)
        self._bt_cap   = self._entry(fr, "1000000000", 1, 8, 16)
        self._lbl(fr, "Estrategia",     0, 10)
        self._bt_strat = self._combo(fr, [
            "Market Making (MM)",
            "Rota (compra/venda 7d)",
            "Ciclo (fase compra/venda)",
        ], 1, 10, 24)

        # Parametros extras por estrategia
        self._lbl(fr, "Spread min % (MM)", 0, 12)
        self._bt_sp_thr = self._entry(fr, "5.0", 1, 12, 7)
        self._lbl(fr, "Net min % (MM)",   0, 14)
        self._bt_net    = self._entry(fr, "1.0", 1, 14, 7)

        self._bt_btn = tk.Button(fr, text="  EXECUTAR BACKTEST  ",
                                  command=self._start_bt,
                                  bg="#7f4adf", fg="#fff",
                                  font=("Courier", 11, "bold"),
                                  relief="flat", cursor="hand2")
        self._bt_btn.grid(row=1, column=16, padx=8)
        tk.Button(fr, text=" EXPORTAR ",
                  command=self._export_bt,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=17, padx=4)
        self._bt_spin = ttk.Progressbar(fr, mode="indeterminate", length=80)
        self._bt_spin.grid(row=1, column=18, padx=6)

        tk.Label(fr,
                 text="Usa historico do banco SQLite local — sem requests ESI",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 9)
                 ).grid(row=0, column=10, columnspan=8, padx=4, sticky="w")

        # Painel de metricas
        met_fr = tk.Frame(p, bg=COLORS["panel"],
                          highlightbackground=COLORS["border"], highlightthickness=1)
        met_fr.pack(fill="x", padx=6, pady=(2, 2))
        self._bt_meta_lbl = tk.Label(
            met_fr,
            text="Configure e execute o backtest.",
            bg=COLORS["panel"], fg=COLORS["subtext"],
            font=("Courier", 11), justify="left", anchor="w")
        self._bt_meta_lbl.pack(fill="x", padx=10, pady=5)

        # Frame do grafico
        self._bt_chart_frame = tk.Frame(p, bg=COLORS["bg"])
        self._bt_chart_frame.pack(fill="both", expand=True, padx=4, pady=2)

        # Tabela de trades
        tbl_fr = tk.Frame(p, bg=COLORS["bg"])
        tbl_fr.pack(fill="x", padx=6, pady=(0, 6))
        self._bt_tree = self._make_tree(tbl_fr, [
            ("buy_date",  "Compra em",   100),
            ("sell_date", "Venda em",    115),
            ("buy_price", "Preco compra", 120),
            ("sell_price","Preco venda",  120),
            ("qty",       "Qtd",          80),
            ("pnl",       "P&L (ISK)",   120),
            ("pnl_pct",   "P&L %",        75),
        ], height=7)
        for cid, lbl in [
            ("buy_date","Compra em"),("sell_date","Venda em"),
            ("buy_price","Preco compra"),("sell_price","Preco venda"),
            ("qty","Qtd"),("pnl","P&L (ISK)"),("pnl_pct","P&L %")]:
            self._bt_tree.heading(cid, text=lbl)
        self._bt_tree.tag_configure("win",  foreground=COLORS["green"])
        self._bt_tree.tag_configure("loss", foreground=COLORS["red"])

        # Auto-lookup nome pelo Type ID (busca no cache local + ESI)
        self._bind_tid_lookup(self._bt_tid, self._bt_name, self._bt_reg)

    def _start_bt(self):
        if self._loading_bt: return
        try:
            tid = int(self._bt_tid.get().strip())
        except ValueError:
            messagebox.showerror("Erro", "Type ID invalido."); return
        params = dict(
            rname    = self._bt_reg.get(),
            days     = max(14, int(self._fread(self._bt_days, 90))),
            capital  = self._fread(self._bt_cap, 1_000_000_000),
            strategy_raw = self._bt_strat.get(),
            sp_thr   = self._fread(self._bt_sp_thr, 5.0),
            net_min  = self._fread(self._bt_net, 1.0),
            name     = self._bt_name.get().strip() or str(tid),
        )
        br, ac = self._gskills()
        params["br"] = br; params["ac"] = ac
        self._loading_bt = True
        self._bt_btn.config(state="disabled")
        self._bt_spin.start(10)
        threading.Thread(target=self._worker_bt,
                         args=(tid, params), daemon=True).start()

    def _worker_bt(self, tid: int, params: dict):
        try:
            rid  = REGIONS.get(params["rname"], 10000002)
            _warm_hist_cache(rid)  # pré-aquece cache RAM antes de _get_hist_cached
            hist = _get_hist_cached(rid, tid)
            if not hist or len(hist) < 14:
                self.after(0, lambda: messagebox.showinfo(
                    "Sem dados", f"Historico insuficiente ({len(hist)} dias)."))
                return

            raw_strat = params["strategy_raw"]
            if "MM" in raw_strat:
                strategy = "mm"
            elif "Ciclo" in raw_strat:
                strategy = "cycle_buy"
            else:
                strategy = "route"

            result = backtest_strategy(
                hist,
                strategy             = strategy,
                capital              = params["capital"],
                days_back            = params["days"],
                spread_threshold_pct = params["sp_thr"],
                target_net_pct       = params["net_min"],
                broker_rel           = params["br"],
                accounting           = params["ac"],
            )
            if not result:
                self.after(0, lambda: messagebox.showinfo(
                    "Sem resultado", "Historico insuficiente para backtest.")); return

            result["item_name"] = params["name"]
            result["type_id"]   = tid
            self._bt_result = result
            self.after(0, self._pop_bt)
        except Exception as exc:
            log.exception("worker_bt")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro Backtest", e))
        finally:
            self._loading_bt = False
            self.after(0, lambda: (self._bt_btn.config(state="normal"),
                                   self._bt_spin.stop()))

    def _pop_bt(self):
        r = self._bt_result
        if not r: return

        pnl_col = COLORS["green"] if r["total_pnl"] >= 0 else COLORS["red"]
        strat_labels = {
            "mm": "Market Making", "route": "Rota 7 dias",
            "cycle_buy": "Ciclo compra/venda"
        }
        strat_lbl = strat_labels.get(r["strategy"], r["strategy"])

        meta = (
            f"{r.get('item_name','')}  |  "
            f"Estrategia: {strat_lbl}  |  "
            f"Periodo: {r['days_back']} dias  |  "
            f"Trades: {r['n_trades']}\n"
            f"Capital inicial: {r['capital_start']:>14,.0f} ISK  →  "
            f"Final: {r['capital_end']:>14,.0f} ISK  |  "
            f"P&L: {r['total_pnl']:>+14,.0f} ISK ({r['total_pnl_pct']:>+.2f}%)\n"
            f"Win rate: {r['win_rate']:.1f}%  |  "
            f"Avg win: {r['avg_win']:,.0f}  |  "
            f"Avg loss: {r['avg_loss']:,.0f}  |  "
            f"Profit factor: {r['profit_factor']:.2f}  |  "
            f"Max drawdown: {r['max_drawdown']:.1f}%  |  "
            f"Sharpe: {r['sharpe']:.3f}")
        self._bt_meta_lbl.config(text=meta, fg=pnl_col)

        # Tabela de trades
        for row in self._bt_tree.get_children():
            self._bt_tree.delete(row)
        for t in r["trades"]:
            tag = ("win",) if t["pnl"] >= 0 else ("loss",)
            self._bt_tree.insert("", "end", tags=tag, values=(
                t["buy_date"], t["sell_date"],
                f"{t['buy_price']:>12,.2f}",
                f"{t['sell_price']:>12,.2f}",
                f"{t['qty']:.4f}",
                f"{t['pnl']:>+12,.0f}",
                f"{t['pnl_pct']:>+7.3f}%",
            ))

        # Grafico de equity curve
        if HAS_MATPLOTLIB and r["equity_curve"]:
            self._draw_bt_chart(r)

    def _draw_bt_chart(self, r: dict):
        for w in self._bt_chart_frame.winfo_children():
            w.destroy()
        dates_eq  = [d for d, _ in r["equity_curve"]]
        equity_eq = [e for _, e in r["equity_curve"]]
        n = len(dates_eq)
        xs = list(range(n))
        bg  = COLORS["bg"]
        pan = COLORS["panel"]

        fig = Figure(figsize=(13, 4.5), facecolor=bg)
        fig.subplots_adjust(left=0.07, right=0.96, top=0.88, bottom=0.14)

        ax = fig.add_subplot(1, 1, 1)
        ax.set_facecolor(pan)
        ax.tick_params(colors=COLORS["subtext"], labelsize=7)
        for sp in ax.spines.values():
            sp.set_edgecolor(COLORS["border"])

        # Linha de equity
        ax.plot(xs, equity_eq, color=COLORS["accent"], lw=2, label="Equity")
        ax.axhline(r["capital_start"], color=COLORS["subtext"],
                   lw=0.8, ls="--", alpha=0.6, label="Capital inicial")
        ax.fill_between(xs, r["capital_start"], equity_eq,
                         where=[e >= r["capital_start"] for e in equity_eq],
                         alpha=0.15, color=COLORS["green"])
        ax.fill_between(xs, r["capital_start"], equity_eq,
                         where=[e < r["capital_start"] for e in equity_eq],
                         alpha=0.15, color=COLORS["red"])

        # Marca trades como pontos
        for t in r["trades"]:
            if t["buy_date"] in dates_eq:
                xi = dates_eq.index(t["buy_date"])
                ax.scatter(xi, equity_eq[xi],
                           color=COLORS["green"] if t["pnl"] >= 0 else COLORS["red"],
                           s=30, zorder=5)

        pnl_col = COLORS["green"] if r["total_pnl"] >= 0 else COLORS["red"]
        ax.set_title(
            f"{r.get('item_name','')}  |  P&L: {r['total_pnl']:+,.0f} ISK "
            f"({r['total_pnl_pct']:+.2f}%)  |  "
            f"Sharpe: {r['sharpe']:.3f}  |  MaxDD: {r['max_drawdown']:.1f}%",
            color=pnl_col, fontsize=9, pad=5)
        ax.set_ylabel("ISK", color=COLORS["subtext"], fontsize=7)
        ax.legend(loc="upper left", fontsize=6,
                  facecolor=pan, labelcolor=COLORS["text"], framealpha=0.8)

        step = max(1, n // 12)
        ax.set_xticks(xs[::step])
        ax.set_xticklabels(dates_eq[::step], rotation=30,
                            ha="right", fontsize=6, color=COLORS["subtext"])

        canvas = FigureCanvasTkAgg(fig, master=self._bt_chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def _export_bt(self):
        if not self._bt_result:
            messagebox.showinfo("Sem dados", "Execute o backtest primeiro."); return
        r    = self._bt_result
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=([("Excel (.xlsx)","*.xlsx"),("CSV","*.csv")] if HAS_OPENPYXL else [("CSV","*.csv")]),
            initialfile=f"backtest_{r.get('item_name','').replace(' ','_')}.csv")
        if not path: return
        if path.endswith(".xlsx") and HAS_OPENPYXL:
            self._export_table("Backtest", [], [], "backtest")
            messagebox.showinfo("Dica", "Para exportação completa deste relatório\nuse CSV — o Excel exportará apenas as linhas tabulares.\nAbra o CSV no Excel para ver todos os dados.")
            return
        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["item", r.get("item_name","")])
                w.writerow(["estrategia", r["strategy"]])
                w.writerow(["periodo_dias", r["days_back"]])
                w.writerow(["capital_inicial", r["capital_start"]])
                w.writerow(["capital_final", r["capital_end"]])
                w.writerow(["pnl_total", r["total_pnl"]])
                w.writerow(["pnl_pct", r["total_pnl_pct"]])
                w.writerow(["n_trades", r["n_trades"]])
                w.writerow(["win_rate_pct", r["win_rate"]])
                w.writerow(["profit_factor", r["profit_factor"]])
                w.writerow(["max_drawdown_pct", r["max_drawdown"]])
                w.writerow(["sharpe", r["sharpe"]])
                w.writerow([])
                w.writerow(["buy_date","sell_date","buy_price","sell_price",
                             "qty","pnl","pnl_pct"])
                for t in r["trades"]:
                    w.writerow([t["buy_date"], t["sell_date"],
                                 t["buy_price"], t["sell_price"],
                                 t["qty"], t["pnl"], t["pnl_pct"]])
            messagebox.showinfo("Exportado", f"Salvo em:\n{path}")
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))
    # ══════════════════════════════════════════════════════════════════════════
    #  ABA HEATMAP — grade visual hub x grupo de itens
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_hm(self):
        p = self._tab_hm

        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(6, 2))

        self._hm_btn = tk.Button(fr, text="  GERAR HEATMAP  ",
                                  command=self._start_hm,
                                  bg=COLORS["orange"], fg="#000",
                                  font=("Courier", 12, "bold"),
                                  relief="flat", cursor="hand2")
        self._hm_btn.grid(row=1, column=0, padx=8, pady=4)

        tk.Button(fr, text=" DETALHAR CELULA ",
                  command=self._hm_drill_down,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=1, padx=4)

        self._hm_prog = ttk.Progressbar(fr, mode="determinate",
                                         length=140, maximum=100)
        self._hm_prog.grid(row=1, column=2, padx=8)
        self._hm_pct = tk.Label(fr, text="", bg=COLORS["panel"],
                                 fg=COLORS["subtext"], font=("Courier", 10), width=5)
        self._hm_pct.grid(row=1, column=3, padx=2)
        tk.Label(fr,
                 text="Score 0-100 por celula  |  Verde = boa oportunidade  |  "
                      "Clique numa celula para detalhar",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 9)).grid(row=0, column=0, columnspan=6,
                                            padx=8, sticky="w")

        # Canvas para o heatmap visual
        self._hm_canvas_frame = tk.Frame(p, bg=COLORS["bg"])
        self._hm_canvas_frame.pack(fill="both", expand=True, padx=6, pady=4)

        # Label de status / cell info
        self._hm_info_lbl = tk.Label(
            p,
            text="Clique em GERAR HEATMAP para calcular todos os hubs e grupos.",
            bg=COLORS["bg"], fg=COLORS["subtext"],
            font=("Courier", 10), justify="left")
        self._hm_info_lbl.pack(fill="x", padx=8, pady=(0, 4))

    def _start_hm(self):
        if self._loading_hm: return
        self._loading_hm = True
        self._hm_btn.config(state="disabled")
        self._hm_prog.config(value=0); self._hm_pct.config(text="")
        br, ac = self._gskills()
        threading.Thread(target=self._worker_hm,
                         args=(br, ac), daemon=True).start()

    def _worker_hm(self, br: int, ac: int):
        prog = self._make_prog(self._hm_prog, self._hm_pct)
        try:
            result = compute_full_heatmap(
                broker_rel  = br,
                accounting  = ac,
                progress_cb = prog,
            )
            self._hm_result = result
            self.after(0, lambda: (self._hm_prog.config(value=100),
                                   self._hm_pct.config(text="100%")))
            self.after(0, self._draw_heatmap)
        except Exception as exc:
            log.exception("worker_hm")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro Heatmap", e))
        finally:
            self._loading_hm = False
            self.after(0, lambda: self._hm_btn.config(state="normal"))

    def _draw_heatmap(self):
        """Desenha o heatmap como grade colorida de células tkinter com scroll."""
        for w in self._hm_canvas_frame.winfo_children():
            w.destroy()

        data    = self._hm_result
        hubs    = list(HEATMAP_HUBS.keys())
        groups  = list(HEATMAP_ITEM_GROUPS.keys())

        if not data:
            tk.Label(self._hm_canvas_frame,
                     text="Sem dados.", bg=COLORS["bg"],
                     fg=COLORS["subtext"], font=("Courier", 12)).pack(pady=20)
            return

        # ── Canvas com scrollbar vertical ─────────────────────────────────────
        hscroll = tk.Scrollbar(self._hm_canvas_frame, orient="horizontal")
        vscroll = tk.Scrollbar(self._hm_canvas_frame, orient="vertical")
        canvas  = tk.Canvas(self._hm_canvas_frame, bg=COLORS["bg"],
                            xscrollcommand=hscroll.set,
                            yscrollcommand=vscroll.set,
                            highlightthickness=0)
        hscroll.config(command=canvas.xview)
        vscroll.config(command=canvas.yview)
        vscroll.pack(side="right", fill="y")
        hscroll.pack(side="bottom", fill="x")
        canvas.pack(side="left", fill="both", expand=True)

        # Frame interno que vai dentro do canvas
        outer = tk.Frame(canvas, bg=COLORS["bg"])
        win_id = canvas.create_window((0, 0), window=outer, anchor="nw")

        def _on_frame_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(e):
            canvas.itemconfig(win_id, width=e.width)

        outer.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # Scroll com roda do mouse
        def _on_mousewheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>",   lambda e: canvas.yview_scroll(-1, "units"))
        canvas.bind_all("<Button-5>",   lambda e: canvas.yview_scroll( 1, "units"))

        def score_to_color(score: float) -> str:
            """Interpola de vermelho (0) → amarelo (50) → verde (100)."""
            s = max(0, min(100, score))
            if s < 50:
                r = 200
                g = int(s / 50 * 200)
            else:
                r = int((100 - s) / 50 * 200)
                g = 200
            return f"#{r:02x}{g:02x}30"

        # Linha de cabeçalho dos hubs
        tk.Label(outer, text="", width=14, bg=COLORS["bg"]).grid(
            row=0, column=0, padx=2, pady=2)
        for j, hub in enumerate(hubs):
            tk.Label(outer, text=hub,
                     bg=COLORS["border"], fg=COLORS["header"],
                     font=("Courier", 11, "bold"),
                     width=18, height=2,
                     relief="flat").grid(row=0, column=j+1, padx=2, pady=2)

        # Separador visual entre categorias de grupo
        _CATEGORY_BORDERS = {
            "Minerais Raros": True, "PLEX/Injector": True,
            "PI Tier 1 (P1)": True, "Moon Ore R4/R8": True,
            "Fragatas T1": True,    "Módulos Escudo": True,
            "Munições T2": True,    "Estruturas": True,
        }

        # Linhas de grupos
        self._hm_selected = None   # (hub, group)
        for i, group in enumerate(groups):
            pady_top = 6 if group in _CATEGORY_BORDERS else 2
            tk.Label(outer, text=group,
                     bg=COLORS["border"], fg=COLORS["accent"],
                     font=("Courier", 10), width=14, height=4,
                     wraplength=100, justify="center",
                     relief="flat").grid(row=i+1, column=0, padx=2,
                                         pady=(pady_top, 2))

            for j, hub in enumerate(hubs):
                cell = data.get(hub, {}).get(group, {})
                score    = cell.get("score", 0)
                net_m    = cell.get("net_margin", 0)
                vol_s    = cell.get("vol_score", 0)
                trend    = cell.get("trend", 0)
                bg_color = score_to_color(score)

                cell_txt = (
                    f"Score: {score:.0f}\n"
                    f"Net:   {net_m:+.1f}%\n"
                    f"Vol:   {vol_s:.2f}x\n"
                    f"Trend: {trend:+.3f}")

                lbl = tk.Label(outer,
                               text=cell_txt,
                               bg=bg_color,
                               fg="#ffffff",
                               font=("Courier", 9),
                               width=18, height=4,
                               justify="left",
                               relief="groove", bd=1,
                               cursor="hand2")
                lbl.grid(row=i+1, column=j+1, padx=2, pady=(pady_top, 2))

                # Click handler
                def on_click(h=hub, g=group, c=cell, s=score):
                    self._hm_selected = (h, g)
                    trend_dir = "▲" if c.get("trend",0)>0.01 else ("▼" if c.get("trend",0)<-0.01 else "◇")
                    self._hm_info_lbl.config(
                        fg=COLORS["green"] if s >= 50 else COLORS["yellow"],
                        text=(f"► {h} × {g}  |  Score: {s:.0f}/100  |  "
                              f"Net margin: {c.get('net_margin',0):+.2f}%  |  "
                              f"Vol score: {c.get('vol_score',0):.2f}x  |  "
                              f"Trend: {c.get('trend',0):+.4f} {trend_dir}  |  "
                              f"Itens com dados: {c.get('n_items',0)}  |  "
                              f"Double-click = DETALHAR"))
                lbl.bind("<Button-1>", lambda e, h=hub, g=group, c=cell, s=score:
                         on_click(h, g, c, s))
                lbl.bind("<Double-Button-1>", lambda e: self._hm_drill_down())

    def _hm_drill_down(self):
        """Abre análise detalhada para a célula selecionada do heatmap."""
        if not self._hm_result:
            messagebox.showinfo("Sem dados", "Gere o heatmap primeiro."); return
        sel = getattr(self, "_hm_selected", None)
        if not sel:
            messagebox.showinfo("Selecione", "Clique numa célula do heatmap."); return
        hub_name, group_name = sel
        hub_info = HEATMAP_HUBS.get(hub_name, {})
        rid      = hub_info.get("rid", 10000002)
        tids_raw = HEATMAP_ITEM_GROUPS.get(group_name, [])
        # Suporta tanto list quanto dict como valores de HEATMAP_ITEM_GROUPS
        tids = list(tids_raw.keys() if isinstance(tids_raw, dict) else tids_raw)
        if not tids:
            return

        # Abre janela de detalhe com scatter plot Score vs Vol
        if not HAS_MATPLOTLIB:
            messagebox.showinfo("Falta matplotlib", "pip install matplotlib"); return
        if self._chart_win and self._chart_win.winfo_exists():
            self._chart_win.destroy()
        win = tk.Toplevel(self)
        win.title(f"Detalhe: {hub_name} × {group_name}")
        win.configure(bg=COLORS["bg"])
        win.geometry("900x560")
        self._chart_win = win

        # Carrega dados em thread
        def worker():
            # Cache já aquecido por compute_full_heatmap — O(1) por item
            _warm_hist_cache(rid)   # garante que o cache está quente
            rows = []
            for tid in tids:
                hist = _get_hist_cached(rid, tid)
                name = self._item_cache.get(tid, {}).get("name", str(tid))
                if not hist or len(hist) < 7:
                    continue
                # hist já ordenado por date (ORDER BY na query bulk)
                hist_r = hist[-30:]
                prices = [float(d.get("average", 0)) for d in hist_r]
                highs  = [float(d.get("highest", 0)) for d in hist_r]
                lows   = [float(d.get("lowest",  0)) for d in hist_r]
                vols   = [int(d.get("volume", 0))    for d in hist_r]
                if not prices or prices[-1] <= 0:
                    continue
                spreads = [(hi-lo)/hi*100 if hi>0 else 0
                           for hi, lo in zip(highs, lows)]
                avg_sp  = sum(spreads)/len(spreads)
                avg_vol = sum(vols)/len(vols) if vols else 1
                rows.append({
                    "name":   name[:18],
                    "price":  prices[-1],
                    "spread": avg_sp,
                    "vol":    avg_vol,
                    "chg30":  ((prices[-1]-prices[0])/prices[0]*100
                                if prices[0]>0 else 0),
                })
            self.after(0, lambda: _draw(rows))

        def _draw(rows):
            if not rows:
                tk.Label(win, text="Sem dados suficientes.",
                         bg=COLORS["bg"], fg=COLORS["red"],
                         font=("Courier", 12)).pack(pady=20); return
            bg  = COLORS["bg"]
            pan = COLORS["panel"]
            fig = Figure(figsize=(10, 6), facecolor=bg)
            fig.subplots_adjust(left=0.08, right=0.96, top=0.88, bottom=0.14)
            ax  = fig.add_subplot(1, 1, 1)
            ax.set_facecolor(pan)
            ax.tick_params(colors=COLORS["subtext"], labelsize=7)
            for sp in ax.spines.values():
                sp.set_edgecolor(COLORS["border"])

            xs = [r["vol"]  for r in rows]
            ys = [r["spread"] for r in rows]
            cs = [COLORS["green"] if r["chg30"] >= 0 else COLORS["red"]
                  for r in rows]
            ax.scatter(xs, ys, c=cs, s=80, alpha=0.8, zorder=5)
            for r in rows:
                ax.annotate(r["name"], (r["vol"], r["spread"]),
                            fontsize=6, color=COLORS["text"],
                            xytext=(4, 4), textcoords="offset points")
            ax.set_xlabel("Volume médio 30d", color=COLORS["subtext"], fontsize=8)
            ax.set_ylabel("Spread médio 30d %", color=COLORS["subtext"], fontsize=8)
            ax.set_title(
                f"{hub_name} × {group_name}  |  "
                f"Verde = preco subindo 30d  |  Vermelho = caindo",
                color=COLORS["header"], fontsize=9, pad=5)
            canvas = FigureCanvasTkAgg(fig, master=win)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)

        threading.Thread(target=worker, daemon=True).start()

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA CATALOGO — Type ID, nome e grupo de todos os itens do cache
    # ══════════════════════════════════════════════════════════════════════════

    _CAT_COLS = {
        "type_id":       "Type ID",
        "name":          "Nome",
        "group_name":    "Grupo",
        "category_name": "Categoria",
        "volume":        "Volume (m³)",
    }

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA STATION TRADER — scanner inspirado no EVE Trade Guru
    # ══════════════════════════════════════════════════════════════════════════

    _ST_COLS = {
        "name":             "Item",
        "purchase_units":   "Compra (un.)",
        "daily_tradable":   "Vol/Dia",
        "sell_units_day":   "Venda/Dia",
        "bvs_ratio":        "B/S Ratio",
        "supply_units":     "Oferta (un.)",
        "dos":              "D.O.S.",
        "now_buy_price":    "Buy Atual",
        "now_sell_price":   "Sell Atual",
        "period_avg_price": "Preco Medio",
        "now_profit":       "Lucro Atual",
        "period_profit":    "Lucro Periodo",
        "capital_required": "Capital Req.",
        "now_roi":          "ROI Atual %",
        "period_roi":       "ROI Periodo %",
        "now_margin":       "Margem % (now)",
        "n_buy_orders":     "Buy Orders",
        "n_sell_orders":    "Sell Orders",
        "risk_flag":        "Risco",
        "group_name":       "Grupo",
        "category_name":    "Categoria",
    }

    def _build_tab_st(self):
        p = self._tab_st

        # ── Sub-notebook com 3 abas de filtro (igual ao EVE Guru) ─────────────
        filter_nb = ttk.Notebook(p)
        filter_nb.pack(fill="x", padx=6, pady=(6, 2))

        tab_pv  = tk.Frame(filter_nb, bg=COLORS["panel"])
        tab_pf  = tk.Frame(filter_nb, bg=COLORS["panel"])
        tab_rp  = tk.Frame(filter_nb, bg=COLORS["panel"])
        filter_nb.add(tab_pv,  text="  Purchase Volume  ")
        filter_nb.add(tab_pf,  text="  Profit Filters  ")
        filter_nb.add(tab_rp,  text="  Risk Profile  ")

        pad = {"padx": 8, "pady": 4}

        # helpers locais
        def lbl(parent, text, col):
            tk.Label(parent, text=text, bg=COLORS["panel"],
                     fg=COLORS["subtext"], font=("Courier", 10)
                     ).grid(row=0, column=col, padx=6, sticky="w")

        def sv_entry(parent, val, col, w=10):
            v = tk.StringVar(value=val)
            tk.Entry(parent, textvariable=v, width=w,
                     bg=COLORS["border"], fg=COLORS["text"],
                     insertbackground=COLORS["text"],
                     font=("Courier", 11), relief="flat", bd=2
                     ).grid(row=1, column=col, padx=6, pady=(0,4), sticky="w")
            return v

        def chkbox(parent, text, col, default=True):
            v = tk.BooleanVar(value=default)
            tk.Checkbutton(parent, text=text, variable=v,
                           bg=COLORS["panel"], fg=COLORS["text"],
                           selectcolor=COLORS["border"],
                           activebackground=COLORS["panel"],
                           font=("Courier", 10)
                           ).grid(row=1, column=col, padx=6, pady=(0,4), sticky="w")
            return v

        # ── Purchase Volume ────────────────────────────────────────────────────
        r0 = tk.Frame(tab_pv, bg=COLORS["panel"]); r0.pack(fill="x", **pad)
        r1 = tk.Frame(tab_pv, bg=COLORS["panel"]); r1.pack(fill="x", **pad)

        lbl(r0, "Days of Trade",      0)
        lbl(r0, "Regiao",             2)
        lbl(r0, "Fonte",              4)
        lbl(r0, "Max Capital (ISK)",  6)
        self._st_days    = sv_entry(r0, "3",  0, 7)
        self._st_reg     = self._combo(r0, list(REGIONS.keys()), 1, 2, width=22)
        self._st_mode    = self._combo(r0, ["CSV (cache)", "API (live)"], 1, 4, width=14)
        self._csv_age_label(r0, self._st_mode, [self._st_reg], 2, 4)
        self._st_max_cap = sv_entry(r0, "0",  6, 16)
        self._st_skip_exp = chkbox(r1, "Ignorar itens sem historico", 0, True)
        self._st_ranged   = chkbox(r1, "Ranged Buy Orders",           2, True)

        # ── Profit Filters ─────────────────────────────────────────────────────
        r2 = tk.Frame(tab_pf, bg=COLORS["panel"]); r2.pack(fill="x", **pad)
        r3 = tk.Frame(tab_pf, bg=COLORS["panel"]); r3.pack(fill="x", **pad)

        lbl(r2, "Min. Item Profit (ISK)",  0)
        lbl(r2, "Min. Order Margin %",     2)
        lbl(r2, "Min. Daily Tradable",     4)
        self._st_min_profit = sv_entry(r2, "5000000", 0, 16)
        self._st_min_margin = sv_entry(r2, "15",      2,  8)
        self._st_min_daily  = sv_entry(r2, "5",       4,  8)
        self._st_use_profit = chkbox(r3, "Ativar Min. Profit",  0, True)
        self._st_use_margin = chkbox(r3, "Ativar Min. Margin",  2, True)
        self._st_use_daily  = chkbox(r3, "Ativar Min. Daily",   4, True)

        # ── Risk Profile ───────────────────────────────────────────────────────
        r4 = tk.Frame(tab_rp, bg=COLORS["panel"]); r4.pack(fill="x", **pad)
        r5 = tk.Frame(tab_rp, bg=COLORS["panel"]); r5.pack(fill="x", **pad)

        lbl(r4, "Avg Price Period (dias)", 0)
        lbl(r4, "Min. Period ROI %",       2)
        lbl(r4, "Min. B/S Ratio",          4)
        lbl(r4, "Max. B/S Ratio",          6)
        lbl(r4, "Limit Buy to P.Low %",    8)
        lbl(r4, "Limit Sell to P.High %", 10)
        self._st_avg_period  = sv_entry(r4, "14",   0, 7)
        self._st_min_proi    = sv_entry(r4, "15",   2, 7)
        self._st_min_bvs     = sv_entry(r4, "0.7",  4, 7)
        self._st_max_bvs     = sv_entry(r4, "1.5",  6, 7)
        self._st_buy_plow    = sv_entry(r4, "-10",   8, 7)
        self._st_sell_phigh  = sv_entry(r4, "-5",   10, 7)
        self._st_use_proi    = chkbox(r5, "Ativar ROI Periodo",  0, True)
        self._st_use_bvs     = chkbox(r5, "Ativar B/S Filter",   2, True)
        self._st_use_plimits = chkbox(r5, "Ativar P.Low/High",   4, True)

        # ── Barra de acoes ─────────────────────────────────────────────────────
        act = tk.Frame(p, bg=COLORS["bg"])
        act.pack(fill="x", padx=6, pady=(2, 2))

        self._st_btn = tk.Button(act, text="  BUSCAR OPORTUNIDADES  ",
                                  command=self._start_st,
                                  bg=COLORS["accent"], fg="#000",
                                  font=("Courier", 12, "bold"),
                                  relief="flat", cursor="hand2")
        self._st_btn.pack(side="left", padx=4)
        tk.Button(act, text=" + WATCHLIST ",
                  command=self._st_add_wl,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)
        tk.Button(act, text=" ABRIR NO JOGO ",
                  command=self._st_open_game,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)
        tk.Button(act, text=" COPIAR NOME ",
                  command=self._st_copy_name,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)
        tk.Button(act, text=" PREVER ",
                  command=self._st_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)
        tk.Button(act, text=" BACKTEST ",
                  command=self._st_open_bt,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)
        tk.Button(act, text=" EXPORTAR ",
                  command=self._st_export,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)

        self._st_prog = ttk.Progressbar(act, mode="determinate",
                                         length=130, maximum=100)
        self._st_prog.pack(side="left", padx=8)
        self._st_pct = tk.Label(act, text="", bg=COLORS["bg"],
                                 fg=COLORS["subtext"], font=("Courier", 10), width=5)
        self._st_pct.pack(side="left")

        # ── Tabela principal ───────────────────────────────────────────────────
        self._st_tree = self._make_tree(p, [
            ("name",             "Item",              200),
            ("purchase_units",   "Compra (un.)",       95),
            ("daily_tradable",   "Vol/Dia",             75),
            ("sell_units_day",   "Venda/Dia",           80),
            ("bvs_ratio",        "B/S",                 55),
            ("supply_units",     "Oferta",              75),
            ("dos",              "D.O.S.",              60),
            ("now_buy_price",    "Buy (ISK)",          115),
            ("now_sell_price",   "Sell (ISK)",         115),
            ("period_avg_price", "Preco Med.",         115),
            ("now_profit",       "Lucro Atual",        110),
            ("period_profit",    "Lucro Period.",      110),
            ("capital_required", "Capital Req.",       115),
            ("now_roi",          "ROI Now %",           80),
            ("period_roi",       "ROI Period %",        88),
            ("now_margin",       "Margem %",            78),
            ("n_buy_orders",     "Buy Ord.",             68),
            ("n_sell_orders",    "Sell Ord.",            68),
            ("risk_flag",        "Risco",              210),
            ("group_name",       "Grupo",              150),
            ("category_name",    "Categoria",          120),
        ])
        for cid in self._ST_COLS:
            self._st_tree.heading(cid, text=self._ST_COLS[cid],
                                   command=lambda c=cid: self._sort_st(c))
        self._st_tree.bind("<Double-1>", self._st_on_dbl)
        self._st_tree.bind("<Button-3>", self._st_context_menu)
        self._st_tree.tag_configure("hot",    foreground=COLORS["green"])
        self._st_tree.tag_configure("warn",   foreground=COLORS["yellow"])
        self._st_tree.tag_configure("danger", foreground=COLORS["red"])

        bar = tk.Frame(p, bg=COLORS["bg"])
        bar.pack(fill="x", padx=6, pady=(0, 4))
        self._st_count = tk.Label(bar, text="0 oportunidades",
                                   fg=COLORS["subtext"], bg=COLORS["bg"],
                                   font=("Courier", 10))
        self._st_count.pack(side="left")
        tk.Label(bar,
                 text="Verde = ROI≥25%  |  Amarelo = risco  |  "
                      "Double-click = copiar Type ID  |  Dir = menu (copiar nome, prever, backtest...)",
                 fg=COLORS["subtext"], bg=COLORS["bg"],
                 font=("Courier", 10)).pack(side="right")

    def _start_st(self):
        if self._loading_st: return
        rname   = self._st_reg.get()
        br, ac  = self._gskills()
        params  = dict(
            rname      = rname,
            use_csv    = self._st_mode.get().startswith("CSV"),
            days       = max(0.5, self._fread(self._st_days, 3.0)),
            max_cap    = self._fread(self._st_max_cap, 0.0),
            skip_exp   = self._st_skip_exp.get(),
            ranged     = self._st_ranged.get(),
            min_profit = self._fread(self._st_min_profit, 0.0) if self._st_use_profit.get() else 0.0,
            min_margin = self._fread(self._st_min_margin, 0.0) if self._st_use_margin.get() else 0.0,
            min_daily  = self._fread(self._st_min_daily,  0.0) if self._st_use_daily.get()  else 0.0,
            avg_period = max(1, int(self._fread(self._st_avg_period, 14))),
            min_proi   = self._fread(self._st_min_proi,   0.0) if self._st_use_proi.get()    else 0.0,
            min_bvs    = self._fread(self._st_min_bvs,    0.0) if self._st_use_bvs.get()     else 0.0,
            max_bvs    = self._fread(self._st_max_bvs, 9999.0) if self._st_use_bvs.get()     else 9999.0,
            buy_plow   = self._fread(self._st_buy_plow,   0.0) if self._st_use_plimits.get() else 0.0,
            sell_phigh = self._fread(self._st_sell_phigh, 0.0) if self._st_use_plimits.get() else 0.0,
            br=br, ac=ac,
        )
        self._loading_st = True
        self._st_btn.config(state="disabled")
        self._st_prog.config(value=0); self._st_pct.config(text="")
        threading.Thread(target=self._worker_st, args=(params,), daemon=True).start()

    def _worker_st(self, params: dict):
        prog = self._make_prog(self._st_prog, self._st_pct)
        try:
            rname = params["rname"]
            rid   = REGIONS.get(rname, 10000002)
            if params["use_csv"]:
                orders = load_orders_from_csv(rname)
                if not orders:
                    self.after(0, lambda: messagebox.showerror(
                        "CSV nao encontrado",
                        "Baixe os dados na aba DADOS & CACHE.")); return
            else:
                prog("Baixando ordens via API...")
                orders = download_region_orders(rid, rname, prog)
                _orders_mem_cache[rname] = (time.time(), orders)

            prog("Analisando oportunidades...")
            results = station_trader_scan(
                orders            = orders,
                region_id         = rid,
                item_cache        = self._item_cache,
                broker_rel        = params["br"],
                accounting        = params["ac"],
                days_of_trade     = params["days"],
                max_capital       = params["max_cap"],
                skip_expired      = params["skip_exp"],
                min_item_profit   = params["min_profit"],
                min_order_margin  = params["min_margin"],
                min_daily_tradable= params["min_daily"],
                use_ranged_buy    = params["ranged"],
                avg_price_period  = params["avg_period"],
                min_period_roi    = params["min_proi"],
                min_bvs           = params["min_bvs"],
                max_bvs           = params["max_bvs"],
                limit_buy_plow    = params["buy_plow"],
                limit_sell_phigh  = params["sell_phigh"],
                blacklist         = self._blacklist,
                progress_cb       = prog,
            )
            self._st_results = results
            self.after(0, lambda: (self._st_prog.config(value=100),
                                   self._st_pct.config(text="100%")))
            self.after(0, self._pop_st)
            self.after(0, self._dash_refresh)
        except Exception as exc:
            log.exception("worker_st")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro Station Trader", e))
        finally:
            self._loading_st = False
            self.after(0, lambda: self._st_btn.config(state="normal"))

    def _pop_st(self, reorder_only=False):
        if reorder_only:
            for i, r in enumerate(self._st_results):
                self._st_tree.move(str(r["type_id"]), "", i)
            return
        for row in self._st_tree.get_children():
            self._st_tree.delete(row)

        def _fi(v):  # format ISK
            if v >= 1_000_000_000: return f"{v/1e9:.2f}B"
            if v >= 1_000_000:     return f"{v/1e6:.2f}M"
            if v >= 1_000:         return f"{v/1e3:.1f}K"
            return f"{v:,.0f}"

        for i, r in enumerate(self._st_results):
            proi = r["period_roi"]
            risk = r["risk_flag"]
            tag  = ("danger",) if risk else \
                   ("hot",)    if proi >= 25 else \
                   ("warn",)   if proi >= 10 else \
                   ("even" if i % 2 == 0 else "odd",)
            self._st_tree.insert("", "end", iid=str(r["type_id"]), tags=tag, values=(
                r["name"],
                f"{r['purchase_units']:,}",
                f"{r['daily_tradable']:,.0f}",
                f"{r['sell_units_day']:,.0f}",
                f"{r['bvs_ratio']:.2f}",
                f"{r['supply_units']:,}",
                f"{r['dos']:.1f}d",
                _fi(r["now_buy_price"]),
                _fi(r["now_sell_price"]),
                _fi(r["period_avg_price"]),
                _fi(r["now_profit"]),
                _fi(r["period_profit"]),
                _fi(r["capital_required"]),
                f"{r['now_roi']:.1f}%",
                f"{r['period_roi']:.1f}%",
                f"{r['now_margin']:.1f}%",
                str(r["n_buy_orders"]),
                str(r["n_sell_orders"]),
                r["risk_flag"],
                r["group_name"],
                r["category_name"],
            ))

        n_hot  = sum(1 for r in self._st_results if r["period_roi"] >= 25 and not r["risk_flag"])
        n_risk = sum(1 for r in self._st_results if r["risk_flag"])
        self._st_count.config(
            text=(f"{len(self._st_results)} oportunidades  |  "
                  f"{n_hot} excelentes (ROI≥25%)  |  {n_risk} com alerta"))

    def _sort_st(self, col):
        num = {"purchase_units","daily_tradable","sell_units_day","bvs_ratio",
               "supply_units","dos","now_buy_price","now_sell_price",
               "period_avg_price","now_profit","period_profit","capital_required",
               "now_roi","period_roi","now_margin","n_buy_orders","n_sell_orders"}
        rev = not self._st_sort_state.get(col, col in num)
        self._st_sort_state = {col: rev}
        self._st_results.sort(
            key=lambda r: (r.get(col) or 0) if col in num else r.get(col,""),
            reverse=rev)
        self._set_heading_arrow(self._st_tree, self._ST_COLS, col, rev)
        self._pop_st(reorder_only=True)

    def _st_selected(self):
        sel = self._st_tree.selection()
        if not sel: return None
        try:
            tid = int(sel[0])
            return next((r for r in self._st_results if r["type_id"] == tid), None)
        except (ValueError, TypeError):
            return None

    def _st_on_dbl(self, _):
        """Double-click: copia o Type ID do item selecionado."""
        r = self._st_selected()
        if not r: return
        self.clipboard_clear()
        self.clipboard_append(str(r["type_id"]))
        self._set_status(f"✓ Type ID {r['type_id']} copiado — {r['name']}")

    def _st_copy_name(self):
        """Copia o Nome do item selecionado para a área de transferência."""
        r = self._st_selected()
        if not r:
            messagebox.showinfo("Selecione", "Clique num item."); return
        self.clipboard_clear()
        self.clipboard_append(r["name"])
        self._set_status(f"✓ Nome copiado: {r['name']} (ID {r['type_id']})")

    def _st_context_menu(self, event):
        iid = self._st_tree.identify_row(event.y)
        if iid: self._st_tree.selection_set(iid)
        r = self._st_selected()
        menu = tk.Menu(self, tearoff=0,
                       bg=COLORS["panel"], fg=COLORS["text"],
                       activebackground=COLORS["border"],
                       activeforeground=COLORS["accent"],
                       font=("Courier", 11))
        if r:
            menu.add_command(label=f"  Copiar Type ID ({r['type_id']})", command=self._st_on_dbl)
            menu.add_command(label=f"  Copiar Nome ({r['name'][:24]})", command=self._st_copy_name)
            menu.add_separator()
            menu.add_command(label="  + Watchlist",  command=self._st_add_wl)
            menu.add_command(label="  Previsao",     command=self._st_open_fcast)
            menu.add_command(label="  Backtesting",  command=self._st_open_bt)
            menu.add_command(label="  Spread Hist.", command=self._st_open_sph)
            menu.add_command(label="  AI Analista",  command=self._st_open_ai)
        menu.tk_popup(event.x_root, event.y_root)

    def _st_add_wl(self):
        r = self._st_selected()
        if not r: messagebox.showinfo("Selecione", "Clique num item."); return
        if any(w["type_id"] == r["type_id"] for w in self._watchlist):
            messagebox.showinfo("Já existe", f"{r['name']} já está na Watchlist."); return
        self._watchlist.append({
            "type_id": r["type_id"], "name": r["name"],
            "alert_below": None, "alert_above": None,
            "last_sell": None, "last_buy": None, "last_net": None, "last_score": None,
        })
        save_watchlist(self._watchlist)
        self._pop_wl()
        messagebox.showinfo("Adicionado",
            f"{r['name']} → Watchlist\nROI Período: {r['period_roi']:.1f}%")

    def _st_open_game(self):
        r = self._st_selected()
        if not r: messagebox.showinfo("Selecione", "Clique num item."); return
        self.clipboard_clear(); self.clipboard_append(str(r["type_id"]))
        self._set_status(f"Type ID {r['type_id']} copiado. Cole no campo de busca do jogo.")

    def _st_open_fcast(self):
        r = self._st_selected()
        if not r: messagebox.showinfo("Selecione", "Clique num item."); return
        self._fcast_tid_e.set(str(r["type_id"]))
        self._fcast_name_e.set(r["name"])
        self._fcast_reg.set(self._st_reg.get())
        self._nb.select(self._tab_fcast)
        self.after(50, self._start_fcast)

    def _st_open_bt(self):
        r = self._st_selected()
        if not r: messagebox.showinfo("Selecione", "Clique num item."); return
        self._bt_tid.set(str(r["type_id"]))
        self._bt_name.set(r["name"])
        self._bt_reg.set(self._st_reg.get())
        self._nb.select(self._tab_bt)
        self.after(50, self._start_bt)

    def _st_open_sph(self):
        r = self._st_selected()
        if not r: messagebox.showinfo("Selecione", "Clique num item."); return
        self._sph_tid.set(str(r["type_id"]))
        self._sph_name.set(r["name"])
        self._sph_reg.set(self._st_reg.get())
        self._nb.select(self._tab_sph)
        self.after(50, self._start_sph)

    def _st_open_ai(self):
        r = self._st_selected()
        if not r: messagebox.showinfo("Selecione", "Clique num item."); return
        self._ai_tid.set(str(r["type_id"]))
        self._ai_name_e.set(r["name"])
        self._ai_reg.set(self._st_reg.get())
        self._nb.select(self._tab_ai)
        self.after(50, self._start_ai)

    def _st_export(self):
        headers = [(k, v) for k,v in self._ST_COLS.items()]
        self._export_table("Station Trader", headers,
                           self._st_results or [], "station_trader")
    # ── Helpers de seleção e navegação: MAX MARGEM, ESCASSEZ, DEMANDA/OFERTA, MOV. PREÇOS, ORÇAMENTO ──

    def _make_tab_handlers(prefix, tree_attr, results_attr, reg_attr):
        """Factory — não chamada diretamente; os métodos são definidos individualmente abaixo."""

    # ── MAX MARGEM ────────────────────────────────────────────────────────────
    def _mm_on_select(self, _e=None):
        sel = self._mm_tree.selection()
        if sel:
            try: self._mm_cur_tid = int(sel[0])
            except ValueError: pass

    def _mm_on_dbl(self, _e=None):
        w = self._mm_get_selected()
        if not w: return
        self.clipboard_clear(); self.clipboard_append(str(w["type_id"]))
        self._set_status(f"✓ Type ID {w['type_id']} copiado — {w['item']}")

    def _mm_get_selected(self):
        sel = self._mm_tree.selection()
        tid = None
        if sel:
            try: tid = int(sel[0])
            except ValueError: pass
        if tid is None: tid = self._mm_cur_tid
        if tid is None: return None
        return next((r for r in self._mm_results if r.get("type_id") == tid), None)

    def _mm_open_chart(self):
        r = self._mm_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        rid = REGIONS.get(self._mm_reg.get(), 10000002)
        self._open_candle_chart(r["type_id"], r["item"], rid)

    def _mm_open_fcast(self):
        r = self._mm_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_fcast: messagebox.showinfo("Aguarde", "Previsão em execução."); return
        self._fcast_tid_e.set(str(r["type_id"])); self._fcast_name_e.set(r["item"])
        self._fcast_reg.set(self._mm_reg.get())
        self._nb.select(self._tab_fcast); self.after(50, self._start_fcast)

    def _mm_open_bt(self):
        r = self._mm_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_bt: messagebox.showinfo("Aguarde", "Backtest em execução."); return
        self._bt_tid.set(str(r["type_id"])); self._bt_name.set(r["item"])
        self._bt_reg.set(self._mm_reg.get())
        self._nb.select(self._tab_bt); self.after(50, self._start_bt)

    def _mm_open_ai(self):
        r = self._mm_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_ai: messagebox.showinfo("Aguarde", "AI em execução."); return
        self._ai_tid.set(str(r["type_id"])); self._ai_name_e.set(r["item"])
        self._ai_reg.set(self._mm_reg.get())
        self._nb.select(self._tab_ai); self.after(50, self._start_ai)

    # ── ESCASSEZ ──────────────────────────────────────────────────────────────
    def _sc_on_select(self, _e=None):
        sel = self._sc_tree.selection()
        if sel:
            try: self._sc_cur_tid = int(sel[0])
            except ValueError: pass

    def _sc_on_dbl(self, _e=None):
        w = self._sc_get_selected()
        if not w: return
        self.clipboard_clear(); self.clipboard_append(str(w["type_id"]))
        self._set_status(f"✓ Type ID {w['type_id']} copiado — {w['item']}")

    def _sc_get_selected(self):
        sel = self._sc_tree.selection()
        tid = None
        if sel:
            try: tid = int(sel[0])
            except ValueError: pass
        if tid is None: tid = self._sc_cur_tid
        if tid is None: return None
        return next((r for r in self._sc_results if r.get("type_id") == tid), None)

    def _sc_open_chart(self):
        r = self._sc_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        rid = REGIONS.get(self._sc_reg.get(), 10000002)
        self._open_candle_chart(r["type_id"], r["item"], rid)

    def _sc_open_fcast(self):
        r = self._sc_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_fcast: messagebox.showinfo("Aguarde", "Previsão em execução."); return
        self._fcast_tid_e.set(str(r["type_id"])); self._fcast_name_e.set(r["item"])
        self._fcast_reg.set(self._sc_reg.get())
        self._nb.select(self._tab_fcast); self.after(50, self._start_fcast)

    def _sc_open_bt(self):
        r = self._sc_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_bt: messagebox.showinfo("Aguarde", "Backtest em execução."); return
        self._bt_tid.set(str(r["type_id"])); self._bt_name.set(r["item"])
        self._bt_reg.set(self._sc_reg.get())
        self._nb.select(self._tab_bt); self.after(50, self._start_bt)

    def _sc_open_ai(self):
        r = self._sc_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_ai: messagebox.showinfo("Aguarde", "AI em execução."); return
        self._ai_tid.set(str(r["type_id"])); self._ai_name_e.set(r["item"])
        self._ai_reg.set(self._sc_reg.get())
        self._nb.select(self._tab_ai); self.after(50, self._start_ai)

    # ── DEMANDA/OFERTA ────────────────────────────────────────────────────────
    def _ds_on_select(self, _e=None):
        sel = self._ds_tree.selection()
        if sel:
            try: self._ds_cur_tid = int(sel[0])
            except ValueError: pass

    def _ds_on_dbl(self, _e=None):
        w = self._ds_get_selected()
        if not w: return
        self.clipboard_clear(); self.clipboard_append(str(w["type_id"]))
        self._set_status(f"✓ Type ID {w['type_id']} copiado — {w['item']}")

    def _ds_get_selected(self):
        sel = self._ds_tree.selection()
        tid = None
        if sel:
            try: tid = int(sel[0])
            except ValueError: pass
        if tid is None: tid = self._ds_cur_tid
        if tid is None: return None
        return next((r for r in self._ds_results if r.get("type_id") == tid), None)

    def _ds_open_chart(self):
        r = self._ds_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        rid = REGIONS.get(self._ds_reg.get(), 10000002)
        self._open_candle_chart(r["type_id"], r["item"], rid)

    def _ds_open_fcast(self):
        r = self._ds_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_fcast: messagebox.showinfo("Aguarde", "Previsão em execução."); return
        self._fcast_tid_e.set(str(r["type_id"])); self._fcast_name_e.set(r["item"])
        self._fcast_reg.set(self._ds_reg.get())
        self._nb.select(self._tab_fcast); self.after(50, self._start_fcast)

    def _ds_open_bt(self):
        r = self._ds_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_bt: messagebox.showinfo("Aguarde", "Backtest em execução."); return
        self._bt_tid.set(str(r["type_id"])); self._bt_name.set(r["item"])
        self._bt_reg.set(self._ds_reg.get())
        self._nb.select(self._tab_bt); self.after(50, self._start_bt)

    def _ds_open_ai(self):
        r = self._ds_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_ai: messagebox.showinfo("Aguarde", "AI em execução."); return
        self._ai_tid.set(str(r["type_id"])); self._ai_name_e.set(r["item"])
        self._ai_reg.set(self._ds_reg.get())
        self._nb.select(self._tab_ai); self.after(50, self._start_ai)

    # ── MOV. PREÇOS ───────────────────────────────────────────────────────────
    def _pm_on_select(self, _e=None):
        sel = self._pm_tree.selection()
        if sel:
            try: self._pm_cur_tid = int(sel[0])
            except ValueError: pass

    def _pm_on_dbl(self, _e=None):
        w = self._pm_get_selected()
        if not w: return
        self.clipboard_clear(); self.clipboard_append(str(w["type_id"]))
        self._set_status(f"✓ Type ID {w['type_id']} copiado — {w['item']}")

    def _pm_get_selected(self):
        sel = self._pm_tree.selection()
        tid = None
        if sel:
            try: tid = int(sel[0])
            except ValueError: pass
        if tid is None: tid = self._pm_cur_tid
        if tid is None: return None
        return next((r for r in self._pm_results if r.get("type_id") == tid), None)

    def _pm_open_chart(self):
        r = self._pm_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        rid = REGIONS.get(self._pm_reg.get(), 10000002)
        self._open_candle_chart(r["type_id"], r["item"], rid)

    def _pm_open_fcast(self):
        r = self._pm_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_fcast: messagebox.showinfo("Aguarde", "Previsão em execução."); return
        self._fcast_tid_e.set(str(r["type_id"])); self._fcast_name_e.set(r["item"])
        self._fcast_reg.set(self._pm_reg.get())
        self._nb.select(self._tab_fcast); self.after(50, self._start_fcast)

    def _pm_open_bt(self):
        r = self._pm_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_bt: messagebox.showinfo("Aguarde", "Backtest em execução."); return
        self._bt_tid.set(str(r["type_id"])); self._bt_name.set(r["item"])
        self._bt_reg.set(self._pm_reg.get())
        self._nb.select(self._tab_bt); self.after(50, self._start_bt)

    def _pm_open_ai(self):
        r = self._pm_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_ai: messagebox.showinfo("Aguarde", "AI em execução."); return
        self._ai_tid.set(str(r["type_id"])); self._ai_name_e.set(r["item"])
        self._ai_reg.set(self._pm_reg.get())
        self._nb.select(self._tab_ai); self.after(50, self._start_ai)

    # ── ORÇAMENTO ─────────────────────────────────────────────────────────────
    def _bud_on_select(self, _e=None):
        sel = self._bud_tree.selection()
        if sel:
            try: self._bud_cur_tid = int(sel[0])
            except ValueError: pass

    def _bud_on_dbl(self, _e=None):
        w = self._bud_get_selected()
        if not w: return
        self.clipboard_clear(); self.clipboard_append(str(w["type_id"]))
        self._set_status(f"✓ Type ID {w['type_id']} copiado — {w['item']}")

    def _bud_get_selected(self):
        sel = self._bud_tree.selection()
        tid = None
        if sel:
            try: tid = int(sel[0])
            except ValueError: pass
        if tid is None: tid = self._bud_cur_tid
        if tid is None: return None
        return next((r for r in self._bud_results if r.get("type_id") == tid), None)

    def _bud_open_chart(self):
        r = self._bud_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        rid = REGIONS.get(self._bud_reg.get(), 10000002)
        self._open_candle_chart(r["type_id"], r["item"], rid)

    def _bud_open_fcast(self):
        r = self._bud_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_fcast: messagebox.showinfo("Aguarde", "Previsão em execução."); return
        self._fcast_tid_e.set(str(r["type_id"])); self._fcast_name_e.set(r["item"])
        self._fcast_reg.set(self._bud_reg.get())
        self._nb.select(self._tab_fcast); self.after(50, self._start_fcast)

    def _bud_open_bt(self):
        r = self._bud_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_bt: messagebox.showinfo("Aguarde", "Backtest em execução."); return
        self._bt_tid.set(str(r["type_id"])); self._bt_name.set(r["item"])
        self._bt_reg.set(self._bud_reg.get())
        self._nb.select(self._tab_bt); self.after(50, self._start_bt)

    def _bud_open_ai(self):
        r = self._bud_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_ai: messagebox.showinfo("Aguarde", "AI em execução."); return
        self._ai_tid.set(str(r["type_id"])); self._ai_name_e.set(r["item"])
        self._ai_reg.set(self._bud_reg.get())
        self._nb.select(self._tab_ai); self.after(50, self._start_ai)

    # ── Helpers de seleção/navegação: ROTAS, ARBITRAGEM, MULTI-HOP,
    #    CICLOS, CORRELAÇÃO, OTIMIZADOR, INDÚSTRIA ─────────────────────────────

    # ── ROTAS ─────────────────────────────────────────────────────────────────
    def _r_on_select(self, _e=None):
        sel = self._r_tree.selection()
        if sel:
            try: self._r_cur_tid = int(sel[0])
            except ValueError: pass

    def _r_on_dbl(self, _e=None):
        r = self._r_get_selected()
        if not r: return
        self.clipboard_clear(); self.clipboard_append(str(r["type_id"]))
        self._set_status(f"✓ Type ID {r['type_id']} copiado — {r['name']}")

    def _r_get_selected(self):
        sel = self._r_tree.selection()
        tid = None
        if sel:
            try: tid = int(sel[0])
            except ValueError: pass
        if tid is None: tid = self._r_cur_tid
        if tid is None: return None
        return next((r for r in self._routes if r.get("type_id") == tid), None)

    def _r_open_chart(self):
        r = self._r_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        rid = REGIONS.get(self._r_src.get(), 10000002)
        self._open_candle_chart(r["type_id"], r["name"], rid)

    def _r_open_fcast(self):
        r = self._r_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_fcast: messagebox.showinfo("Aguarde", "Previsão em execução."); return
        self._fcast_tid_e.set(str(r["type_id"])); self._fcast_name_e.set(r["name"])
        self._fcast_reg.set(self._r_src.get())
        self._nb.select(self._tab_fcast); self.after(50, self._start_fcast)

    def _r_open_bt(self):
        r = self._r_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_bt: messagebox.showinfo("Aguarde", "Backtest em execução."); return
        self._bt_tid.set(str(r["type_id"])); self._bt_name.set(r["name"])
        self._bt_reg.set(self._r_src.get())
        self._nb.select(self._tab_bt); self.after(50, self._start_bt)

    def _r_open_ai(self):
        r = self._r_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_ai: messagebox.showinfo("Aguarde", "AI em execução."); return
        self._ai_tid.set(str(r["type_id"])); self._ai_name_e.set(r["name"])
        self._ai_reg.set(self._r_src.get())
        self._nb.select(self._tab_ai); self.after(50, self._start_ai)

    # ── ARBITRAGEM ────────────────────────────────────────────────────────────
    def _a_on_select(self, _e=None):
        sel = self._a_tree.selection()
        if sel:
            try:
                idx = int(sel[0])
                self._a_cur_idx = idx
                # Atualiza _a_cur_tid para uso pelas funções de navegação
                if self._arb_opps and idx < len(self._arb_opps):
                    self._a_cur_tid = self._arb_opps[idx].get("type_id")
            except ValueError:
                pass

    def _a_on_dbl(self, _e=None):
        r = self._a_get_selected()
        if not r: return
        self.clipboard_clear(); self.clipboard_append(str(r["type_id"]))
        self._set_status(f"✓ Type ID {r['type_id']} copiado — {r['item']}")

    def _a_get_selected(self):
        sel = self._a_tree.selection()
        idx = None
        if sel:
            try: idx = int(sel[0])
            except ValueError: pass
        if idx is None: idx = getattr(self, '_a_cur_idx', None)
        if idx is not None and self._arb_opps and idx < len(self._arb_opps):
            return self._arb_opps[idx]
        # Fallback por type_id (compatibilidade)
        tid = self._a_cur_tid
        if tid is None: return None
        return next((r for r in self._arb_opps if r.get("type_id") == tid), None)

    def _a_open_chart(self):
        r = self._a_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        rid = REGIONS.get(self._a_src.get(), 10000002)
        self._open_candle_chart(r["type_id"], r["item"], rid)

    def _a_open_fcast(self):
        r = self._a_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_fcast: messagebox.showinfo("Aguarde", "Previsão em execução."); return
        self._fcast_tid_e.set(str(r["type_id"])); self._fcast_name_e.set(r["item"])
        self._fcast_reg.set(self._a_src.get())
        self._nb.select(self._tab_fcast); self.after(50, self._start_fcast)

    def _a_open_bt(self):
        r = self._a_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_bt: messagebox.showinfo("Aguarde", "Backtest em execução."); return
        self._bt_tid.set(str(r["type_id"])); self._bt_name.set(r["item"])
        self._bt_reg.set(self._a_src.get())
        self._nb.select(self._tab_bt); self.after(50, self._start_bt)

    def _a_open_ai(self):
        r = self._a_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_ai: messagebox.showinfo("Aguarde", "AI em execução."); return
        self._ai_tid.set(str(r["type_id"])); self._ai_name_e.set(r["item"])
        self._ai_reg.set(self._a_src.get())
        self._nb.select(self._tab_ai); self.after(50, self._start_ai)

    # ── MULTI-HOP ─────────────────────────────────────────────────────────────
    def _mh_on_select(self, _e=None):
        sel = self._mh_tree.selection()
        if sel:
            try: self._mh_cur_tid = int(sel[0])
            except ValueError: pass

    def _mh_on_dbl(self, _e=None):
        r = self._mh_get_selected()
        if not r: return
        self.clipboard_clear(); self.clipboard_append(str(r["type_id"]))
        self._set_status(f"✓ Type ID {r['type_id']} copiado — {r['item']}")

    def _mh_get_selected(self):
        sel = self._mh_tree.selection()
        tid = None
        if sel:
            try: tid = int(sel[0])
            except ValueError: pass
        if tid is None: tid = self._mh_cur_tid
        if tid is None: return None
        return next((r for r in self._mh_results if r.get("type_id") == tid), None)

    def _mh_open_chart(self):
        r = self._mh_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        rid = REGIONS.get(self._mh_a.get(), 10000002)
        self._open_candle_chart(r["type_id"], r["item"], rid)

    def _mh_open_fcast(self):
        r = self._mh_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_fcast: messagebox.showinfo("Aguarde", "Previsão em execução."); return
        self._fcast_tid_e.set(str(r["type_id"])); self._fcast_name_e.set(r["item"])
        self._fcast_reg.set(self._mh_a.get())
        self._nb.select(self._tab_fcast); self.after(50, self._start_fcast)

    def _mh_open_bt(self):
        r = self._mh_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_bt: messagebox.showinfo("Aguarde", "Backtest em execução."); return
        self._bt_tid.set(str(r["type_id"])); self._bt_name.set(r["item"])
        self._bt_reg.set(self._mh_a.get())
        self._nb.select(self._tab_bt); self.after(50, self._start_bt)

    def _mh_open_ai(self):
        r = self._mh_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_ai: messagebox.showinfo("Aguarde", "AI em execução."); return
        self._ai_tid.set(str(r["type_id"])); self._ai_name_e.set(r["item"])
        self._ai_reg.set(self._mh_a.get())
        self._nb.select(self._tab_ai); self.after(50, self._start_ai)

    # ── CICLOS ────────────────────────────────────────────────────────────────
    def _cyc_on_select(self, _e=None):
        sel = self._cyc_tree.selection()
        if sel:
            try: self._cyc_cur_tid = int(sel[0])
            except ValueError: pass

    def _cyc_on_dbl(self, _e=None):
        r = self._cyc_get_selected()
        if not r: return
        self.clipboard_clear(); self.clipboard_append(str(r["type_id"]))
        self._set_status(f"✓ Type ID {r['type_id']} copiado — {r['name']}")

    def _cyc_get_selected(self):
        sel = self._cyc_tree.selection()
        tid = None
        if sel:
            try: tid = int(sel[0])
            except ValueError: pass
        if tid is None: tid = self._cyc_cur_tid
        if tid is None: return None
        items = (self._cycle_result or {}).get("items", {})
        return items.get(tid)

    def _cyc_open_chart(self):
        r = self._cyc_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        rid = REGIONS.get(self._cyc_reg.get(), 10000002)
        self._open_candle_chart(r["type_id"], r.get("name", str(r["type_id"])), rid)

    def _cyc_open_fcast(self):
        r = self._cyc_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_fcast: messagebox.showinfo("Aguarde", "Previsão em execução."); return
        self._fcast_tid_e.set(str(r["type_id"]))
        self._fcast_name_e.set(r.get("name", str(r["type_id"])))
        self._fcast_reg.set(self._cyc_reg.get())
        self._nb.select(self._tab_fcast); self.after(50, self._start_fcast)

    def _cyc_open_bt(self):
        r = self._cyc_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_bt: messagebox.showinfo("Aguarde", "Backtest em execução."); return
        self._bt_tid.set(str(r["type_id"]))
        self._bt_name.set(r.get("name", str(r["type_id"])))
        self._bt_reg.set(self._cyc_reg.get())
        self._nb.select(self._tab_bt); self.after(50, self._start_bt)

    def _cyc_open_ai(self):
        r = self._cyc_get_selected()
        if not r: messagebox.showinfo("Selecione", "Clique em um item primeiro."); return
        if self._loading_ai: messagebox.showinfo("Aguarde", "AI em execução."); return
        self._ai_tid.set(str(r["type_id"]))
        self._ai_name_e.set(r.get("name", str(r["type_id"])))
        self._ai_reg.set(self._cyc_reg.get())
        self._nb.select(self._tab_ai); self.after(50, self._start_ai)

    # ── CORRELAÇÃO ────────────────────────────────────────────────────────────
    # Nota: cada linha da correlação representa um PAR de itens, não um único
    # type_id. Para ações, usa o item_a do par selecionado como referência.

    def _corr_open_chart(self):
        """Abre gráfico para o item_a do par selecionado."""
        sel = self._corr_tree.selection()
        idx = int(sel[0]) if sel else getattr(self, '_corr_cur_row', 0)
        if not self._corr_rows or idx >= len(self._corr_rows):
            messagebox.showinfo("Selecione", "Clique em um par primeiro."); return
        row = self._corr_rows[idx]
        name_map = getattr(self, '_corr_name_to_tid', {})
        tid = name_map.get(row["item_a"])
        if not tid: messagebox.showinfo("Sem dados", "Execute CALCULAR primeiro."); return
        rid = getattr(self, '_corr_rid', REGIONS.get(self._corr_reg.get(), 10000002))
        self._open_candle_chart(tid, row["item_a"], rid)

    def _corr_open_fcast(self):
        sel = self._corr_tree.selection()
        idx = int(sel[0]) if sel else getattr(self, '_corr_cur_row', 0)
        if not self._corr_rows or idx >= len(self._corr_rows):
            messagebox.showinfo("Selecione", "Clique em um par primeiro."); return
        row = self._corr_rows[idx]
        name_map = getattr(self, '_corr_name_to_tid', {})
        tid = name_map.get(row["item_a"])
        if not tid: messagebox.showinfo("Sem dados", "Execute CALCULAR primeiro."); return
        if self._loading_fcast: messagebox.showinfo("Aguarde", "Previsão em execução."); return
        self._fcast_tid_e.set(str(tid)); self._fcast_name_e.set(row["item_a"])
        self._fcast_reg.set(self._corr_reg.get())
        self._nb.select(self._tab_fcast); self.after(50, self._start_fcast)

    def _corr_open_bt(self):
        sel = self._corr_tree.selection()
        idx = int(sel[0]) if sel else getattr(self, '_corr_cur_row', 0)
        if not self._corr_rows or idx >= len(self._corr_rows):
            messagebox.showinfo("Selecione", "Clique em um par primeiro."); return
        row = self._corr_rows[idx]
        name_map = getattr(self, '_corr_name_to_tid', {})
        tid = name_map.get(row["item_a"])
        if not tid: messagebox.showinfo("Sem dados", "Execute CALCULAR primeiro."); return
        if self._loading_bt: messagebox.showinfo("Aguarde", "Backtest em execução."); return
        self._bt_tid.set(str(tid)); self._bt_name.set(row["item_a"])
        self._bt_reg.set(self._corr_reg.get())
        self._nb.select(self._tab_bt); self.after(50, self._start_bt)

    def _corr_open_ai(self):
        sel = self._corr_tree.selection()
        idx = int(sel[0]) if sel else getattr(self, '_corr_cur_row', 0)
        if not self._corr_rows or idx >= len(self._corr_rows):
            messagebox.showinfo("Selecione", "Clique em um par primeiro."); return
        row = self._corr_rows[idx]
        name_map = getattr(self, '_corr_name_to_tid', {})
        tid = name_map.get(row["item_a"])
        if not tid: messagebox.showinfo("Sem dados", "Execute CALCULAR primeiro."); return
        if self._loading_ai: messagebox.showinfo("Aguarde", "AI em execução."); return
        self._ai_tid.set(str(tid)); self._ai_name_e.set(row["item_a"])
        self._ai_reg.set(self._corr_reg.get())
        self._nb.select(self._tab_ai); self.after(50, self._start_ai)

    # ── OTIMIZADOR ────────────────────────────────────────────────────────────
    # Nota: _opt_plan usa "Item" (nome) como chave principal, sem type_id direto.
    # Resolve o type_id via item_cache antes de navegar.

    def _opt_on_select(self, _e=None):
        sel = self._o_tree.selection()
        if sel:
            try: self._opt_cur_idx = int(sel[0])
            except ValueError: pass

    def _opt_on_dbl(self, _e=None):
        idx = self._opt_cur_idx
        if not self._opt_plan or idx >= len(self._opt_plan): return
        r = self._opt_plan[idx]
        name = r.get("Item", "")
        # Resolve type_id pelo nome
        tid = next((tid for tid, v in self._item_cache.items()
                    if v.get("name","") == name), None)
        if tid:
            self.clipboard_clear(); self.clipboard_append(str(tid))
            self._set_status(f"✓ Type ID {tid} copiado — {name}")

    def _opt_get_selected_tid_name(self):
        """Retorna (tid, name) do item selecionado no Otimizador, ou (None, None)."""
        idx = self._opt_cur_idx
        if not self._opt_plan or idx >= len(self._opt_plan):
            return None, None
        name = self._opt_plan[idx].get("Item", "")
        tid  = next((tid for tid, v in self._item_cache.items()
                     if v.get("name","") == name), None)
        return tid, name

    def _opt_open_chart(self):
        tid, name = self._opt_get_selected_tid_name()
        if not tid: messagebox.showinfo("Selecione", "Clique num item e aguarde a resolução do Type ID."); return
        # Usa a região de origem do plano
        row = self._opt_plan[self._opt_cur_idx]
        rid = REGIONS.get(row.get("Origem", ""), 10000002)
        self._open_candle_chart(tid, name, rid)

    def _opt_open_fcast(self):
        tid, name = self._opt_get_selected_tid_name()
        if not tid: messagebox.showinfo("Selecione", "Clique num item primeiro."); return
        if self._loading_fcast: messagebox.showinfo("Aguarde", "Previsão em execução."); return
        row = self._opt_plan[self._opt_cur_idx]
        self._fcast_tid_e.set(str(tid)); self._fcast_name_e.set(name)
        self._fcast_reg.set(row.get("Origem", self._fcast_reg.get()))
        self._nb.select(self._tab_fcast); self.after(50, self._start_fcast)

    def _opt_open_bt(self):
        tid, name = self._opt_get_selected_tid_name()
        if not tid: messagebox.showinfo("Selecione", "Clique num item primeiro."); return
        if self._loading_bt: messagebox.showinfo("Aguarde", "Backtest em execução."); return
        row = self._opt_plan[self._opt_cur_idx]
        self._bt_tid.set(str(tid)); self._bt_name.set(name)
        self._bt_reg.set(row.get("Origem", self._bt_reg.get()))
        self._nb.select(self._tab_bt); self.after(50, self._start_bt)

    def _opt_open_ai(self):
        tid, name = self._opt_get_selected_tid_name()
        if not tid: messagebox.showinfo("Selecione", "Clique num item primeiro."); return
        if self._loading_ai: messagebox.showinfo("Aguarde", "AI em execução."); return
        row = self._opt_plan[self._opt_cur_idx]
        self._ai_tid.set(str(tid)); self._ai_name_e.set(name)
        self._ai_reg.set(row.get("Origem", self._ai_reg.get()))
        self._nb.select(self._tab_ai); self.after(50, self._start_ai)

    # ── INDÚSTRIA ─────────────────────────────────────────────────────────────
    # Usa o produto fabricado (self._ind_prod_tid + self._ind_prod_name)
    # e a região configurada em CONFIGURAÇÕES (_g_reg).

    def _ind_open_chart(self):
        try:
            tid = int(self._ind_prod_tid.get().strip())
        except (ValueError, AttributeError):
            messagebox.showinfo("Sem produto", "Configure um produto primeiro."); return
        name = self._ind_prod_name.get().strip() or str(tid)
        rid  = REGIONS.get(self._g_reg.get(), 10000002) if hasattr(self, '_g_reg') else 10000002
        self._open_candle_chart(tid, name, rid)

    def _ind_open_fcast(self):
        try:
            tid = int(self._ind_prod_tid.get().strip())
        except (ValueError, AttributeError):
            messagebox.showinfo("Sem produto", "Configure um produto primeiro."); return
        if self._loading_fcast: messagebox.showinfo("Aguarde", "Previsão em execução."); return
        name = self._ind_prod_name.get().strip() or str(tid)
        self._fcast_tid_e.set(str(tid)); self._fcast_name_e.set(name)
        self._nb.select(self._tab_fcast); self.after(50, self._start_fcast)

    def _ind_open_bt(self):
        try:
            tid = int(self._ind_prod_tid.get().strip())
        except (ValueError, AttributeError):
            messagebox.showinfo("Sem produto", "Configure um produto primeiro."); return
        if self._loading_bt: messagebox.showinfo("Aguarde", "Backtest em execução."); return
        name = self._ind_prod_name.get().strip() or str(tid)
        self._bt_tid.set(str(tid)); self._bt_name.set(name)
        self._nb.select(self._tab_bt); self.after(50, self._start_bt)

    def _ind_open_ai(self):
        try:
            tid = int(self._ind_prod_tid.get().strip())
        except (ValueError, AttributeError):
            messagebox.showinfo("Sem produto", "Configure um produto primeiro."); return
        if self._loading_ai: messagebox.showinfo("Aguarde", "AI em execução."); return
        name = self._ind_prod_name.get().strip() or str(tid)
        self._ai_tid.set(str(tid)); self._ai_name_e.set(name)
        self._nb.select(self._tab_ai); self.after(50, self._start_ai)

    def _build_tab_catalog(self):
        p = self._tab_catalog

        # ── Barra de busca e filtros ──────────────────────────────────────────
        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(6, 2))

        tk.Label(fr, text="Buscar nome:",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 10)).grid(row=0, column=0, padx=6, sticky="w")
        self._cat_q = tk.Entry(fr, width=32,
                                bg=COLORS["border"], fg=COLORS["text"],
                                insertbackground=COLORS["text"],
                                font=("Courier", 12), relief="flat", bd=2)
        self._cat_q.grid(row=1, column=0, padx=6, pady=(0, 6), sticky="w")
        self._cat_q.bind("<Return>", lambda e: self._cat_apply_filter())
        self._cat_q.bind("<KeyRelease>", self._cat_on_keyrelease)

        tk.Label(fr, text="Filtrar grupo:",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 10)).grid(row=0, column=1, padx=6, sticky="w")
        self._cat_grp = tk.Entry(fr, width=24,
                                  bg=COLORS["border"], fg=COLORS["text"],
                                  insertbackground=COLORS["text"],
                                  font=("Courier", 12), relief="flat", bd=2)
        self._cat_grp.grid(row=1, column=1, padx=6, pady=(0, 6), sticky="w")
        self._cat_grp.bind("<Return>", lambda e: self._cat_apply_filter())

        tk.Label(fr, text="Filtrar categoria:",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 10)).grid(row=0, column=2, padx=6, sticky="w")
        self._cat_cat = tk.Entry(fr, width=24,
                                  bg=COLORS["border"], fg=COLORS["text"],
                                  insertbackground=COLORS["text"],
                                  font=("Courier", 12), relief="flat", bd=2)
        self._cat_cat.grid(row=1, column=2, padx=6, pady=(0, 6), sticky="w")
        self._cat_cat.bind("<Return>", lambda e: self._cat_apply_filter())

        tk.Button(fr, text="  BUSCAR  ",
                  command=self._cat_apply_filter,
                  bg=COLORS["accent"], fg="#000",
                  font=("Courier", 11, "bold"), relief="flat", cursor="hand2"
                  ).grid(row=1, column=3, padx=8, pady=(0, 6))

        tk.Button(fr, text=" LIMPAR ",
                  command=self._cat_clear,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=4, padx=4, pady=(0, 6))

        tk.Button(fr, text=" ENRIQUECER CACHE ",
                  command=self._start_cat_enrich,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=5, padx=4, pady=(0, 6))

        tk.Button(fr, text=" COPIAR TYPE ID ",
                  command=self._cat_copy_tid,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=6, padx=4, pady=(0, 6))

        tk.Button(fr, text=" USAR NAS ABAS ",
                  command=self._cat_use_in_tabs,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=7, padx=4, pady=(0, 6))

        self._cat_enrich_spin = ttk.Progressbar(fr, mode="determinate",
                                                  length=100, maximum=100)
        self._cat_enrich_spin.grid(row=1, column=8, padx=6, pady=(0, 6))
        self._cat_enrich_pct = tk.Label(fr, text="", bg=COLORS["panel"],
                                         fg=COLORS["subtext"],
                                         font=("Courier", 10), width=5)
        self._cat_enrich_pct.grid(row=1, column=9, padx=2, pady=(0, 6))

        tk.Label(fr,
                 text="Enter ou KeyRelease aplica filtro automaticamente  |  "
                      "ENRIQUECER baixa grupo/categoria dos itens no cache via ESI",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 9)
                 ).grid(row=0, column=3, columnspan=7, padx=4, sticky="w")

        # ── Tabela principal ──────────────────────────────────────────────────
        self._cat_tree = self._make_tree(p, [
            ("type_id",       "Type ID",    90),
            ("name",          "Nome",      260),
            ("group_name",    "Grupo",     180),
            ("category_name", "Categoria", 160),
            ("volume",        "Vol (m³)",   85),
        ])
        for cid in self._CAT_COLS:
            self._cat_tree.heading(cid, text=self._CAT_COLS[cid],
                                    command=lambda c=cid: self._sort_cat(c))
        self._cat_tree.bind("<Double-1>", self._cat_on_dbl)
        self._cat_tree.bind("<Button-3>", self._cat_context_menu)

        # Barra de rodapé
        bar = tk.Frame(p, bg=COLORS["bg"])
        bar.pack(fill="x", padx=6, pady=(0, 4))
        self._cat_count = tk.Label(bar, text="0 itens",
                                    fg=COLORS["subtext"], bg=COLORS["bg"],
                                    font=("Courier", 10))
        self._cat_count.pack(side="left")
        self._cat_sel_lbl = tk.Label(bar, text="",
                                      fg=COLORS["accent"], bg=COLORS["bg"],
                                      font=("Courier", 10))
        self._cat_sel_lbl.pack(side="left", padx=16)
        tk.Label(bar,
                 text="Double-click = copiar Type ID  |  "
                      "Botão direito = menu  |  "
                      "USAR NAS ABAS = preenche Previsao, Backtest, Spread Hist., AI",
                 fg=COLORS["subtext"], bg=COLORS["bg"],
                 font=("Courier", 10)).pack(side="right")

        # Popula com o cache existente
        self.after(200, self._cat_apply_filter)

    # ── Handlers de busca e filtro ────────────────────────────────────────────

    _cat_after_id = None

    def _cat_on_keyrelease(self, _):
        """Debounce de 400ms antes de aplicar o filtro (evita travar enquanto digita)."""
        if self._cat_after_id:
            try: self.after_cancel(self._cat_after_id)
            except Exception: pass
        self._cat_after_id = self.after(400, self._cat_apply_filter)

    def _cat_apply_filter(self):
        """Aplica busca/filtro e repopula a tabela."""
        query = self._cat_q.get()
        grp   = self._cat_grp.get()
        cat   = self._cat_cat.get()
        rows  = search_item_catalog(self._item_cache, query, grp, cat)
        self._cat_rows = rows
        self._pop_cat()

    def _cat_clear(self):
        self._cat_q.delete(0, "end")
        self._cat_grp.delete(0, "end")
        self._cat_cat.delete(0, "end")
        self._cat_apply_filter()

    def _pop_cat(self, reorder_only=False):
        if reorder_only:
            for i in range(len(self._cat_rows)):
                self._cat_tree.move(str(i), "", i)
            return
        for row in self._cat_tree.get_children():
            self._cat_tree.delete(row)
        for i, r in enumerate(self._cat_rows):
            has_group = bool(r.get("group_name"))
            tag = ("even" if i % 2 == 0 else "odd",)
            if not has_group:
                tag = ("dim",)
            self._cat_tree.insert("", "end", iid=str(i), tags=tag, values=(
                r["type_id"],
                r["name"],
                r["group_name"]    or "(sem grupo — clique ENRIQUECER)",
                r["category_name"] or "",
                f"{r['volume']:.2f}" if r["volume"] else "--",
            ))
        self._cat_tree.tag_configure("dim", foreground=COLORS["subtext"])
        n_sem = sum(1 for r in self._cat_rows if not r.get("group_name"))
        self._cat_count.config(
            text=(f"{len(self._cat_rows):,} itens  "
                  f"({n_sem:,} sem grupo — clique ENRIQUECER CACHE)"))

    def _sort_cat(self, col):
        num = {"type_id", "volume"}
        rev = not self._cat_sort_state.get(col, col == "type_id")
        self._cat_sort_state = {col: rev}
        self._cat_rows.sort(
            key=lambda r: (r.get(col) or 0) if col in num
            else (r.get(col) or "").lower(),
            reverse=rev)
        self._set_heading_arrow(self._cat_tree, self._CAT_COLS, col, rev)
        self._pop_cat(reorder_only=True)

    # ── Enriquecimento de cache ───────────────────────────────────────────────

    def _start_cat_enrich(self):
        """Baixa grupo/categoria dos itens que ainda nao tem via ESI."""
        if self._loading_cat:
            return
        # Determina quais enricher: itens sem group_id no cache
        sem_grupo = [tid for tid, info in self._item_cache.items()
                     if not info.get("group_id")]
        if not sem_grupo:
            messagebox.showinfo("Cache completo",
                f"Todos os {len(self._item_cache):,} itens já têm grupo/categoria.")
            return
        self._loading_cat = True
        self._cat_enrich_spin.config(value=0)
        self._cat_enrich_pct.config(text="")
        self._set_status(f"Enriquecendo {len(sem_grupo):,} itens sem grupo...")
        threading.Thread(target=self._worker_cat_enrich,
                         args=(sem_grupo,), daemon=True).start()

    def _worker_cat_enrich(self, tids: list):
        prog = self._make_prog(self._cat_enrich_spin, self._cat_enrich_pct)
        try:
            n = enrich_item_cache_groups(self._item_cache, tids, prog)
            self.after(0, lambda: (
                self._cat_enrich_spin.config(value=100),
                self._cat_enrich_pct.config(text="100%"),
                self._set_status(f"Catalogo enriquecido: {n:,} itens atualizados."),
                self._cat_apply_filter()))
        except Exception as exc:
            log.exception("worker_cat_enrich")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro", e))
        finally:
            self._loading_cat = False

    # ── Ações na seleção ─────────────────────────────────────────────────────

    def _cat_selected_row(self):
        """Retorna o dict da linha selecionada ou None."""
        sel = self._cat_tree.selection()
        if not sel:
            return None
        try:
            idx = int(sel[0])
            return self._cat_rows[idx]
        except (ValueError, IndexError):
            return None

    def _cat_copy_tid(self):
        """Copia o Type ID do item selecionado para a área de transferência."""
        row = self._cat_selected_row()
        if not row:
            messagebox.showinfo("Selecione", "Clique num item."); return
        tid = str(row["type_id"])
        self.clipboard_clear()
        self.clipboard_append(tid)
        self._cat_sel_lbl.config(
            text=f"✓ Type ID {tid} copiado  ({row['name']})")

    def _cat_on_dbl(self, _):
        """Double-click: copia Type ID e mostra confirmação."""
        self._cat_copy_tid()

    def _cat_context_menu(self, event):
        """Menu de contexto com botão direito."""
        row = self._cat_selected_row()

        # Seleciona o item sob o cursor
        iid = self._cat_tree.identify_row(event.y)
        if iid:
            self._cat_tree.selection_set(iid)
            try:
                row = self._cat_rows[int(iid)]
            except (ValueError, IndexError):
                row = None

        menu = tk.Menu(self, tearoff=0,
                       bg=COLORS["panel"], fg=COLORS["text"],
                       activebackground=COLORS["border"],
                       activeforeground=COLORS["accent"],
                       font=("Courier", 11))
        if row:
            tid  = row["type_id"]
            name = row["name"]
            menu.add_command(
                label=f"  Copiar Type ID ({tid})",
                command=self._cat_copy_tid)
            menu.add_command(
                label=f"  Copiar Nome",
                command=lambda: (self.clipboard_clear(),
                                 self.clipboard_append(name),
                                 self._cat_sel_lbl.config(text=f"✓ Nome copiado: {name}")))
            menu.add_separator()
            menu.add_command(
                label="  Abrir em Previsao",
                command=lambda r=row: self._cat_jump("fcast", r))
            menu.add_command(
                label="  Abrir em Backtesting",
                command=lambda r=row: self._cat_jump("bt", r))
            menu.add_command(
                label="  Abrir em Spread Hist.",
                command=lambda r=row: self._cat_jump("sph", r))
            menu.add_command(
                label="  Abrir em AI Analista",
                command=lambda r=row: self._cat_jump("ai", r))
            menu.add_command(
                label="  Abrir em Market Making (Simulador)",
                command=lambda r=row: self._cat_jump("mm2_sim", r))
            menu.add_command(
                label="  Abrir em Station Trader",
                command=lambda r=row: self._cat_jump("st", r))
            menu.add_command(
                label="  Abrir em Indústria (calcular custo)",
                command=lambda r=row: self._cat_jump("ind", r))
            menu.add_command(
                label="  Adicionar a Watchlist",
                command=lambda r=row: self._cat_jump("wl", r))
        else:
            menu.add_command(label="  (nenhum item selecionado)", state="disabled")

        menu.tk_popup(event.x_root, event.y_root)

    def _cat_use_in_tabs(self):
        """Preenche o Type ID e Nome do item selecionado em todas as abas relevantes."""
        row = self._cat_selected_row()
        if not row:
            messagebox.showinfo("Selecione", "Clique num item."); return
        tid  = row["type_id"]
        name = row["name"]

        # Previsao
        self._fcast_tid_e.set(str(tid))
        self._fcast_name_e.set(name)
        # Backtesting
        self._bt_tid.set(str(tid))
        self._bt_name.set(name)
        # Spread Hist.
        self._sph_tid.set(str(tid))
        self._sph_name.set(name)
        # AI Analista
        self._ai_tid.set(str(tid))
        self._ai_name_e.set(name)
        # Market Making Simulador
        self._mm2_sim_tid.set(str(tid))
        self._mm2_sim_name.set(name)
        # Industria (produto)
        self._ind_prod_tid.set(str(tid))
        self._ind_prod_name.set(name)
        # Hub comparador
        self._hub_tid.set(str(tid))
        self._hub_name.set(name)
        # Station Trader — mostra referência na status bar (aba ST não tem campo tid fixo)
        self._set_status(f"Station Trader — item referência: {name} (ID {tid})")

        # Copia para clipboard também
        self.clipboard_clear()
        self.clipboard_append(str(tid))
        self._cat_sel_lbl.config(
            text=f"✓ {name} (ID {tid}) → preenchido em 8 abas  |  Type ID copiado")

    def _cat_jump(self, dest: str, row: dict):
        """Preenche uma aba especifica e navega ate ela."""
        tid  = row["type_id"]
        name = row["name"]
        if dest == "fcast":
            self._fcast_tid_e.set(str(tid))
            self._fcast_name_e.set(name)
            self._nb.select(self._tab_fcast)
        elif dest == "bt":
            self._bt_tid.set(str(tid))
            self._bt_name.set(name)
            self._nb.select(self._tab_bt)
        elif dest == "sph":
            self._sph_tid.set(str(tid))
            self._sph_name.set(name)
            self._nb.select(self._tab_sph)
            self.after(50, self._start_sph)
        elif dest == "ai":
            self._ai_tid.set(str(tid))
            self._ai_name_e.set(name)
            self._nb.select(self._tab_ai)
        elif dest == "mm2_sim":
            self._mm2_sim_tid.set(str(tid))
            self._mm2_sim_name.set(name)
            self._nb.select(self._tab_mm2)
            self._mm2_nb.select(1)  # sub-aba Simulador
        elif dest == "st":
            # Station Trader — preenche filtro de busca rapida (nao ha campo tid)
            # Navega para a aba e coloca o nome no status para referência
            self._nb.select(self._tab_st)
            self._set_status(f"Station Trader — item referência: {name} (ID {tid})")
        elif dest == "ind":
            self._ind_prod_tid.set(str(tid))
            self._ind_prod_name.set(name)
            self._nb.select(self._tab_ind)
            # Dispara auto-preenchimento automaticamente, como as outras abas
            self.after(80, self._ind_auto_load)
        elif dest == "wl":
            if any(w["type_id"] == tid for w in self._watchlist):
                messagebox.showinfo("Já existe", f"{name} já está na Watchlist.")
                return
            self._watchlist.append({
                "type_id": tid, "name": name,
                "alert_below": None, "alert_above": None,
                "last_sell": None, "last_buy": None,
                "last_net": None, "last_score": None,
            })
            save_watchlist(self._watchlist)
            self._pop_wl()
            messagebox.showinfo("Adicionado", f"{name} adicionado à Watchlist.")

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA 4 — DADOS & CACHE (original)
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_dados(self):
        p = self._tab_dados

        sec1 = tk.LabelFrame(p, text=" Download de Ordens por Regiao ",
                             bg=COLORS["panel"], fg=COLORS["header"],
                             font=("Courier", 11, "bold"),
                             highlightbackground=COLORS["border"], highlightthickness=1)
        sec1.pack(fill="x", padx=10, pady=8)

        tk.Label(sec1, text=f"CSVs em: {DATA_DIR}",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 10)).pack(anchor="w", padx=8, pady=(4, 0))

        leg = tk.Frame(sec1, bg=COLORS["panel"])
        leg.pack(anchor="w", padx=8, pady=(2, 0))
        for cor, txt in [(COLORS["green"],"< 1h"),
                         (COLORS["yellow"],"< 24h"),
                         (COLORS["red"],"≥ 24h")]:
            tk.Label(leg, text="●", fg=cor, bg=COLORS["panel"],
                     font=("Courier", 11)).pack(side="left")
            tk.Label(leg, text=f" {txt}   ", fg=COLORS["subtext"],
                     bg=COLORS["panel"], font=("Courier", 10)).pack(side="left")

        chk_frame = tk.Frame(sec1, bg=COLORS["panel"])
        chk_frame.pack(fill="x", padx=8, pady=4)

        self._region_chk_vars = {}
        self._chk_labels      = {}
        for i, rname in enumerate(REGIONS):
            label_txt, color = cache_file_info(rname)
            var  = tk.BooleanVar(value=True)
            cell = tk.Frame(chk_frame, bg=COLORS["panel"])
            cell.grid(row=i//4, column=i%4, sticky="w", padx=10, pady=2)
            tk.Checkbutton(cell, text=rname, variable=var,
                           bg=COLORS["panel"], fg=COLORS["text"],
                           selectcolor=COLORS["border"],
                           activebackground=COLORS["panel"],
                           activeforeground=COLORS["accent"],
                           font=("Courier", 11)).pack(side="left")
            ts_lbl = tk.Label(cell, text=f"[{label_txt}]", fg=color,
                              bg=COLORS["panel"], font=("Courier", 10))
            ts_lbl.pack(side="left", padx=(2, 0))
            self._chk_labels[rname]      = ts_lbl
            self._region_chk_vars[rname] = var

        btn_row = tk.Frame(sec1, bg=COLORS["panel"])
        btn_row.pack(fill="x", padx=8, pady=8)
        tk.Button(btn_row, text="  BAIXAR ORDENS SELECIONADAS  ",
                  command=self._start_download,
                  bg=COLORS["accent"], fg="#000",
                  font=("Courier", 11, "bold"), relief="flat", cursor="hand2"
                  ).pack(side="left")
        self._prog_d = ttk.Progressbar(btn_row, mode="indeterminate", length=140)
        self._prog_d.pack(side="left", padx=12)

        sec2 = tk.LabelFrame(p, text=" Cache de Estacoes / Estruturas ",
                             bg=COLORS["panel"], fg=COLORS["header"],
                             font=("Courier", 11, "bold"),
                             highlightbackground=COLORS["border"], highlightthickness=1)
        sec2.pack(fill="x", padx=10, pady=4)
        tk.Label(sec2, text=f"Resolve nomes NPC -> {STATION_CACHE_FILE}",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 10)).pack(anchor="w", padx=8, pady=4)
        btn_row2 = tk.Frame(sec2, bg=COLORS["panel"])
        btn_row2.pack(fill="x", padx=8, pady=6)
        tk.Button(btn_row2, text="  GERAR CACHE DE ESTACOES  ",
                  command=self._start_station_cache,
                  bg=COLORS["green"], fg="#000",
                  font=("Courier", 11, "bold"), relief="flat", cursor="hand2"
                  ).pack(side="left")
        self._cache_lbl = tk.Label(
            btn_row2,
            text=f"{len(self._station_cache)} estacoes | {len(self._item_cache)} itens",
            bg=COLORS["panel"], fg=COLORS["subtext"], font=("Courier", 11))
        self._cache_lbl.pack(side="left", padx=16)

        log_frame = tk.Frame(p, bg=COLORS["bg"])
        log_frame.pack(fill="both", expand=True, padx=10, pady=(4, 8))
        tk.Label(log_frame, text="Log:", bg=COLORS["bg"],
                 fg=COLORS["subtext"], font=("Courier", 10)).pack(anchor="w")
        self._log_w = tk.Text(log_frame, bg=COLORS["panel"], fg=COLORS["text"],
                              font=("Courier", 10), height=10, relief="flat",
                              state="disabled", wrap="word")
        sb = ttk.Scrollbar(log_frame, command=self._log_w.yview)
        self._log_w.configure(yscrollcommand=sb.set)
        self._log_w.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def _start_download(self):
        if getattr(self, '_downloading', False): return
        selected = [r for r, v in self._region_chk_vars.items() if v.get()]
        if not selected:
            messagebox.showinfo("Aviso", "Selecione ao menos uma regiao."); return
        self._downloading = True
        self._prog_d.start(10)
        self._log_clear()
        threading.Thread(target=self._worker_download, args=(selected,),
                         daemon=True).start()

    def _worker_download(self, regions):
        def cb(msg):
            self._log_append(msg)
            self._set_status(msg)
        try:
            t0 = time.time()
            cb(f"Iniciando download de {len(regions)} regiao(oes) em paralelo...")

            def _dl(rname):
                orders = download_region_orders(REGIONS[rname], rname, cb)
                # M5: invalida cache em memoria para forcar releitura do CSV fresco
                _orders_mem_cache.pop(rname, None)
                return orders

            with concurrent.futures.ThreadPoolExecutor(max_workers=len(regions)) as ex:
                list(ex.map(_dl, regions))

            elapsed = time.time() - t0
            cb(f"Concluido em {elapsed:.1f}s")
            self.after(0, lambda: self._set_status(f"Download concluido em {elapsed:.1f}s"))
        except Exception as exc:
            log.exception("worker_download")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro Download", e))
        finally:
            self._downloading = False
            self.after(0, self._prog_d.stop)
            self.after(0, self._refresh_labels)

    def _refresh_labels(self):
        for rname, lbl in self._chk_labels.items():
            label_txt, color = cache_file_info(rname)
            lbl.config(text=f"[{label_txt}]", fg=color)
        for upd in getattr(self, "_csv_age_updaters", []):
            upd()

    def _start_station_cache(self):
        if getattr(self, '_station_caching', False):
            messagebox.showinfo("Aguarde", "Cache de estações já está sendo atualizado."); return
        self._station_caching = True
        self._log_clear()
        threading.Thread(target=self._worker_station_cache, daemon=True).start()

    def _worker_station_cache(self):
        def cb(msg):
            self._log_append(msg)
            self._set_status(msg)
        try:
            cache = build_station_cache(list(REGIONS.keys()), progress_cb=cb)
            self._station_cache = cache
            n, ni = len(cache), len(self._item_cache)
            self.after(0, lambda: self._cache_lbl.config(
                text=f"{n} estacoes | {ni} itens"))
        except Exception as exc:
            log.exception("worker_station_cache")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro Cache", e))
        finally:
            self._station_caching = False

    def _log_append(self, msg):
        def _do():
            self._log_w.config(state="normal")
            self._log_w.insert("end", str(msg) + "\n")
            self._log_w.see("end")
            self._log_w.config(state="disabled")
        self.after(0, _do)

    def _log_clear(self):
        self._log_w.config(state="normal")
        self._log_w.delete("1.0", "end")
        self._log_w.config(state="disabled")

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA DASHBOARD (Tier 2) — resumo operacional em tempo real
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_dash(self):
        p = self._tab_dash

        hdr = tk.Frame(p, bg=COLORS["bg"])
        hdr.pack(fill="x", padx=16, pady=(10, 4))
        tk.Label(hdr, text="DASHBOARD  v10",
                 bg=COLORS["bg"], fg=COLORS["accent"],
                 font=("Courier", 16, "bold")).pack(side="left")
        ts_lbl = tk.Label(hdr, text="", bg=COLORS["bg"],
                          fg=COLORS["subtext"], font=("Courier", 10))
        ts_lbl.pack(side="right")
        self._dash_ts_lbl = ts_lbl
        tk.Button(hdr, text=" ATUALIZAR AGORA ",
                  command=self._dash_refresh,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).pack(side="right", padx=8)

        grid = tk.Frame(p, bg=COLORS["bg"])
        grid.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        for col in range(4):
            grid.columnconfigure(col, weight=1, uniform="dash")
        for row in range(3):
            grid.rowconfigure(row, weight=1, uniform="dashrow")

        def _card(row, col, title, nav_cmd=None, colspan=1):
            fr = tk.LabelFrame(grid, text=f"  {title}  ",
                               bg=COLORS["panel"], fg=COLORS["header"],
                               font=("Courier", 10, "bold"),
                               highlightbackground=COLORS["border"],
                               highlightthickness=1)
            fr.grid(row=row, column=col, columnspan=colspan,
                    sticky="nsew", padx=4, pady=4)
            if nav_cmd:
                tk.Button(fr, text="►", command=nav_cmd,
                          bg=COLORS["panel"], fg=COLORS["accent"],
                          font=("Courier", 9), relief="flat", cursor="hand2",
                          bd=0).place(relx=1.0, rely=0, anchor="ne",
                                      x=-2, y=2)
            return fr

        def _lbl(parent, **kw):
            lbl = tk.Label(parent, bg=COLORS["panel"], justify="left",
                           font=("Courier", 10), **kw)
            lbl.pack(anchor="w", padx=6, pady=3, fill="x")
            return lbl

        def _nav(tab):
            return lambda: self._nb.select(tab)

        # ── Linha 0 ──────────────────────────────────────────────────────────
        c_paper  = _card(0, 0, "Paper Trading",    _nav(self._tab_paper))
        c_wl     = _card(0, 1, "Watchlist",         _nav(self._tab_wl))
        c_db     = _card(0, 2, "Banco Historico",   _nav(self._tab_pm))
        c_skills = _card(0, 3, "Skills Globais",    _nav(self._tab_cfg))

        # ── Linha 1 ──────────────────────────────────────────────────────────
        c_mm     = _card(1, 0, "Top Max Margem",    _nav(self._tab_mm))
        c_arb    = _card(1, 1, "Top Arbitragem",    _nav(self._tab_arb))
        c_bud    = _card(1, 2, "Orcamento",         _nav(self._tab_bud))
        c_hub    = _card(1, 3, "Ultimo Comparador", _nav(self._tab_hub))

        # ── Linha 2 ──────────────────────────────────────────────────────────
        c_alert  = _card(2, 0, "Alertas Recentes",  _nav(self._tab_wl))
        c_mh     = _card(2, 1, "Station Trader",    _nav(self._tab_st))
        c_cache  = _card(2, 2, "Industria",         _nav(self._tab_ind))
        c_pm     = _card(2, 3, "Ciclos de Mercado", _nav(self._tab_cycle))

        self._dash_paper_lbl  = _lbl(c_paper,  fg=COLORS["green"],   text="...")
        self._dash_wl_lbl     = _lbl(c_wl,     fg=COLORS["text"],    text="...")
        self._dash_db_lbl     = _lbl(c_db,     fg=COLORS["accent"],  text="...")
        self._dash_skills_lbl = _lbl(c_skills, fg=COLORS["accent"],  text="...")
        self._dash_mm_lbl     = _lbl(c_mm,     fg=COLORS["subtext"], text="(calcule Max Margem)")
        self._dash_arb_lbl    = _lbl(c_arb,    fg=COLORS["subtext"], text="(calcule Arbitragem)")
        self._dash_bud_lbl    = _lbl(c_bud,    fg=COLORS["subtext"], text="(calcule Orcamento)")
        self._dash_hub_lbl    = _lbl(c_hub,    fg=COLORS["subtext"], text="(use Comparador)")
        self._dash_alert_lbl  = _lbl(c_alert,  fg=COLORS["yellow"],  text="Nenhum alerta.")
        self._dash_mh_lbl     = _lbl(c_mh,     fg=COLORS["subtext"], text="(escanear Station Trader)")
        self._dash_cache_lbl  = _lbl(c_cache,  fg=COLORS["subtext"], text="(calcule Industria)")
        self._dash_pm_lbl     = _lbl(c_pm,     fg=COLORS["subtext"], text="(calcule Ciclos)")

        self.after(300, self._dash_refresh)

    def _dash_refresh(self):
        """Atualiza todos os widgets do Dashboard com dados em memoria."""
        now_ts = datetime.now().strftime("%d/%m  %H:%M:%S")
        self._dash_ts_lbl.config(text=f"atualizado {now_ts}")

        # ── Paper Trading ─────────────────────────────────────────────────────
        bal   = self._paper.get("balance", 1e9)
        hist  = self._paper.get("history", [])
        pos   = self._paper.get("positions", [])
        pnls  = [h.get("pnl", 0) or 0 for h in hist]
        total_pnl = sum(pnls)
        wins  = sum(1 for p in pnls if p >= 0)
        wr    = wins / len(pnls) * 100 if pnls else 0
        self._dash_paper_lbl.config(
            fg=COLORS["green"] if total_pnl >= 0 else COLORS["red"],
            text=(f"Saldo:    {bal:>16,.0f} ISK\n"
                  f"P&L:      {total_pnl:>+16,.0f} ISK\n"
                  f"Trades: {len(hist):>3}  Win: {wr:.0f}%  Pos: {len(pos)}"))

        # ── Watchlist ─────────────────────────────────────────────────────────
        n_wl = len(self._watchlist)
        alertas_wl = sum(
            1 for w in self._watchlist
            if (w.get("last_sell") and w.get("alert_below") and
                w["last_sell"] <= w["alert_below"]) or
               (w.get("last_sell") and w.get("alert_above") and
                w["last_sell"] >= w["alert_above"]))
        wl_lines = [f"Monitorados: {n_wl}  Alertas: {alertas_wl}"]
        for w in self._watchlist[:4]:
            ls = w.get("last_sell")
            if ls:
                wl_lines.append(f"  {w['name'][:14]:<14} {ls:>13,.0f}")
        self._dash_wl_lbl.config(text="\n".join(wl_lines))

        # ── Banco SQLite (thread) ─────────────────────────────────────────────
        def _db_info():
            try:
                conn = _get_db()
                row  = conn.execute(
                    "SELECT COUNT(DISTINCT region_id), COUNT(DISTINCT type_id),"
                    " MIN(date), MAX(date) FROM history").fetchone()
                n_reg, n_tid = row[0] or 0, row[1] or 0
                dmin, dmax   = row[2] or "--", row[3] or "--"
                txt = (f"Regioes: {n_reg}  Items: {n_tid:,}\n"
                       f"De: {dmin}\nAte: {dmax}")
                col = COLORS["green"] if n_tid > 0 else COLORS["subtext"]
            except Exception:
                txt = "BD nao inicializado"; col = COLORS["subtext"]
            self.after(0, lambda: self._dash_db_lbl.config(text=txt, fg=col))
        threading.Thread(target=_db_info, daemon=True).start()

        # ── Skills Globais ────────────────────────────────────────────────────
        br, ac = self._gskills()
        rb, rt = _calc_fee_rates(br, ac)
        self._dash_skills_lbl.config(
            text=(f"Broker Rel: {br}  → {rb*100:.2f}%\n"
                  f"Accounting: {ac}  → {rt*100:.3f}%\n"
                  f"Blacklist:  {len(self._blacklist)} itens"))

        # ── Top Max Margem ────────────────────────────────────────────────────
        if self._mm_results:
            self._dash_mm_lbl.config(
                fg=COLORS["green"],
                text="\n".join(
                    f"{r['item'][:16]:<16} {r['net_margin']:>5.1f}%"
                    for r in self._mm_results[:4]))
        else:
            self._dash_mm_lbl.config(fg=COLORS["subtext"],
                                      text="(calcule Max Margem)")

        # ── Top Arbitragem ────────────────────────────────────────────────────
        if self._arb_opps:
            top4 = sorted(self._arb_opps, key=lambda x: x["roi_total"],
                          reverse=True)[:4]
            self._dash_arb_lbl.config(
                fg=COLORS["accent"],
                text="\n".join(
                    f"{r['item'][:13]:<13} {r['margem_pct']:>5.1f}%"
                    f" {r['roi_total']:>11,.0f}"
                    for r in top4))
        else:
            self._dash_arb_lbl.config(fg=COLORS["subtext"],
                                       text="(calcule Arbitragem)")

        # ── Orcamento ─────────────────────────────────────────────────────────
        if self._bud_results:
            st = self._bud_stats
            self._dash_bud_lbl.config(
                fg=COLORS["yellow"],
                text=(f"Itens: {st.get('n_items',0)}"
                      f"  Util: {st.get('utilization',0):.1f}%\n"
                      f"Gasto:  {st.get('total_spent',0):>14,.0f} ISK\n"
                      f"Lucro:  {st.get('total_profit',0):>14,.0f} ISK"))
        else:
            self._dash_bud_lbl.config(fg=COLORS["subtext"],
                                       text="(calcule Orcamento)")

        # ── Ultimo Comparador ─────────────────────────────────────────────────
        if self._hub_results:
            name = getattr(self, "_hub_name_cur", "")
            best = self._hub_results[0]
            self._dash_hub_lbl.config(
                fg=COLORS["green"],
                text=(f"{name[:20]}\n"
                      f"Melhor: {best['hub']}\n"
                      f"Margem: {best['net_pct']:.2f}%\n"
                      f"Sell:   {best['sell_min']:,.0f} ISK"))
        else:
            self._dash_hub_lbl.config(fg=COLORS["subtext"],
                                       text="(use Comparador)")

        # ── Alertas recentes ──────────────────────────────────────────────────
        if self._alert_history:
            self._dash_alert_lbl.config(
                fg=COLORS["yellow"],
                text="\n".join(f"{ts}  {msg}"
                               for ts, msg in self._alert_history[-4:]))
        else:
            self._dash_alert_lbl.config(fg=COLORS["subtext"],
                                         text="Nenhum alerta.")

        # ── Station Trader ────────────────────────────────────────────────────
        if self._st_results:
            top4 = self._st_results[:4]
            total_cap = sum(r.get("capital_required", 0) for r in self._st_results)
            n_hot = sum(1 for r in self._st_results
                        if r["period_roi"] >= 25 and not r["risk_flag"])
            self._dash_mh_lbl.config(
                fg=COLORS["green"],
                text=("\n".join(
                    f"{r['name'][:15]:<15} {r['period_roi']:>5.1f}%"
                    for r in top4) +
                    f"\n{len(self._st_results)} oport. | {n_hot} excelentes"))
        else:
            self._dash_mh_lbl.config(fg=COLORS["subtext"],
                                      text="(escanear Station Trader)")

        # ── Industria ─────────────────────────────────────────────────────────
        if self._ind_result:
            ir = self._ind_result
            col = COLORS["green"] if ir["profit_unit"] > 0 else COLORS["red"]
            self._dash_cache_lbl.config(
                fg=col,
                text=(f"{ir['verdict']}\n"
                      f"Custo/un: {ir['cost_per_unit']:>12,.0f} ISK\n"
                      f"Mercado:  {ir['market_price']:>12,.0f} ISK\n"
                      f"Margem:   {ir['profit_pct']:>+11.2f}%"))
        else:
            self._dash_cache_lbl.config(fg=COLORS["subtext"],
                                         text="(calcule Industria)")

        # ── Ciclos de Mercado ─────────────────────────────────────────────────
        if self._cycle_result:
            cr = self._cycle_result
            ph = cr["region_phase"]
            pc = cr["region_color"]
            n_an = len(cr["anomalies"])
            cyc_txt = (f"{CYCLE_PHASES[ph]['emoji']} {ph}\n"
                       f"Confianca: {cr['confidence']:.0f}%\n"
                       f"Trend: {cr['avg_trend']:+.4f}  Vol: {cr['avg_vol']:.2f}x\n"
                       f"Anomalias: {n_an}"
                       + (f"\n⚠ EVENTO SISTEMICO" if cr["systemic_event"] else ""))
            # Se ha previsao recente, adiciona resumo
            if self._fcast_result and self._fcast_name:
                fr = self._fcast_result
                last_p = fr["prices_hist"][-1] if fr["prices_hist"] else 0
                fcast7 = fr["fcast_central"][6] if len(fr["fcast_central"]) >= 7 else (
                          fr["fcast_central"][-1] if fr["fcast_central"] else 0)
                chg7   = (fcast7 - last_p) / last_p * 100 if last_p else 0
                cyc_txt += (f"\n\nPrevisao {self._fcast_name[:12]}:\n"
                            f"  +7d: {fcast7:,.0f} ISK ({chg7:+.1f}%)\n"
                            f"  Conf: {fr['confidence']:.0f}%")
            self._dash_pm_lbl.config(fg=pc, text=cyc_txt)
        elif self._fcast_result and self._fcast_name:
            fr = self._fcast_result
            last_p = fr["prices_hist"][-1] if fr["prices_hist"] else 0
            fc     = fr["fcast_central"]
            fcast7 = fc[6]  if len(fc) >= 7  else fc[-1] if fc else 0
            fcast14= fc[13] if len(fc) >= 14 else fc[-1] if fc else 0
            chg7   = (fcast7  - last_p) / last_p * 100 if last_p else 0
            chg14  = (fcast14 - last_p) / last_p * 100 if last_p else 0
            col    = (COLORS["green"] if chg7 > 0 else
                      COLORS["red"]   if chg7 < 0 else COLORS["subtext"])
            self._dash_pm_lbl.config(
                fg=col,
                text=(f"{self._fcast_name[:18]}\n"
                      f"+7d:  {fcast7:>12,.0f} ({chg7:+.1f}%)\n"
                      f"+14d: {fcast14:>12,.0f} ({chg14:+.1f}%)\n"
                      f"Conf: {fr['confidence']:.0f}%  "
                      f"Vol: {fr['volatility']*100:.2f}%"))
        else:
            self._dash_pm_lbl.config(fg=COLORS["subtext"],
                                      text="(calcule Ciclos ou Previsao)")

        # Reagenda refresh automatico (60s)
        if self._dash_after_id:
            try: self.after_cancel(self._dash_after_id)
            except Exception: pass
        self._dash_after_id = self.after(60_000, self._dash_refresh)

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA MULTI-HOP (Tier 2) — arbitragem encadeada A→B→C
    # ══════════════════════════════════════════════════════════════════════════

    _MH_COL_NAMES = {
        "item":          "Item",
        "route":         "Rota",
        "net_pct_total": "Margem Total %",
        "profit_unit":   "Lucro/un (ISK)",
        "capital_entry": "Capital Entrada",
        "n_hops":        "Pulos",
    }

    def _build_tab_mh(self):
        p = self._tab_mh

        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(6, 2))

        regions = list(REGIONS.keys())
        self._lbl(fr, "Hub A",           0, 0)
        self._mh_a = self._combo(fr, regions, 1, 0)
        self._lbl(fr, "Hub B",           0, 2)
        self._mh_b = self._combo(fr, regions[1:]+[regions[0]], 1, 2)
        self._lbl(fr, "Hub C (opcional)",0, 4)
        c3_opts = ["(sem terceiro hub)"] + regions
        self._mh_c = self._combo(fr, c3_opts, 1, 4, 20)
        self._lbl(fr, "Margem min %",    0, 6)
        self._mh_mmin = self._entry(fr, "3.0", 1, 6, 7)
        self._lbl(fr, "Top N",           0, 8)
        self._mh_topn = self._entry(fr, "200", 1, 8, 6)
        self._lbl(fr, "Fonte",           0, 10)
        self._mh_mode = self._combo(fr, ["CSV (cache)", "API (live)"], 1, 10, 13)
        self._csv_age_label(fr, self._mh_mode, [self._mh_a, self._mh_b, self._mh_c], 2, 10)

        self._mh_btn = tk.Button(fr, text="  CALCULAR MULTI-HOP  ",
                                  command=self._start_mh,
                                  bg=COLORS["orange"], fg="#000",
                                  font=("Courier", 11, "bold"),
                                  relief="flat", cursor="hand2")
        self._mh_btn.grid(row=1, column=12, padx=8)
        tk.Button(fr, text=" EXPORTAR ", command=self._export_mh,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=13, padx=4)
        tk.Button(fr, text=" + WATCHLIST ", command=self._mh_add_wl,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=14, padx=4)
        tk.Button(fr, text=" SPREAD HIST. ", command=self._mh_open_spread,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=15, padx=4)
        self._mh_prog = ttk.Progressbar(fr, mode="indeterminate", length=90)
        self._mh_prog.grid(row=1, column=16, padx=6)

        self._mh_tree = self._make_tree(p, [
            ("item",          "Item",           220),
            ("route",         "Rota",           280),
            ("net_pct_total", "Margem Total %",  110),
            ("profit_unit",   "Lucro/un (ISK)",  130),
            ("capital_entry", "Capital Entrada", 130),
            ("n_hops",        "Pulos",            60),
        ])
        for cid in self._MH_COL_NAMES:
            self._mh_tree.heading(cid, text=self._MH_COL_NAMES[cid],
                                   command=lambda c=cid: self._sort_mh(c))
        self._mh_tree.bind("<Double-1>", self._on_dbl_mh)

        bar = tk.Frame(p, bg=COLORS["bg"])
        bar.pack(fill="x", padx=6, pady=(0, 4))
        self._mh_count = tk.Label(bar, text="0 oportunidades",
                                   fg=COLORS["subtext"], bg=COLORS["bg"],
                                   font=("Courier", 10))
        self._mh_count.pack(side="left")
        self._mh_tree.bind("<<TreeviewSelect>>", self._mh_on_select)
        self._mh_tree.bind("<Double-1>",         self._mh_on_dbl)
        tk.Button(bar, text=" GRÁFICO ",     command=self._mh_open_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" PREVER ",      command=self._mh_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" BACKTEST ",    command=self._mh_open_bt,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" ANALISAR IA ", command=self._mh_open_ai,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Label(bar,
                 text="Double-click = detalhe dos pulos  |  "
                      "Margem calculada com skills globais (aba CONFIGURACOES)",
                 fg=COLORS["subtext"], bg=COLORS["bg"],
                 font=("Courier", 10)).pack(side="right")

    def _start_mh(self):
        if self._loading_mh: return
        self._loading_mh = True
        self._mh_btn.config(state="disabled")
        self._mh_prog.start(10)
        hub_a   = self._mh_a.get()
        hub_b   = self._mh_b.get()
        hub_c   = self._mh_c.get()
        hops    = [hub_a, hub_b]
        if hub_c and not hub_c.startswith("("):
            hops.append(hub_c)
        br, ac = self._gskills()
        params = dict(
            hops     = hops,
            use_live = self._mh_mode.get().startswith("API"),
            mmin     = self._fread(self._mh_mmin, 3.0),
            topn     = int(self._fread(self._mh_topn, 200)),
            br=br, ac=ac,
        )
        threading.Thread(target=self._worker_mh, args=(params,), daemon=True).start()

    def _worker_mh(self, params: dict):
        try:
            hops     = params["hops"]
            use_live = params["use_live"]
            orders_map = {}
            for rname in hops:
                if use_live:
                    orders_map[rname] = download_region_orders(
                        REGIONS[rname], rname,
                        progress_cb=lambda m: self.after(
                            0, lambda t=m: self._set_status(str(t)[:180])))
                    _orders_mem_cache[rname] = (time.time(), orders_map[rname])
                else:
                    o = load_orders_from_csv(rname)
                    if not o:
                        self.after(0, lambda r=rname: messagebox.showerror(
                            "CSV ausente",
                            f"Baixe ordens de {r} na aba DADOS & CACHE."))
                        return
                    orders_map[rname] = o

            results = run_multihop(
                hops, orders_map, self._item_cache,
                broker_rel  = params["br"],
                accounting  = params["ac"],
                min_net_pct = params["mmin"],
                top_n       = params["topn"],
                progress_cb = lambda m: self.after(
                    0, lambda t=m: self._set_status(str(t)[:180])),
            )
            self._mh_results = results
            self.after(0, self._pop_mh)
        except Exception as exc:
            log.exception("worker_mh")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro", e))
        finally:
            self._loading_mh = False
            self.after(0, lambda: (self._mh_btn.config(state="normal"),
                                   self._mh_prog.stop()))

    def _pop_mh(self, reorder_only=False):
        if reorder_only:
            for i in range(len(self._mh_results)):
                self._mh_tree.move(str(i), "", i)
            return
        for row in self._mh_tree.get_children():
            self._mh_tree.delete(row)
        for i, r in enumerate(self._mh_results):
            tag = ("hot",) if r["net_pct_total"] >= 10 else (
                "even" if i % 2 == 0 else "odd",)
            self._mh_tree.insert("", "end", iid=str(r["type_id"]), tags=tag, values=(
                r["item"],
                r["route"],
                f"{r['net_pct_total']:.2f}%",
                f"{r['profit_unit']:,.2f}",
                f"{r['capital_entry']:,.2f}",
                str(r["n_hops"]),
            ))
        self._mh_count.config(
            text=f"{len(self._mh_results)} oportunidades multi-hop")

    def _sort_mh(self, col):
        num = {"net_pct_total","profit_unit","capital_entry","n_hops"}
        rev = not self._mh_sort_state.get(col, False)
        self._mh_sort_state = {col: rev}
        self._mh_results.sort(
            key=lambda r: (r.get(col) or 0) if col in num else r.get(col, ""),
            reverse=rev)
        self._set_heading_arrow(self._mh_tree, self._MH_COL_NAMES, col, rev)
        self._pop_mh(reorder_only=True)

    def _on_dbl_mh(self, _):
        sel = self._mh_tree.selection()
        if not sel: return
        r = self._mh_results[int(sel[0])]
        detail = f"{r['item']}\n\n"
        for h in r["hops"]:
            detail += (f"  {h['from'][:20]}\n"
                       f"    Compra: {h['buy_p']:>16,.2f} ISK\n"
                       f"    Venda:  {h['sell_p']:>16,.2f} ISK\n"
                       f"    Margem perna: {h['net_pct']:+.2f}%\n"
                       f"    → {h['to'][:20]}\n\n")
        detail += f"Margem encadeada total: {r['net_pct_total']:.2f}%\n"
        detail += f"Lucro estimado / un:    {r['profit_unit']:,.2f} ISK"
        messagebox.showinfo(f"Detalhe Multi-hop — {r['item']}", detail)

    def _export_mh(self):
        headers = [("item","Item"),("route","Rota"),("net_pct_total","Margem Total %"),
                   ("profit_unit","Lucro/un"),("capital_entry","Capital"),("n_hops","Pulos")]
        self._export_table("Multi-Hop", headers, self._mh_results or [], "multi_hop")
    def _mh_add_wl(self):
        """Adiciona o item selecionado no multi-hop à Watchlist."""
        sel = self._mh_tree.selection()
        if not sel:
            messagebox.showinfo("Selecione", "Clique num item primeiro."); return
        r   = self._mh_results[int(sel[0])]
        tid = r["type_id"]
        if any(w["type_id"] == tid for w in self._watchlist):
            messagebox.showinfo("Já existe", f"{r['item']} já está na Watchlist."); return
    def _export_mh(self):
        headers = [("item","Item"),("src","Origem"),("dst","Destino"),
                   ("net_pct_total","Margem Total %"),("lucro_dia","Lucro/Dia"),
                   ("vol","Volume"),("hops","Pulos")]
        self._export_table("Multi-Hop", headers, self._mh_results or [], "multi_hop")
    def _mh_open_spread(self):
        """Abre Spread Hist. para o item selecionado no multi-hop."""
        sel = self._mh_tree.selection()
        if not sel:
            messagebox.showinfo("Selecione", "Clique num item primeiro."); return
        r    = self._mh_results[int(sel[0])]
        tid  = r["type_id"]
        name = r["item"]
        self._sph_tid.set(str(tid))
        self._sph_name.set(name)
        self._nb.select(self._tab_sph)
        self.after(50, self._start_sph)

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA CORRELACAO (Tier 2) — correlacao de preco entre grupos de itens
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_corr(self):
        p = self._tab_corr

        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(6, 2))

        self._lbl(fr, "Regiao",           0, 0)
        self._corr_reg = self._combo(fr, list(REGIONS.keys()), 1, 0)
        self._lbl(fr, "Type IDs (virgula)", 0, 2)
        self._corr_tids = self._entry(fr, "34,35,36,37,38,39,40", 1, 2, 38)
        self._lbl(fr, "Dias min comuns",  0, 4)
        self._corr_days = self._entry(fr, "30", 1, 4, 6)

        self._corr_btn = tk.Button(fr, text="  CALCULAR CORRELACAO  ",
                                    command=self._start_corr,
                                    bg=COLORS["yellow"], fg="#000",
                                    font=("Courier", 11, "bold"),
                                    relief="flat", cursor="hand2")
        self._corr_btn.grid(row=1, column=6, padx=8)
        tk.Button(fr, text=" EXPORTAR ", command=self._export_corr,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=7, padx=4)
        tk.Button(fr, text=" GRAFICO SERIES ", command=self._corr_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=8, padx=4)
        self._corr_spin = ttk.Progressbar(fr, mode="indeterminate", length=80)
        self._corr_spin.grid(row=1, column=9, padx=6)

        tk.Label(fr,
                 text="Use type_ids separados por virgula (ex: 34=Tritanium, "
                      "35=Pyerite, 36=Mexallon...)",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 9)).grid(row=0, column=8, columnspan=6,
                                           padx=8, sticky="w")

        # Tabela de resultados (pares de itens x correlacao)
        self._corr_tree = self._make_tree(p, [
            ("item_a", "Item A",        220),
            ("item_b", "Item B",        220),
            ("r",      "Pearson r",     100),
            ("signal", "Sinal",         140),
            ("days",   "Dias comuns",    90),
        ])
        for cid, label in [("item_a","Item A"),("item_b","Item B"),
                            ("r","Pearson r"),("signal","Sinal"),("days","Dias comuns")]:
            self._corr_tree.heading(
                cid, text=label,
                command=lambda c=cid: self._sort_corr(c))
        self._corr_tree.tag_configure("pos_strong", foreground=COLORS["green"])
        self._corr_tree.tag_configure("neg_strong", foreground=COLORS["red"])
        self._corr_tree.tag_configure("neutral",    foreground=COLORS["subtext"])
        self._corr_tree.bind("<Double-1>", self._on_dbl_corr)
        self._corr_tree.bind("<<TreeviewSelect>>",
                            lambda e: setattr(self, "_corr_cur_row",
                                int(self._corr_tree.selection()[0])
                                if self._corr_tree.selection() else
                                getattr(self, "_corr_cur_row", 0)))

        bar = tk.Frame(p, bg=COLORS["bg"])
        bar.pack(fill="x", padx=6, pady=(0, 4))
        self._corr_count = tk.Label(bar, text="0 pares",
                                     fg=COLORS["subtext"], bg=COLORS["bg"],
                                     font=("Courier", 10))
        self._corr_count.pack(side="left")
        tk.Button(bar, text=" GRÁFICO ",     command=self._corr_open_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" PREVER ",      command=self._corr_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" BACKTEST ",    command=self._corr_open_bt,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" ANALISAR IA ", command=self._corr_open_ai,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Label(bar,
                 text="Verde = alta correlacao positiva  |  "
                      "Vermelho = correlacao negativa  |  "
                      "Util para hedge e diversificacao",
                 fg=COLORS["subtext"], bg=COLORS["bg"],
                 font=("Courier", 10)).pack(side="right")

    def _start_corr(self):
        if self._loading_corr: return
        try:
            raw  = self._corr_tids.get().strip()
            tids = [int(x.strip()) for x in raw.split(",") if x.strip().isdigit()]
            if len(tids) < 2:
                messagebox.showinfo("Poucos itens",
                    "Insira pelo menos 2 type_ids separados por virgula."); return
        except ValueError:
            messagebox.showerror("Erro", "Type IDs invalidos."); return
        self._loading_corr = True
        self._corr_btn.config(state="disabled")
        self._corr_spin.start(10)
        params = dict(
            rname = self._corr_reg.get(),
            mdays = int(self._fread(self._corr_days, 30)),
        )
        threading.Thread(target=self._worker_corr, args=(tids, params), daemon=True).start()

    def _worker_corr(self, tids: list, params: dict):
        try:
            rname = params["rname"]
            rid   = REGIONS.get(rname, 10000002)
            mdays = params["mdays"]

            self.after(0, lambda: self._set_status(
                f"Buscando historico de {len(tids)} itens..."))

            # Pré-aquece cache RAM com 1 query SQL antes das threads paralelas
            _warm_hist_cache(rid)
            # Busca historico em paralelo via cache
            hist_map = {}
            def _fetch(tid):
                hist_map[tid] = _get_hist_cached(rid, tid)
            with concurrent.futures.ThreadPoolExecutor(max_workers=10) as pool:
                list(pool.map(_fetch, tids))

            # Resolve nomes e salva mapa tid->nome para uso no grafico
            _ensure_names(tids, self._item_cache)
            self._corr_tid_map = {tid: self._item_cache.get(tid, {}).get("name", str(tid))
                                  for tid in tids}
            # Mapa inverso nome->tid para _corr_chart
            self._corr_name_to_tid = {v: k for k, v in self._corr_tid_map.items()}
            self._corr_rid = rid

            self.after(0, lambda: self._set_status("Calculando correlacoes..."))
            matrix = compute_price_correlation(hist_map, min_days=mdays)

            # Monta rows com nomes
            rows = []
            for (a, b), r in matrix.items():
                name_a = self._item_cache.get(a, {}).get("name", str(a))
                name_b = self._item_cache.get(b, {}).get("name", str(b))
                # conta dias comuns
                sa = {d["date"] for d in hist_map.get(a, [])}
                sb = {d["date"] for d in hist_map.get(b, [])}
                days = len(sa & sb)
                signal = ("Forte positiva" if r >= 0.7
                          else "Fraca positiva" if r >= 0.3
                          else "Fraca negativa" if r >= -0.3
                          else "Forte negativa")
                rows.append({
                    "item_a": name_a, "item_b": name_b,
                    "r": r, "signal": signal, "days": days,
                })
            rows.sort(key=lambda x: abs(x["r"]), reverse=True)
            self.after(0, lambda: self._pop_corr(rows))
        except Exception as exc:
            log.exception("worker_corr")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro", e))
        finally:
            self._loading_corr = False
            self.after(0, lambda: (self._corr_btn.config(state="normal"),
                                   self._corr_spin.stop()))

    def _pop_corr(self, rows, reorder_only=False):
        if reorder_only:
            for i in range(len(rows)):
                self._corr_tree.move(str(i), "", i)
            return
        self._corr_rows = rows   # store for export/chart
        for row in self._corr_tree.get_children():
            self._corr_tree.delete(row)
        for i, r in enumerate(rows):
            if r["r"] >= 0.7:
                tag = ("pos_strong",)
            elif r["r"] <= -0.3:
                tag = ("neg_strong",)
            else:
                tag = ("neutral",)
            self._corr_tree.insert("", "end", iid=str(i), tags=tag, values=(
                r["item_a"], r["item_b"],
                f"{r['r']:+.4f}",
                r["signal"],
                str(r["days"]),
            ))
        self._corr_count.config(text=f"{len(rows)} pares calculados")

    def _sort_corr(self, col):
        if not hasattr(self, "_corr_rows") or not self._corr_rows:
            return
        if not hasattr(self, "_corr_sort_state"):
            self._corr_sort_state = {}
        if not hasattr(self, "_corr_cur_row"):
            self._corr_cur_row = 0
        num = {"r", "days"}
        rev = not self._corr_sort_state.get(col, col in num)
        self._corr_sort_state = {col: rev}
        self._corr_rows.sort(
            key=lambda r: (r.get(col) or 0) if col in num else r.get(col, ""),
            reverse=rev)
        # P6 — move() em vez de delete+reinsert
        self._pop_corr(self._corr_rows)

    def _on_dbl_corr(self, _):
        """Double-click: mostra series de preco dos dois itens."""
        sel = self._corr_tree.selection()
        if not sel or not hasattr(self, "_corr_rows"):
            return
        self._corr_chart(row_idx=int(sel[0]))

    def _corr_chart(self, row_idx: int = 0):
        """Abre grafico com as series de preco normalizadas dos dois itens selecionados."""
        if not HAS_MATPLOTLIB:
            messagebox.showinfo("Falta matplotlib", "pip install matplotlib"); return
        if not hasattr(self, "_corr_rows") or not self._corr_rows:
            messagebox.showinfo("Sem dados", "Calcule a correlacao primeiro."); return

        # Usa _corr_cur_row (armazenado no <<TreeviewSelect>>) como fallback
        # robusto quando selection() retorna vazio após perda de foco (Linux/GTK)
        sel = self._corr_tree.selection()
        if sel:
            try:
                row_idx = int(sel[0])
            except ValueError as _e:
                log.debug("ValueError silenced: %s", _e)
        elif hasattr(self, '_corr_cur_row') and self._corr_cur_row is not None:
            row_idx = self._corr_cur_row
        if row_idx >= len(self._corr_rows):
            row_idx = 0
        row = self._corr_rows[row_idx]

        # Resolve type_ids usando o mapa salvo pelo worker (O(1))
        name_map = getattr(self, "_corr_name_to_tid", {})
        rid      = getattr(self, "_corr_rid", REGIONS.get(self._corr_reg.get(), 10000002))
        tid_a = name_map.get(row["item_a"])
        tid_b = name_map.get(row["item_b"])
        if not tid_a or not tid_b:
            messagebox.showinfo("Sem Type ID",
                "Execute CALCULAR CORRELACAO primeiro."); return

        def worker():
            h_a = _get_hist_cached(rid, tid_a)
            h_b = _get_hist_cached(rid, tid_b)
            self.after(0, lambda: _draw(h_a, h_b))

        def _draw(h_a, h_b):
            if self._chart_win and self._chart_win.winfo_exists():
                self._chart_win.destroy()
            win = tk.Toplevel(self)
            win.title(f"Series: {row['item_a']}  vs  {row['item_b']}")
            win.configure(bg=COLORS["bg"])
            win.geometry("1000x560")
            self._chart_win = win

            # Datas comuns
            by_a = {d["date"]: float(d.get("average",0)) for d in h_a}
            by_b = {d["date"]: float(d.get("average",0)) for d in h_b}
            dates = sorted(set(by_a) & set(by_b))
            if len(dates) < 5:
                tk.Label(win, text="Dados insuficientes.",
                         bg=COLORS["bg"], fg=COLORS["red"],
                         font=("Courier", 12)).pack(pady=20)
                return

            xs  = list(range(len(dates)))
            pa  = [by_a[d] for d in dates]
            pb  = [by_b[d] for d in dates]
            # Normaliza (base 100)
            pa_n = [v / pa[0] * 100  if pa[0] else 0 for v in pa]
            pb_n = [v / pb[0] * 100  if pb[0] else 0 for v in pb]

            bg  = COLORS["bg"]
            pan = COLORS["panel"]
            fig = Figure(figsize=(12, 6), facecolor=bg)
            fig.subplots_adjust(left=0.07, right=0.96, top=0.90, bottom=0.12)
            ax  = fig.add_subplot(1, 1, 1)
            ax.set_facecolor(pan)
            ax.tick_params(colors=COLORS["subtext"], labelsize=7)
            for sp in ax.spines.values():
                sp.set_edgecolor(COLORS["border"])

            ax.plot(xs, pa_n, color=COLORS["accent"], lw=1.5,
                    label=row["item_a"])
            ax.plot(xs, pb_n, color=COLORS["green"],  lw=1.5,
                    label=row["item_b"])
            ax.axhline(100, color=COLORS["subtext"], lw=0.5, ls=":")
            ax.set_title(
                f"{row['item_a']} vs {row['item_b']}  |  "
                f"Pearson r = {row['r']:+.4f}  |  {row['signal']}  |  "
                f"{len(dates)} dias comuns  (base = 100)",
                color=COLORS["header"], fontsize=9, pad=6)
            ax.set_ylabel("Preco normalizado (base 100)",
                          color=COLORS["subtext"], fontsize=8)
            ax.legend(loc="upper left", fontsize=8,
                      facecolor=pan, labelcolor=COLORS["text"], framealpha=0.8)

            step = max(1, len(dates) // 10)
            ax.set_xticks(xs[::step])
            ax.set_xticklabels(dates[::step], rotation=30, ha="right",
                               fontsize=6, color=COLORS["subtext"])

            canvas = FigureCanvasTkAgg(fig, master=win)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)

        threading.Thread(target=worker, daemon=True).start()

    def _export_corr(self):
        headers = [("item_a","Item A"),("item_b","Item B"),("r","Pearson r"),
                   ("signal","Sinal"),("days","Dias Comuns")]
        rows = self._corr_rows if hasattr(self, "_corr_rows") else []
        self._export_table("Correlação", headers, rows, "correlacao")
    # ══════════════════════════════════════════════════════════════════════════
    #  ABA SPREAD HISTORICO (Tier 2) — (sell_min - buy_max) ao longo do tempo
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_sph(self):
        p = self._tab_sph

        # ── Barra de filtros ──────────────────────────────────────────────────
        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(6, 2))

        self._lbl(fr, "Type ID",     0, 0)
        self._sph_tid  = self._entry(fr, "", 1, 0, 12)
        self._lbl(fr, "Nome",        0, 2)
        self._sph_name = self._entry(fr, "", 1, 2, 22)
        self._lbl(fr, "Região",      0, 4)
        self._sph_reg  = self._combo(fr, list(REGIONS.keys()), 1, 4)
        self._lbl(fr, "Dias",        0, 6)
        self._sph_days = self._entry(fr, "180", 1, 6, 6)

        self._sph_btn = tk.Button(
            fr, text="  ANALISAR SPREAD  ",
            command=self._start_sph,
            bg=COLORS["orange"], fg="#000",
            font=("Courier", 11, "bold"), relief="flat", cursor="hand2")
        self._sph_btn.grid(row=1, column=8, padx=8)

        tk.Button(fr, text=" BACKTEST ",
                  command=self._sph_open_backtest,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 11, "bold"), relief="flat", cursor="hand2"
                  ).grid(row=1, column=9, padx=4)

        tk.Button(fr, text=" GRÁFICO ",
                  command=self._sph_open_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=10, padx=4)

        tk.Button(fr, text=" PREVER ",
                  command=self._sph_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 11, "bold"), relief="flat", cursor="hand2"
                  ).grid(row=1, column=11, padx=4)

        tk.Button(fr, text=" ANALISAR IA ",
                  command=self._sph_open_ai,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 11, "bold"), relief="flat", cursor="hand2"
                  ).grid(row=1, column=12, padx=4)

        self._sph_spin = ttk.Progressbar(fr, mode="indeterminate", length=80)
        self._sph_spin.grid(row=1, column=13, padx=6)

        tk.Label(fr,
                 text="Spread = highest − lowest por dia  |  Verde = acima da média  |  "
                      "Use Dias para controlar o período",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 9)
                 ).grid(row=0, column=4, columnspan=10, padx=4, sticky="w")

        # ── Info / rodapé ─────────────────────────────────────────────────────
        self._sph_info = tk.Label(
            p,
            text="Digite um Type ID e clique ANALISAR SPREAD.",
            bg=COLORS["bg"], fg=COLORS["subtext"],
            font=("Courier", 10), justify="left", anchor="w")
        self._sph_info.pack(fill="x", padx=8, pady=(2, 0))

        # ── Área do gráfico ───────────────────────────────────────────────────
        self._sph_frame = tk.Frame(p, bg=COLORS["bg"])
        self._sph_frame.pack(fill="both", expand=True, padx=4, pady=4)

        # Auto-lookup nome pelo Type ID
        self._bind_tid_lookup(self._sph_tid, self._sph_name, None)

    def _sph_open_chart(self):
        """Abre gráfico técnico (candlestick) para o item atual do Spread Hist."""
        try:
            tid = int(self._sph_tid.get().strip())
        except (ValueError, AttributeError):
            messagebox.showerror("Erro", "Digite um Type ID válido."); return
        name = self._sph_name.get().strip() or str(tid)
        rid  = REGIONS.get(self._sph_reg.get(), 10000002)
        self._open_candle_chart(tid, name, rid)

    def _sph_open_fcast(self):
        """Abre aba PREVISÃO pré-preenchida com o item atual do Spread Hist."""
        try:
            tid = int(self._sph_tid.get().strip())
        except (ValueError, AttributeError):
            messagebox.showerror("Erro", "Digite um Type ID válido."); return
        if self._loading_fcast: messagebox.showinfo("Aguarde", "Previsão em execução."); return
        name = self._sph_name.get().strip() or str(tid)
        self._fcast_tid_e.set(str(tid)); self._fcast_name_e.set(name)
        self._fcast_reg.set(self._sph_reg.get())
        self._nb.select(self._tab_fcast); self.after(50, self._start_fcast)

    def _sph_open_ai(self):
        """Abre aba AI ANALISTA pré-preenchida com o item atual do Spread Hist."""
        try:
            tid = int(self._sph_tid.get().strip())
        except (ValueError, AttributeError):
            messagebox.showerror("Erro", "Digite um Type ID válido."); return
        if self._loading_ai: messagebox.showinfo("Aguarde", "AI em execução."); return
        name = self._sph_name.get().strip() or str(tid)
        self._ai_tid.set(str(tid)); self._ai_name_e.set(name)
        self._ai_reg.set(self._sph_reg.get())
        self._nb.select(self._tab_ai); self.after(50, self._start_ai)


        headers = [("pair","Par"),("r","Correlação r"),("p","p-valor"),
                   ("n","N pontos"),("label","Interpretação")]
        self._export_table("Correlação", headers, self._corr_rows or [], "correlacao")
    def _start_sph(self):
        if self._loading_sph: return
        try:
            tid = int(self._sph_tid.get().strip())
        except ValueError:
            messagebox.showerror("Erro", "Type ID invalido."); return
        self._loading_sph = True
        self._sph_btn.config(state="disabled")
        self._sph_spin.start(10)
        params = dict(
            rname = self._sph_reg.get(),
            days  = max(7, int(self._fread(self._sph_days, 180))),
            name  = self._sph_name.get().strip() or str(tid),
        )
        threading.Thread(target=self._worker_sph, args=(tid, params), daemon=True).start()

    def _worker_sph(self, tid: int, params: dict):
        try:
            rname = params["rname"]
            rid   = REGIONS.get(rname, 10000002)
            days  = params["days"]
            name  = params["name"]

            _warm_hist_cache(rid)  # pré-aquece cache RAM
            hist = _get_hist_cached(rid, tid)
            if not hist:
                self.after(0, lambda: messagebox.showinfo(
                    "Sem dados", "Sem historico disponivel para este item/regiao."))
                return

            # Calcula spread diario = highest - lowest (proxy do spread intraday)
            # e variacao do preco medio
            cutoff = (datetime.now(timezone.utc) - timedelta(days=days)).strftime("%Y-%m-%d")
            hist_f = sorted([d for d in hist if d.get("date","") >= cutoff],
                            key=lambda x: x["date"])
            if len(hist_f) < 5:
                self.after(0, lambda: messagebox.showinfo(
                    "Poucos dados",
                    f"Apenas {len(hist_f)} registros no periodo."))
                return

            dates    = [d["date"]               for d in hist_f]
            spreads  = [float(d.get("highest",0)) - float(d.get("lowest",0))
                        for d in hist_f]
            prices   = [float(d.get("average",0)) for d in hist_f]
            spreads_pct = [(s / p * 100) if p > 0 else 0
                           for s, p in zip(spreads, prices)]
            vols     = [int(d.get("volume",0))   for d in hist_f]

            self.after(0, lambda: self._draw_sph(
                name, dates, spreads, spreads_pct, prices, vols))
        except Exception as exc:
            log.exception("worker_sph")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro", e))
        finally:
            self._loading_sph = False
            self.after(0, lambda: (self._sph_btn.config(state="normal"),
                                   self._sph_spin.stop()))

    def _draw_sph(self, name, dates, spreads, spreads_pct, prices, vols):
        if not HAS_MATPLOTLIB:
            messagebox.showinfo("Falta matplotlib", "pip install matplotlib")
            return
        # Limpa frame anterior
        for w in self._sph_frame.winfo_children():
            w.destroy()

        xs  = list(range(len(dates)))
        bg  = COLORS["bg"]
        pan = COLORS["panel"]

        fig = Figure(figsize=(13, 7), facecolor=bg)
        fig.subplots_adjust(left=0.07, right=0.96, hspace=0.35,
                            top=0.92, bottom=0.10)

        def _style(ax):
            ax.set_facecolor(pan)
            ax.tick_params(colors=COLORS["subtext"], labelsize=7)
            for sp in ax.spines.values():
                sp.set_edgecolor(COLORS["border"])

        # Subplot 1 — Preco medio diario
        ax1 = fig.add_subplot(3, 1, 1)
        _style(ax1)
        ax1.plot(xs, prices, color=COLORS["accent"], lw=1.5)
        ax1.set_title(f"{name} — Preco medio diario",
                      color=COLORS["header"], fontsize=9, pad=4)
        ax1.set_ylabel("ISK", color=COLORS["subtext"], fontsize=7)
        ax1.tick_params(labelbottom=False)

        # Subplot 2 — Spread absoluto (high - low)
        ax2 = fig.add_subplot(3, 1, 2)
        _style(ax2)
        avg_spread = sum(spreads) / len(spreads)
        above = [s >= avg_spread for s in spreads]
        cols_s = [COLORS["green"] if a else COLORS["subtext"] for a in above]
        ax2.bar(xs, spreads, color=cols_s, alpha=0.75, width=0.8)
        ax2.axhline(avg_spread, color=COLORS["yellow"], lw=1, ls="--",
                    label=f"Media {avg_spread:,.0f} ISK")
        ax2.set_title("Spread diario absoluto (high - low)",
                      color=COLORS["header"], fontsize=9, pad=4)
        ax2.set_ylabel("ISK", color=COLORS["subtext"], fontsize=7)
        ax2.legend(loc="upper left", fontsize=6,
                   facecolor=pan, labelcolor=COLORS["text"], framealpha=0.8)
        ax2.tick_params(labelbottom=False)

        # Subplot 3 — Spread % sobre preco medio
        ax3 = fig.add_subplot(3, 1, 3)
        _style(ax3)
        avg_pct = sum(spreads_pct) / len(spreads_pct)
        cols_p = [COLORS["green"] if s >= avg_pct else COLORS["subtext"]
                  for s in spreads_pct]
        ax3.bar(xs, spreads_pct, color=cols_p, alpha=0.75, width=0.8)
        ax3.axhline(avg_pct, color=COLORS["yellow"], lw=1, ls="--",
                    label=f"Media {avg_pct:.2f}%")
        ax3.set_title("Spread % sobre preco medio",
                      color=COLORS["header"], fontsize=9, pad=4)
        ax3.set_ylabel("%", color=COLORS["subtext"], fontsize=7)
        ax3.legend(loc="upper left", fontsize=6,
                   facecolor=pan, labelcolor=COLORS["text"], framealpha=0.8)

        # Eixo X
        step = max(1, len(dates) // 12)
        ax3.set_xticks(xs[::step])
        ax3.set_xticklabels(dates[::step], rotation=30, ha="right",
                             fontsize=6, color=COLORS["subtext"])

        max_spread = max(spreads)
        max_idx    = spreads.index(max_spread)
        self._sph_info.config(
            text=(f"{name}  |  {len(dates)} dias  |  "
                  f"Spread max: {max_spread:,.0f} ISK ({spreads_pct[max_idx]:.2f}%) "
                  f"em {dates[max_idx]}  |  Media: {avg_spread:,.0f} ISK ({avg_pct:.2f}%)"))

        canvas = FigureCanvasTkAgg(fig, master=self._sph_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    # ── Helper: skills globais ────────────────────────────────────────────────

    def _gskills(self) -> tuple[int, int]:
        """Retorna (broker_rel, accounting) dos campos globais (CONFIGURACOES)."""
        try:
            br = max(0, min(5, int(float(self._g_brel.get() or "0"))))
        except (ValueError, AttributeError):
            br = 0
        try:
            ac = max(0, min(5, int(float(self._g_acc.get() or "0"))))
        except (ValueError, AttributeError):
            ac = 0
        return br, ac

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA 13 — COMPARADOR DE HUBS (Tier 1)
    # ══════════════════════════════════════════════════════════════════════════

    _HUB_COLS = {
        "hub":         "Hub",
        "sell_min":    "Sell Min (ISK)",
        "buy_max":     "Buy Max (ISK)",
        "spread_pct":  "Spread %",
        "net_pct":     "Margem Líq %",
        "sell_orders": "Ordens Venda",
        "buy_orders":  "Ordens Compra",
    }

    def _build_tab_hub(self):
        p = self._tab_hub

        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(6, 2))

        self._lbl(fr, "Type ID",     0, 0)
        self._hub_tid  = self._entry(fr, "", 1, 0, 12)
        self._lbl(fr, "Nome",        0, 2)
        self._hub_name = self._entry(fr, "", 1, 2, 22)
        self._lbl(fr, "Broker Rel",  0, 4)
        tk.Label(fr, text="(global)", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 9)
                 ).grid(row=0, column=5, padx=0)

        self._hub_btn = tk.Button(fr, text="  COMPARAR HUBS  ",
                                   command=self._start_hub,
                                   bg=COLORS["accent"], fg="#000",
                                   font=("Courier", 11, "bold"),
                                   relief="flat", cursor="hand2")
        self._hub_btn.grid(row=1, column=6, padx=8)

        tk.Button(fr, text=" + WATCHLIST ", command=self._hub_add_wl,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=7, padx=4)

        tk.Button(fr, text=" GRÁFICO ", command=self._hub_chart,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=8, padx=4)

        tk.Button(fr, text=" SPREAD HIST. ", command=self._hub_open_spread,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=9, padx=4)

        self._hub_spin = ttk.Progressbar(fr, mode="indeterminate", length=80)
        self._hub_spin.grid(row=1, column=10, padx=6)

        # Tabela principal — hubs como linhas
        self._hub_tree = self._make_tree(p, [
            ("hub",         "Hub",           120),
            ("sell_min",    "Sell Min (ISK)", 140),
            ("buy_max",     "Buy Max (ISK)",  140),
            ("spread_pct",  "Spread %",        90),
            ("net_pct",     "Margem Líq %",   100),
            ("sell_orders", "Ordens Venda",    100),
            ("buy_orders",  "Ordens Compra",   100),
        ])
        for cid in self._HUB_COLS:
            self._hub_tree.heading(cid, text=self._HUB_COLS[cid],
                                    command=lambda c=cid: self._sort_hub(c))
        self._hub_tree.tag_configure("best",  foreground=COLORS["green"])
        self._hub_tree.tag_configure("worst", foreground=COLORS["red"])

        # Barra de info
        bar = tk.Frame(p, bg=COLORS["bg"])
        bar.pack(fill="x", padx=6, pady=(0, 4))
        self._hub_info = tk.Label(bar, text="Digite um Type ID e clique COMPARAR HUBS",
                                   fg=COLORS["subtext"], bg=COLORS["bg"],
                                   font=("Courier", 10))
        self._hub_info.pack(side="left")
        tk.Label(bar,
                 text="Verde = melhor margem  |  Skills globais (aba CONFIGURACOES)",
                 fg=COLORS["subtext"], bg=COLORS["bg"],
                 font=("Courier", 10)).pack(side="right")

        # Auto-lookup nome pelo Type ID (busca no cache local + ESI)
        self._bind_tid_lookup(self._hub_tid, self._hub_name, None)

    def _start_hub(self):
        if self._loading_hub: return
        try:
            tid = int(self._hub_tid.get().strip())
        except ValueError:
            messagebox.showerror("Erro", "Type ID invalido."); return
        self._loading_hub = True
        self._hub_btn.config(state="disabled")
        self._hub_spin.start(10)
        name = self._hub_name.get().strip() or str(tid)
        br, ac = self._gskills()
        threading.Thread(target=self._worker_hub,
                         args=(tid, name, br, ac), daemon=True).start()

    def _worker_hub(self, tid: int, name: str, br: int, ac: int):
        try:
            results = compare_hubs(tid, broker_rel=br, accounting=ac,
                                   progress_cb=self._set_status)
            self._hub_results = results
            self._hub_tid_cur  = tid
            self._hub_name_cur = name
            self.after(0, self._pop_hub)
        except Exception as exc:
            log.exception("worker_hub")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro", e))
        finally:
            self._loading_hub = False
            self.after(0, lambda: (self._hub_btn.config(state="normal"),
                                   self._hub_spin.stop()))

    def _pop_hub(self, reorder_only=False):
        if reorder_only:
            for i in range(len(self._hub_results)):
                self._hub_tree.move(str(i), "", i)
            return
        for row in self._hub_tree.get_children():
            self._hub_tree.delete(row)
        best_net = max((r["net_pct"] for r in self._hub_results), default=0)
        worst_net= min((r["net_pct"] for r in self._hub_results), default=0)
        for i, r in enumerate(self._hub_results):
            if r["net_pct"] == best_net and best_net > 0:
                tag = ("best",)
            elif r["net_pct"] == worst_net and len(self._hub_results) > 1:
                tag = ("worst",)
            else:
                tag = ("even" if i % 2 == 0 else "odd",)
            self._hub_tree.insert("", "end", iid=str(i), tags=tag, values=(
                r["hub"],
                f"{r['sell_min']:,.2f}" if r["sell_min"] else "--",
                f"{r['buy_max']:,.2f}"  if r["buy_max"]  else "--",
                f"{r['spread_pct']:.2f}%",
                f"{r['net_pct']:.2f}%",
                f"{r['sell_orders']:,}",
                f"{r['buy_orders']:,}",
            ))
        name = getattr(self, "_hub_name_cur", str(getattr(self, "_hub_tid_cur", "")))
        self._hub_info.config(
            text=f"{name}  |  {len(self._hub_results)} hubs comparados  "
                 f"|  Melhor margem: {best_net:.2f}%")

    def _sort_hub(self, col):
        num = {"sell_min","buy_max","spread_pct","net_pct","sell_orders","buy_orders"}
        rev = not self._hub_sort_state.get(col, False)
        self._hub_sort_state = {col: rev}
        self._hub_results.sort(
            key=lambda r: (r.get(col) or 0) if col in num else r.get(col, ""),
            reverse=rev)
        self._set_heading_arrow(self._hub_tree, self._HUB_COLS, col, rev)
        self._pop_hub(reorder_only=True)

    def _hub_add_wl(self):
        """Adiciona o item atual à Watchlist."""
        tid = getattr(self, "_hub_tid_cur", None)
        if not tid:
            messagebox.showinfo("Sem item", "Compare um item primeiro."); return
        if any(w["type_id"] == tid for w in self._watchlist):
            messagebox.showinfo("Já existe", "Item já está na Watchlist."); return
        name = getattr(self, "_hub_name_cur", str(tid))
        self._watchlist.append({
            "type_id": tid, "name": name,
            "alert_below": None, "alert_above": None,
            "last_sell": None, "last_buy": None,
            "last_net": None, "last_score": None,
        })
        save_watchlist(self._watchlist)
        self._pop_wl()
        messagebox.showinfo("Adicionado", f"{name} adicionado à Watchlist.")

    def _hub_open_spread(self):
        """Preenche a aba SPREAD HIST. com o item atual do Comparador e calcula."""
        tid = getattr(self, "_hub_tid_cur", None)
        if not tid:
            messagebox.showinfo("Sem item", "Compare um item primeiro."); return
        name = getattr(self, "_hub_name_cur", str(tid))
        self._sph_tid.set(str(tid))
        self._sph_name.set(name)
        if self._hub_results:
            best_hub = self._hub_results[0]["hub"]
            rname    = next((k for k in REGIONS if best_hub in k),
                            self._sph_reg.get())
            self._sph_reg.set(rname)
        self._nb.select(self._tab_sph)
        self.after(50, self._start_sph)

    def _hub_chart(self):
        """Abre gráfico técnico (candlestick) do item atual no hub de maior margem."""
        if not HAS_MATPLOTLIB:
            messagebox.showinfo("Falta matplotlib", "pip install matplotlib"); return
        tid = getattr(self, "_hub_tid_cur", None)
        if not tid:
            messagebox.showinfo("Sem item", "Compare um item primeiro."); return
        if self._hub_results:
            best_hub = self._hub_results[0]["hub"]
            rid      = HUB_REGIONS.get(best_hub, 10000002)
        else:
            rid = 10000002
        name = getattr(self, "_hub_name_cur", str(tid))
        self._open_candle_chart(tid, name, rid)

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA 14 — CONFIGURAÇÕES GLOBAIS + BLACKLIST (Tier 1)
    # ══════════════════════════════════════════════════════════════════════════

    # ── Recursos experimentais (P0.4) ─────────────────────────────────────────

    def _setup_experimental_gate(self):
        """Esconde Previsão e AI Analista quando 'experimental' está desligado.

        Os ~40 atalhos que fazem self._nb.select(tab) passam pelo wrapper abaixo
        e mostram um aviso em vez de abrir a aba escondida.
        """
        self._gated_tabs = (self._tab_fcast, self._tab_ai)
        for tab in self._gated_tabs:
            tk.Label(tab, text="⚠ " + EXPERIMENTAL_NOTICE, bg=COLORS["yellow"],
                     fg="#000", font=("Courier", 10, "bold"), anchor="w", padx=8
                     ).pack(fill="x", before=tab.winfo_children()[0])
        nb = self._nb
        orig_select = nb.select

        def gated_select(tab_id=None):
            if tab_id is not None and not self._experimental.get() and \
                    str(tab_id) in {str(t) for t in self._gated_tabs}:
                self._experimental_ok()
                return None
            return orig_select(tab_id) if tab_id is not None else orig_select()
        nb.select = gated_select
        self._apply_experimental()

    def _apply_experimental(self):
        for tab in self._gated_tabs:
            if self._experimental.get():
                self._nb.add(tab)       # reexibe aba escondida
            else:
                self._nb.hide(tab)

    def _experimental_ok(self) -> bool:
        if self._experimental.get():
            return True
        messagebox.showinfo(
            "Recurso experimental desligado",
            "Previsão e AI Analista não têm validação estatística.\n\n"
            "Para usar mesmo assim, ative 'Recursos experimentais' na aba "
            "CONFIGURACOES.")
        return False

    def _toggle_experimental(self):
        self._apply_experimental()
        self._persist_config()

    def _build_tab_cfg(self):
        p = self._tab_cfg

        # ── Recursos experimentais ────────────────────────────────────────────
        exp_fr = tk.LabelFrame(p, text=" Recursos experimentais ",
                               bg=COLORS["panel"], fg=COLORS["yellow"],
                               font=("Courier", 11, "bold"),
                               highlightbackground=COLORS["border"], highlightthickness=1)
        exp_fr.pack(fill="x", padx=10, pady=(10, 4))
        tk.Checkbutton(exp_fr, text="Mostrar Previsão e AI Analista",
                       variable=self._experimental, command=self._toggle_experimental,
                       bg=COLORS["panel"], fg=COLORS["text"],
                       selectcolor=COLORS["border"], activebackground=COLORS["panel"],
                       font=("Courier", 11)).pack(anchor="w", padx=8, pady=(4, 0))
        tk.Label(exp_fr, text=EXPERIMENTAL_NOTICE, bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 9)
                 ).pack(anchor="w", padx=8, pady=(0, 4))

        # ── Skills globais ────────────────────────────────────────────────────
        sk_fr = tk.LabelFrame(p, text=" Skills de Trading Globais ",
                              bg=COLORS["panel"], fg=COLORS["header"],
                              font=("Courier", 11, "bold"),
                              highlightbackground=COLORS["border"], highlightthickness=1)
        sk_fr.pack(fill="x", padx=10, pady=(10, 4))

        tk.Label(sk_fr,
                 text="Estes valores propagam automaticamente para Max Margem, "
                      "Orçamento, Paper Trading e Comparador de Hubs.",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 10)).grid(row=0, column=0, columnspan=8,
                                           padx=8, pady=(4, 2), sticky="w")

        def _skill_lbl(col, txt):
            tk.Label(sk_fr, text=txt, bg=COLORS["panel"],
                     fg=COLORS["subtext"], font=("Courier", 10)
                     ).grid(row=1, column=col, padx=10, pady=2, sticky="w")

        def _skill_spin(col, var, lo=0, hi=5):
            sb = tk.Spinbox(sk_fr, from_=lo, to=hi, textvariable=var,
                            width=4, bg=COLORS["border"], fg=COLORS["text"],
                            buttonbackground=COLORS["border"],
                            insertbackground=COLORS["text"],
                            font=("Courier", 12, "bold"), relief="flat")
            sb.grid(row=2, column=col, padx=10, pady=4)
            return sb

        _skill_lbl(0, "Broker Relations (0-5)")
        _skill_lbl(1, "Accounting (0-5)")
        _skill_spin(0, self._g_brel)
        _skill_spin(1, self._g_acc)

        # Preview de taxas
        self._cfg_fee_lbl = tk.Label(sk_fr, text="",
                                      bg=COLORS["panel"], fg=COLORS["accent"],
                                      font=("Courier", 11, "bold"))
        self._cfg_fee_lbl.grid(row=2, column=2, columnspan=4, padx=16, sticky="w")

        def _update_preview(*_):
            try:
                br = max(0, min(5, int(float(self._g_brel.get() or "0"))))
                ac = max(0, min(5, int(float(self._g_acc.get() or "0"))))
            except ValueError:
                return
            rb, rt = _calc_fee_rates(br, ac)
            total  = rb * 2 + rt   # buy broker + sell broker + tax
            self._cfg_fee_lbl.config(
                text=f"Broker fee: {rb*100:.2f}%  |  "
                     f"Sales tax: {rt*100:.3f}%  |  "
                     f"Custo total por trade: {total*100:.3f}%")

        self._g_brel.trace_add("write", _update_preview)
        self._g_acc.trace_add("write",  _update_preview)
        _update_preview()

        tk.Button(sk_fr, text=" SALVAR SKILLS ",
                  command=self._cfg_save_skills,
                  bg=COLORS["green"], fg="#000",
                  font=("Courier", 11, "bold"), relief="flat", cursor="hand2"
                  ).grid(row=2, column=6, padx=12)

        # ── Blacklist ─────────────────────────────────────────────────────────
        bl_fr = tk.LabelFrame(p, text=" Blacklist de Itens ",
                              bg=COLORS["panel"], fg=COLORS["header"],
                              font=("Courier", 11, "bold"),
                              highlightbackground=COLORS["border"], highlightthickness=1)
        bl_fr.pack(fill="both", expand=True, padx=10, pady=(4, 10))

        tk.Label(bl_fr,
                 text="Itens na blacklist são ignorados por todas as análises "
                      "(Max Margem, Escassez, DSIF, Orçamento).",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 10)).pack(anchor="w", padx=8, pady=(4, 2))

        # Controles
        add_bar = tk.Frame(bl_fr, bg=COLORS["panel"])
        add_bar.pack(fill="x", padx=8, pady=(2, 4))

        tk.Label(add_bar, text="Type ID:", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 10)).pack(side="left")
        self._bl_tid = tk.Entry(add_bar, width=12,
                                 bg=COLORS["border"], fg=COLORS["text"],
                                 insertbackground=COLORS["text"],
                                 font=("Courier", 11), relief="flat", bd=2)
        self._bl_tid.pack(side="left", padx=4)
        tk.Label(add_bar, text="Nome:", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 10)).pack(side="left")
        self._bl_name = tk.Entry(add_bar, width=24,
                                  bg=COLORS["border"], fg=COLORS["text"],
                                  insertbackground=COLORS["text"],
                                  font=("Courier", 11), relief="flat", bd=2)
        self._bl_name.pack(side="left", padx=4)
        tk.Button(add_bar, text=" + ADICIONAR ",
                  command=self._bl_add,
                  bg=COLORS["red"], fg="#fff",
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)
        tk.Button(add_bar, text=" REMOVER SELECIONADO ",
                  command=self._bl_remove,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)
        tk.Button(add_bar, text=" LIMPAR TUDO ",
                  command=self._bl_clear,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)

        # Lista da blacklist
        self._bl_tree = self._make_tree(bl_fr, [
            ("type_id", "Type ID", 90),
            ("name",    "Nome",   360),
        ])

        self._bl_count = tk.Label(bl_fr, text="0 itens na blacklist",
                                   bg=COLORS["panel"], fg=COLORS["subtext"],
                                   font=("Courier", 10))
        self._bl_count.pack(anchor="w", padx=8, pady=(0, 4))

        self._pop_bl()

    def _cfg_save_skills(self):
        self._persist_config()
        br, ac = self._gskills()
        rb, rt = _calc_fee_rates(br, ac)
        messagebox.showinfo("Skills Salvos",
            f"Broker Relations: {br}  →  {rb*100:.2f}% por ordem\n"
            f"Accounting:        {ac}  →  {rt*100:.3f}% sales tax\n\n"
            f"Todas as abas usarão esses valores automaticamente.")

    def _pop_bl(self):
        for row in self._bl_tree.get_children():
            self._bl_tree.delete(row)
        # Resolve nomes dos type_ids na blacklist
        names = {tid: self._item_cache.get(tid, {}).get("name", "--")
                 for tid in self._blacklist}
        for i, tid in enumerate(sorted(self._blacklist)):
            tag = "even" if i % 2 == 0 else "odd"
            self._bl_tree.insert("", "end", iid=str(tid), tags=(tag,),
                                  values=(tid, names.get(tid, "--")))
        self._bl_count.config(
            text=f"{len(self._blacklist)} iten(s) na blacklist")

    def _bl_add(self):
        try:
            tid = int(self._bl_tid.get().strip())
        except ValueError:
            messagebox.showerror("Erro", "Type ID deve ser um número inteiro."); return
        name = self._bl_name.get().strip()
        if name and tid not in self._item_cache:
            self._item_cache[tid] = {"name": name, "volume": 0}
            save_item_cache(self._item_cache)
        self._blacklist = blacklist_add(self._blacklist, tid)
        self._bl_tid.delete(0, "end")
        self._bl_name.delete(0, "end")
        self._pop_bl()

    def _bl_remove(self):
        sel = self._bl_tree.selection()
        if not sel: return
        try:
            tid = int(sel[0])
        except ValueError:
            return
        self._blacklist = blacklist_remove(self._blacklist, tid)
        self._pop_bl()

    def _bl_clear(self):
        if not messagebox.askyesno("Limpar", "Remover TODOS os itens da blacklist?"):
            return
        self._blacklist = set()
        save_blacklist(self._blacklist)
        self._pop_bl()

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA 9 — MERCADO LIVE
    # ══════════════════════════════════════════════════════════════════════════

    _LIVE_COLS = {
        "name":       "Item",
        "sell_min":   "Sell Min (ISK)",
        "buy_max":    "Buy Max (ISK)",
        "spread_pct": "Spread %",
        "net_pct":    "Margem Líq %",
        "vol_7d":     "Vol 7d/dia",
        "score":      "Score",
    }

    def _build_tab_live(self):
        p = self._tab_live

        # ── barra de controle ─────────────────────────────────────────────────
        top = tk.Frame(p, bg=COLORS["panel"],
                       highlightbackground=COLORS["border"], highlightthickness=1)
        top.pack(fill="x", padx=6, pady=(6, 2))

        regions = list(REGIONS.keys())
        self._lbl(top, "Regiao",      0, 0)
        self._live_reg_var = self._combo(top, regions, 1, 0)
        self._lbl(top, "Intervalo(s)",0, 2)
        self._live_interval = self._entry(top, "60", 1, 2, 5)

        self._live_start_btn = tk.Button(
            top, text=" ▶  INICIAR LIVE ",
            command=self._live_toggle,
            bg=COLORS["green"], fg="#000",
            font=("Courier", 11, "bold"), relief="flat", cursor="hand2")
        self._live_start_btn.grid(row=1, column=4, padx=8)

        tk.Button(top, text=" GRÁFICO ", command=self._live_open_chart,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=5, padx=4)

        tk.Button(top, text=" + WATCHLIST ", command=self._live_add_to_wl,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=6, padx=4)

        tk.Button(top, text=" AI ", command=self._live_ai_quick,
                  bg=COLORS["orange"], fg="#000",
                  font=("Courier", 11, "bold"), relief="flat", cursor="hand2"
                  ).grid(row=1, column=7, padx=4)

        tk.Button(top, text=" PREVER ", command=self._live_open_forecast,
                  bg=COLORS["green"], fg="#000",
                  font=("Courier", 11, "bold"), relief="flat", cursor="hand2"
                  ).grid(row=1, column=8, padx=4)

        self._live_status = tk.Label(top, text="Parado", bg=COLORS["panel"],
                                     fg=COLORS["subtext"], font=("Courier", 10))
        self._live_status.grid(row=1, column=9, padx=10)

        # ── adicionar item custom ─────────────────────────────────────────────
        add_bar = tk.Frame(p, bg=COLORS["bg"])
        add_bar.pack(fill="x", padx=6, pady=(2, 0))
        tk.Label(add_bar, text="Type ID:", bg=COLORS["bg"],
                 fg=COLORS["subtext"], font=("Courier", 10)).pack(side="left")
        self._live_add_tid = tk.Entry(add_bar, width=10,
                                       bg=COLORS["border"], fg=COLORS["text"],
                                       insertbackground=COLORS["text"],
                                       font=("Courier", 11), relief="flat", bd=2)
        self._live_add_tid.pack(side="left", padx=4)
        tk.Label(add_bar, text="Nome:", bg=COLORS["bg"],
                 fg=COLORS["subtext"], font=("Courier", 10)).pack(side="left")
        self._live_add_name = tk.Entry(add_bar, width=20,
                                        bg=COLORS["border"], fg=COLORS["text"],
                                        insertbackground=COLORS["text"],
                                        font=("Courier", 11), relief="flat", bd=2)
        self._live_add_name.pack(side="left", padx=4)
        tk.Button(add_bar, text="ADD",
                  command=self._live_add_item,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=2)
        tk.Button(add_bar, text="REMOVER",
                  command=self._live_remove_item,
                  bg=COLORS["border"], fg=COLORS["red"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=2)

        # ── tabela ────────────────────────────────────────────────────────────
        self._live_tree = self._make_tree(p, [
            ("name",       "Item",           200),
            ("sell_min",   "Sell Min (ISK)", 130),
            ("buy_max",    "Buy Max (ISK)",  130),
            ("spread_pct", "Spread %",        80),
            ("net_pct",    "Margem Líq %",    90),
            ("vol_7d",     "Vol 7d/dia",       95),
            ("score",      "Score",            68),
        ])
        self._live_tree.bind("<<TreeviewSelect>>", self._live_on_select)
        self._live_tree.bind("<Double-1>", self._live_on_dbl)
        for cid in self._LIVE_COLS:
            self._live_tree.heading(cid, text=self._LIVE_COLS[cid],
                                     command=lambda c=cid: self._sort_live(c))
        self._live_tree.tag_configure("bull", foreground=COLORS["green"])
        self._live_tree.tag_configure("bear", foreground=COLORS["red"])

        bar = tk.Frame(p, bg=COLORS["bg"])
        bar.pack(fill="x", padx=6, pady=(0, 4))
        self._live_count = tk.Label(bar, text="--",
                                     fg=COLORS["subtext"], bg=COLORS["bg"],
                                     font=("Courier", 10))
        self._live_count.pack(side="left")
        tk.Label(bar,
                 text="Score 0-10 = margem liq (50%) + volume (50%)  |  Verde = spread positivo",
                 fg=COLORS["subtext"], bg=COLORS["bg"],
                 font=("Courier", 10)).pack(side="right")

        # Preenche tabela com items da lista mesmo sem dados live
        self._pop_live_empty()

    def _pop_live_empty(self):
        for row in self._live_tree.get_children():
            self._live_tree.delete(row)
        for i, (tid, name) in enumerate(self._live_items):
            tag = "even" if i % 2 == 0 else "odd"
            self._live_tree.insert("", "end", iid=str(tid), tags=(tag,), values=(
                name, "--", "--", "--", "--", "--", "--"))
        self._live_count.config(text=f"{len(self._live_items)} itens monitorados")

    def _pop_live(self, reorder_only=False):
        """Atualiza celulas da tabela com dados em _live_data (update in-place via .item())."""
        # reorder_only ignorado: esta funcao sempre faz update in-place, nunca delete+reinsert
        for tid, name in self._live_items:
            d = self._live_data.get(tid)
            if not d:
                continue
            net = d["net_pct"]
            tag = "bull" if net >= 5 else ("bear" if net < 0 else "even")
            score_str = f"{d['score']:.1f}/10"
            self._live_tree.item(str(tid), tags=(tag,), values=(
                name,
                f"{d['sell_min']:,.2f}",
                f"{d['buy_max']:,.2f}",
                f"{d['spread_pct']:.2f}%",
                f"{net:.2f}%",
                f"{int(d['vol_7d']):,}",
                score_str,
            ))
        now = datetime.now().strftime("%H:%M:%S")
        self._live_status.config(text=f"LIVE ● {now}", fg=COLORS["green"])
        self._live_count.config(
            text=f"{len(self._live_items)} itens  |  atualizado {now}")

    def _sort_live(self, col):
        rev = not (self._live_sort_col == col and self._live_sort_rev)
        self._live_sort_col = col
        self._live_sort_rev = rev
        num = {"sell_min","buy_max","spread_pct","net_pct","vol_7d","score"}
        def key(item):
            tid, name = item
            d = self._live_data.get(tid, {})
            if col == "name": return name
            return d.get(col, 0) if col in num else 0
        self._live_items.sort(key=key, reverse=rev)
        # P6 — move() reordena visualmente sem delete+reinsert
        for i, (tid, _) in enumerate(self._live_items):
            self._live_tree.move(str(tid), "", i)
        if self._live_data:
            self._pop_live(reorder_only=True)

    def _live_toggle(self):
        if self._live_running:
            self._live_running = False
            self._live_start_btn.config(text=" ▶  INICIAR LIVE ", bg=COLORS["green"])
            self._live_status.config(text="Parado", fg=COLORS["subtext"])
        else:
            self._live_running = True
            self._live_region  = self._live_reg_var.get()
            self._live_start_btn.config(text=" ■  PARAR LIVE  ", bg=COLORS["red"])
            self._live_status.config(text="Iniciando...", fg=COLORS["yellow"])
            self._live_tick()

    def _live_tick(self):
        # C4 — ler widget no thread principal, passar como argumento
        if not self._live_running:
            return
        if getattr(self, "_live_busy", False):   # M2 — guard contra double-fire
            return
        try:
            interval = max(10, int(self._fread(self._live_interval, 60)))
        except Exception:
            interval = 60
        self._live_busy = True
        threading.Thread(target=self._live_worker,
                         args=(interval,), daemon=True).start()

    def _live_worker(self, interval: int):  # C4 — interval via argumento
        rid = REGIONS.get(self._live_region, 10000002)

        def _fetch_one(item):
            tid, name = item
            d = fetch_live_price(rid, tid)   # usa _get_hist_cached internamente (C1)
            if d:
                d["name"] = name
                self._live_data[tid] = d

        # C3 — max_workers=4 (historico ja vem do cache, so ordens em paralelo)
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
            list(pool.map(_fetch_one, self._live_items))

        self._live_busy = False          # M2 — libera guard
        self.after(0, self._pop_live)
        if self._live_running:
            self._live_after_id = self.after(interval * 1000, self._live_tick)

    def _live_add_item(self):
        try:
            tid  = int(self._live_add_tid.get().strip())
            name = self._live_add_name.get().strip() or str(tid)
        except ValueError:
            messagebox.showerror("Erro", "Type ID deve ser um numero inteiro."); return
        if any(t == tid for t, _ in self._live_items):
            messagebox.showinfo("Ja existe", f"Type ID {tid} ja esta na lista."); return
        self._live_items.append((tid, name))
        self._live_add_tid.delete(0, "end")
        self._live_add_name.delete(0, "end")
        self._pop_live_empty()

    def _live_on_select(self, _event=None):
        """Captura seleção no momento do clique — robusto para Linux/GTK."""
        sel = self._live_tree.selection()
        if sel:
            try:
                self._live_cur_tid = int(sel[0])
            except ValueError as _e:
                log.debug("ValueError silenced: %s", _e)

    def _live_on_dbl(self, _event=None):
        """Double-click: copia Type ID para o clipboard."""
        tid = self._live_get_tid()
        if tid is None: return
        self.clipboard_clear()
        self.clipboard_append(str(tid))
        self._set_status(f"✓ Type ID {tid} copiado")

    def _live_get_tid(self):
        """Retorna o type_id selecionado com fallback para _live_cur_tid."""
        sel = self._live_tree.selection()
        if sel:
            try: return int(sel[0])
            except ValueError: pass
        return self._live_cur_tid

    def _live_remove_item(self):
        tid = self._live_get_tid()
        if tid is None: return
        self._live_items  = [(t, n) for t, n in self._live_items if t != tid]
        self._live_data.pop(tid, None)
        self._live_cur_tid = None
        self._pop_live_empty()

    def _live_open_chart(self):
        """Abre gráfico de candlestick+indicadores para o item selecionado."""
        tid = self._live_get_tid()
        if tid is None:
            messagebox.showinfo("Selecione", "Clique num item primeiro."); return
        if not HAS_MATPLOTLIB:
            messagebox.showinfo("Falta matplotlib", "pip install matplotlib"); return
        name = next((n for t, n in self._live_items if t == tid), str(tid))
        rid  = REGIONS.get(self._live_region, 10000002)
        self._open_candle_chart(tid, name, rid)

    def _live_ai_quick(self):
        """Analisa item selecionado direto pela AI sem trocar de aba."""
        tid = self._live_get_tid()
        if tid is None:
            messagebox.showinfo("Selecione", "Clique num item primeiro."); return
        name = next((n for t, n in self._live_items if t == tid), str(tid))
        d    = self._live_data.get(tid)
        if not d:
            messagebox.showinfo("Sem dados", "Inicie o LIVE primeiro."); return
        key  = self._ai_api_key.get().strip()
        rid  = REGIONS.get(self._live_region, 10000002)

        def worker():
            hist = _get_hist_cached(rid, tid)
            ind  = compute_indicators(hist)
            result = ai_analyze(name, ind, d["sell_min"], d["buy_max"],
                                d["net_pct"], key)
            self.after(0, lambda: messagebox.showinfo(f"AI — {name}", result))

        threading.Thread(target=worker, daemon=True).start()

    def _live_open_forecast(self):
        """Abre aba PREVISAO pre-preenchida com o item selecionado no Mercado Live."""
        tid = self._live_get_tid()
        if tid is None:
            messagebox.showinfo("Selecione", "Clique num item primeiro."); return
        name = next((n for t, n in self._live_items if t == tid), str(tid))
        self._fcast_tid_e.set(str(tid))
        self._fcast_name_e.set(name)
        self._fcast_reg.set(self._live_region)
        self._nb.select(self._tab_fcast)
        self.after(50, self._start_fcast)

    def _live_add_to_wl(self):
        tid = self._live_get_tid()
        if tid is None:
            messagebox.showinfo("Selecione", "Clique num item primeiro."); return
        if any(w["type_id"] == tid for w in self._watchlist):
            messagebox.showinfo("Ja existe", "Item ja esta na watchlist."); return
        name = next((n for t, n in self._live_items if t == tid), str(tid))
        self._watchlist.append({
            "type_id": tid, "name": name,
            "alert_below": None, "alert_above": None,
            "last_sell": None, "last_buy": None,
            "last_net": None, "last_score": None,
        })
        save_watchlist(self._watchlist)
        self._pop_wl()
        messagebox.showinfo("Adicionado", f"{name} adicionado à Watchlist.")
    # ══════════════════════════════════════════════════════════════════════════
    #  ABA 10 — WATCHLIST
    # ══════════════════════════════════════════════════════════════════════════

    _WL_COLS = {
        "name":        "Item",
        "last_sell":   "Sell (ISK)",
        "last_buy":    "Buy (ISK)",
        "last_net":    "Margem Líq %",
        "last_score":  "Score",
        "alert_below": "Alerta ↓",
        "alert_above": "Alerta ↑",
    }

    def _build_tab_wl(self):
        p = self._tab_wl

        top = tk.Frame(p, bg=COLORS["panel"],
                       highlightbackground=COLORS["border"], highlightthickness=1)
        top.pack(fill="x", padx=6, pady=(6, 2))

        regions = list(REGIONS.keys())
        self._lbl(top, "Regiao",      0, 0)
        self._wl_reg_var = self._combo(top, regions, 1, 0)
        self._lbl(top, "Intervalo(s)",0, 2)
        self._wl_interval = self._entry(top, "120", 1, 2, 5)

        self._wl_start_btn = tk.Button(
            top, text=" ▶  MONITORAR ",
            command=self._wl_toggle,
            bg=COLORS["accent"], fg="#000",
            font=("Courier", 11, "bold"), relief="flat", cursor="hand2")
        self._wl_start_btn.grid(row=1, column=4, padx=8)

        self._wl_status = tk.Label(top, text="Parado",
                                    bg=COLORS["panel"], fg=COLORS["subtext"],
                                    font=("Courier", 10))
        self._wl_status.grid(row=1, column=5, padx=10)

        # Alertas
        alert_bar = tk.Frame(p, bg=COLORS["bg"])
        alert_bar.pack(fill="x", padx=6, pady=(2,0))
        tk.Label(alert_bar, text="Alerta ↓:", bg=COLORS["bg"],
                 fg=COLORS["subtext"], font=("Courier", 10)).pack(side="left")
        self._wl_ab = tk.Entry(alert_bar, width=14,
                                bg=COLORS["border"], fg=COLORS["text"],
                                insertbackground=COLORS["text"],
                                font=("Courier", 11), relief="flat", bd=2)
        self._wl_ab.pack(side="left", padx=2)
        tk.Label(alert_bar, text="Alerta ↑:", bg=COLORS["bg"],
                 fg=COLORS["subtext"], font=("Courier", 10)).pack(side="left")
        self._wl_aa = tk.Entry(alert_bar, width=14,
                                bg=COLORS["border"], fg=COLORS["text"],
                                insertbackground=COLORS["text"],
                                font=("Courier", 11), relief="flat", bd=2)
        self._wl_aa.pack(side="left", padx=2)
        tk.Button(alert_bar, text="SET ALERTAS",
                  command=self._wl_set_alerts,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)
        tk.Button(alert_bar, text="REMOVER",
                  command=self._wl_remove,
                  bg=COLORS["border"], fg=COLORS["red"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=2)
        tk.Button(alert_bar, text="GRÁFICO",
                  command=self._wl_chart,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=2)
        tk.Button(alert_bar, text="PREVER",
                  command=self._wl_open_forecast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=2)
        tk.Button(alert_bar, text="BACKTEST",
                  command=self._wl_open_backtest,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=2)
        tk.Button(alert_bar, text="ANALISAR IA",
                  command=self._wl_open_ai,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=2)

        self._wl_tree = self._make_tree(p, [
            ("name",        "Item",          200),
            ("last_sell",   "Sell (ISK)",    120),
            ("last_buy",    "Buy (ISK)",     120),
            ("last_net",    "Margem Líq %",   90),
            ("last_score",  "Score",           65),
            ("alert_below", "Alerta ↓",      110),
            ("alert_above", "Alerta ↑",      110),
        ])
        # Armazena o item selecionado em _wl_cur_tid para evitar perda de seleção
        # quando o foco muda para o botão (comportamento Linux/GTK)
        self._wl_tree.bind("<<TreeviewSelect>>", self._wl_on_select)
        self._wl_tree.bind("<Double-1>",         self._wl_on_dbl)
        self._pop_wl()

    def _pop_wl(self):
        for row in self._wl_tree.get_children():
            self._wl_tree.delete(row)
        for i, w in enumerate(self._watchlist):
            tag = "even" if i % 2 == 0 else "odd"
            sell = f"{w['last_sell']:,.2f}"  if w.get("last_sell") else "--"
            buy  = f"{w['last_buy']:,.2f}"   if w.get("last_buy")  else "--"
            net  = f"{w['last_net']:.2f}%"   if w.get("last_net") is not None else "--"
            sc   = f"{w['last_score']:.1f}"  if w.get("last_score") is not None else "--"
            ab   = f"{w['alert_below']:,.0f}" if w.get("alert_below") else "--"
            aa   = f"{w['alert_above']:,.0f}" if w.get("alert_above") else "--"
            self._wl_tree.insert("", "end", iid=str(w["type_id"]),
                                  tags=(tag,), values=(
                w["name"], sell, buy, net, sc, ab, aa))

    def _wl_toggle(self):
        if self._wl_running:
            self._wl_running = False
            self._wl_start_btn.config(text=" ▶  MONITORAR ", bg=COLORS["accent"])
            self._wl_status.config(text="Parado", fg=COLORS["subtext"])
        else:
            self._wl_running = True
            self._wl_start_btn.config(text=" ■  PARAR     ", bg=COLORS["red"])
            self._wl_status.config(text="Monitorando...", fg=COLORS["yellow"])
            self._wl_tick()

    def _wl_tick(self):
        if not self._wl_running: return
        # M3/C4 — ler widget no thread principal, passar como argumento
        try:
            interval = max(10, int(self._fread(self._wl_interval, 120)))
        except Exception:
            interval = 120
        rid = REGIONS.get(self._wl_reg_var.get(), 10000002)
        threading.Thread(target=self._wl_worker,
                         args=(interval, rid), daemon=True).start()

    def _wl_worker(self, interval: int, rid: int):   # C4 — args via parametro
        alerts = check_watchlist_alerts(self._watchlist, rid)  # C2 — paralelo
        self.after(0, self._pop_wl)
        if alerts:
            msgs = "\n".join(
                f"⚠ {a['name']} {a['kind']} {a['threshold']:,.0f} ISK  →  atual: {a['current_price']:,.2f}"
                for a in alerts)
            self.after(0, lambda m=msgs: messagebox.showwarning("Alertas de Preço!", m))
            # Tier 2 — registra no historico de alertas e notifica desktop
            ts = datetime.now().strftime("%H:%M")
            for a in alerts:
                entry = (ts, f"{a['name']} {a['kind']} {a['threshold']:,.0f} ISK")
                self._alert_history.append(entry)
                if len(self._alert_history) > 50:
                    self._alert_history = self._alert_history[-50:]
            # Notificacao desktop (nao bloqueia)
            notif_body = "\n".join(
                f"{a['name']}: {a['current_price']:,.0f} ISK" for a in alerts[:5])
            threading.Thread(
                target=_desktop_notify,
                args=("EVE Trade — Alerta de Preco", notif_body),
                daemon=True).start()
            # Atualiza dashboard se estiver visivel
            self.after(0, self._dash_refresh)
        now = datetime.now().strftime("%H:%M:%S")
        self.after(0, lambda: self._wl_status.config(
            text=f"● {now}", fg=COLORS["green"]))
        if self._wl_running:
            self._wl_after_id = self.after(interval * 1000, self._wl_tick)

    def _wl_set_alerts(self):
        w = self._wl_get_selected()
        if not w:
            messagebox.showinfo("Selecione", "Clique em um item da lista primeiro."); return
        tid = w["type_id"]
        try:
            ab_txt = self._wl_ab.get().strip().replace(",","")
            aa_txt = self._wl_aa.get().strip().replace(",","")
            w["alert_below"] = float(ab_txt) if ab_txt else None
            w["alert_above"] = float(aa_txt) if aa_txt else None
        except ValueError:
            messagebox.showerror("Erro", "Valor invalido. Use numeros."); return
        save_watchlist(self._watchlist)
        self._pop_wl()

    def _wl_remove(self):
        w = self._wl_get_selected()
        if not w: return
        tid = w["type_id"]
        self._watchlist = [x for x in self._watchlist if x["type_id"] != tid]
        self._wl_cur_tid = None
        save_watchlist(self._watchlist)
        self._pop_wl()

    def _wl_chart(self):
        w = self._wl_get_selected()
        if not w:
            messagebox.showinfo("Selecione", "Clique em um item da lista primeiro."); return
        if not HAS_MATPLOTLIB:
            messagebox.showinfo("Falta matplotlib", "pip install matplotlib"); return
        tid = w["type_id"]
        rid = REGIONS.get(self._wl_reg_var.get(), 10000002)
        self._open_candle_chart(tid, w["name"], rid)

    def _wl_on_select(self, _event=None):
        """Captura o item selecionado no momento do clique na tree.
        Armazena em _wl_cur_tid para uso pelos botões mesmo após perda de foco."""
        sel = self._wl_tree.selection()
        if sel:
            try:
                self._wl_cur_tid = int(sel[0])
            except ValueError as _e:
                log.debug("ValueError silenced: %s", _e)

    def _wl_on_dbl(self, _event=None):
        """Double-click na Watchlist: copia o Type ID para o clipboard."""
        w = self._wl_get_selected()
        if not w: return
        self.clipboard_clear()
        self.clipboard_append(str(w["type_id"]))
        self._set_status(f"✓ Type ID {w['type_id']} copiado — {w['name']}")

    def _wl_get_selected(self):
        """Retorna o dict do item selecionado da watchlist.
        Usa _wl_cur_tid (armazenado no clique) como fallback robusto
        quando selection() retorna vazio (perda de foco no Linux/GTK)."""
        # Tenta selection() primeiro
        sel = self._wl_tree.selection()
        tid = None
        if sel:
            try:
                tid = int(sel[0])
            except ValueError as _e:
                log.debug("ValueError silenced: %s", _e)
        # Fallback: usa o último item clicado
        if tid is None:
            tid = self._wl_cur_tid
        if tid is None:
            return None
        return next((x for x in self._watchlist if x["type_id"] == tid), None)

    def _wl_open_forecast(self):
        """Abre aba PREVISAO pre-preenchida com o item selecionado na Watchlist."""
        w = self._wl_get_selected()
        if not w:
            messagebox.showinfo("Selecione", "Clique em um item da lista primeiro."); return
        if self._loading_fcast:
            messagebox.showinfo("Aguarde", "Previsão já está em execução."); return
        self._fcast_tid_e.set(str(w["type_id"]))
        self._fcast_name_e.set(w["name"])
        self._fcast_reg.set(self._wl_reg_var.get())
        self._nb.select(self._tab_fcast)
        self.after(50, self._start_fcast)

    def _wl_open_backtest(self):
        """Abre aba BACKTESTING pre-preenchida com o item selecionado na Watchlist."""
        w = self._wl_get_selected()
        if not w:
            messagebox.showinfo("Selecione", "Clique em um item da lista primeiro."); return
        if self._loading_bt:
            messagebox.showinfo("Aguarde", "Backtesting já está em execução."); return
        self._bt_tid.set(str(w["type_id"]))
        self._bt_name.set(w["name"])
        self._bt_reg.set(self._wl_reg_var.get())
        self._nb.select(self._tab_bt)
        self.after(50, self._start_bt)

    def _wl_open_ai(self):
        """Abre aba AI ANALISTA pre-preenchida com o item selecionado na Watchlist."""
        w = self._wl_get_selected()
        if not w:
            messagebox.showinfo("Selecione", "Clique em um item da lista primeiro."); return
        if self._loading_ai:
            messagebox.showinfo("Aguarde", "AI Analista já está em execução."); return
        self._ai_tid.set(str(w["type_id"]))
        self._ai_name_e.set(w["name"])
        self._ai_reg.set(self._wl_reg_var.get())
        self._nb.select(self._tab_ai)
        self.after(50, self._start_ai)

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA 11 — PAPER TRADING
    # ══════════════════════════════════════════════════════════════════════════

    _PT_POS_COLS = {
        "name":       "Item",
        "qty":        "Qtd",
        "buy_price":  "Compra (ISK)",
        "broker_fee": "Taxa Corret.",
        "total_cost": "Custo Total",
        "opened_at":  "Aberto em",
    }
    _PT_HIST_COLS = {
        "name":      "Item",
        "qty":       "Qtd",
        "buy_price": "Compra",
        "sell_price":"Venda",
        "pnl":       "P&L (ISK)",
        "pnl_pct":   "P&L %",
        "sales_tax": "Imposto",
        "closed_at": "Fechado em",
    }

    def _build_tab_paper(self):
        p = self._tab_paper

        # Barra de controle
        top = tk.Frame(p, bg=COLORS["panel"],
                       highlightbackground=COLORS["border"], highlightthickness=1)
        top.pack(fill="x", padx=6, pady=(6, 2))

        self._lbl(top, "Type ID",       0, 0)
        self._pt_tid   = self._entry(top, "",          1, 0, 10)
        self._lbl(top, "Nome",          0, 2)
        self._pt_name  = self._entry(top, "",          1, 2, 18)
        self._lbl(top, "Qtd",           0, 4)
        self._pt_qty   = self._entry(top, "1",         1, 4, 7)
        self._lbl(top, "Preco Compra",  0, 6)
        self._pt_buy   = self._entry(top, "",          1, 6, 14)
        # Skills globais (lidos via _gskills())
        self._lbl(top, "Skills",        0, 8)
        tk.Label(top, textvariable=self._g_brel,
                 bg=COLORS["border"], fg=COLORS["accent"],
                 font=("Courier", 11, "bold"), width=3, relief="flat"
                 ).grid(row=1, column=8, padx=2, pady=2)
        tk.Label(top, text="/", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 11)
                 ).grid(row=1, column=9, padx=0)
        tk.Label(top, textvariable=self._g_acc,
                 bg=COLORS["border"], fg=COLORS["accent"],
                 font=("Courier", 11, "bold"), width=3, relief="flat"
                 ).grid(row=1, column=10, padx=2, pady=2)
        tk.Label(top, text="(global)", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 9)
                 ).grid(row=0, column=9, columnspan=3, padx=2)

        tk.Button(top, text=" COMPRAR (ABRIR) ",
                  command=self._paper_buy,
                  bg=COLORS["green"], fg="#000",
                  font=("Courier", 11, "bold"), relief="flat", cursor="hand2"
                  ).grid(row=1, column=12, padx=8)

        self._lbl(top, "Preco Venda",   0, 14)
        self._pt_sell  = self._entry(top, "",          1, 14, 14)
        tk.Button(top, text=" VENDER (FECHAR) ",
                  command=self._paper_sell,
                  bg=COLORS["red"], fg="#fff",
                  font=("Courier", 11, "bold"), relief="flat", cursor="hand2"
                  ).grid(row=1, column=16, padx=6)

        tk.Button(top, text=" RESETAR CONTA ",
                  command=self._paper_reset,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).grid(row=1, column=17, padx=4)
        tk.Button(top, text=" P&L GRAFICO ",
                  command=self._paper_pnl_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2"
                  ).grid(row=1, column=18, padx=4)

        # Barra de saldo
        self._pt_balance_lbl = tk.Label(p, text="",
                                         bg=COLORS["panel"], fg=COLORS["green"],
                                         font=("Courier", 12, "bold"))
        self._pt_balance_lbl.pack(fill="x", padx=6, pady=(2, 0))

        # Posicoes abertas
        tk.Label(p, text="  POSIÇÕES ABERTAS",
                 bg=COLORS["bg"], fg=COLORS["header"],
                 font=("Courier", 11, "bold")).pack(anchor="w", padx=6)
        self._pt_pos_tree = self._make_tree(p, [
            ("name",       "Item",          200),
            ("qty",        "Qtd",            58),
            ("buy_price",  "Compra (ISK)",  130),
            ("broker_fee", "Taxa Corret.",  110),
            ("total_cost", "Custo Total",   130),
            ("opened_at",  "Aberto em",     140),
        ])

        # Historico
        tk.Label(p, text="  HISTÓRICO DE TRADES",
                 bg=COLORS["bg"], fg=COLORS["header"],
                 font=("Courier", 11, "bold")).pack(anchor="w", padx=6, pady=(4,0))
        self._pt_hist_tree = self._make_tree(p, [
            ("name",       "Item",          180),
            ("qty",        "Qtd",            58),
            ("buy_price",  "Compra",        110),
            ("sell_price", "Venda",         110),
            ("pnl",        "P&L (ISK)",     120),
            ("pnl_pct",    "P&L %",          75),
            ("sales_tax",  "Imposto",        95),
            ("closed_at",  "Fechado em",    140),
        ])
        self._pt_hist_tree.tag_configure("profit", foreground=COLORS["green"])
        self._pt_hist_tree.tag_configure("loss",   foreground=COLORS["red"])

        self._pop_paper()

        # Auto-lookup nome pelo Type ID
        self._bind_tid_lookup(self._pt_tid, self._pt_name, None)

    def _pop_paper(self):
        bal = self._paper.get("balance", 1e9)
        self._pt_balance_lbl.config(
            text=f"  Saldo disponível: {bal:,.2f} ISK")

        for row in self._pt_pos_tree.get_children():
            self._pt_pos_tree.delete(row)
        for i, pos in enumerate(self._paper.get("positions", [])):
            tag = "even" if i % 2 == 0 else "odd"
            self._pt_pos_tree.insert("", "end", iid=str(pos["id"]),
                                      tags=(tag,), values=(
                pos["name"],
                f"{pos['qty']:,}",
                f"{pos['buy_price']:,.2f}",
                f"{pos['broker_fee']:,.2f}",
                f"{pos['total_cost']:,.2f}",
                pos["opened_at"][:16],
            ))

        for row in self._pt_hist_tree.get_children():
            self._pt_hist_tree.delete(row)
        for i, h in enumerate(reversed(self._paper.get("history", []))):
            pnl = h.get("pnl", 0) or 0
            tag = "profit" if pnl >= 0 else "loss"
            self._pt_hist_tree.insert("", "end", iid=f"h{h['id']}",
                                       tags=(tag,), values=(
                h["name"],
                f"{h['qty']:,}",
                f"{h['buy_price']:,.2f}",
                f"{h.get('sell_price',0):,.2f}",
                f"{pnl:+,.2f}",
                f"{h.get('pnl_pct',0):+.2f}%",
                f"{h.get('sales_tax',0):,.2f}",
                h.get("closed_at","--")[:16],
            ))

    def _paper_buy(self):
        try:
            tid  = int(self._pt_tid.get().strip() or "0")
            name = self._pt_name.get().strip() or f"Item {tid}"
            qty  = int(self._pt_qty.get().strip())
            bp   = float(self._pt_buy.get().strip().replace(",",""))
            brel, acc = self._gskills()
        except ValueError:
            messagebox.showerror("Erro", "Preencha todos os campos corretamente."); return
        result = paper_open(self._paper, tid, name, qty, bp, brel, acc)
        if not result["ok"]:
            messagebox.showerror("Erro", result["msg"]); return
        pos = result["pos"]
        messagebox.showinfo("Posição Aberta",
            f"✓ {name} x{qty} comprado por {bp:,.2f} ISK\n"
            f"Custo total: {pos['total_cost']:,.2f} ISK\n"
            f"Saldo restante: {self._paper['balance']:,.2f} ISK")
        self._pop_paper()

    def _paper_sell(self):
        sel = self._pt_pos_tree.selection()
        if not sel:
            messagebox.showinfo("Selecione", "Clique numa posição aberta."); return
        try:
            pos_id = int(sel[0])
            sp     = float(self._pt_sell.get().strip().replace(",",""))
            brel, acc = self._gskills()
        except ValueError:
            messagebox.showerror("Erro", "Preencha o Preco Venda corretamente."); return
        result = paper_close(self._paper, pos_id, sp, brel, acc)
        if not result["ok"]:
            messagebox.showerror("Erro", result["msg"]); return
        pnl = result["pnl"]
        color_txt = "LUCRO" if pnl >= 0 else "PREJUIZO"
        messagebox.showinfo("Posição Fechada",
            f"{color_txt}: {pnl:+,.2f} ISK  ({result['pnl_pct']:+.2f}%)\n"
            f"Saldo atual: {self._paper['balance']:,.2f} ISK")
        self._pop_paper()

    def _paper_reset(self):
        if not messagebox.askyesno("Resetar", "Zerar conta de paper trading?"):
            return
        self._paper = {"positions": [], "history": [], "balance": 1_000_000_000.0,
                       "next_id": 1}
        save_paper_trades(self._paper)
        self._pop_paper()

    def _paper_pnl_chart(self):
        """
        Abre janela matplotlib com:
          - Curva de equity acumulada (saldo ao longo do tempo)
          - Barras de P&L por trade (verde=lucro, vermelho=prejuizo)
          - Estatisticas: win rate, drawdown maximo, melhor/pior trade
        """
        if not HAS_MATPLOTLIB:
            messagebox.showinfo("Falta matplotlib", "pip install matplotlib"); return
        hist = self._paper.get("history", [])
        if not hist:
            messagebox.showinfo("Sem historico",
                "Feche pelo menos uma posicao para ver o grafico P&L."); return

        # Ordena por data de fechamento
        trades = sorted(hist, key=lambda h: h.get("closed_at", ""))
        pnls   = [t.get("pnl", 0) or 0 for t in trades]
        labels = [t.get("name", f"#{t.get('id',i)}")[:14]
                  for i, t in enumerate(trades)]
        dates  = [t.get("closed_at","")[:10] for t in trades]

        # Curva de equity: começa em 1 bilhão e aplica cada P&L
        initial = 1_000_000_000.0
        equity  = []
        running = initial
        for p in pnls:
            running += p
            equity.append(running)

        # Drawdown maximo
        peak    = initial
        max_dd  = 0.0
        for e in equity:
            if e > peak:
                peak = e
            dd = (peak - e) / peak * 100 if peak > 0 else 0
            if dd > max_dd:
                max_dd = dd

        wins     = sum(1 for p in pnls if p >= 0)
        losses   = len(pnls) - wins
        win_rate = wins / len(pnls) * 100 if pnls else 0
        best     = max(pnls)
        worst    = min(pnls)
        total_pnl= sum(pnls)

        if self._chart_win and self._chart_win.winfo_exists():
            self._chart_win.destroy()
        win = tk.Toplevel(self)
        win.title("Paper Trading — P&L Acumulado")
        win.configure(bg=COLORS["bg"])
        win.geometry("1000x640")
        self._chart_win = win

        bg  = COLORS["bg"]
        pan = COLORS["panel"]

        fig = Figure(figsize=(12, 7.5), facecolor=bg)
        fig.subplots_adjust(left=0.08, right=0.96, hspace=0.35,
                            top=0.88, bottom=0.12)

        xs = list(range(len(trades)))

        # Subplot 1 — Curva de equity
        ax1 = fig.add_subplot(2, 1, 1)
        ax1.set_facecolor(pan)
        ax1.tick_params(colors=COLORS["subtext"], labelsize=7)
        for sp in ax1.spines.values():
            sp.set_edgecolor(COLORS["border"])

        ax1.axhline(initial, color=COLORS["subtext"], lw=0.5, ls=":")
        ax1.plot(xs, equity, color=COLORS["accent"], lw=2, zorder=3)
        ax1.fill_between(xs, equity, initial,
                         where=[e >= initial for e in equity],
                         alpha=0.2, color=COLORS["green"])
        ax1.fill_between(xs, equity, initial,
                         where=[e < initial for e in equity],
                         alpha=0.2, color=COLORS["red"])
        ax1.set_ylabel("Saldo (ISK)", color=COLORS["subtext"], fontsize=8)
        ax1.set_title(
            f"Equity  |  {len(trades)} trades  |  Win rate: {win_rate:.1f}%  |  "
            f"Max drawdown: {max_dd:.2f}%  |  P&L total: {total_pnl:+,.0f} ISK",
            color=COLORS["header"], fontsize=9, pad=6)
        ax1.tick_params(labelbottom=False)

        # Marcador do ponto atual
        if equity:
            ax1.scatter([xs[-1]], [equity[-1]], color=COLORS["accent"],
                        s=40, zorder=5)

        # Subplot 2 — Barras de P&L por trade
        ax2 = fig.add_subplot(2, 1, 2)
        ax2.set_facecolor(pan)
        ax2.tick_params(colors=COLORS["subtext"], labelsize=7)
        for sp in ax2.spines.values():
            sp.set_edgecolor(COLORS["border"])

        colors = [COLORS["green"] if p >= 0 else COLORS["red"] for p in pnls]
        ax2.bar(xs, pnls, color=colors, alpha=0.8, width=0.7)
        ax2.axhline(0, color=COLORS["subtext"], lw=0.5)
        ax2.set_ylabel("P&L (ISK)", color=COLORS["subtext"], fontsize=8)
        ax2.set_title(
            f"P&L por trade  |  Melhor: {best:+,.0f}  |  "
            f"Pior: {worst:+,.0f}  |  Vitórias: {wins}  |  Derrotas: {losses}",
            color=COLORS["header"], fontsize=9, pad=6)

        # Eixo X com nomes dos itens
        step = max(1, len(xs) // 12)
        ax2.set_xticks(xs[::step])
        ax2.set_xticklabels(
            [f"{labels[i]}\n{dates[i]}" for i in xs[::step]],
            rotation=30, ha="right", fontsize=6, color=COLORS["subtext"])

        canvas = FigureCanvasTkAgg(fig, master=win)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA 12 — AI ANALISTA
    # ══════════════════════════════════════════════════════════════════════════

    def _build_tab_ai(self):
        p = self._tab_ai

        # Configuracao da API Key
        key_fr = tk.Frame(p, bg=COLORS["panel"],
                          highlightbackground=COLORS["border"], highlightthickness=1)
        key_fr.pack(fill="x", padx=6, pady=(6,2))

        tk.Label(key_fr, text="Anthropic API Key:",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 11)).grid(row=0, column=0, padx=6, pady=4, sticky="w")
        tk.Entry(key_fr, textvariable=self._ai_api_key, show="●",
                 width=52, bg=COLORS["border"], fg=COLORS["text"],
                 insertbackground=COLORS["text"],
                 font=("Courier", 11), relief="flat", bd=2
                 ).grid(row=0, column=1, padx=4, pady=4, sticky="w")
        tk.Label(key_fr,
                 text=f"({api_key_storage_status()})",
                 bg=COLORS["panel"],
                 fg=COLORS["subtext"] if HAS_KEYRING or os.environ.get(API_KEY_ENV)
                    else COLORS["yellow"],
                 font=("Courier", 9)).grid(row=0, column=2, padx=8)

        # Item a analisar
        item_fr = tk.Frame(p, bg=COLORS["panel"],
                           highlightbackground=COLORS["border"], highlightthickness=1)
        item_fr.pack(fill="x", padx=6, pady=(2,4))

        self._lbl(item_fr, "Regiao",      0, 0)
        self._ai_reg = self._combo(item_fr, list(REGIONS.keys()), 1, 0)
        self._lbl(item_fr, "Type ID",     0, 2)
        self._ai_tid = self._entry(item_fr, "", 1, 2, 10)
        self._lbl(item_fr, "Nome",        0, 4)
        self._ai_name_e = self._entry(item_fr, "", 1, 4, 20)
        self._lbl(item_fr, "Broker Rel",  0, 6)
        self._ai_brel = self._entry(item_fr, "0", 1, 6, 4)
        self._lbl(item_fr, "Accounting",  0, 8)
        self._ai_acc  = self._entry(item_fr, "0", 1, 8, 4)

        self._ai_btn = tk.Button(item_fr, text="  ANALISAR COM IA  ",
                                  command=self._start_ai,
                                  bg=COLORS["orange"], fg="#000",
                                  font=("Courier", 11, "bold"),
                                  relief="flat", cursor="hand2")
        self._ai_btn.grid(row=1, column=10, padx=10)

        tk.Button(item_fr, text=" GRÁFICO ",
                  command=self._ai_open_chart,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=11, padx=4)

        self._ai_spin = ttk.Progressbar(item_fr, mode="indeterminate", length=80)
        self._ai_spin.grid(row=1, column=12, padx=6)

        # Painel de resultado
        res_fr = tk.Frame(p, bg=COLORS["panel"],
                          highlightbackground=COLORS["border"], highlightthickness=1)
        res_fr.pack(fill="both", expand=True, padx=6, pady=(0,4))

        # Header com indicadores numericos
        ind_top = tk.Frame(res_fr, bg=COLORS["panel"])
        ind_top.pack(fill="x", padx=8, pady=(6,2))

        self._ai_ind_labels = {}
        for col, label in enumerate(
                ["Sell Min", "Buy Max", "Spread%", "Margem Líq%",
                 "RSI(14)", "MACD", "SMA7", "SMA26"]):
            tk.Label(ind_top, text=label, bg=COLORS["panel"],
                     fg=COLORS["subtext"], font=("Courier", 9)).grid(
                row=0, column=col, padx=10)
            v = tk.Label(ind_top, text="--", bg=COLORS["panel"],
                         fg=COLORS["accent"], font=("Courier", 12, "bold"))
            v.grid(row=1, column=col, padx=10)
            self._ai_ind_labels[label] = v

        # Recomendacao em destaque
        rec_fr = tk.Frame(res_fr, bg=COLORS["bg"])
        rec_fr.pack(fill="x", padx=8, pady=(4,2))
        self._ai_rec_lbl = tk.Label(rec_fr, text="—",
                                     bg=COLORS["bg"], fg=COLORS["yellow"],
                                     font=("Courier", 20, "bold"))
        self._ai_rec_lbl.pack(side="left", padx=6)

        # Texto completo da analise
        self._ai_text = tk.Text(res_fr, bg=COLORS["panel"],
                                 fg=COLORS["text"], font=("Courier", 11),
                                 wrap="word", relief="flat",
                                 state="disabled", height=10)
        self._ai_text.pack(fill="both", expand=True, padx=8, pady=(0,6))

        # Tags de cor para o texto
        self._ai_text.tag_configure("buy",    foreground=COLORS["green"],
                                              font=("Courier", 11, "bold"))
        self._ai_text.tag_configure("sell",   foreground=COLORS["red"],
                                              font=("Courier", 11, "bold"))
        self._ai_text.tag_configure("wait",   foreground=COLORS["yellow"],
                                              font=("Courier", 11, "bold"))
        self._ai_text.tag_configure("header", foreground=COLORS["accent"],
                                              font=("Courier", 11, "bold"))

        # Auto-lookup nome pelo Type ID (busca no cache local + ESI)
        self._bind_tid_lookup(self._ai_tid, self._ai_name_e, self._ai_reg)

    def _start_ai(self):
        if not self._experimental_ok(): return
        if self._loading_ai: return
        try:
            tid  = int(self._ai_tid.get().strip())
        except ValueError:
            messagebox.showerror("Erro", "Type ID invalido."); return
        name = self._ai_name_e.get().strip() or str(tid)
        # C2 — lê widgets no thread principal
        params = dict(
            rname = self._ai_reg.get(),
            brel  = int(self._fread(self._ai_brel, 0)),
            acc   = int(self._fread(self._ai_acc,  0)),
            key   = self._ai_api_key.get().strip(),
        )
        self._loading_ai = True
        self._ai_btn.config(state="disabled")
        self._ai_spin.start(10)
        threading.Thread(target=self._worker_ai,
                         args=(tid, name, params), daemon=True).start()

    def _worker_ai(self, tid: int, name: str, params: dict):
        """
        O2 — reutiliza _live_data[tid] se disponivel (evita fetch_live_price extra).
        A6 — usa _ind_last() em vez de filtrar listas de indicadores duas vezes.
        O6 — streaming: tokens aparecem no widget a medida que chegam da API.
        """
        try:
            rname = params["rname"]
            rid   = REGIONS.get(rname, 10000002)
            brel  = params["brel"]
            acc   = params["acc"]

            _warm_hist_cache(rid)  # pré-aquece cache RAM: 1 query SQL → O(1) no _get_hist_cached
            # O2 — reusa dado do Live se disponivel e recente (< 5 min)
            cached = self._live_data.get(tid)
            if cached:
                live = cached
                hist = _get_hist_cached(rid, tid)   # O1 — cache de historico
            else:
                # Busca preco e historico em paralelo
                with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
                    f_live = pool.submit(fetch_live_price, rid, tid)
                    f_hist = pool.submit(_get_hist_cached, rid, tid)
                    live   = f_live.result()
                    hist   = f_hist.result()

            ind = compute_indicators(hist)

            sell_min = live["sell_min"] if live else 0
            buy_max  = live["buy_max"]  if live else 0
            rb, rt   = _calc_fee_rates(brel, acc)
            cost     = buy_max + buy_max*rb + sell_min*rb + sell_min*rt
            net_pct  = (sell_min - cost) / cost * 100 if cost > 0 else 0

            # A6 — usa _ind_last() em vez de re-filtrar listas
            closes = ind.get("closes", [])
            last_rsi  = _ind_last(ind, "rsi")       or 0.0
            last_macd = _ind_last(ind, "macd_line") or 0.0
            last_sm7  = _ind_last(ind, "sma7")      or 0.0
            last_sm26 = _ind_last(ind, "sma26")     or 0.0

            nums = {
                "Sell Min":    f"{sell_min:,.2f}" if sell_min else "--",
                "Buy Max":     f"{buy_max:,.2f}"  if buy_max  else "--",
                "Spread%":     f"{((sell_min-buy_max)/sell_min*100):.2f}%" if sell_min else "--",
                "Margem Líq%": f"{net_pct:.2f}%",
                "RSI(14)":     f"{last_rsi:.1f}",
                "MACD":        f"{last_macd:.4f}",
                "SMA7":        f"{last_sm7:,.2f}"  if last_sm7  else "--",
                "SMA26":       f"{last_sm26:,.2f}" if last_sm26 else "--",
            }

            rsi_col = COLORS["text"]
            if last_rsi > 70:   rsi_col = COLORS["red"]
            elif last_rsi < 30: rsi_col = COLORS["green"]

            self.after(0, lambda: self._show_ai_result("", nums, rsi_col))
            self.after(0, lambda: self._ai_rec_lbl.config(
                text="Analisando...", fg=COLORS["yellow"]))

            key = params["key"]

            # O6 — streaming: atualiza Text widget token a token
            full_text = []

            def _stream_cb(token: str):
                full_text.append(token)
                # Atualiza UI no thread principal
                self.after(0, lambda t=token: self._append_ai_token(t))

            result = ai_analyze(name, ind, sell_min, buy_max, net_pct,
                                key, stream_cb=_stream_cb)

            if not result:            # streaming acumulou em full_text
                result = "".join(full_text)

            self._ai_last = {
                "tid": tid, "name": name, "rid": rid,
                "ind": ind, "result": result,
            }
            # Atualiza recomendacao em destaque apos texto completo
            self.after(0, lambda: self._finalize_ai_rec(result, rsi_col))

        except Exception as exc:
            log.exception("worker_ai")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro AI", e))
        finally:
            self._loading_ai = False
            self.after(0, lambda: (self._ai_btn.config(state="normal"),
                                   self._ai_spin.stop()))

    def _append_ai_token(self, token: str):
        """O6 — insere token de streaming no widget de texto da AI."""
        self._ai_text.config(state="normal")
        self._ai_text.insert("end", token)
        self._ai_text.see("end")
        self._ai_text.config(state="disabled")

    def _show_ai_result(self, text: str, nums: dict, rsi_col: str):
        """Atualiza painel de indicadores numericos. Se text nao vazio, renderiza texto."""
        # Atualiza labels de indicadores
        for label, val in nums.items():
            lbl = self._ai_ind_labels.get(label)
            if lbl:
                col = rsi_col if label == "RSI(14)" else COLORS["accent"]
                lbl.config(text=val, fg=col)
        # Renderiza texto se fornecido (modo sem streaming)
        if text:
            self._finalize_ai_rec(text, rsi_col)

    def _finalize_ai_rec(self, text: str, rsi_col: str):
        """Apos streaming completo, re-renderiza com coloracao e extrai REC."""
        # Recolore o texto completo
        self._ai_text.config(state="normal")
        self._ai_text.delete("1.0", "end")
        for line in text.splitlines():
            ul = line.upper()
            if ul.startswith("RECOMENDACAO:"):
                self._ai_text.insert("end", line + "\n", "header")
            elif "COMPRAR" in ul:
                self._ai_text.insert("end", line + "\n", "buy")
            elif "VENDER" in ul:
                self._ai_text.insert("end", line + "\n", "sell")
            elif "ESPERAR" in ul:
                self._ai_text.insert("end", line + "\n", "wait")
            else:
                self._ai_text.insert("end", line + "\n")
        self._ai_text.config(state="disabled")
        # Atualiza label de recomendacao
        rec = "—"
        for line in text.splitlines():
            if line.upper().startswith("RECOMENDACAO:"):
                rec = line.split(":", 1)[1].strip().upper()
                break
        rec_col = (COLORS["green"] if "COMPRAR" in rec
                   else (COLORS["red"] if "VENDER" in rec
                         else COLORS["yellow"]))
        self._ai_rec_lbl.config(text=rec, fg=rec_col)

    def _ai_open_chart(self):
        if not HAS_MATPLOTLIB:
            messagebox.showinfo("Falta matplotlib", "pip install matplotlib"); return
        last = self._ai_last
        if not last:
            messagebox.showinfo("Sem dados", "Execute uma analise primeiro."); return
        self._open_candle_chart(last["tid"], last["name"], last["rid"])

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA 4 — MAX MARGEM
    # ══════════════════════════════════════════════════════════════════════════

    _MM_COLS = {
        "item":         "Item",
        "ps":           "Sell (ISK)",
        "pb":           "Buy (ISK)",
        "gross_margin": "Margem Bruta %",
        "net_margin":   "Margem Liq %",
        "profit_unit":  "Lucro/un (ISK)",
        "avg_vol_90":   "Vol/Dia (90d)",
        "edp":          "EDP (ISK/dia)",
        "sell_vol":     "Vol Venda",
        "buy_vol":      "Vol Compra",
    }

    def _build_tab_mm(self):
        p = self._tab_mm
        # ── filtros ──────────────────────────────────────────────────────────
        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(6, 2))

        regions = list(REGIONS.keys())
        self._lbl(fr, "Regiao",          0, 0)
        self._mm_reg  = self._combo(fr, regions, 1, 0)
        self._lbl(fr, "Fonte",           0, 2)
        self._mm_mode = self._combo(fr, ["CSV (cache)", "API (live)"], 1, 2, 13)
        self._csv_age_label(fr, self._mm_mode, [self._mm_reg], 2, 2)
        # Skills vem da aba CONFIGURACOES (globais)
        self._lbl(fr, "Skills",          0, 4)
        tk.Label(fr, textvariable=self._g_brel,
                 bg=COLORS["border"], fg=COLORS["accent"],
                 font=("Courier", 11, "bold"), width=3, relief="flat"
                 ).grid(row=1, column=4, padx=2, pady=2)
        tk.Label(fr, text="/", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 11)
                 ).grid(row=1, column=5, padx=0)
        tk.Label(fr, textvariable=self._g_acc,
                 bg=COLORS["border"], fg=COLORS["accent"],
                 font=("Courier", 11, "bold"), width=3, relief="flat"
                 ).grid(row=1, column=6, padx=2, pady=2)
        tk.Label(fr, text="(global)", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 9)
                 ).grid(row=0, column=5, columnspan=3, padx=2)
        self._lbl(fr, "Margem Liq min %",0, 8)
        self._mm_mmin = self._entry(fr, "5", 1, 8, 7)
        self._lbl(fr, "Max Sell (0=∞)",  0, 10)
        self._mm_maxs = self._entry(fr, "0", 1, 10, 14)
        self._lbl(fr, "Vol/Dia min",     0, 12)
        self._mm_vmin = self._entry(fr, "1", 1, 12, 6)
        self._lbl(fr, "Top N",           0, 14)
        self._mm_topn = self._entry(fr, "200", 1, 14, 6)

        self._mm_btn = tk.Button(fr, text="  CALCULAR  ",
                                 command=self._start_mm,
                                 bg=COLORS["accent"], fg="#000",
                                 font=("Courier", 11, "bold"),
                                 relief="flat", cursor="hand2")
        self._mm_btn.grid(row=1, column=16, padx=8)
        tk.Button(fr, text=" EXPORTAR ", command=self._export_mm,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=17, padx=4)
        self._mm_prog = ttk.Progressbar(fr, mode="determinate",
                                         length=120, maximum=100)
        self._mm_prog.grid(row=1, column=18, padx=6)
        self._mm_pct = tk.Label(fr, text="", bg=COLORS["panel"],
                                fg=COLORS["subtext"], font=("Courier", 10), width=5)
        self._mm_pct.grid(row=1, column=19, padx=2)

        # ── tree ─────────────────────────────────────────────────────────────
        self._mm_tree = self._make_tree(p, [
            ("item",         "Item",              220),
            ("ps",           "Sell (ISK)",        120),
            ("pb",           "Buy (ISK)",         120),
            ("gross_margin", "Margem Bruta %",     95),
            ("net_margin",   "Margem Liq %",       90),
            ("profit_unit",  "Lucro/un (ISK)",    120),
            ("avg_vol_90",   "Vol/Dia (90d)",       88),
            ("edp",          "EDP (ISK/dia)",      130),
            ("sell_vol",     "Vol Venda",           80),
            ("buy_vol",      "Vol Compra",          80),
        ])
        for cid in self._MM_COLS:
            self._mm_tree.heading(cid, text=self._MM_COLS[cid],
                                   command=lambda c=cid: self._sort_mm(c))


        # Seleção robusta (Linux/GTK) + double-click
        self._mm_cur_tid = None
        self._mm_tree.bind("<<TreeviewSelect>>", self._mm_on_select)
        self._mm_tree.bind("<Double-1>",         self._mm_on_dbl)
        bar = tk.Frame(p, bg=COLORS["bg"])
        bar.pack(fill="x", padx=6, pady=(0, 4))
        self._mm_count = tk.Label(bar, text="0 itens",
                                   fg=COLORS["subtext"], bg=COLORS["bg"],
                                   font=("Courier", 10))
        self._mm_count.pack(side="left")
        tk.Button(bar, text=" GRÁFICO ",     command=self._mm_open_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" PREVER ",      command=self._mm_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" BACKTEST ",    command=self._mm_open_bt,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" ANALISAR IA ", command=self._mm_open_ai,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Label(bar, text="EDP = lucro/un × vol/dia  |  Ranked por EDP desc",
                 fg=COLORS["subtext"], bg=COLORS["bg"],
                 font=("Courier", 10)).pack(side="right")

    def _start_mm(self):
        if self._loading_mm: return
        self._loading_mm = True
        self._mm_btn.config(state="disabled")
        self._mm_prog.config(value=0); self._mm_pct.config(text="")
        br, ac = self._gskills()
        params = dict(
            rname    = self._mm_reg.get(),
            use_csv  = self._mm_mode.get().startswith("CSV"),
            min_margin    = self._fread(self._mm_mmin, 5.0),
            max_sell_price= self._fread(self._mm_maxs, 0.0),
            min_vol       = int(self._fread(self._mm_vmin, 1)),
            top_n         = int(self._fread(self._mm_topn, 200)),
            br=br, ac=ac,
        )
        threading.Thread(target=self._worker_mm, args=(params,), daemon=True).start()

    def _worker_mm(self, params: dict):
        prog = self._make_prog(self._mm_prog, self._mm_pct)
        try:
            rname = params["rname"]
            if params["use_csv"]:
                orders = load_orders_from_csv(rname)
                if not orders:
                    self.after(0, lambda: messagebox.showerror(
                        "CSV nao encontrado",
                        "Baixe os dados primeiro na aba DADOS & CACHE.")); return
            else:
                orders = download_region_orders(REGIONS[rname], rname, prog)
                _orders_mem_cache[rname] = (time.time(), orders)

            results = analyze_max_margin(
                orders, rname, self._item_cache,
                broker_rel    = params["br"],
                accounting    = params["ac"],
                min_margin    = params["min_margin"],
                max_sell_price= params["max_sell_price"],
                min_vol       = params["min_vol"],
                top_n         = params["top_n"],
                blacklist     = self._blacklist,
                progress_cb   = prog,
            )
            self._mm_results = results
            self.after(0, lambda: (self._mm_prog.config(value=100),
                                   self._mm_pct.config(text="100%")))
            self.after(0, self._pop_mm)
            self.after(0, self._dash_refresh)   # Tier 2 — atualiza Dashboard
        except Exception as exc:
            log.exception("worker_mm")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro", e))
        finally:
            self._loading_mm = False
            self.after(0, lambda: self._mm_btn.config(state="normal"))

    def _pop_mm(self, reorder_only=False):
        if reorder_only:
            for i in range(len(self._mm_results)):
                self._mm_tree.move(str(i), "", i)
            return
        for row in self._mm_tree.get_children(): self._mm_tree.delete(row)
        for i, r in enumerate(self._mm_results):
            hot = (r.get("net_margin", 0) >= 15)
            tag = ("hot",) if hot else ("even" if i % 2 == 0 else "odd",)
            edp = self._fmt(r.get("edp"), "isk") if r.get("edp") is not None else "--"
            self._mm_tree.insert("", "end", iid=str(r["type_id"]), tags=tag, values=(
                r["item"],
                self._fmt(r["ps"],           "isk"),
                self._fmt(r["pb"],           "isk"),
                self._fmt(r["gross_margin"], "pct"),
                self._fmt(r["net_margin"],   "pct"),
                self._fmt(r["profit_unit"],  "isk"),
                self._fmt(r.get("avg_vol_90"), "vol"),
                edp,
                self._fmt(r["sell_vol"],     "int"),
                self._fmt(r["buy_vol"],      "int"),
            ))
        self._mm_count.config(text=f"{len(self._mm_results)} itens")

    def _sort_mm(self, col):
        num = {"ps","pb","gross_margin","net_margin","profit_unit",
               "avg_vol_90","edp","sell_vol","buy_vol"}
        rev = not self._mm_sort_state.get(col, False)
        self._mm_sort_state = {col: rev}
        self._mm_results.sort(
            key=lambda r: (r.get(col) or 0) if col in num else r.get(col, ""),
            reverse=rev)
        self._set_heading_arrow(self._mm_tree, self._MM_COLS, col, rev)
        self._pop_mm(reorder_only=True)

    def _export_mm(self):
        headers = [(k, v) for k,v in self._MM_COLS.items()] + [("region","Região")]
        self._export_table("Max Margem", headers, self._mm_results, "max_margem")

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA 5 — ESCASSEZ
    # ══════════════════════════════════════════════════════════════════════════

    _SC_COLS = {
        "item":       "Item",
        "ps":         "Sell (ISK)",
        "pb":         "Buy (ISK)",
        "margin":     "Margem %",
        "sell_vol":   "Vol Venda",
        "avg_vol_90": "Vol/Dia (90d)",
        "dos":        "Dias Suprimento",
        "score":      "Score",
    }

    def _build_tab_sc(self):
        p = self._tab_sc
        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(6, 2))

        regions = list(REGIONS.keys())
        self._lbl(fr, "Regiao",          0, 0)
        self._sc_reg  = self._combo(fr, regions, 1, 0)
        self._lbl(fr, "Fonte",           0, 2)
        self._sc_mode = self._combo(fr, ["CSV (cache)", "API (live)"], 1, 2, 13)
        self._csv_age_label(fr, self._sc_mode, [self._sc_reg], 2, 2)
        self._lbl(fr, "Max DoS (dias)",  0, 4)
        self._sc_dos  = self._entry(fr, "30",  1, 4, 6)
        self._lbl(fr, "Margem min %",    0, 6)
        self._sc_mmin = self._entry(fr, "0",   1, 6, 7)
        self._lbl(fr, "Vol/Dia min",     0, 8)
        self._sc_vmin = self._entry(fr, "5",   1, 8, 6)
        self._lbl(fr, "Top N",           0, 10)
        self._sc_topn = self._entry(fr, "200", 1, 10, 6)

        self._sc_btn = tk.Button(fr, text="  CALCULAR  ",
                                 command=self._start_sc,
                                 bg=COLORS["yellow"], fg="#000",
                                 font=("Courier", 11, "bold"),
                                 relief="flat", cursor="hand2")
        self._sc_btn.grid(row=1, column=12, padx=8)
        tk.Button(fr, text=" EXPORTAR ", command=self._export_sc,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=13, padx=4)
        self._sc_prog = ttk.Progressbar(fr, mode="determinate",
                                         length=120, maximum=100)
        self._sc_prog.grid(row=1, column=14, padx=6)
        self._sc_pct = tk.Label(fr, text="", bg=COLORS["panel"],
                                fg=COLORS["subtext"], font=("Courier", 10), width=5)
        self._sc_pct.grid(row=1, column=15, padx=2)

        self._sc_tree = self._make_tree(p, [
            ("item",       "Item",               240),
            ("ps",         "Sell (ISK)",          120),
            ("pb",         "Buy (ISK)",           120),
            ("margin",     "Margem %",             82),
            ("sell_vol",   "Vol Venda",             80),
            ("avg_vol_90", "Vol/Dia (90d)",         90),
            ("dos",        "Dias Suprimento",      105),
            ("score",      "Score",                 88),
        ])
        for cid in self._SC_COLS:
            self._sc_tree.heading(cid, text=self._SC_COLS[cid],
                                   command=lambda c=cid: self._sort_sc(c))


        # Seleção robusta (Linux/GTK) + double-click
        self._sc_cur_tid = None
        self._sc_tree.bind("<<TreeviewSelect>>", self._sc_on_select)
        self._sc_tree.bind("<Double-1>",         self._sc_on_dbl)
        bar = tk.Frame(p, bg=COLORS["bg"])
        bar.pack(fill="x", padx=6, pady=(0, 4))
        self._sc_count = tk.Label(bar, text="0 itens",
                                   fg=COLORS["subtext"], bg=COLORS["bg"],
                                   font=("Courier", 10))
        self._sc_count.pack(side="left")
        tk.Button(bar, text=" GRÁFICO ",     command=self._sc_open_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" PREVER ",      command=self._sc_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" BACKTEST ",    command=self._sc_open_bt,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" ANALISAR IA ", command=self._sc_open_ai,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Label(bar,
                 text="DoS = Vol_Venda / Vol_Dia  |  Score = Margem × log10(Vol) / (DoS+1)",
                 fg=COLORS["subtext"], bg=COLORS["bg"],
                 font=("Courier", 10)).pack(side="right")

    def _start_sc(self):
        if self._loading_sc: return
        self._loading_sc = True
        self._sc_btn.config(state="disabled")
        self._sc_prog.config(value=0); self._sc_pct.config(text="")
        br, ac = self._gskills()
        params = dict(
            rname     = self._sc_reg.get(),
            use_csv   = self._sc_mode.get().startswith("CSV"),
            max_dos   = self._fread(self._sc_dos,  30.0),
            min_margin= self._fread(self._sc_mmin,  0.0),
            min_vol   = int(self._fread(self._sc_vmin, 5)),
            top_n     = int(self._fread(self._sc_topn, 200)),
            br=br, ac=ac,
        )
        threading.Thread(target=self._worker_sc, args=(params,), daemon=True).start()

    def _worker_sc(self, params: dict):
        prog = self._make_prog(self._sc_prog, self._sc_pct)
        try:
            rname = params["rname"]
            if params["use_csv"]:
                orders = load_orders_from_csv(rname)
                if not orders:
                    self.after(0, lambda: messagebox.showerror(
                        "CSV nao encontrado",
                        "Baixe os dados na aba DADOS & CACHE.")); return
            else:
                orders = download_region_orders(REGIONS[rname], rname, prog)
                _orders_mem_cache[rname] = (time.time(), orders)

            results = analyze_supply_shortage(
                orders, rname, self._item_cache,
                max_dos    = params["max_dos"],
                min_margin = params["min_margin"],
                min_vol    = params["min_vol"],
                top_n      = params["top_n"],
                blacklist  = self._blacklist,
                progress_cb= prog,
            )
            self._sc_results = results
            self.after(0, lambda: (self._sc_prog.config(value=100),
                                   self._sc_pct.config(text="100%")))
            self.after(0, self._pop_sc)
        except Exception as exc:
            log.exception("worker_sc")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro", e))
        finally:
            self._loading_sc = False
            self.after(0, lambda: self._sc_btn.config(state="normal"))

    def _pop_sc(self, reorder_only=False):
        if reorder_only:
            for i in range(len(self._sc_results)):
                self._sc_tree.move(str(i), "", i)
            return
        for row in self._sc_tree.get_children(): self._sc_tree.delete(row)
        for i, r in enumerate(self._sc_results):
            dos = r.get("dos", 9999)
            hot = dos <= 3
            tag = ("hot",) if hot else ("even" if i % 2 == 0 else "odd",)
            self._sc_tree.insert("", "end", iid=str(r["type_id"]), tags=tag, values=(
                r["item"],
                self._fmt(r["ps"],           "isk"),
                self._fmt(r["pb"],           "isk"),
                self._fmt(r["margin"],       "pct"),
                self._fmt(r["sell_vol"],     "int"),
                self._fmt(r.get("avg_vol_90", 0), "vol"),
                self._fmt(dos,               "vol"),
                self._fmt(r.get("score", 0), "vol"),
            ))
        self._sc_count.config(text=f"{len(self._sc_results)} itens em escassez")

    def _sort_sc(self, col):
        num = {"ps","pb","margin","sell_vol","avg_vol_90","dos","score"}
        rev = not self._sc_sort_state.get(col, False)
        self._sc_sort_state = {col: rev}
        self._sc_results.sort(
            key=lambda r: (r.get(col) or 0) if col in num else r.get(col, ""),
            reverse=rev)
        self._set_heading_arrow(self._sc_tree, self._SC_COLS, col, rev)
        self._pop_sc(reorder_only=True)

    def _export_sc(self):
        headers = [(k, v) for k,v in self._SC_COLS.items()] + [("region","Região")]
        self._export_table("Escassez", headers, self._sc_results, "escassez")

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA 6 — DEMANDA / OFERTA  (DSIF)
    # ══════════════════════════════════════════════════════════════════════════

    _DS_COLS = {
        "item":       "Item",
        "ps":         "Sell (ISK)",
        "pb":         "Buy (ISK)",
        "margin":     "Margem %",
        "buy_vol":    "Vol Compra",
        "sell_vol":   "Vol Venda",
        "ratio":      "Ratio B/S",
        "avg_vol_90": "Vol/Dia (90d)",
        "dsif":       "DSIF Score",
    }

    def _build_tab_ds(self):
        p = self._tab_ds
        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(6, 2))

        regions = list(REGIONS.keys())
        self._lbl(fr, "Regiao",        0, 0)
        self._ds_reg  = self._combo(fr, regions, 1, 0)
        self._lbl(fr, "Fonte",         0, 2)
        self._ds_mode = self._combo(fr, ["CSV (cache)", "API (live)"], 1, 2, 13)
        self._csv_age_label(fr, self._ds_mode, [self._ds_reg], 2, 2)
        self._lbl(fr, "Margem min %",  0, 4)
        self._ds_mmin = self._entry(fr, "0",   1, 4, 7)
        self._lbl(fr, "Vol/Dia min",   0, 6)
        self._ds_vmin = self._entry(fr, "5",   1, 6, 6)
        self._lbl(fr, "Top N",         0, 8)
        self._ds_topn = self._entry(fr, "200", 1, 8, 6)

        self._ds_btn = tk.Button(fr, text="  CALCULAR  ",
                                 command=self._start_ds,
                                 bg=COLORS["green"], fg="#000",
                                 font=("Courier", 11, "bold"),
                                 relief="flat", cursor="hand2")
        self._ds_btn.grid(row=1, column=10, padx=8)
        tk.Button(fr, text=" EXPORTAR ", command=self._export_ds,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=11, padx=4)
        self._ds_prog = ttk.Progressbar(fr, mode="determinate",
                                         length=120, maximum=100)
        self._ds_prog.grid(row=1, column=12, padx=6)
        self._ds_pct = tk.Label(fr, text="", bg=COLORS["panel"],
                                fg=COLORS["subtext"], font=("Courier", 10), width=5)
        self._ds_pct.grid(row=1, column=13, padx=2)

        self._ds_tree = self._make_tree(p, [
            ("item",       "Item",              240),
            ("ps",         "Sell (ISK)",         120),
            ("pb",         "Buy (ISK)",          120),
            ("margin",     "Margem %",            82),
            ("buy_vol",    "Vol Compra",           85),
            ("sell_vol",   "Vol Venda",            80),
            ("ratio",      "Ratio B/S",            80),
            ("avg_vol_90", "Vol/Dia (90d)",         90),
            ("dsif",       "DSIF Score",            95),
        ])
        for cid in self._DS_COLS:
            self._ds_tree.heading(cid, text=self._DS_COLS[cid],
                                   command=lambda c=cid: self._sort_ds(c))


        # Seleção robusta (Linux/GTK) + double-click
        self._ds_cur_tid = None
        self._ds_tree.bind("<<TreeviewSelect>>", self._ds_on_select)
        self._ds_tree.bind("<Double-1>",         self._ds_on_dbl)
        bar = tk.Frame(p, bg=COLORS["bg"])
        bar.pack(fill="x", padx=6, pady=(0, 4))
        self._ds_count = tk.Label(bar, text="0 itens",
                                   fg=COLORS["subtext"], bg=COLORS["bg"],
                                   font=("Courier", 10))
        self._ds_count.pack(side="left")
        tk.Button(bar, text=" GRÁFICO ",     command=self._ds_open_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" PREVER ",      command=self._ds_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" BACKTEST ",    command=self._ds_open_bt,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" ANALISAR IA ", command=self._ds_open_ai,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Label(bar,
                 text="DSIF = (Bv/Sv)² × (Vd/Sv) × log10(Ps−Pb+1)  |  Alto DSIF = alta demanda relativa",
                 fg=COLORS["subtext"], bg=COLORS["bg"],
                 font=("Courier", 10)).pack(side="right")

    def _start_ds(self):
        if self._loading_ds: return
        self._loading_ds = True
        self._ds_btn.config(state="disabled")
        self._ds_prog.config(value=0); self._ds_pct.config(text="")
        br, ac = self._gskills()
        params = dict(
            rname     = self._ds_reg.get(),
            use_csv   = self._ds_mode.get().startswith("CSV"),
            min_margin= self._fread(self._ds_mmin, 0.0),
            min_vol   = int(self._fread(self._ds_vmin, 5)),
            top_n     = int(self._fread(self._ds_topn, 200)),
            br=br, ac=ac,
        )
        threading.Thread(target=self._worker_ds, args=(params,), daemon=True).start()

    def _worker_ds(self, params: dict):
        prog = self._make_prog(self._ds_prog, self._ds_pct)
        try:
            rname = params["rname"]
            if params["use_csv"]:
                orders = load_orders_from_csv(rname)
                if not orders:
                    self.after(0, lambda: messagebox.showerror(
                        "CSV nao encontrado",
                        "Baixe os dados na aba DADOS & CACHE.")); return
            else:
                orders = download_region_orders(REGIONS[rname], rname, prog)
                _orders_mem_cache[rname] = (time.time(), orders)

            results = analyze_demand_supply(
                orders, rname, self._item_cache,
                min_margin = params["min_margin"],
                min_vol    = params["min_vol"],
                top_n      = params["top_n"],
                blacklist  = self._blacklist,
                progress_cb= prog,
            )
            self._ds_results = results
            self.after(0, lambda: (self._ds_prog.config(value=100),
                                   self._ds_pct.config(text="100%")))
            self.after(0, self._pop_ds)
        except Exception as exc:
            log.exception("worker_ds")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro", e))
        finally:
            self._loading_ds = False
            self.after(0, lambda: self._ds_btn.config(state="normal"))

    def _pop_ds(self, reorder_only=False):
        if reorder_only:
            for i in range(len(self._ds_results)):
                self._ds_tree.move(str(i), "", i)
            return
        for row in self._ds_tree.get_children(): self._ds_tree.delete(row)
        for i, r in enumerate(self._ds_results):
            dsif = r.get("dsif", 0) or 0
            hot  = dsif >= 10
            tag  = ("hot",) if hot else ("even" if i % 2 == 0 else "odd",)
            self._ds_tree.insert("", "end", iid=str(r["type_id"]), tags=tag, values=(
                r["item"],
                self._fmt(r["ps"],           "isk"),
                self._fmt(r["pb"],           "isk"),
                self._fmt(r["margin"],       "pct"),
                self._fmt(r["buy_vol"],      "int"),
                self._fmt(r["sell_vol"],     "int"),
                self._fmt(r["ratio"],        "vol"),
                self._fmt(r.get("avg_vol_90", 0), "vol"),
                self._fmt(dsif,              "vol"),
            ))
        self._ds_count.config(text=f"{len(self._ds_results)} itens")

    def _sort_ds(self, col):
        num = {"ps","pb","margin","buy_vol","sell_vol","ratio","avg_vol_90","dsif"}
        rev = not self._ds_sort_state.get(col, False)
        self._ds_sort_state = {col: rev}
        self._ds_results.sort(
            key=lambda r: (r.get(col) or 0) if col in num else r.get(col, ""),
            reverse=rev)
        self._set_heading_arrow(self._ds_tree, self._DS_COLS, col, rev)
        self._pop_ds(reorder_only=True)

    def _export_ds(self):
        headers = [(k, v) for k,v in self._DS_COLS.items()] + [("region","Região")]
        self._export_table("Demanda/Oferta", headers, self._ds_results, "demanda_oferta")

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA 7 — MOVIMENTACAO DE PRECOS
    # ══════════════════════════════════════════════════════════════════════════

    _PM_COLS = {
        "item":        "Item",
        "avg_p_week":  "Preco 7d (ISK)",
        "avg_p_month": "Preco 30d (ISK)",
        "price_chg":   "Var. Preco %",
        "avg_v_week":  "Vol 7d/dia",
        "avg_v_month": "Vol 30d/dia",
        "vol_chg":     "Var. Volume %",
        "margin":      "Margem %",
        "correlation": "Correlacao",
        "avg_vol_90":  "Vol/Dia (90d)",
        "score":       "Score",
    }

    def _build_tab_pm(self):
        p = self._tab_pm

        # ── Painel de sincronizacao do banco ──────────────────────────────────
        sync_fr = tk.Frame(p, bg=COLORS["panel"],
                           highlightbackground=COLORS["border"], highlightthickness=1)
        sync_fr.pack(fill="x", padx=6, pady=(6, 1))

        tk.Label(sync_fr, text="BANCO DE HISTORICO LOCAL",
                 bg=COLORS["panel"], fg=COLORS["header"],
                 font=("Courier", 11, "bold")).grid(row=0, column=0, padx=8,
                                                   pady=(4,0), sticky="w",
                                                   columnspan=10)

        regions = list(REGIONS.keys())
        self._lbl(sync_fr, "Regiao",   1, 0)
        self._pm_reg  = self._combo(sync_fr, regions, 2, 0)

        self._pm_sync_btn = tk.Button(
            sync_fr, text=" ⬇  SINCRONIZAR BD ",
            command=self._start_pm_sync,
            bg=COLORS["accent"], fg="#000",
            font=("Courier", 11, "bold"), relief="flat", cursor="hand2")
        self._pm_sync_btn.grid(row=2, column=2, padx=8)

        self._pm_stop_btn = tk.Button(
            sync_fr, text=" ■  PARAR ",
            command=self._stop_pm_sync,
            bg=COLORS["border"], fg=COLORS["red"],
            font=("Courier", 11), relief="flat", cursor="hand2",
            state="disabled")
        self._pm_stop_btn.grid(row=2, column=3, padx=4)

        self._pm_sync_prog = ttk.Progressbar(sync_fr, mode="determinate",
                                              length=200, maximum=100)
        self._pm_sync_prog.grid(row=2, column=4, padx=6)
        self._pm_sync_pct = tk.Label(sync_fr, text="", bg=COLORS["panel"],
                                     fg=COLORS["subtext"], font=("Courier", 10), width=5)
        self._pm_sync_pct.grid(row=2, column=5, padx=2)

        self._pm_db_lbl = tk.Label(sync_fr, text="BD: --",
                                    bg=COLORS["panel"], fg=COLORS["subtext"],
                                    font=("Courier", 10))
        self._pm_db_lbl.grid(row=2, column=6, padx=12, sticky="w")

        tk.Label(sync_fr,
                 text="Primeira vez: baixa historico completo de todos os itens.\n"
                      "Dias seguintes: baixa so os dias novos (~segundos).",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 9), justify="left"
                 ).grid(row=0, column=6, rowspan=2, padx=12, sticky="w")

        # ── Painel de filtros e calculo ───────────────────────────────────────
        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(2, 2))

        self._lbl(fr, "Fonte de itens", 0, 0)
        self._pm_mode = self._combo(fr,
            ["BD local (recomendado)", "CSV (cache)", "API (live)"], 1, 0, 20)
        self._lbl(fr, "Vol/Dia min",    0, 2)
        self._pm_vmin = self._entry(fr, "5",   1, 2, 6)
        self._lbl(fr, "Margem min %",   0, 4)
        self._pm_mmin = self._entry(fr, "0",   1, 4, 7)
        self._lbl(fr, "Top N",          0, 6)
        self._pm_topn = self._entry(fr, "500", 1, 6, 6)

        self._pm_btn = tk.Button(fr, text="  CALCULAR  ",
                                 command=self._start_pm,
                                 bg=COLORS["orange"], fg="#000",
                                 font=("Courier", 11, "bold"),
                                 relief="flat", cursor="hand2")
        self._pm_btn.grid(row=1, column=8, padx=8)
        tk.Button(fr, text=" EXPORTAR ", command=self._export_pm,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=9, padx=4)
        self._pm_prog = ttk.Progressbar(fr, mode="determinate",
                                         length=120, maximum=100)
        self._pm_prog.grid(row=1, column=10, padx=6)
        self._pm_pct = tk.Label(fr, text="", bg=COLORS["panel"],
                                fg=COLORS["subtext"], font=("Courier", 10), width=5)
        self._pm_pct.grid(row=1, column=11, padx=2)

        # ── Tabela de resultados ──────────────────────────────────────────────
        self._pm_tree = self._make_tree(p, [
            ("item",        "Item",              220),
            ("avg_p_week",  "Preco 7d (ISK)",    130),
            ("avg_p_month", "Preco 30d (ISK)",   130),
            ("price_chg",   "Var. Preco %",        90),
            ("avg_v_week",  "Vol 7d/dia",          82),
            ("avg_v_month", "Vol 30d/dia",          88),
            ("vol_chg",     "Var. Volume %",        88),
            ("margin",      "Margem %",             78),
            ("correlation", "Correlacao",           88),
            ("avg_vol_90",  "Vol/Dia (90d)",        90),
            ("score",       "Score",                72),
        ])
        for cid in self._PM_COLS:
            self._pm_tree.heading(cid, text=self._PM_COLS[cid],
                                   command=lambda c=cid: self._sort_pm(c))
        self._pm_tree.tag_configure("up",   foreground=COLORS["green"])
        self._pm_tree.tag_configure("down", foreground=COLORS["red"])
        # Seleção robusta (Linux/GTK) + double-click
        self._pm_tree.bind("<<TreeviewSelect>>", self._pm_on_select)
        self._pm_tree.bind("<Double-1>",         self._pm_on_dbl)

        bar = tk.Frame(p, bg=COLORS["bg"])
        bar.pack(fill="x", padx=6, pady=(0, 4))
        self._pm_count = tk.Label(bar, text="0 itens",
                                   fg=COLORS["subtext"], bg=COLORS["bg"],
                                   font=("Courier", 10))
        self._pm_count.pack(side="left")
        tk.Button(bar, text=" GRÁFICO ",     command=self._pm_open_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" PREVER ",      command=self._pm_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" BACKTEST ",    command=self._pm_open_bt,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" ANALISAR IA ", command=self._pm_open_ai,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Label(bar,
                 text="Verde = preco subindo  |  Vermelho = preco caindo  |  Score = log10(Vol/Dia+1)",
                 fg=COLORS["subtext"], bg=COLORS["bg"],
                 font=("Courier", 10)).pack(side="right")

        # Atualiza label de BD ao abrir
        self._pm_refresh_db_lbl()

    # ── Sincronizacao do banco ─────────────────────────────────────────────────

    def _pm_refresh_db_lbl(self):
        """Atualiza o label de estatisticas do banco na aba MOV. PRECOS."""
        rname = self._pm_reg.get()
        rid   = REGIONS.get(rname, 10000002)
        def _do():
            st = _db_stats(rid)
            self.after(0, lambda: self._pm_db_lbl.config(
                fg=COLORS["green"] if st["n_items"] > 0 else COLORS["subtext"],
                text=(f"BD: {st['n_items']:,} itens  |  "
                      f"{st['min_date']} → {st['max_date']}"
                      if st["n_items"] > 0 else "BD: vazio (clique SINCRONIZAR)")))
        threading.Thread(target=_do, daemon=True).start()

    def _start_pm_sync(self):
        """Inicia sincronizacao incremental do banco de historico."""
        if getattr(self, "_pm_syncing", False):
            return
        self._pm_syncing  = True
        self._pm_stop_flag = [False]
        self._pm_sync_btn.config(state="disabled")
        self._pm_stop_btn.config(state="normal")
        self._pm_sync_prog.config(value=0)
        self._pm_sync_pct.config(text="")
        rname = self._pm_reg.get()   # C2 — lê no thread principal
        threading.Thread(target=self._worker_pm_sync,
                         args=(rname,), daemon=True).start()

    def _stop_pm_sync(self):
        """Sinaliza para o worker de sync parar apos o item atual."""
        if hasattr(self, "_pm_stop_flag"):
            self._pm_stop_flag[0] = True
        self._pm_stop_btn.config(state="disabled")
        self._set_status("Parando sincronizacao...")

    def _worker_pm_sync(self, rname: str):
        """
        Worker de sincronizacao incremental do SQLite.
        Descobre os type_ids da regiao a partir do CSV de ordens (se disponivel)
        ou do item_cache, depois chama sync_history_db().
        """
        prog = self._make_prog(self._pm_sync_prog, self._pm_sync_pct)
        try:
            rid   = REGIONS.get(rname, 10000002)

            # Descobre type_ids: CSV de ordens e a fonte mais completa
            orders = load_orders_from_csv(rname)
            if orders:
                idx      = build_order_index(orders)
                type_ids = list(idx.keys())
                prog(f"Fonte: CSV ({len(type_ids):,} itens da regiao)")
            elif self._item_cache:
                type_ids = list(self._item_cache.keys())
                prog(f"Fonte: item_cache ({len(type_ids):,} itens)")
            else:
                self.after(0, lambda: messagebox.showinfo(
                    "Sem dados de origem",
                    "Baixe as ordens da regiao na aba DADOS & CACHE primeiro.\n"
                    "Os type_ids sao extraidos do CSV de ordens."))
                return

            prog(f"Sincronizando {len(type_ids):,} itens para {rname}...")
            stats = sync_history_db(rid, type_ids, prog, self._pm_stop_flag)

            summary = (f"Sincronizacao concluida!\n\n"
                       f"Total:     {stats['total']:,} itens\n"
                       f"Atualizados: {stats['updated']:,}\n"
                       f"Ja ok:     {stats['skipped']:,}\n"
                       f"Erros:     {stats['errors']:,}")
            self.after(0, lambda: (
                messagebox.showinfo("Sincronizacao", summary),
                self._pm_refresh_db_lbl()))
            self.after(0, lambda: (self._pm_sync_prog.config(value=100),
                                   self._pm_sync_pct.config(text="100%")))

        except Exception as exc:
            log.exception("worker_pm_sync")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro Sync", e))
        finally:
            self._pm_syncing = False
            self.after(0, lambda: (
                self._pm_sync_btn.config(state="normal"),
                self._pm_stop_btn.config(state="disabled")))

    def _start_pm(self):
        if self._loading_pm: return
        self._loading_pm = True
        self._pm_btn.config(state="disabled")
        self._pm_prog.config(value=0); self._pm_pct.config(text="")
        params = dict(
            rname     = self._pm_reg.get(),
            mode      = self._pm_mode.get(),
            min_vol   = int(self._fread(self._pm_vmin, 5)),
            min_margin= self._fread(self._pm_mmin, 0.0),
            top_n     = int(self._fread(self._pm_topn, 500)),
        )
        threading.Thread(target=self._worker_pm, args=(params,), daemon=True).start()

    def _worker_pm(self, params: dict):
        prog = self._make_prog(self._pm_prog, self._pm_pct)
        try:
            rname  = params["rname"]
            rid    = REGIONS.get(rname, 10000002)
            mode   = params["mode"]
            orders = None
            type_ids = None

            if mode.startswith("BD"):
                # Usa os type_ids que ja estao no banco local — sem request de ordens
                n = _db_item_count(rid)
                if n == 0:
                    self.after(0, lambda: messagebox.showinfo(
                        "Banco vazio",
                        "Clique SINCRONIZAR BD primeiro para baixar o historico."))
                    return
                conn     = _get_db()
                rows     = conn.execute(
                    "SELECT DISTINCT type_id FROM history WHERE region_id=?",
                    (rid,)).fetchall()
                type_ids = [r[0] for r in rows]
                prog(f"BD local: {len(type_ids):,} itens para {rname}")

            elif mode.startswith("CSV"):
                orders = load_orders_from_csv(rname)
                if not orders:
                    self.after(0, lambda: messagebox.showerror(
                        "CSV nao encontrado",
                        "Baixe os dados na aba DADOS & CACHE.")); return

            else:  # API live
                orders = download_region_orders(rid, rname, prog)
                _orders_mem_cache[rname] = (time.time(), orders)

            results = analyze_price_movement(
                rname, self._item_cache,
                min_vol     = params["min_vol"],
                min_margin  = params["min_margin"],
                top_n       = params["top_n"],
                orders      = orders,
                type_ids    = type_ids,
                progress_cb = prog,
            )
            self._pm_results = results
            self.after(0, lambda: (self._pm_prog.config(value=100),
                                   self._pm_pct.config(text="100%")))
            self.after(0, self._pop_pm)
            self.after(0, self._pm_refresh_db_lbl)
            self.after(0, self._dash_refresh)   # Tier 2 — atualiza Dashboard
        except Exception as exc:
            log.exception("worker_pm")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro", e))
        finally:
            self._loading_pm = False
            self.after(0, lambda: self._pm_btn.config(state="normal"))

    def _pop_pm(self, reorder_only=False):
        if reorder_only:
            for i in range(len(self._pm_results)):
                self._pm_tree.move(str(i), "", i)
            return
        for row in self._pm_tree.get_children(): self._pm_tree.delete(row)
        for i, r in enumerate(self._pm_results):
            pc   = r.get("price_chg", 0) or 0
            base = "even" if i % 2 == 0 else "odd"
            tag  = ("up",) if pc > 0 else (("down",) if pc < 0 else (base,))
            self._pm_tree.insert("", "end", iid=str(r["type_id"]), tags=tag, values=(
                r["item"],
                self._fmt(r["avg_p_week"],  "isk"),
                self._fmt(r["avg_p_month"], "isk"),
                self._fmt(pc,               "pct"),
                self._fmt(r["avg_v_week"],  "vol"),
                self._fmt(r["avg_v_month"], "vol"),
                self._fmt(r.get("vol_chg",0),"pct"),
                self._fmt(r.get("margin",0), "pct"),
                r.get("correlation", "--"),
                self._fmt(r.get("avg_vol_90",0),"vol"),
                self._fmt(r.get("score",0), "vol"),
            ))
        self._pm_count.config(text=f"{len(self._pm_results)} itens")

    def _sort_pm(self, col):
        num = {"avg_p_week","avg_p_month","price_chg","avg_v_week",
               "avg_v_month","vol_chg","margin","avg_vol_90","score"}
        rev = not self._pm_sort_state.get(col, False)
        self._pm_sort_state = {col: rev}
        self._pm_results.sort(
            key=lambda r: (r.get(col) or 0) if col in num else r.get(col, ""),
            reverse=rev)
        self._set_heading_arrow(self._pm_tree, self._PM_COLS, col, rev)
        self._pop_pm(reorder_only=True)

    def _export_pm(self):
        headers = [(k, v) for k,v in self._PM_COLS.items()] + [("region","Região")]
        self._export_table("Mov. Preços", headers, self._pm_results, "mov_precos")

    # ══════════════════════════════════════════════════════════════════════════
    #  ABA 8 — ORCAMENTO
    # ══════════════════════════════════════════════════════════════════════════

    _BUD_COLS = {
        "item":         "Item",
        "pb":           "Buy (ISK)",
        "ps":           "Sell (ISK)",
        "qty":          "Qtd",
        "u_max":        "Umax",
        "cost_unit":    "Custo/un",
        "pi_net":       "Lucro/un",
        "net_margin":   "Margem Liq %",
        "total_buy":    "Capital (ISK)",
        "total_profit": "Lucro Total (ISK)",
        "avg_vol_90":   "Vol/Dia (90d)",
    }

    def _build_tab_bud(self):
        p = self._tab_bud
        fr = tk.Frame(p, bg=COLORS["panel"],
                      highlightbackground=COLORS["border"], highlightthickness=1)
        fr.pack(fill="x", padx=6, pady=(6, 2))

        regions = list(REGIONS.keys())
        self._lbl(fr, "Regiao",          0, 0)
        self._bud_reg  = self._combo(fr, regions, 1, 0)
        self._lbl(fr, "Fonte",           0, 2)
        self._bud_mode = self._combo(fr, ["CSV (cache)", "API (live)"], 1, 2, 13)
        self._csv_age_label(fr, self._bud_mode, [self._bud_reg], 2, 2)
        self._lbl(fr, "Orcamento (ISK)", 0, 4)
        self._bud_isk  = self._entry(fr, "500000000", 1, 4, 16)
        # Skills globais
        self._lbl(fr, "Skills",          0, 6)
        tk.Label(fr, textvariable=self._g_brel,
                 bg=COLORS["border"], fg=COLORS["accent"],
                 font=("Courier", 11, "bold"), width=3, relief="flat"
                 ).grid(row=1, column=6, padx=2, pady=2)
        tk.Label(fr, text="/", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 11)
                 ).grid(row=1, column=7, padx=0)
        tk.Label(fr, textvariable=self._g_acc,
                 bg=COLORS["border"], fg=COLORS["accent"],
                 font=("Courier", 11, "bold"), width=3, relief="flat"
                 ).grid(row=1, column=8, padx=2, pady=2)
        tk.Label(fr, text="(global)", bg=COLORS["panel"],
                 fg=COLORS["subtext"], font=("Courier", 9)
                 ).grid(row=0, column=7, columnspan=3, padx=2)
        self._lbl(fr, "Risco",           0, 10)
        self._bud_risk = self._combo(fr,
            list(_RISK_DAYS.keys()), 1, 10, 20)
        self._lbl(fr, "Vol/Dia min",     0, 12)
        self._bud_vmin = self._entry(fr, "1", 1, 12, 6)

        self._bud_btn = tk.Button(fr, text="  CALCULAR  ",
                                  command=self._start_bud,
                                  bg=COLORS["accent"], fg="#000",
                                  font=("Courier", 11, "bold"),
                                  relief="flat", cursor="hand2")
        self._bud_btn.grid(row=1, column=14, padx=8)
        tk.Button(fr, text=" EXPORTAR ", command=self._export_bud,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).grid(row=1, column=15, padx=4)
        self._bud_prog = ttk.Progressbar(fr, mode="determinate",
                                          length=120, maximum=100)
        self._bud_prog.grid(row=1, column=16, padx=6)
        self._bud_pct = tk.Label(fr, text="", bg=COLORS["panel"],
                                 fg=COLORS["subtext"], font=("Courier", 10), width=5)
        self._bud_pct.grid(row=1, column=17, padx=2)

        self._bud_tree = self._make_tree(p, [
            ("item",         "Item",               220),
            ("pb",           "Buy (ISK)",           120),
            ("ps",           "Sell (ISK)",          120),
            ("qty",          "Qtd",                  58),
            ("u_max",        "Umax",                 58),
            ("cost_unit",    "Custo/un",            118),
            ("pi_net",       "Lucro/un",            115),
            ("net_margin",   "Margem Liq %",         95),
            ("total_buy",    "Capital (ISK)",        130),
            ("total_profit", "Lucro Total (ISK)",   140),
            ("avg_vol_90",   "Vol/Dia (90d)",         90),
        ])
        for cid in self._BUD_COLS:
            self._bud_tree.heading(cid, text=self._BUD_COLS[cid],
                                    command=lambda c=cid: self._sort_bud(c))

        self._bud_stats_lbl = tk.Label(p, text="--",
                                        bg=COLORS["panel"], fg=COLORS["green"],
                                        font=("Courier", 11, "bold"))
        self._bud_stats_lbl.pack(fill="x", padx=6, pady=(2, 4))
        # Seleção robusta (Linux/GTK) + double-click
        self._bud_tree.bind("<<TreeviewSelect>>", self._bud_on_select)
        self._bud_tree.bind("<Double-1>",         self._bud_on_dbl)

        bar = tk.Frame(p, bg=COLORS["bg"])
        bar.pack(fill="x", padx=6, pady=(0, 2))
        tk.Button(bar, text=" GRÁFICO ",     command=self._bud_open_chart,
                  bg=COLORS["border"], fg=COLORS["accent"],
                  font=("Courier", 10), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" PREVER ",      command=self._bud_open_fcast,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" BACKTEST ",    command=self._bud_open_bt,
                  bg=COLORS["border"], fg="#7f4adf",
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Button(bar, text=" ANALISAR IA ", command=self._bud_open_ai,
                  bg=COLORS["border"], fg=COLORS["yellow"],
                  font=("Courier", 10, "bold"), relief="flat", cursor="hand2").pack(side="left", padx=2)
        tk.Label(bar,
                 text="Greedy knapsack por eficiencia (lucro/ISK)  |  "
                      "Umax = Vol/Dia × Dias_Risco  |  max 150 ordens",
                 fg=COLORS["subtext"], bg=COLORS["bg"],
                 font=("Courier", 10)).pack(side="left")

    def _start_bud(self):
        if self._loading_bud: return
        self._loading_bud = True
        self._bud_btn.config(state="disabled")
        self._bud_prog.config(value=0); self._bud_pct.config(text="")
        br, ac = self._gskills()
        params = dict(
            rname    = self._bud_reg.get(),
            use_csv  = self._bud_mode.get().startswith("CSV"),
            budget   = self._fread(self._bud_isk,  500_000_000),
            risk_mode= self._bud_risk.get(),
            min_vol  = int(self._fread(self._bud_vmin, 1)),
            br=br, ac=ac,
        )
        threading.Thread(target=self._worker_bud, args=(params,), daemon=True).start()

    def _worker_bud(self, params: dict):
        prog = self._make_prog(self._bud_prog, self._bud_pct)
        try:
            rname = params["rname"]
            if params["use_csv"]:
                orders = load_orders_from_csv(rname)
                if not orders:
                    self.after(0, lambda: messagebox.showerror(
                        "CSV nao encontrado",
                        "Baixe os dados na aba DADOS & CACHE.")); return
            else:
                orders = download_region_orders(REGIONS[rname], rname, prog)
                _orders_mem_cache[rname] = (time.time(), orders)

            basket, stats = analyze_budget(
                orders, rname, self._item_cache,
                budget     = params["budget"],
                broker_rel = params["br"],
                accounting = params["ac"],
                risk_mode  = params["risk_mode"],
                min_vol    = params["min_vol"],
                blacklist  = self._blacklist,
                progress_cb= prog,
            )
            self._bud_results = basket
            self._bud_stats   = stats
            self.after(0, lambda: (self._bud_prog.config(value=100),
                                   self._bud_pct.config(text="100%")))
            self.after(0, self._pop_bud)
            self.after(0, self._dash_refresh)   # Tier 2 — atualiza Dashboard
        except Exception as exc:
            log.exception("worker_bud")
            self.after(0, lambda e=str(exc): messagebox.showerror("Erro", e))
        finally:
            self._loading_bud = False
            self.after(0, lambda: self._bud_btn.config(state="normal"))

    def _pop_bud(self, reorder_only=False):
        if reorder_only:
            for i in range(len(self._bud_results)):
                self._bud_tree.move(str(i), "", i)
            return
        for row in self._bud_tree.get_children(): self._bud_tree.delete(row)
        for i, r in enumerate(self._bud_results):
            hot = r.get("net_margin", 0) >= 15
            tag = ("hot",) if hot else ("even" if i % 2 == 0 else "odd",)
            self._bud_tree.insert("", "end", iid=str(r["type_id"]), tags=tag, values=(
                r["item"],
                self._fmt(r["pb"],           "isk"),
                self._fmt(r["ps"],           "isk"),
                self._fmt(r["qty"],          "int"),
                self._fmt(r["u_max"],        "int"),
                self._fmt(r["cost_unit"],    "isk"),
                self._fmt(r["pi_net"],       "isk"),
                self._fmt(r["net_margin"],   "pct"),
                self._fmt(r["total_buy"],    "isk"),
                self._fmt(r["total_profit"], "isk"),
                self._fmt(r.get("avg_vol_90", 0), "vol"),
            ))
        # Stats bar
        s = self._bud_stats
        if s:
            util  = s.get("utilization", 0)
            limit = "  ⚠ Limite 150 ordens atingido!" if s.get("limit_hit") else ""
            roi   = (s["total_profit"] / s["total_spent"] * 100
                     if s.get("total_spent", 0) > 0 else 0)
            self._bud_stats_lbl.config(
                fg=COLORS["green"],
                text=(f"  {s['n_items']} itens  |  "
                      f"Capital: {s['total_spent']:,.0f} / {s['budget']:,.0f} ISK "
                      f"({util:.1f}%)  |  "
                      f"Lucro est.: {s['total_profit']:,.0f} ISK  |  "
                      f"ROI: {roi:.1f}%"
                      f"{limit}"))

    def _sort_bud(self, col):
        num = {"pb","ps","qty","u_max","cost_unit","pi_net",
               "net_margin","total_buy","total_profit","avg_vol_90"}
        rev = not self._bud_sort_state.get(col, False)
        self._bud_sort_state = {col: rev}
        self._bud_results.sort(
            key=lambda r: (r.get(col) or 0) if col in num else r.get(col, ""),
            reverse=rev)
        self._set_heading_arrow(self._bud_tree, self._BUD_COLS, col, rev)
        self._pop_bud(reorder_only=True)

    def _export_bud(self):
        headers = [(k, v) for k,v in self._BUD_COLS.items()] + [("region","Região")]
        self._export_table("Orçamento", headers, self._bud_results, "orcamento")

    # ═══════════════════════════════════════════════════════════════════════════
    #  MAPA DA ROTA — integração com eve_route_map.py
    # ═══════════════════════════════════════════════════════════════════════════

    def _open_route_map(self, segments: list):
        """
        Exporta os segmentos para route_export.json e abre o eve_route_map.py.
        Funciona tanto para a Arbitragem quanto para o Otimizador.
        """
        import subprocess
        import sys
        if not segments:
            messagebox.showinfo("Mapa da Rota",
                                "Calcule a Arbitragem ou o Otimizador primeiro.")
            return

        # Serializa os segmentos para JSON (converte tipos não-serializáveis)
        route_file = os.path.join(DATA_DIR, "route_export.json")
        try:
            os.makedirs(DATA_DIR, exist_ok=True)
            with open(route_file, "w", encoding="utf-8") as f:
                json.dump(segments, f, ensure_ascii=False, indent=2, default=str)
        except Exception as exc:
            messagebox.showerror("Erro ao exportar rota", str(exc))
            return

        # Lança o mapa em processo separado
        map_script = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "eve_route_map.py")
        if not os.path.exists(map_script):
            messagebox.showerror(
                "Arquivo não encontrado",
                f"eve_route_map.py não encontrado em:\n{map_script}\n\n"
                "Certifique-se de que ele está na mesma pasta que eve_trade.py.")
            return

        try:
            subprocess.Popen([sys.executable, map_script, route_file],
                             cwd=os.path.dirname(map_script))
            self._set_status(f"Mapa da rota aberto ({len(segments)} segmentos exportados).")
        except Exception as exc:
            messagebox.showerror("Erro ao abrir mapa", str(exc))

    def _arb_open_map(self):
        """Abre o mapa a partir das oportunidades de Arbitragem atuais."""
        if not self._arb_opps:
            messagebox.showinfo("Mapa da Rota",
                                "Calcule a Arbitragem primeiro.")
            return
        # Converte estrutura de Arbitragem para formato esperado pelo mapa
        segments = []
        for opp in self._arb_opps:
            segments.append({
                "Item":         opp.get("item", ""),
                "Origem":       opp.get("loc_origem", ""),
                "Destino":      opp.get("loc_destino", ""),
                "Lucro Total":  opp.get("roi_total", 0),
                "Qtd":          opp.get("qtd_origem", 1),
                "Vol Total (m3)": opp.get("vol_total_m3", 0),
                "Margem (%)":   opp.get("margem_pct", 0),
            })
        self._open_route_map(segments)

    def _opt_open_map(self):
        """Abre o mapa a partir do plano do Otimizador."""
        source = self._opt_plan or []
        if not source:
            messagebox.showinfo("Mapa da Rota",
                                "Execute o Otimizador primeiro.")
            return
        self._open_route_map(source)

    # ═══════════════════════════════════════════════════════════════════════════
    #  ABA 28 — IND. SCANNER  (scanner de lucratividade de fabricação)
    #  Varre todos os blueprints disponíveis, calcula custo de fabricação e
    #  compara com o maior buy order do produto → mostra o que vale fabricar.
    # ═══════════════════════════════════════════════════════════════════════════

    _IND_SCAN_COLS = {
        "product":      "Produto",
        "cost_unit":    "Custo/un (ISK)",
        "buy_max":      "Buy Max (ISK)",
        "profit_unit":  "Lucro/un (ISK)",
        "roi":          "ROI %",
        "mat_cost":     "Mat. Total (ISK)",
        "runs":         "Runs",
        "qty_total":    "Qtd Prod.",
        "source":       "Fonte BP",
    }

    def _build_tab_ind_scan(self):
        p = self._tab_ind_scan

        # ── Painel de parâmetros ──────────────────────────────────────────────
        top = tk.Frame(p, bg=COLORS["panel"],
                       highlightbackground=COLORS["border"], highlightthickness=1)
        top.pack(fill="x", padx=6, pady=(6, 2))

        # Linha 0: parâmetros básicos
        self._lbl(top, "Região",          0, 0)
        self._iscan_reg    = self._combo(top, list(REGIONS.keys()), 0, 1)
        self._lbl(top, "ME (0-10)",       0, 2)
        self._iscan_me     = self._entry(top, "0",  0, 3, 5)
        self._lbl(top, "Runs",            0, 4)
        self._iscan_runs   = self._entry(top, "1",  0, 5, 5)
        self._lbl(top, "Max blueprints",  0, 6)
        self._iscan_max    = self._entry(top, "200", 0, 7, 6)

        # Linha 1: parâmetros financeiros e filtros
        self._lbl(top, "Facility Tax %",  1, 0)
        self._iscan_tax    = self._entry(top, "5",  1, 1, 5)
        self._lbl(top, "Cost Index %",    1, 2)
        self._iscan_ci     = self._entry(top, "5",  1, 3, 5)
        self._lbl(top, "ROI mín %",       1, 4)
        self._iscan_roi    = self._entry(top, "5",  1, 5, 5)
        self._lbl(top, "Lucro mín (ISK)", 1, 6)
        self._iscan_pmin   = self._entry(top, "100000", 1, 7, 12)

        # ── Linha 2: Filtro de Categoria ──────────────────────────────────────
        cat_fr = tk.Frame(top, bg=COLORS["panel"])
        cat_fr.grid(row=2, column=0, columnspan=14, sticky="ew", padx=4, pady=(6, 2))

        tk.Frame(cat_fr, bg=COLORS["border"], height=1).pack(fill="x", pady=(0, 6))

        flt_row = tk.Frame(cat_fr, bg=COLORS["panel"])
        flt_row.pack(fill="x")

        tk.Label(flt_row, text="Categoria do produto:",
                 bg=COLORS["panel"], fg=COLORS["header"],
                 font=("Courier", 10, "bold")).pack(side="left", padx=(0, 6))

        _CAT_OPTIONS = [
            "(Todas as categorias)",
            "Ship",
            "Module",
            "Charge",
            "Drone",
            "Implant",
            "Subsystem",
            "Deployable",
            "Structure",
            "Fighter",
            "Commodity",
            "Ancient Relic",
            "Reaction Formulas",
        ]
        self._iscan_cat_var = tk.StringVar(value="(Todas as categorias)")
        self._iscan_cat_cb  = ttk.Combobox(
            flt_row, textvariable=self._iscan_cat_var,
            values=_CAT_OPTIONS,
            state="readonly", width=28, font=("Courier", 11))
        self._iscan_cat_cb.pack(side="left", padx=4)

        tk.Label(flt_row,
                 text="  Filtra pelo tipo do produto fabricado  |  "
                      "Requer cache enriquecido (CATALOGO → ENRIQUECER CACHE). "
                      "Itens sem categoria registrada sao incluidos mesmo com filtro ativo.",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 9)).pack(side="left", padx=8)

        # ── Linha 3: checkboxes de fonte ──────────────────────────────────────
        chk_fr = tk.Frame(top, bg=COLORS["panel"])
        chk_fr.grid(row=3, column=0, columnspan=14, sticky="w", padx=4, pady=(4, 6))
        tk.Label(chk_fr, text="Fontes de blueprints:",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier", 10)).pack(side="left", padx=(0, 8))
        self._iscan_src_cache  = tk.BooleanVar(value=True)
        self._iscan_src_sde    = tk.BooleanVar(value=True)
        self._iscan_src_item   = tk.BooleanVar(value=True)
        for var, lbl in [
            (self._iscan_src_cache, "Cache SQLite"),
            (self._iscan_src_sde,   "Indice SDE local"),
            (self._iscan_src_item,  "Item Cache (nomes)"),
        ]:
            tk.Checkbutton(chk_fr, text=lbl, variable=var,
                           bg=COLORS["panel"], fg=COLORS["text"],
                           selectcolor=COLORS["border"],
                           activebackground=COLORS["panel"],
                           font=("Courier", 10)).pack(side="left", padx=6)


        # ── Barra de ações ────────────────────────────────────────────────────
        act = tk.Frame(p, bg=COLORS["bg"])
        act.pack(fill="x", padx=6, pady=(3, 1))

        self._iscan_btn = tk.Button(
            act, text="  🔍 ESCANEAR BLUEPRINTS  ",
            command=self._start_ind_scan,
            bg=COLORS["accent"], fg="#000",
            font=("Courier", 12, "bold"), relief="flat", cursor="hand2")
        self._iscan_btn.pack(side="left", padx=4)

        self._iscan_stop_btn = tk.Button(
            act, text=" ⏹ PARAR ",
            command=self._stop_ind_scan,
            bg=COLORS["red"], fg="#fff",
            font=("Courier", 11, "bold"), relief="flat", cursor="hand2",
            state="disabled")
        self._iscan_stop_btn.pack(side="left", padx=4)

        tk.Button(act, text=" EXPORTAR ",
                  command=self._export_ind_scan,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)

        tk.Button(act, text=" ABRIR NA ABA INDUSTRIA ",
                  command=self._iscan_open_ind,
                  bg=COLORS["border"], fg=COLORS["green"],
                  font=("Courier", 11), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)

        self._iscan_spin = ttk.Progressbar(act, mode="indeterminate", length=100)
        self._iscan_spin.pack(side="left", padx=8)

        self._iscan_status = tk.Label(act, text="Pronto. Configure os parâmetros e clique ESCANEAR.",
                                      bg=COLORS["bg"], fg=COLORS["subtext"],
                                      font=("Courier", 10))
        self._iscan_status.pack(side="left", padx=4)

        # ── Barra de estatísticas ─────────────────────────────────────────────
        self._iscan_stats_lbl = tk.Label(
            p, text="",
            bg=COLORS["panel"], fg=COLORS["green"],
            font=("Courier", 10), anchor="w")
        self._iscan_stats_lbl.pack(fill="x", padx=6, pady=(1, 1))

        # ── Treeview de resultados ────────────────────────────────────────────
        cols = [
            ("product",     "Produto",           220),
            ("cost_unit",   "Custo/un (ISK)",     140),
            ("buy_max",     "Buy Max (ISK)",       140),
            ("profit_unit", "Lucro/un (ISK)",      140),
            ("roi",         "ROI %",                80),
            ("mat_cost",    "Mat. Total (ISK)",    150),
            ("runs",        "Runs",                 55),
            ("qty_total",   "Qtd Prod.",             80),
            ("source",      "Fonte BP",             90),
        ]
        res_fr = tk.LabelFrame(p, text="  Resultados  ",
                               bg=COLORS["panel"], fg=COLORS["header"],
                               font=("Courier", 11, "bold"),
                               highlightbackground=COLORS["border"], highlightthickness=1)
        res_fr.pack(fill="both", expand=True, padx=6, pady=(2, 6))

        self._iscan_tree = self._make_tree(res_fr, cols)

        # Sorting ao clicar no cabeçalho
        self._iscan_sort_state: dict = {}
        self._iscan_results:    list = []
        self._iscan_running = False

        for cid, _, _ in cols:
            self._iscan_tree.heading(
                cid, text=dict(self._IND_SCAN_COLS).get(cid, cid),
                command=lambda c=cid: self._sort_ind_scan(c))

        # Tag extra para itens muito lucrativos (ROI >= 20%)
        self._iscan_tree.tag_configure("great", foreground=COLORS["yellow"])
        self._iscan_tree.tag_configure("hot",   foreground=COLORS["green"])

    # ── Controle de execução ──────────────────────────────────────────────────

    def _start_ind_scan(self):
        if self._iscan_running:
            return
        try:
            me_level   = int(self._iscan_me.get().strip())
            runs       = max(1, int(self._iscan_runs.get().strip()))
            fac_tax    = float(self._iscan_tax.get().strip()) / 100.0
            cost_idx   = float(self._iscan_ci.get().strip())  / 100.0
            min_roi    = float(self._iscan_roi.get().strip())
            min_profit = float(self._iscan_pmin.get().strip().replace(",", "").replace(".", "").replace(" ", "") or "0")
            max_bps    = max(1, int(self._iscan_max.get().strip()))
            region_name= self._iscan_reg.get()
            region_id  = REGIONS.get(region_name, 10000002)
        except ValueError as e:
            messagebox.showerror("Parâmetros inválidos", str(e))
            return

        # Lê filtro de categoria
        raw_cat    = self._iscan_cat_var.get()
        filter_cat = "" if raw_cat == "(Todas as categorias)" else raw_cat

        self._iscan_running = True
        self._iscan_results = []
        for row in self._iscan_tree.get_children():
            self._iscan_tree.delete(row)
        self._iscan_stats_lbl.config(text="")
        self._iscan_btn.config(state="disabled")
        self._iscan_stop_btn.config(state="normal")
        self._iscan_spin.start(12)
        self._iscan_status.config(
            text="Coletando lista de blueprints...", fg=COLORS["yellow"])

        threading.Thread(
            target=self._worker_ind_scan,
            args=(me_level, runs, fac_tax, cost_idx,
                  min_roi, min_profit, max_bps,
                  region_id, region_name, filter_cat),
            daemon=True).start()

    def _stop_ind_scan(self):
        self._iscan_running = False
        self._iscan_status.config(text="Parando...", fg=COLORS["red"])

    # ── Worker principal ──────────────────────────────────────────────────────

    def _worker_ind_scan(self, me_level, runs, fac_tax, cost_idx,
                          min_roi, min_profit, max_bps,
                          region_id, region_name, filter_cat):
        """
        Thread de varredura de blueprints.

        1. Coleta IDs de blueprints de 3 fontes (cache SQLite, índice SDE, item_cache)
        2. Para cada blueprint:
           a. fetch_blueprint_materials  → materiais + qty_per_run
           b. fetch_material_prices_bulk → preço de venda dos materiais
           c. _get_product_tid           → type_id do produto fabricado
           d. Filtra por category_name do PRODUTO (se filter_cat ativo)
              — itens sem categoria registrada no cache são incluídos por padrão
           e. fetch_live_price           → buy_max do produto
           f. calculate_industry_cost    → custo total + custo/un
           g. Filtra por ROI e lucro mínimo
        3. Atualiza a tabela progressivamente a cada resultado válido encontrado
        """
        try:
            # ── 1. Coleta blueprints de todas as fontes ────────────────────────
            bp_ids: dict = {}   # {bp_tid: source_label}

            if self._iscan_src_cache.get():
                try:
                    conn = _get_bp_db()
                    rows = conn.execute(
                        "SELECT DISTINCT lookup_tid FROM bp_materials LIMIT ?",
                        (max_bps * 5,)).fetchall()
                    for (tid,) in rows:
                        bp_ids[tid] = "cache"
                except Exception:
                    pass

            if self._iscan_src_sde.get():
                try:
                    conn = _get_bp_db()
                    rows = conn.execute(
                        "SELECT DISTINCT bp_tid FROM hoboleaks_idx "
                        "WHERE activity_id=1 LIMIT ?",
                        (max_bps * 5,)).fetchall()
                    for (tid,) in rows:
                        bp_ids.setdefault(tid, "sde")
                except Exception:
                    pass

            if self._iscan_src_item.get():
                for tid, info in self._item_cache.items():
                    if info.get("name", "").endswith(" Blueprint"):
                        bp_ids.setdefault(tid, "item_cache")

            if not bp_ids:
                self.after(0, lambda: (
                    self._iscan_status.config(
                        text=("Nenhum blueprint encontrado. "
                              "Use a aba INDUSTRIA ao menos uma vez para popular o cache, "
                              "ou baixe o SDE em DADOS & CACHE."),
                        fg=COLORS["red"]),
                    self._iscan_spin.stop(),
                    self._iscan_btn.config(state="normal"),
                    self._iscan_stop_btn.config(state="disabled"),
                ))
                self._iscan_running = False
                return

            all_bps = list(bp_ids.items())[:max_bps]
            total   = len(all_bps)

            cat_lbl = f" [{filter_cat}]" if filter_cat else ""
            self.after(0, lambda n=total, cl=cat_lbl: self._iscan_status.config(
                text=f"Varrendo {n} blueprints{cl} na regiao {region_name}...",
                fg=COLORS["yellow"]))

            # ── 2. Preços ajustados (EIV) — uma requisição para toda a sessão
            try:
                adjusted_prices = fetch_adjusted_prices()
            except Exception:
                adjusted_prices = {}

            # ── 3. Processa cada blueprint ─────────────────────────────────────
            n_done = 0
            n_ok   = 0
            n_skip = 0
            n_err  = 0
            lock   = threading.Lock()

            def _process_one(item):
                nonlocal n_done, n_ok, n_skip, n_err
                if not self._iscan_running:
                    return

                bp_tid, src = item
                try:
                    # 3a. Materiais
                    bp_data = fetch_blueprint_materials(bp_tid, activity_id=1)
                    mats    = bp_data.get("materials", [])
                    qpr     = bp_data.get("qty_per_run", 1)
                    bp_src  = bp_data.get("source", src)

                    if not mats:
                        with lock:
                            n_skip += 1; n_done += 1
                        return

                    # 3b. Resolve produto
                    _ensure_names([bp_tid], self._item_cache)
                    prod_tid = _get_product_tid(bp_tid, bp_data, self._item_cache)
                    if not prod_tid:
                        with lock:
                            n_skip += 1; n_done += 1
                        return

                    # 3c. Filtro de categoria — checado no produto, não no blueprint
                    if filter_cat:
                        prod_cat = self._item_cache.get(prod_tid, {}).get("category_name", "")
                        # Se a categoria do produto é conhecida e não bate → pula
                        # Se é desconhecida (cache não enriquecido) → inclui por padrão
                        if prod_cat and prod_cat != filter_cat:
                            with lock:
                                n_skip += 1; n_done += 1
                            return

                    # 3d. Preços dos materiais (sell_min)
                    mat_tids = [m["type_id"] for m in mats]
                    prices   = fetch_material_prices_bulk(mat_tids, region_id)

                    materials_with_price = []
                    for m in mats:
                        sell = prices.get(m["type_id"], {}).get("sell_min", 0)
                        materials_with_price.append({
                            "type_id":    m["type_id"],
                            "quantity":   m["quantity"],
                            "unit_price": sell,
                        })

                    # 3e. Buy max do produto (receita de venda)
                    prod_price_data = fetch_live_price(region_id, prod_tid)
                    buy_max = prod_price_data.get("buy_max", 0) if prod_price_data else 0
                    if buy_max <= 0:
                        with lock:
                            n_skip += 1; n_done += 1
                        return

                    # 3f. Custo de fabricação (fórmula CCP)
                    result = calculate_industry_cost(
                        materials=materials_with_price,
                        product_qty=qpr,
                        facility_tax=fac_tax,
                        system_cost_index=cost_idx,
                        adjusted_prices=adjusted_prices,
                        me_level=me_level,
                        runs=runs,
                        structure_mat_bonus=0.0,
                        structure_job_bonus=0.0,
                    )

                    cost_unit = result["cost_per_unit"]
                    if cost_unit <= 0:
                        with lock:
                            n_skip += 1; n_done += 1
                        return

                    # 3g. Lucratividade
                    profit_unit = buy_max - cost_unit
                    roi         = (profit_unit / cost_unit) * 100.0

                    if roi < min_roi or profit_unit < min_profit:
                        with lock:
                            n_skip += 1; n_done += 1
                        return

                    # 3h. Nome do produto
                    _ensure_names([prod_tid], self._item_cache)
                    prod_name = self._item_cache.get(prod_tid, {}).get("name", str(prod_tid))

                    row_data = {
                        "bp_tid":      bp_tid,
                        "prod_tid":    prod_tid,
                        "product":     prod_name,
                        "cost_unit":   cost_unit,
                        "buy_max":     buy_max,
                        "profit_unit": profit_unit,
                        "roi":         roi,
                        "mat_cost":    result["total_mat_cost"],
                        "total_cost":  result["total_cost"],
                        "runs":        runs,
                        "qty_total":   result["product_qty_total"],
                        "source":      bp_src,
                        "region":      region_name,
                        "category":    self._item_cache.get(prod_tid, {}).get("category_name", ""),
                    }

                    with lock:
                        n_ok   += 1
                        n_done += 1
                        self._iscan_results.append(row_data)

                    self.after(0, lambda r=row_data, nd=n_done, t=total, ok=n_ok:
                        self._iscan_insert_row(r, nd, t, ok))

                except Exception:
                    with lock:
                        n_err += 1; n_done += 1

            with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
                list(pool.map(_process_one, all_bps))

            # ── 4. Finaliza ────────────────────────────────────────────────────
            def _done(ok=n_ok, skip=n_skip, err=n_err, t=total):
                if ok == 0:
                    msg = (f"Nenhum item lucrativo em {t} blueprints "
                           f"({skip} sem produto/preco, {err} erros). "
                           "Tente reduzir ROI min ou Lucro min.")
                    self._iscan_status.config(text=msg, fg=COLORS["red"])
                else:
                    self._iscan_status.config(
                        text=(f"OK {ok} lucrativos | {t} varridos | "
                              f"{skip} sem dados | {err} erros"),
                        fg=COLORS["green"])
                    self._iscan_update_stats()
                self._iscan_spin.stop()
                self._iscan_btn.config(state="normal")
                self._iscan_stop_btn.config(state="disabled")

            self.after(0, _done)

        except Exception as exc:
            log.exception("worker_ind_scan")
            _err = str(exc)
            self.after(0, lambda e=_err: (
                messagebox.showerror("Erro Scanner", e),
                self._iscan_status.config(text=f"Erro: {e[:80]}", fg=COLORS["red"]),
                self._iscan_spin.stop(),
                self._iscan_btn.config(state="normal"),
                self._iscan_stop_btn.config(state="disabled"),
            ))
        finally:
            self._iscan_running = False

    # ── Atualização de UI ─────────────────────────────────────────────────────

    def _iscan_insert_row(self, r, n_done, total, n_ok):
        """Insere uma linha de resultado na treeview e atualiza status."""
        roi = r["roi"]
        if roi >= 20:
            tag = ("great",)
        elif roi >= 10:
            tag = ("hot",)
        elif n_ok % 2 == 0:
            tag = ("even",)
        else:
            tag = ("odd",)

        iid = str(r["prod_tid"])
        # Remove linha duplicada se o mesmo produto aparecer de outra fonte
        try:
            self._iscan_tree.delete(iid)
        except Exception:
            pass

        self._iscan_tree.insert("", "end", iid=iid, tags=tag, values=(
            r["product"],
            self._fmt(r["cost_unit"],   "isk"),
            self._fmt(r["buy_max"],     "isk"),
            self._fmt(r["profit_unit"], "isk"),
            f"{roi:.1f}%",
            self._fmt(r["mat_cost"],    "isk"),
            r["runs"],
            self._fmt(r["qty_total"],   "int"),
            r["source"],
        ))

        pct = int(n_done / max(total, 1) * 100)
        self._iscan_status.config(
            text=(f"Varrendo... {n_done}/{total} ({pct}%) | "
                  f"{n_ok} lucrativos encontrados"),
            fg=COLORS["yellow"])

    def _iscan_update_stats(self):
        """Atualiza a barra de estatísticas com resumo dos resultados."""
        if not self._iscan_results:
            return
        best   = max(self._iscan_results, key=lambda r: r["roi"])
        avg_roi= sum(r["roi"] for r in self._iscan_results) / len(self._iscan_results)
        total_p= sum(r["profit_unit"] * r["qty_total"] for r in self._iscan_results)
        self._iscan_stats_lbl.config(
            text=(f"  {len(self._iscan_results)} itens lucrativos  |  "
                  f"Melhor ROI: {best['roi']:.1f}% ({best['product']})  |  "
                  f"ROI médio: {avg_roi:.1f}%  |  "
                  f"Lucro potencial total: {total_p:,.0f} ISK"),
            fg=COLORS["green"])

    # ── Ordenação ─────────────────────────────────────────────────────────────

    def _sort_ind_scan(self, col):
        num_cols = {"cost_unit", "buy_max", "profit_unit", "roi",
                    "mat_cost", "runs", "qty_total"}
        rev = not self._iscan_sort_state.get(col, False)
        self._iscan_sort_state = {col: rev}
        self._iscan_results.sort(
            key=lambda r: (r.get(col) or 0) if col in num_cols else r.get(col, ""),
            reverse=rev)
        # Reordena apenas posições (sem recriar linhas — O(n), não O(n²))
        for i, r in enumerate(self._iscan_results):
            self._iscan_tree.move(str(r["prod_tid"]), "", i)
        arrow = " ▼" if rev else " ▲"
        for cid, label in self._IND_SCAN_COLS.items():
            self._iscan_tree.heading(
                cid, text=label + (arrow if cid == col else ""))

    # ── Abrir item selecionado na aba Indústria ───────────────────────────────

    def _iscan_open_ind(self):
        sel = self._iscan_tree.selection()
        if not sel:
            messagebox.showinfo("IND. SCANNER",
                                "Selecione um item na tabela primeiro.")
            return
        iid = sel[0]
        # Encontra o blueprint correspondente ao produto selecionado
        prod_tid = int(iid)
        row = next((r for r in self._iscan_results if r["prod_tid"] == prod_tid), None)
        if not row:
            return
        bp_tid = row.get("bp_tid", prod_tid)
        # Preenche campos da aba Indústria e navega para ela
        self._ind_prod_tid.set(str(bp_tid))
        self._ind_want_qty.set(str(row["qty_total"]))
        self._nb.select(self._tab_ind)
        self.after(80, self._ind_auto_load)

    # ── Exportar ──────────────────────────────────────────────────────────────

    def _export_ind_scan(self):
        if not self._iscan_results:
            messagebox.showinfo("IND. SCANNER", "Nenhum resultado para exportar.")
            return
        headers = [(k, v) for k, v in self._IND_SCAN_COLS.items()] + [("region", "Região")]
        self._export_table("IND. SCANNER", headers, self._iscan_results, "ind_scanner")

if __name__ == "__main__":
    app = EVETradeApp()
    app.mainloop()