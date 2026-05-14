"""
eve_route_map.py -- Mapa interativo de rotas EVE Online com otimizador de entregas.

Funcionalidades:
  - Visualiza rota entre origem e destino no mapa estelar
  - OTIMIZAR ENTREGAS: calcula a ordem mais eficiente para visitar
    todos os sistemas de destino, minimizando pulos totais (nearest-neighbor TSP)
  - Exibe painel de entregas sequencial com itens por parada
  - Aceita JSON (route_export), CSV e XLSX do eve_trade

Dependencias: requests  (pip install requests)
              openpyxl  (pip install openpyxl) -- para XLSX
"""

import concurrent.futures
import csv
import json
import math
import os
import sqlite3
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ── Caminhos ──────────────────────────────────────────────────────────────────
ESI_BASE       = "https://esi.evetech.net/latest"
DATA_DIR       = os.path.join(os.path.dirname(os.path.abspath(__file__)), "eve_data")
ROUTE_FILE     = os.path.join(DATA_DIR, "route_export.json")
JUMP_CACHE_DB  = os.path.join(DATA_DIR, "jump_cache.db")
JUMP_CACHE_TTL = 30 * 86_400   # 30 dias — stargates quase nunca mudam

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
    "route":     "#ff8c00",
    "gate":      "#1e3a5a",
    "sys_orig":  "#00ff88",
    "sys_dest":  "#ff4455",
    "sys_stop":  "#aa44ff",
    "sys_hi":    "#ffd700",
    "canvas_bg": "#060c16",
}

_SEC_COLORS = [
    (0.9, "#2255ff"), (0.7, "#00aaff"), (0.5, "#00ccaa"),
    (0.4, "#ffdd00"), (0.3, "#ff8800"), (0.2, "#ff4400"), (0.0, "#ff2222"),
]
def sec_color(sec: float) -> str:
    for threshold, color in _SEC_COLORS:
        if sec >= threshold:
            return color
    return "#ff2222"

# ── Cache em memória ──────────────────────────────────────────────────────────
_sys_cache:   dict = {}
_gate_cache:  dict = {}
_name_id:     dict = {}
_route_cache: dict = {}
_jump_mem:    dict = {}   # cache RAM de pulos (espelho do SQLite para a sessão)
_session = None

# ── Cache SQLite de pulos ─────────────────────────────────────────────────────
_jdb_lock = threading.Lock()
_jdb_conn: sqlite3.Connection | None = None

def _get_jump_db() -> sqlite3.Connection:
    """
    Retorna conexão singleton ao banco SQLite de pulos.
    Cria o banco e a tabela se ainda não existirem.
    Thread-safe via _jdb_lock.
    """
    global _jdb_conn
    if _jdb_conn is not None:
        return _jdb_conn
    os.makedirs(DATA_DIR, exist_ok=True)
    conn = sqlite3.connect(JUMP_CACHE_DB, check_same_thread=False)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS jump_cache (
            origin_id  INTEGER NOT NULL,
            dest_id    INTEGER NOT NULL,
            jumps      INTEGER NOT NULL,
            cached_at  REAL    NOT NULL,
            PRIMARY KEY (origin_id, dest_id)
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_jc_pair ON jump_cache(origin_id, dest_id)")
    conn.commit()
    _jdb_conn = conn
    return conn

def _jdb_get(a: int, b: int) -> int | None:
    """Lê pulos do SQLite. Retorna None se ausente ou expirado (TTL 30 dias)."""
    lo, hi = (a, b) if a < b else (b, a)
    # Verifica RAM primeiro (mais rápido)
    if (lo, hi) in _jump_mem:
        return _jump_mem[(lo, hi)]
    try:
        with _jdb_lock:
            row = _get_jump_db().execute(
                "SELECT jumps, cached_at FROM jump_cache WHERE origin_id=? AND dest_id=?",
                (lo, hi)).fetchone()
        if row:
            jumps, cached_at = row
            if time.time() - cached_at < JUMP_CACHE_TTL:
                _jump_mem[(lo, hi)] = jumps   # popula RAM
                return jumps
            # Expirado — apaga para forçar nova busca
            with _jdb_lock:
                _get_jump_db().execute(
                    "DELETE FROM jump_cache WHERE origin_id=? AND dest_id=?", (lo, hi))
                _get_jump_db().commit()
    except Exception:
        pass
    return None

def _jdb_set(a: int, b: int, jumps: int):
    """Salva pulos no SQLite e na RAM."""
    lo, hi = (a, b) if a < b else (b, a)
    _jump_mem[(lo, hi)] = jumps
    try:
        with _jdb_lock:
            _get_jump_db().execute(
                "INSERT OR REPLACE INTO jump_cache (origin_id, dest_id, jumps, cached_at) "
                "VALUES (?, ?, ?, ?)",
                (lo, hi, jumps, time.time()))
            _get_jump_db().commit()
    except Exception:
        pass

def get_cache_stats() -> dict:
    """Retorna estatísticas do cache SQLite para exibição."""
    try:
        with _jdb_lock:
            conn  = _get_jump_db()
            total = conn.execute("SELECT COUNT(*) FROM jump_cache").fetchone()[0]
            fresh = conn.execute(
                "SELECT COUNT(*) FROM jump_cache WHERE cached_at > ?",
                (time.time() - JUMP_CACHE_TTL,)).fetchone()[0]
            oldest = conn.execute(
                "SELECT MIN(cached_at) FROM jump_cache").fetchone()[0]
        return {
            "total":  total,
            "fresh":  fresh,
            "oldest": oldest,
        }
    except Exception:
        return {"total": 0, "fresh": 0, "oldest": None}

def _get_session():
    global _session
    if _session is None:
        _session = requests.Session()
        _session.headers.update({"User-Agent": "EVE Trade Suite / route_map"})
    return _session

def esi_get(path, params=None):
    if not HAS_REQUESTS:
        return None
    try:
        r = _get_session().get(ESI_BASE + path, params=params, timeout=20)
        if r.status_code == 200:
            return r.json()
    except Exception:
        pass
    return None

def esi_post(path, body):
    if not HAS_REQUESTS:
        return None
    try:
        r = _get_session().post(ESI_BASE + path, json=body, timeout=20)
        if r.status_code == 200:
            return r.json()
    except Exception:
        pass
    return None

# ── Nomes de sistemas ─────────────────────────────────────────────────────────

def station_to_system_name(station: str) -> str:
    """
    "Jita IV - Moon 4 - Caldari Navy Assembly Plant" -> "Jita"
    "Airaken VI - Moon 1 - Sisters of EVE Bureau"    -> "Airaken"
    """
    if not station:
        return ""
    return station.split(" - ")[0].split(" ")[0].strip()

def resolve_system_name(raw_station: str) -> int:
    sys_name = station_to_system_name(raw_station)
    if not sys_name:
        return 0
    if sys_name in _name_id:
        return _name_id[sys_name]
    data = esi_post("/universe/ids/", [sys_name])
    if data and "systems" in data and data["systems"]:
        sid = data["systems"][0]["id"]
        _name_id[sys_name] = sid
        return sid
    return 0

def fetch_system(system_id: int) -> dict:
    if system_id in _sys_cache:
        return _sys_cache[system_id]
    data = esi_get(f"/universe/systems/{system_id}/")
    if not data:
        return {}
    pos = data.get("position", {})
    info = {
        "id":        system_id,
        "name":      data.get("name", str(system_id)),
        "sec":       round(data.get("security_status", 0.0), 1),
        "x":         pos.get("x", 0),
        "y":         pos.get("y", 0),
        "z":         pos.get("z", 0),
        "stargates": data.get("stargates", []),
        "planets":   len(data.get("planets", [])),
    }
    _sys_cache[system_id] = info
    _name_id[info["name"]] = system_id
    return info

def fetch_stargate_dest(gate_id: int) -> int:
    if gate_id in _gate_cache:
        return _gate_cache[gate_id]
    data = esi_get(f"/universe/stargates/{gate_id}/")
    if data:
        dest = data.get("destination", {}).get("system_id", 0)
        _gate_cache[gate_id] = dest
        return dest
    return 0

def fetch_route_ids(origin_id: int, dest_id: int) -> list:
    key = (origin_id, dest_id)
    if key in _route_cache:
        return _route_cache[key]
    data = esi_get(f"/route/{origin_id}/{dest_id}/", {"flag": "shortest"})
    route = data if isinstance(data, list) else []
    _route_cache[key] = route
    return route

def fetch_jumps(origin_id: int, dest_id: int) -> int:
    """
    Retorna numero de pulos entre dois sistemas.

    Ordem de consulta:
      1. Cache RAM (_jump_mem) — O(1), sem I/O
      2. Cache SQLite (jump_cache.db, TTL 30 dias) — evita ESI em execuções futuras
      3. ESI /route/ — apenas quando o par não está em cache
         Salva resultado no SQLite para próximas execuções.

    Retorna 999 se o caminho não for encontrado.
    """
    if origin_id == dest_id:
        return 0

    # 1 + 2: RAM e SQLite
    cached = _jdb_get(origin_id, dest_id)
    if cached is not None:
        return cached

    # 3: ESI
    route  = fetch_route_ids(origin_id, dest_id)
    jumps  = max(0, len(route) - 1) if route else 999
    _jdb_set(origin_id, dest_id, jumps)
    return jumps

# ── Normalizacao de dados ─────────────────────────────────────────────────────

def _normalise(row: dict) -> dict:
    def g(*keys):
        for k in keys:
            v = row.get(k)
            if v not in (None, "", "─── TOTAIS ───"):
                return v
        return ""
    return {
        "Item":           g("Item", "item", "name"),
        "Origem":         g("Origem", "loc_origem", "Estacao Origem"),
        "Destino":        g("Destino", "loc_destino", "Estacao Destino"),
        "Pulos":          g("Pulos", "pulos"),
        "Qtd":            g("Qtd", "qtd_origem"),
        "Vol Total (m3)": g("Vol Total (m3)", "vol_total_m3", "Vol (m\u00b3)", "volume"),
        "Lucro Total":    g("Lucro Total", "roi_total", "Lucro Estimado", "lucro"),
        "Capital":        g("Capital (ISK)", "Capital", "capital"),
        "Preco Compra":   g("Preco Compra", "Preco Compra (ISK)", "preco_origem"),
        "Margem":         g("Margem (%)", "Margem %", "margem_pct"),
    }

def load_route_file(path: str) -> list:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".json":
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        rows = data if isinstance(data, list) else data.get("segments", [])
        return [_normalise(r) for r in rows]

    if ext == ".csv":
        with open(path, newline="", encoding="utf-8-sig") as f:
            return [_normalise(dict(r)) for r in csv.DictReader(f)]

    if ext in (".xlsx", ".xlsm"):
        if not HAS_OPENPYXL:
            raise ImportError("pip install openpyxl")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        it = ws.iter_rows(values_only=True)
        headers = [str(h or "").strip() for h in next(it)]
        rows = []
        for row in it:
            d = {headers[i]: (row[i] if i < len(row) else "") for i in range(len(headers))}
            rows.append(_normalise(d))
        return rows

    raise ValueError(f"Formato nao suportado: {ext}")

# ── Otimizador de rota multi-destino ─────────────────────────────────────────

def _build_distance_matrix(all_ids: list, progress_cb=None) -> dict:
    """
    Constrói a matriz completa de pulos entre todos os pares de system_ids.
    Usa o SQLite como cache primário e consulta a ESI apenas para pares ausentes.
    Faz as consultas ESI em paralelo (8 workers) para minimizar tempo de rede.

    Retorna: {(id_a, id_b): jumps, ...}  — sempre com id_a < id_b
    """
    n      = len(all_ids)
    matrix = {}

    # 1. Carrega o que já está no SQLite (cache hit instantâneo)
    pairs_needed = []
    for i in range(n):
        for j in range(i + 1, n):
            a, b   = all_ids[i], all_ids[j]
            lo, hi = (a, b) if a < b else (b, a)
            cached = _jdb_get(lo, hi)
            if cached is not None:
                matrix[(lo, hi)] = cached
            else:
                pairs_needed.append((lo, hi))

    if not pairs_needed:
        return matrix   # 100% cache hit — zero chamadas ESI

    if progress_cb:
        progress_cb(f"Cache: {len(matrix)} pares prontos | "
                    f"ESI: {len(pairs_needed)} pares novos...")

    # 2. Consulta ESI em paralelo apenas para os pares ausentes
    lock = threading.Lock()

    def _fetch_pair(pair):
        a, b   = pair
        route  = fetch_route_ids(a, b)
        jumps  = max(0, len(route) - 1) if route else 999
        _jdb_set(a, b, jumps)
        with lock:
            matrix[(a, b)] = jumps

    total_esi = len(pairs_needed)
    done      = [0]

    def _fetch_and_count(pair):
        _fetch_pair(pair)
        with lock:
            done[0] += 1
            if progress_cb and done[0] % 10 == 0:
                pct = int(done[0] / total_esi * 100)
                progress_cb(f"Baixando {done[0]}/{total_esi} pares via ESI ({pct}%)...")

    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as pool:
        list(pool.map(_fetch_and_count, pairs_needed))

    return matrix


def _matrix_get(matrix: dict, a: int, b: int) -> int:
    """Lê da matriz simétrica independente da ordem dos IDs."""
    if a == b:
        return 0
    lo, hi = (a, b) if a < b else (b, a)
    return matrix.get((lo, hi), 999)


def nearest_neighbor_tsp(origin_id: int, dest_ids: list,
                          matrix: dict,
                          progress_cb=None) -> list:
    """
    Heurística Nearest-Neighbor: parte da origem e visita sempre
    o destino não visitado com menos pulos do ponto atual.

    Complexidade: O(n²) consultas à matriz (já em memória, sem I/O).
    """
    if not dest_ids:
        return []

    remaining = list(dest_ids)
    ordered   = []
    current   = origin_id
    total     = len(dest_ids)

    while remaining:
        if progress_cb:
            visited = total - len(remaining)
            progress_cb(f"Nearest-Neighbor: {visited}/{total} paradas...")

        nearest = min(remaining,
                      key=lambda s: _matrix_get(matrix, current, s))
        ordered.append(nearest)
        remaining.remove(nearest)
        current = nearest

    return ordered


def two_opt_improve(route: list, origin_id: int,
                    matrix: dict, progress_cb=None) -> tuple[list, int]:
    """
    Melhoria 2-opt: tenta todas as trocas de pares de arestas e aceita
    sempre que o custo total (pulos) diminuir. Itera até não haver melhora.

    Uma troca 2-opt reverte o segmento [i+1 .. j] da rota:
        ... A -> B -> C -> D -> ...
        ... A -> C -> B -> D -> ...   (se dist(A,C) + dist(B,D) < dist(A,B) + dist(C,D))

    Complexidade por passo: O(n²) comparações sobre a matriz em RAM.
    Convergência típica: 2-5 passos para instâncias com 30-80 paradas.

    Retorna: (rota_melhorada, pulos_economizados)
    """
    if len(route) < 4:
        return route, 0

    # Sequência completa incluindo origem no início
    full = [origin_id] + list(route)
    n    = len(full)

    def route_cost(r):
        return sum(_matrix_get(matrix, r[i], r[i+1]) for i in range(len(r)-1))

    best      = list(full)
    best_cost = route_cost(best)
    improved  = True
    passes    = 0
    total_saved = 0

    while improved:
        improved = False
        passes  += 1
        if progress_cb:
            progress_cb(f"2-opt passo {passes}: custo atual = {best_cost} pulos...")

        for i in range(1, n - 2):
            for j in range(i + 1, n - 1):
                # Custo das duas arestas atuais
                cost_before = (_matrix_get(matrix, best[i-1], best[i]) +
                               _matrix_get(matrix, best[j],   best[j+1]))
                # Custo das duas arestas após a reversão
                cost_after  = (_matrix_get(matrix, best[i-1], best[j]) +
                               _matrix_get(matrix, best[i],   best[j+1]))

                if cost_after < cost_before:
                    # Aceita a troca: reverte o segmento [i..j]
                    best[i:j+1] = best[i:j+1][::-1]
                    saving       = cost_before - cost_after
                    best_cost   -= saving
                    total_saved += saving
                    improved     = True

    # Remove a origem do início — o chamador já sabe qual é
    return best[1:], total_saved


def optimise_route(origin_id: int, dest_ids: list,
                   progress_cb=None) -> tuple[list, dict]:
    """
    Pipeline completo de otimização:
      1. Constrói matriz de distâncias (cache SQLite + ESI paralelo)
      2. Nearest-Neighbor para rota inicial
      3. 2-opt para melhoria iterativa

    Retorna: (rota_ordenada, stats)
    stats = {
        "nn_cost":    int,   pulos antes do 2-opt
        "final_cost": int,   pulos após o 2-opt
        "saved":      int,   pulos economizados
        "passes":     int,   iterações do 2-opt
        "cache_hits": int,   pares lidos do SQLite
        "esi_calls":  int,   pares consultados na ESI
    }
    """
    if not dest_ids:
        return [], {}

    all_ids = [origin_id] + list(dest_ids)

    # Estatísticas de cache antes da matriz
    stats_before = get_cache_stats()

    # 1. Matriz completa
    if progress_cb:
        progress_cb("Construindo matriz de distâncias...")
    matrix = _build_distance_matrix(all_ids, progress_cb)

    stats_after  = get_cache_stats()
    cache_hits   = stats_after["total"] - stats_before["total"]

    # 2. Nearest-Neighbor
    if progress_cb:
        progress_cb("Rodando Nearest-Neighbor TSP...")
    nn_route = nearest_neighbor_tsp(origin_id, dest_ids, matrix, progress_cb)
    nn_cost  = sum(_matrix_get(matrix, ([origin_id] + nn_route)[i],
                               ([origin_id] + nn_route)[i+1])
                   for i in range(len(nn_route)))

    # 3. 2-opt
    if progress_cb:
        progress_cb("Rodando melhoria 2-opt...")
    opt_route, saved = two_opt_improve(nn_route, origin_id, matrix, progress_cb)
    final_cost       = nn_cost - saved

    n_pairs     = len(all_ids) * (len(all_ids) - 1) // 2
    esi_calls   = max(0, n_pairs - cache_hits)

    return opt_route, {
        "nn_cost":    nn_cost,
        "final_cost": final_cost,
        "saved":      saved,
        "cache_hits": cache_hits,
        "esi_calls":  esi_calls,
        "matrix_size": n_pairs,
    }


def build_delivery_plan(segments: list, sys_name_map: dict,
                         progress_cb=None) -> dict:
    """
    Agrupa os segmentos por sistema de destino e calcula a rota
    mais eficiente usando NN + 2-opt.

    Retorno:
    {
      "origin_id":    int,
      "origin_name":  str,
      "stops": [
        {
          "seq":               int,
          "system_id":         int,
          "system_name":       str,
          "jumps_from_prev":   int,
          "jumps_from_origin": int,
          "cumulative_jumps":  int,
          "stations":          [str, ...],
          "items":             [dict, ...],
          "total_vol":         float,
          "total_profit":      float,
          "total_capital":     float,
          "n_items":           int,
        }, ...
      ],
      "total_jumps":   int,
      "total_items":   int,
      "total_profit":  float,
      "total_capital": float,
      "opt_stats":     dict,   # estatísticas do otimizador
    }
    """
    valid = [s for s in segments
             if s.get("Origem") and s.get("Destino")
             and s.get("Item") not in (None, "", "\u2500\u2500\u2500 TOTAIS \u2500\u2500\u2500")]

    # Origem = sistema mais frequente entre as origens
    origin_counts: dict = {}
    for s in valid:
        o = station_to_system_name(s["Origem"])
        origin_counts[o] = origin_counts.get(o, 0) + 1
    origin_name = max(origin_counts, key=origin_counts.get)
    origin_id   = sys_name_map.get(origin_name, 0)

    if not origin_id:
        raise ValueError(f"Sistema de origem '{origin_name}' nao resolvido.")

    # Agrupa por sistema de destino
    stops_map: dict = {}
    for s in valid:
        dest_sys = station_to_system_name(s.get("Destino", ""))
        if not dest_sys:
            continue
        if dest_sys not in stops_map:
            stops_map[dest_sys] = {
                "system_name": dest_sys,
                "system_id":   sys_name_map.get(dest_sys, 0),
                "stations":    set(),
                "items":       [],
            }
        stops_map[dest_sys]["stations"].add(s.get("Destino", ""))
        stops_map[dest_sys]["items"].append(s)

    stops_map = {k: v for k, v in stops_map.items() if v["system_id"]}

    if not stops_map:
        raise ValueError("Nenhum sistema de destino resolvido via ESI.")

    dest_ids = [v["system_id"] for v in stops_map.values()]

    # Rota otimizada: NN + 2-opt
    ordered_ids, opt_stats = optimise_route(origin_id, dest_ids, progress_cb)

    # Monta plano de paradas
    id_to_info = {v["system_id"]: v for v in stops_map.values()}
    cumulative = 0
    prev_id    = origin_id
    stops      = []

    def _num(v):
        try: return float(str(v or 0).replace(",", "").replace(" ", ""))
        except: return 0.0

    for seq, sid in enumerate(ordered_ids, 1):
        info   = id_to_info.get(sid, {})
        j_prev = fetch_jumps(prev_id, sid)
        j_orig = fetch_jumps(origin_id, sid)
        cumulative += j_prev
        items  = info.get("items", [])
        stops.append({
            "seq":               seq,
            "system_id":         sid,
            "system_name":       info.get("system_name", str(sid)),
            "jumps_from_prev":   j_prev,
            "jumps_from_origin": j_orig,
            "cumulative_jumps":  cumulative,
            "stations":          sorted(info.get("stations", set())),
            "items":             items,
            "total_vol":         sum(_num(i.get("Vol Total (m3)")) for i in items),
            "total_profit":      sum(_num(i.get("Lucro Total")) for i in items),
            "total_capital":     sum(_num(i.get("Capital")) for i in items),
            "n_items":           len(items),
        })
        prev_id = sid

    return {
        "origin_id":    origin_id,
        "origin_name":  origin_name,
        "stops":        stops,
        "total_jumps":  cumulative,
        "total_items":  sum(s["n_items"] for s in stops),
        "total_profit": sum(s["total_profit"] for s in stops),
        "total_capital":sum(s["total_capital"] for s in stops),
        "opt_stats":    opt_stats,
    }

# ── Aplicacao ─────────────────────────────────────────────────────────────────

class EVERouteMap(tk.Tk):
    def __init__(self, route_file=None):
        super().__init__()
        self.title("EVE Route Map")
        self.configure(bg=COLORS["bg"])
        self.geometry("1500x900")
        self.minsize(1000, 600)

        self._route_data    = []
        self._all_systems   = {}
        self._route_ids     = []
        self._gate_edges    = []
        self._canvas_nodes  = {}
        self._origin_ids    = set()
        self._dest_ids      = set()
        self._norm_pos      = {}
        self._selected_sys  = 0
        self._loading       = False

        # Plano de entregas otimizado
        self._delivery_plan  = None   # dict retornado por build_delivery_plan
        self._delivery_mode  = False  # True = exibe rota de entregas no canvas
        self._selected_stop  = -1    # indice da parada selecionada (-1 = nenhuma)

        self._scale    = 1.0
        self._offset_x = 0.0
        self._offset_y = 0.0
        self._drag_start = None

        self._build_ui()
        self.after(300, lambda: self._load_file(route_file or ROUTE_FILE))

    # ── Construcao da UI ──────────────────────────────────────────────────────

    def _build_ui(self):
        # Barra de controle superior
        top = tk.Frame(self, bg=COLORS["panel"],
                       highlightbackground=COLORS["border"], highlightthickness=1)
        top.pack(fill="x")

        tk.Label(top, text="  EVE ROUTE MAP  ",
                 bg=COLORS["panel"], fg=COLORS["header"],
                 font=("Courier", 13, "bold")).pack(side="left", padx=6, pady=6)

        for lbl, cmd, fg in [
            (" ABRIR ARQUIVO ", self._open_file_dialog, COLORS["accent"]),
            (" RECARREGAR ",    lambda: self._load_file(ROUTE_FILE), COLORS["yellow"]),
            (" CENTRALIZAR ",   self._center_view,      COLORS["text"]),
        ]:
            tk.Button(top, text=lbl, command=cmd,
                      bg=COLORS["border"], fg=fg,
                      font=("Courier", 10), relief="flat", cursor="hand2"
                      ).pack(side="left", padx=4, pady=6)

        # Separador
        tk.Frame(top, bg=COLORS["border"], width=2).pack(side="left", fill="y",
                                                          padx=6, pady=4)

        # Botao principal de otimizacao
        self._opt_btn = tk.Button(
            top, text="  OTIMIZAR ENTREGAS  ",
            command=self._start_optimize,
            bg=COLORS["sys_stop"], fg="#fff",
            font=("Courier", 11, "bold"), relief="flat", cursor="hand2")
        self._opt_btn.pack(side="left", padx=4, pady=6)

        self._view_btn = tk.Button(
            top, text=" VER MAPA SIMPLES ",
            command=self._toggle_view,
            bg=COLORS["border"], fg=COLORS["subtext"],
            font=("Courier", 10), relief="flat", cursor="hand2",
            state="disabled")
        self._view_btn.pack(side="left", padx=4, pady=6)

        # Legenda
        leg = tk.Frame(top, bg=COLORS["panel"])
        leg.pack(side="left", padx=12)
        for col, lbl in [(COLORS["sys_orig"],"Origem"),(COLORS["sys_dest"],"Destino"),
                          (COLORS["sys_stop"],"Parada"),(COLORS["route"],"Rota")]:
            tk.Label(leg, text="O", fg=col, bg=COLORS["panel"],
                     font=("Courier", 10)).pack(side="left", padx=1)
            tk.Label(leg, text=lbl, fg=COLORS["subtext"],
                     bg=COLORS["panel"], font=("Courier", 8)).pack(side="left", padx=(0,5))

        self._spin = ttk.Progressbar(top, mode="indeterminate", length=80)
        self._spin.pack(side="left", padx=8)
        self._status = tk.Label(top, text="Aguardando...",
                                bg=COLORS["panel"], fg=COLORS["subtext"],
                                font=("Courier", 10))
        self._status.pack(side="left", padx=4)

        # Corpo: canvas + painel direito (abas)
        body = tk.Frame(self, bg=COLORS["bg"])
        body.pack(fill="both", expand=True)

        self._canvas = tk.Canvas(body, bg=COLORS["canvas_bg"],
                                  highlightthickness=0, cursor="crosshair")
        self._canvas.pack(side="left", fill="both", expand=True)

        # Painel direito com notebook (Resumo / Entregas)
        right = tk.Frame(body, bg=COLORS["panel"], width=360,
                         highlightbackground=COLORS["border"], highlightthickness=1)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)

        nb = ttk.Notebook(right)
        nb.pack(fill="both", expand=True)

        # Aba 1: Resumo
        tab_resumo = tk.Frame(nb, bg=COLORS["panel"])
        nb.add(tab_resumo, text="  Resumo  ")
        self._build_tab_resumo(tab_resumo)

        # Aba 2: Plano de Entregas
        tab_entrega = tk.Frame(nb, bg=COLORS["panel"])
        nb.add(tab_entrega, text="  Rota de Entregas  ")
        self._build_tab_entregas(tab_entrega)

        self._nb = nb

        self._statusbar = tk.Label(
            self,
            text="Scroll=zoom  |  Arraste=mover  |  Clique no sistema=info  "
                 "|  OTIMIZAR ENTREGAS=calcula rota multi-destino",
            bg=COLORS["border"], fg=COLORS["subtext"],
            font=("Courier", 9), anchor="w")
        self._statusbar.pack(fill="x", side="bottom")

        self._canvas.bind("<ButtonPress-1>",  self._on_click)
        self._canvas.bind("<B1-Motion>",       self._on_drag)
        self._canvas.bind("<ButtonRelease-1>", self._on_release)
        self._canvas.bind("<MouseWheel>",       self._on_scroll)
        self._canvas.bind("<Button-4>",         self._on_scroll)
        self._canvas.bind("<Button-5>",         self._on_scroll)
        self._canvas.bind("<Configure>",        lambda e: self._redraw())

    def _build_tab_resumo(self, parent):
        def sep(): tk.Frame(parent, bg=COLORS["border"], height=1).pack(fill="x")
        def hdr(t):
            tk.Label(parent, text=f"  {t}  ", bg=COLORS["panel"],
                     fg=COLORS["header"], font=("Courier", 10, "bold")).pack(fill="x", pady=(6,2))
        hdr("RESUMO DA ROTA")
        sep()
        self._info_route = tk.Label(parent, text="Nenhuma rota carregada.",
                                     bg=COLORS["panel"], fg=COLORS["subtext"],
                                     font=("Courier", 10), justify="left",
                                     anchor="nw", wraplength=340)
        self._info_route.pack(fill="x", padx=8, pady=6, anchor="nw")
        sep()
        hdr("SISTEMA SELECIONADO")
        self._info_sys = tk.Label(parent, text="Clique em um sistema no mapa.",
                                   bg=COLORS["panel"], fg=COLORS["subtext"],
                                   font=("Courier", 10), justify="left",
                                   anchor="nw", wraplength=340)
        self._info_sys.pack(fill="x", padx=8, pady=4, anchor="nw")
        sep()
        hdr("ITENS NO ARQUIVO")
        lf = tk.Frame(parent, bg=COLORS["panel"])
        lf.pack(fill="both", expand=True, padx=4, pady=4)
        self._item_list = tk.Listbox(lf, bg=COLORS["border"], fg=COLORS["text"],
                                      font=("Courier", 9), relief="flat", bd=0,
                                      selectbackground=COLORS["accent"],
                                      selectforeground="#000", activestyle="none")
        vsb = ttk.Scrollbar(lf, orient="vertical", command=self._item_list.yview)
        self._item_list.configure(yscrollcommand=vsb.set)
        self._item_list.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

    def _build_tab_entregas(self, parent):
        tk.Label(parent, text="  PLANO DE ENTREGAS OTIMIZADO  ",
                 bg=COLORS["panel"], fg=COLORS["header"],
                 font=("Courier", 10, "bold")).pack(fill="x", pady=(8, 2))
        tk.Frame(parent, bg=COLORS["border"], height=1).pack(fill="x")

        self._plan_summary = tk.Label(
            parent,
            text="Clique em  OTIMIZAR ENTREGAS  para calcular\na rota mais eficiente.",
            bg=COLORS["panel"], fg=COLORS["subtext"],
            font=("Courier", 10), justify="left", anchor="nw", wraplength=340)
        self._plan_summary.pack(fill="x", padx=8, pady=6, anchor="nw")

        tk.Frame(parent, bg=COLORS["border"], height=1).pack(fill="x")

        # Lista de paradas
        lf = tk.Frame(parent, bg=COLORS["panel"])
        lf.pack(fill="both", expand=True, padx=4, pady=4)

        self._stop_list = tk.Listbox(
            lf, bg=COLORS["border"], fg=COLORS["text"],
            font=("Courier", 9), relief="flat", bd=0,
            selectbackground=COLORS["sys_stop"],
            selectforeground="#fff", activestyle="none")
        vsb2 = ttk.Scrollbar(lf, orient="vertical", command=self._stop_list.yview)
        self._stop_list.configure(yscrollcommand=vsb2.set)
        self._stop_list.pack(side="left", fill="both", expand=True)
        vsb2.pack(side="right", fill="y")
        self._stop_list.bind("<<ListboxSelect>>", self._on_stop_select)

        # Detalhe da parada selecionada
        tk.Frame(parent, bg=COLORS["border"], height=1).pack(fill="x")
        self._stop_detail = tk.Label(
            parent, text="Selecione uma parada acima.",
            bg=COLORS["panel"], fg=COLORS["subtext"],
            font=("Courier", 9), justify="left", anchor="nw", wraplength=340)
        self._stop_detail.pack(fill="x", padx=8, pady=6, anchor="nw")

    # ── Carregamento ──────────────────────────────────────────────────────────

    def _open_file_dialog(self):
        ftypes = [
            ("Todos os formatos", "*.json *.csv *.xlsx *.xlsm"),
            ("JSON",  "*.json"), ("CSV",   "*.csv"),
            ("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*"),
        ]
        path = filedialog.askopenfilename(title="Abrir arquivo de rota",
                                          filetypes=ftypes)
        if path:
            self._load_file(path)

    def _load_file(self, path):
        if not path or not os.path.exists(path):
            self._status.config(
                text="Aguardando rota (route_export.json nao encontrado).",
                fg=COLORS["subtext"])
            return
        if self._loading:
            return
        self._loading     = True
        self._all_systems = {}
        self._route_ids   = []
        self._gate_edges  = []
        self._origin_ids  = set()
        self._dest_ids    = set()
        self._norm_pos    = {}
        self._delivery_plan  = None
        self._delivery_mode  = False
        self._spin.start(12)
        self._status.config(
            text=f"Carregando {os.path.basename(path)}...", fg=COLORS["yellow"])
        threading.Thread(target=self._worker_load, args=(path,), daemon=True).start()

    def _worker_load(self, path):
        try:
            try:
                segments = load_route_file(path)
            except Exception as exc:
                self.after(0, lambda e=str(exc): (
                    self._status.config(text=f"Erro: {e[:60]}", fg=COLORS["red"]),
                    messagebox.showerror("Erro ao ler arquivo", str(e))))
                return

            segments = [s for s in segments
                        if s.get("Origem") and s.get("Destino")
                        and s.get("Item") not in (None, "", "\u2500\u2500\u2500 TOTAIS \u2500\u2500\u2500")]
            if not segments:
                self.after(0, lambda: messagebox.showwarning(
                    "Arquivo vazio",
                    "Nenhuma linha valida com Origem e Destino.\n\n"
                    "Use MAPA DA ROTA no eve_trade ou abra um CSV/XLSX do Otimizador."))
                return

            self._route_data = segments

            # Resolve sistemas
            unique_stations = set()
            for s in segments:
                unique_stations.add(s["Origem"])
                unique_stations.add(s["Destino"])

            self.after(0, lambda n=len(unique_stations): self._status.config(
                text=f"Resolvendo {n} estacoes via ESI...", fg=COLORS["yellow"]))

            sys_name_map = {}
            failed = []
            for st in unique_stations:
                sn = station_to_system_name(st)
                if not sn or sn in sys_name_map:
                    continue
                sid = resolve_system_name(st)
                if sid:
                    sys_name_map[sn] = sid
                else:
                    failed.append(sn)

            if not sys_name_map:
                self.after(0, lambda f=failed: messagebox.showerror(
                    "Erro ESI",
                    "Nao foi possivel resolver nenhum sistema via ESI.\n\n"
                    "Verifique a conexao com a internet.\n"
                    f"Sistemas tentados: {', '.join(f[:5])}"))
                return

            # Origens e destinos
            origin_ids, dest_ids = set(), set()
            for s in segments:
                on = station_to_system_name(s["Origem"])
                dn = station_to_system_name(s["Destino"])
                if on in sys_name_map: origin_ids.add(sys_name_map[on])
                if dn in sys_name_map: dest_ids.add(sys_name_map[dn])

            # Rota ESI (origem -> primeiro destino para visualizacao inicial)
            all_route_ids = []
            if origin_ids and dest_ids:
                oid = next(iter(origin_ids))
                did = next(iter(dest_ids))
                self.after(0, lambda: self._status.config(
                    text="Buscando rota ESI...", fg=COLORS["yellow"]))
                all_route_ids = fetch_route_ids(oid, did)
            if not all_route_ids:
                all_route_ids = list(origin_ids | dest_ids)

            # Carrega info de todos os sistemas
            all_ids = list(set(all_route_ids) | origin_ids | dest_ids)
            total   = len(all_ids)
            for i, sid in enumerate(all_ids):
                if i % 3 == 0:
                    self.after(0, lambda i=i, t=total: self._status.config(
                        text=f"Carregando sistemas {i+1}/{t}...", fg=COLORS["yellow"]))
                fetch_system(sid)

            systems = {s: _sys_cache[s] for s in all_ids if s in _sys_cache}

            # Stargates
            self.after(0, lambda: self._status.config(
                text="Carregando stargates...", fg=COLORS["yellow"]))
            edges = []
            for sid, info in systems.items():
                for gid in info.get("stargates", []):
                    dst = fetch_stargate_dest(gid)
                    if dst and dst in systems:
                        e = tuple(sorted((sid, dst)))
                        if e not in edges:
                            edges.append(e)

            self._route_ids   = all_route_ids
            self._gate_edges  = edges
            self._all_systems = systems
            self._origin_ids  = origin_ids
            self._dest_ids    = dest_ids
            self._sys_name_map = sys_name_map
            self.after(0, self._on_load_done)

        except Exception as exc:
            err = str(exc)
            self.after(0, lambda e=err: (
                self._status.config(text=f"Erro: {e[:60]}", fg=COLORS["red"]),
                messagebox.showerror("Erro", e)))
        finally:
            self._loading = False
            self.after(0, self._spin.stop)

    def _on_load_done(self):
        n_j = max(0, len(self._route_ids) - 1)
        n_d = len({station_to_system_name(s["Destino"]) for s in self._route_data})
        self._status.config(
            text=(f"OK  {len(self._all_systems)} sistemas | {n_j} pulos ate 1o destino | "
                  f"{n_d} sistemas de destino | {len(self._route_data)} itens"),
            fg=COLORS["green"])
        self._update_info_panel()
        self._populate_item_list()
        self._compute_layout()
        self._center_view()

    # ── Otimizacao de entregas ────────────────────────────────────────────────

    def _start_optimize(self):
        if not self._route_data:
            messagebox.showinfo("Sem dados",
                                "Carregue um arquivo de rota primeiro.")
            return
        if self._loading:
            return
        self._loading = True
        self._spin.start(12)
        self._opt_btn.config(state="disabled")
        self._status.config(
            text="Calculando rota otimizada de entregas...", fg=COLORS["yellow"])
        self._nb.select(1)   # muda para aba de entregas
        threading.Thread(target=self._worker_optimize, daemon=True).start()

    def _worker_optimize(self):
        try:
            def prog(msg):
                self.after(0, lambda m=msg: self._status.config(
                    text=m, fg=COLORS["yellow"]))

            sys_name_map = getattr(self, "_sys_name_map", {})
            if not sys_name_map:
                self.after(0, lambda: messagebox.showerror(
                    "Erro", "Carregue um arquivo primeiro."))
                return

            plan = build_delivery_plan(self._route_data, sys_name_map, prog)

            # Carrega sistemas das paradas que ainda nao foram carregados
            stop_ids = [s["system_id"] for s in plan["stops"]]
            for sid in stop_ids:
                if sid not in self._all_systems:
                    info = fetch_system(sid)
                    if info:
                        self._all_systems[sid] = info

            # Carrega stargates adicionais
            prog("Atualizando stargates...")
            for sid, info in list(self._all_systems.items()):
                for gid in info.get("stargates", []):
                    dst = fetch_stargate_dest(gid)
                    if dst and dst in self._all_systems:
                        e = tuple(sorted((sid, dst)))
                        if e not in self._gate_edges:
                            self._gate_edges.append(e)

            self._delivery_plan = plan
            self._delivery_mode = True
            self._dest_ids = set(stop_ids)
            self._compute_layout()

            self.after(0, self._on_optimize_done)

        except Exception as exc:
            err = str(exc)
            self.after(0, lambda e=err: (
                self._status.config(text=f"Erro: {e[:60]}", fg=COLORS["red"]),
                messagebox.showerror("Erro na otimizacao", e)))
        finally:
            self._loading = False
            self.after(0, lambda: (
                self._spin.stop(),
                self._opt_btn.config(state="normal")))

    def _on_optimize_done(self):
        plan  = self._delivery_plan
        stats = plan.get("opt_stats", {})
        n     = len(plan["stops"])
        tj    = plan["total_jumps"]
        tp    = plan["total_profit"]
        tc    = plan["total_capital"]
        roi   = (tp / tc * 100) if tc else 0

        nn_cost    = stats.get("nn_cost",    tj)
        saved      = stats.get("saved",      0)
        cache_hits = stats.get("cache_hits", 0)
        esi_calls  = stats.get("esi_calls",  0)
        mat_size   = stats.get("matrix_size",0)

        self._status.config(
            text=(f"Rota otimizada: {n} paradas | {tj} pulos "
                  f"(NN={nn_cost}, 2-opt -{saved}) | "
                  f"Cache: {cache_hits}/{mat_size} hits | "
                  f"ESI: {esi_calls} chamadas | "
                  f"Lucro: {self._isk(tp)} | ROI: {roi:.1f}%"),
            fg=COLORS["green"])

        # Resumo detalhado no painel
        opt_line = (f"NN: {nn_cost} pulos  →  2-opt: {tj} pulos  "
                    f"(-{saved} pulos economizados)")
        cache_line = (f"Cache SQLite: {cache_hits}/{mat_size} pares ({int(cache_hits/max(mat_size,1)*100)}% hit rate)"
                      if mat_size else "")

        self._plan_summary.config(
            text=(f"Origem: {plan['origin_name']}\n"
                  f"Paradas: {n} sistemas\n"
                  f"Pulos totais: {tj}\n"
                  f"Itens: {plan['total_items']}\n"
                  f"{'─'*30}\n"
                  f"Lucro est.:  {self._isk(tp)}\n"
                  f"Capital:     {self._isk(tc)}\n"
                  f"ROI:         {roi:.1f}%\n"
                  f"{'─'*30}\n"
                  f"Otimizacao:\n"
                  f"  {opt_line}\n"
                  f"  {cache_line}"),
            fg=COLORS["text"])

        # Preenche lista de paradas
        self._stop_list.delete(0, "end")
        for s in plan["stops"]:
            cj    = s["cumulative_jumps"]
            jprev = s["jumps_from_prev"]
            line  = (f"  {s['seq']:>2}. {s['system_name']:<16}"
                     f"  +{jprev:>2}j ({cj:>3}j total)"
                     f"  {s['n_items']} it")
            self._stop_list.insert("end", line)

        self._view_btn.config(state="normal",
                               text=" VER MAPA SIMPLES ",
                               fg=COLORS["subtext"])
        self._center_view()

    def _on_stop_select(self, event):
        sel = self._stop_list.curselection()
        if not sel or not self._delivery_plan:
            return
        idx = sel[0]
        self._selected_stop = idx
        stop = self._delivery_plan["stops"][idx]

        # Highlight no mapa
        self._selected_sys = stop["system_id"]
        self._redraw()

        # Detalhe
        stations_str = "\n".join(f"  - {s[:38]}" for s in stop["stations"])
        items_lines  = ""
        for it in stop["items"]:
            item = (it.get("Item") or "?")[:28]
            vol  = self._num(it.get("Vol Total (m3)"))
            prof = self._isk(self._num(it.get("Lucro Total")))
            qtd  = it.get("Qtd", 1)
            items_lines += f"  {item}\n    Qtd:{qtd}  Vol:{vol:.1f}m3  L:{prof}\n"

        self._stop_detail.config(
            text=(f"Parada #{stop['seq']}: {stop['system_name']}\n"
                  f"Pulos da origem:    {stop['jumps_from_origin']}\n"
                  f"Pulos do anterior:  {stop['jumps_from_prev']}\n"
                  f"Acumulado:          {stop['cumulative_jumps']}\n"
                  f"Volume total:       {stop['total_vol']:.1f} m3\n"
                  f"Lucro est.:         {self._isk(stop['total_profit'])}\n"
                  f"\nEstacoes:\n{stations_str}\n"
                  f"\nItens a entregar:\n{items_lines}"),
            fg=COLORS["text"])

    def _toggle_view(self):
        self._delivery_mode = not self._delivery_mode
        if self._delivery_mode:
            self._view_btn.config(text=" VER MAPA SIMPLES ", fg=COLORS["subtext"])
        else:
            self._view_btn.config(text=" VER ROTA ENTREGAS ", fg=COLORS["sys_stop"])
        self._redraw()

    # ── Layout e renderizacao ─────────────────────────────────────────────────

    def _compute_layout(self):
        if not self._all_systems:
            return
        pts = {sid: (info["x"], info["z"]) for sid, info in self._all_systems.items()}
        xs  = [p[0] for p in pts.values()]
        zs  = [p[1] for p in pts.values()]
        sx  = (max(xs) - min(xs)) or 1
        sz  = (max(zs) - min(zs)) or 1
        self._norm_pos = {
            sid: ((x - min(xs)) / sx * 2 - 1,
                  (z - min(zs)) / sz * 2 - 1)
            for sid, (x, z) in pts.items()}

    def _w2c(self, nx, nz):
        w  = self._canvas.winfo_width()  or 1000
        h  = self._canvas.winfo_height() or 700
        mg = 70
        uw = (w - mg * 2) * self._scale
        uh = (h - mg * 2) * self._scale
        return (mg + (nx + 1) / 2 * uw + self._offset_x,
                mg + (nz + 1) / 2 * uh + self._offset_y)

    def _redraw(self):
        c = self._canvas
        c.delete("all")
        if not self._all_systems or not self._norm_pos:
            c.create_text(
                (self._canvas.winfo_width() or 900) // 2,
                (self._canvas.winfo_height() or 600) // 2,
                text=("Clique em  MAPA DA ROTA  no eve_trade\n"
                      "ou use  ABRIR ARQUIVO  para abrir JSON, CSV ou XLSX.\n\n"
                      "Depois clique  OTIMIZAR ENTREGAS."),
                fill=COLORS["subtext"], font=("Courier", 12), justify="center")
            return

        if self._delivery_mode and self._delivery_plan:
            self._redraw_delivery()
        else:
            self._redraw_simple()

    def _redraw_simple(self):
        c = self._canvas
        route_set   = set(self._route_ids)
        route_edges = {tuple(sorted((self._route_ids[i], self._route_ids[i+1])))
                       for i in range(len(self._route_ids) - 1)}

        for a, b in self._gate_edges:
            if a not in self._norm_pos or b not in self._norm_pos: continue
            ax, ay = self._w2c(*self._norm_pos[a])
            bx, by = self._w2c(*self._norm_pos[b])
            is_re  = tuple(sorted((a, b))) in route_edges
            c.create_line(ax, ay, bx, by,
                          fill=COLORS["route"] if is_re else COLORS["gate"],
                          width=2.5 if is_re else 0.8,
                          dash=() if is_re else (3, 6))

        self._draw_arrows(self._route_ids)
        self._draw_nodes_simple()

    def _redraw_delivery(self):
        """Renderiza o modo de entregas com numeracao de paradas."""
        c    = self._canvas
        plan = self._delivery_plan
        stops = plan["stops"]

        # Constroi sequencia de sistemas: origem -> parada1 -> parada2 -> ...
        full_seq = [plan["origin_id"]] + [s["system_id"] for s in stops]

        # Arestas de stargates (fundo)
        stop_set = set(full_seq)
        for a, b in self._gate_edges:
            if a not in self._norm_pos or b not in self._norm_pos: continue
            ax, ay = self._w2c(*self._norm_pos[a])
            bx, by = self._w2c(*self._norm_pos[b])
            in_seq = a in stop_set and b in stop_set
            c.create_line(ax, ay, bx, by,
                          fill=COLORS["route"] if in_seq else COLORS["gate"],
                          width=2.0 if in_seq else 0.6,
                          dash=() if in_seq else (3, 8))

        # Linhas da rota de entrega (liga paradas em ordem)
        for i in range(len(full_seq) - 1):
            a = full_seq[i];  b = full_seq[i+1]
            if a not in self._norm_pos or b not in self._norm_pos: continue
            ax, ay = self._w2c(*self._norm_pos[a])
            bx, by = self._w2c(*self._norm_pos[b])
            c.create_line(ax, ay, bx, by,
                          fill=COLORS["sys_stop"], width=2.0,
                          dash=(6, 4), arrow="last",
                          arrowshape=(10, 12, 4))

        # Nos
        nr = max(5, int(8 * self._scale))
        self._canvas_nodes = {}

        for sid, info in self._all_systems.items():
            if sid not in self._norm_pos: continue
            cx, cy = self._w2c(*self._norm_pos[sid])
            self._canvas_nodes[sid] = (cx, cy)

            is_origin  = sid == plan["origin_id"]
            stop_idx   = next((i for i, s in enumerate(stops) if s["system_id"] == sid), -1)
            is_sel     = sid == self._selected_sys

            if is_sel:
                fill, r, outl = COLORS["sys_hi"],   nr+4, "#ffffff"
            elif is_origin:
                fill, r, outl = COLORS["sys_orig"], nr+3, COLORS["yellow"]
            elif stop_idx >= 0:
                fill, r, outl = COLORS["sys_stop"], nr+2, COLORS["border"]
            else:
                fill, r, outl = sec_color(info["sec"]), nr-2, COLORS["gate"]

            c.create_oval(cx-r, cy-r, cx+r, cy+r, fill=fill, outline=outl,
                          width=1.5, tags=("node", f"sys_{sid}"))

            # Numero da parada dentro do no
            if stop_idx >= 0:
                num_lbl = str(stop_idx + 1)
                c.create_text(cx, cy, text=num_lbl,
                              fill="#fff" if not is_sel else "#000",
                              font=("Courier", max(6, int(7 * self._scale)), "bold"))

            # Label do sistema
            if is_origin or stop_idx >= 0 or is_sel or self._scale > 1.5:
                lc = COLORS["yellow"] if is_sel else \
                     COLORS["green"] if is_origin else \
                     COLORS["text"] if stop_idx >= 0 else COLORS["subtext"]
                n_items = stops[stop_idx]["n_items"] if stop_idx >= 0 else ""
                extra   = f" ({n_items}it)" if n_items else ""
                c.create_text(cx, cy+r+9,
                              text=f"{info['name']}{extra}",
                              fill=lc,
                              font=("Courier", max(7, int(8*self._scale))))

        c.tag_bind("node", "<Button-1>", self._on_node_click)
        c.tag_bind("node", "<Enter>",    self._on_node_enter)
        c.tag_bind("node", "<Leave>",    self._on_node_leave)

    def _draw_arrows(self, seq):
        c = self._canvas
        for i in range(len(seq) - 1):
            a = seq[i]; b = seq[i+1]
            if a not in self._norm_pos or b not in self._norm_pos: continue
            ax, ay = self._w2c(*self._norm_pos[a])
            bx, by = self._w2c(*self._norm_pos[b])
            mx, my = (ax+bx)/2, (ay+by)/2
            ang = math.atan2(by-ay, bx-ax); sz = 8
            c.create_polygon(
                mx, my,
                mx - sz*math.cos(ang-0.4), my - sz*math.sin(ang-0.4),
                mx - sz*math.cos(ang+0.4), my - sz*math.sin(ang+0.4),
                fill=COLORS["route"], outline="")

    def _draw_nodes_simple(self):
        c  = self._canvas
        nr = max(4, int(7 * self._scale))
        self._canvas_nodes = {}
        for sid, info in self._all_systems.items():
            if sid not in self._norm_pos: continue
            cx, cy = self._w2c(*self._norm_pos[sid])
            self._canvas_nodes[sid] = (cx, cy)
            is_orig = sid in self._origin_ids
            is_dest = sid in self._dest_ids
            is_rt   = sid in set(self._route_ids)
            is_sel  = sid == self._selected_sys
            if is_sel:
                fill, r, outl = COLORS["sys_hi"],  nr+4, "#ffffff"
            elif is_orig:
                fill, r, outl = COLORS["sys_orig"], nr+3, COLORS["yellow"]
            elif is_dest:
                fill, r, outl = COLORS["sys_dest"], nr+2, COLORS["border"]
            elif is_rt:
                fill, r, outl = COLORS["route"],    nr+1, COLORS["border"]
            else:
                fill, r, outl = sec_color(info["sec"]), nr, COLORS["gate"]
            c.create_oval(cx-r, cy-r, cx+r, cy+r, fill=fill, outline=outl,
                          width=1.5, tags=("node", f"sys_{sid}"))
            if is_rt or is_sel or self._scale > 1.5:
                lc = (COLORS["yellow"] if is_sel else
                      COLORS["text"]   if is_rt  else COLORS["subtext"])
                c.create_text(cx, cy+r+8,
                              text=f"{info['name']}  {info['sec']:.1f}",
                              fill=lc, font=("Courier", max(7, int(8*self._scale))))
        c.tag_bind("node", "<Button-1>", self._on_node_click)
        c.tag_bind("node", "<Enter>",    self._on_node_enter)
        c.tag_bind("node", "<Leave>",    self._on_node_leave)

    # ── Interacao ─────────────────────────────────────────────────────────────

    def _on_click(self, e):   self._drag_start = (e.x, e.y)
    def _on_release(self, e): self._drag_start = None

    def _on_drag(self, e):
        if self._drag_start:
            self._offset_x += e.x - self._drag_start[0]
            self._offset_y += e.y - self._drag_start[1]
            self._drag_start = (e.x, e.y)
            self._redraw()

    def _on_scroll(self, e):
        up = e.num == 4 or (hasattr(e, "delta") and e.delta > 0)
        self._scale = max(0.1, min(12.0, self._scale * (1.15 if up else 1/1.15)))
        self._redraw()

    def _on_node_click(self, e):
        for t in self._canvas.gettags("current"):
            if t.startswith("sys_"):
                sid = int(t[4:])
                self._selected_sys = sid
                self._show_system_info(sid)
                # Se modo entrega, seleciona a parada correspondente
                if self._delivery_plan:
                    for i, s in enumerate(self._delivery_plan["stops"]):
                        if s["system_id"] == sid:
                            self._stop_list.selection_clear(0, "end")
                            self._stop_list.selection_set(i)
                            self._stop_list.see(i)
                            self._on_stop_select(None)
                            break
                self._redraw()
                break

    def _on_node_enter(self, e):
        for t in self._canvas.gettags("current"):
            if t.startswith("sys_"):
                info = self._all_systems.get(int(t[4:]), {})
                self._statusbar.config(
                    text=f"  {info.get('name','?')}  |  Seg: {info.get('sec',0):.1f}"
                         f"  |  Planetas: {info.get('planets',0)}  |  ID: {t[4:]}")
                break

    def _on_node_leave(self, e):
        self._statusbar.config(
            text="Scroll=zoom  |  Arraste=mover  |  Clique no sistema=info  "
                 "|  OTIMIZAR ENTREGAS=calcula rota multi-destino")

    def _center_view(self):
        self._scale = 1.0; self._offset_x = 0.0; self._offset_y = 0.0
        self._redraw()

    # ── Paineis de texto ──────────────────────────────────────────────────────

    def _num(self, v):
        try: return float(str(v or 0).replace(",","").replace(" ",""))
        except: return 0.0

    def _isk(self, v):
        v = self._num(v)
        if v >= 1e9: return f"{v/1e9:.2f} B ISK"
        if v >= 1e6: return f"{v/1e6:.2f} M ISK"
        return f"{v:,.0f} ISK"

    def _update_info_panel(self):
        if not self._route_data: return
        origs = list({station_to_system_name(s["Origem"]) for s in self._route_data if s.get("Origem")})
        dests = {station_to_system_name(s["Destino"]) for s in self._route_data if s.get("Destino")}
        nj    = max(0, len(self._route_ids) - 1)
        tp    = sum(self._num(s.get("Lucro Total")) for s in self._route_data)
        tc    = sum(self._num(s.get("Capital"))     for s in self._route_data)
        tv    = sum(self._num(s.get("Vol Total (m3)")) for s in self._route_data)
        roi   = (tp / tc * 100) if tc else 0
        self._info_route.config(
            text=(f"Origem(ns): {', '.join(origs)}\n"
                  f"Destinos:  {len(dests)} sistemas\n"
                  f"Sistemas:  {len(self._all_systems)}\n"
                  f"Itens:     {len(self._route_data)}\n"
                  f"{'─'*26}\n"
                  f"Lucro est.: {self._isk(tp)}\n"
                  f"Capital:    {self._isk(tc)}\n"
                  f"Volume:     {tv:,.1f} m3\n"
                  f"ROI:        {roi:.1f}%\n"
                  f"\nClique  OTIMIZAR ENTREGAS\npara calcular a rota otima."),
            fg=COLORS["text"])

    def _show_system_info(self, sid):
        info = self._all_systems.get(sid, {})
        if not info: return
        sec  = info.get("sec", 0)
        sl   = "HiSec" if sec >= 0.5 else "LowSec" if sec >= 0.1 else "Null"
        here = [s for s in self._route_data
                if station_to_system_name(s.get("Origem","")) == info["name"]
                or station_to_system_name(s.get("Destino","")) == info["name"]]
        txt  = (f"Sistema: {info['name']}\n"
                f"Seg.:    {sec:.2f} ({sl})\n"
                f"Planetas:{info.get('planets',0)}\n"
                f"ID ESI:  {sid}\n")
        if here:
            tp = sum(self._num(s.get("Lucro Total")) for s in here)
            tv = sum(self._num(s.get("Vol Total (m3)")) for s in here)
            txt += (f"\nItens ({len(here)}):\n"
                    f"  Vol: {tv:.1f} m3\n"
                    f"  Lucro: {self._isk(tp)}\n")
            for s in here[:5]:
                txt += f"  - {(s.get('Item') or '?')[:28]}\n"
            if len(here) > 5:
                txt += f"  ... +{len(here)-5} mais\n"
        self._info_sys.config(text=txt, fg=COLORS["text"])

    def _populate_item_list(self):
        self._item_list.delete(0, "end")
        for s in self._route_data:
            item = (s.get("Item") or "?")[:22]
            try: p = float(str(s.get("Lucro Total") or 0).replace(",",""))
            except: p = 0.0
            self._item_list.insert("end", f"  {item:<22}  {self._isk(p):>14}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if not HAS_REQUESTS:
        messagebox.showerror("Dependencia ausente", "pip install requests")
        return
    rf = sys.argv[1] if len(sys.argv) > 1 and os.path.exists(sys.argv[1]) else None
    EVERouteMap(rf).mainloop()

if __name__ == "__main__":
    main()
