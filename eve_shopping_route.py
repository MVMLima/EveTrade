"""
eve_shopping_route.py — Rota de compras otimizada para EVE Online
=================================================================

Lê uma lista de compras (XLSX/CSV/JSON exportada pelo eve_trade.py),
pede ao usuário a base de partida e de chegada (normalmente ambas Jita),
e calcula a rota mais curta para:

    BASE  →  sistema_compra_1  →  ...  →  sistema_compra_N  →  BASE

Usando Nearest-Neighbor + 2-opt com cache SQLite de pulos (TTL 30 dias).

Formato esperado do arquivo (coluna Origem = onde COMPRAR):
  Item | Origem (compra) | Destino (venda) | Pulos | Qtd | Preco | Capital | Lucro | Margem | Vol

Dependencias:
    pip install requests openpyxl
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
JUMP_CACHE_DB  = os.path.join(DATA_DIR, "jump_cache.db")
JUMP_CACHE_TTL = 30 * 86_400

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
    "route_out": "#ff8c00",   # ida (base → compras)
    "route_ret": "#00aaff",   # volta (compras → base)
    "gate":      "#1a2d40",
    "sys_base":  "#00ff88",   # base de partida/chegada
    "sys_stop":  "#aa44ff",   # parada de compra
    "sys_hi":    "#ffd700",   # selecionado
    "canvas_bg": "#060c16",
}

_SEC = [(0.9,"#2255ff"),(0.7,"#00aaff"),(0.5,"#00ccaa"),
        (0.4,"#ffdd00"),(0.3,"#ff8800"),(0.2,"#ff4400"),(0.0,"#ff2222")]
def sec_color(s):
    for t, c in _SEC:
        if s >= t: return c
    return "#ff2222"

# ── SQLite: cache de pulos ────────────────────────────────────────────────────
_jdb_lock = threading.Lock()
_jdb_conn = None
_jump_mem: dict = {}

def _get_jump_db():
    global _jdb_conn
    if _jdb_conn: return _jdb_conn
    os.makedirs(DATA_DIR, exist_ok=True)
    c = sqlite3.connect(JUMP_CACHE_DB, check_same_thread=False)
    c.execute("""CREATE TABLE IF NOT EXISTS jump_cache(
        origin_id INTEGER, dest_id INTEGER, jumps INTEGER,
        cached_at REAL, PRIMARY KEY(origin_id, dest_id))""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_jc ON jump_cache(origin_id,dest_id)")
    c.commit()
    _jdb_conn = c
    return c

def _jdb_get(a, b):
    lo, hi = (a,b) if a<b else (b,a)
    if (lo,hi) in _jump_mem: return _jump_mem[(lo,hi)]
    try:
        with _jdb_lock:
            row = _get_jump_db().execute(
                "SELECT jumps,cached_at FROM jump_cache WHERE origin_id=? AND dest_id=?",
                (lo,hi)).fetchone()
        if row and time.time()-row[1] < JUMP_CACHE_TTL:
            _jump_mem[(lo,hi)] = row[0]; return row[0]
        if row:
            with _jdb_lock:
                _get_jump_db().execute("DELETE FROM jump_cache WHERE origin_id=? AND dest_id=?",(lo,hi))
                _get_jump_db().commit()
    except Exception: pass
    return None

def _jdb_set(a, b, jumps):
    lo, hi = (a,b) if a<b else (b,a)
    _jump_mem[(lo,hi)] = jumps
    try:
        with _jdb_lock:
            _get_jump_db().execute(
                "INSERT OR REPLACE INTO jump_cache VALUES(?,?,?,?)",
                (lo,hi,jumps,time.time()))
            _get_jump_db().commit()
    except Exception: pass

# ── ESI ───────────────────────────────────────────────────────────────────────
_session = None
_sys_cache: dict  = {}
_name_id:   dict  = {}
_route_cache: dict = {}

def _sess():
    global _session
    if not _session:
        _session = requests.Session()
        _session.headers["User-Agent"] = "EVE Trade Suite / shopping_route"
    return _session

def esi_get(path, params=None):
    try:
        r = _sess().get(ESI_BASE+path, params=params, timeout=20)
        if r.status_code == 200: return r.json()
    except Exception: pass
    return None

def esi_post(path, body):
    try:
        r = _sess().post(ESI_BASE+path, json=body, timeout=20)
        if r.status_code == 200: return r.json()
    except Exception: pass
    return None

def sys_name_from_station(station: str) -> str:
    """'Jita IV - Moon 4 - ...' → 'Jita'"""
    if not station: return ""
    return station.split(" - ")[0].split(" ")[0].strip()

def resolve_name(station: str) -> int:
    n = sys_name_from_station(station)
    if not n: return 0
    if n in _name_id: return _name_id[n]
    d = esi_post("/universe/ids/", [n])
    if d and "systems" in d and d["systems"]:
        sid = d["systems"][0]["id"]
        _name_id[n] = sid
        return sid
    return 0

def fetch_system(sid: int) -> dict:
    if sid in _sys_cache: return _sys_cache[sid]
    d = esi_get(f"/universe/systems/{sid}/")
    if not d: return {}
    pos = d.get("position", {})
    info = {"id":sid,"name":d.get("name",str(sid)),"sec":round(d.get("security_status",0),1),
            "x":pos.get("x",0),"y":pos.get("y",0),"z":pos.get("z",0),
            "stargates":d.get("stargates",[]),"planets":len(d.get("planets",[]))}
    _sys_cache[sid] = info
    _name_id[info["name"]] = sid
    return info

def fetch_gate_dest(gid: int) -> int:
    d = esi_get(f"/universe/stargates/{gid}/")
    return d.get("destination",{}).get("system_id",0) if d else 0

def fetch_route_ids(a: int, b: int) -> list:
    k = (a,b)
    if k in _route_cache: return _route_cache[k]
    d = esi_get(f"/route/{a}/{b}/", {"flag":"shortest"})
    r = d if isinstance(d,list) else []
    _route_cache[k] = r
    return r

def fetch_jumps(a: int, b: int) -> int:
    if a == b: return 0
    c = _jdb_get(a, b)
    if c is not None: return c
    r = fetch_route_ids(a, b)
    j = max(0, len(r)-1) if r else 999
    _jdb_set(a, b, j)
    return j

# ── Leitura do arquivo ────────────────────────────────────────────────────────

def _norm(row: dict) -> dict:
    def g(*ks):
        for k in ks:
            v = row.get(k)
            if v not in (None,"","─── TOTAIS ───"): return v
        return ""
    return {
        "Item":    g("Item","item","name"),
        "Origem":  g("Origem","loc_origem","Estacao Origem"),  # local de COMPRA
        "Destino": g("Destino","loc_destino","Estacao Destino"), # local de VENDA
        "Qtd":     g("Qtd","qtd_origem","quantidade"),
        "Preco":   g("Preco Compra","Preco Compra (ISK)","Preço Compra","preco_origem"),
        "Capital": g("Capital (ISK)","Capital","capital"),
        "Lucro":   g("Lucro Total","Lucro Estimado","roi_total","lucro"),
        "Margem":  g("Margem %","Margem (%)","margem_pct"),
        "Vol":     g("Vol (m³)","Vol Total (m3)","vol_total_m3","volume","Vol"),
    }

def load_file(path: str) -> list:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".json":
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        rows = data if isinstance(data,list) else data.get("segments",[])
        return [_norm(r) for r in rows]
    if ext == ".csv":
        with open(path, newline="", encoding="utf-8-sig") as f:
            return [_norm(dict(r)) for r in csv.DictReader(f)]
    if ext in (".xlsx",".xlsm"):
        if not HAS_OPENPYXL: raise ImportError("pip install openpyxl")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        it = ws.iter_rows(values_only=True)
        hdrs = [str(h or "").strip() for h in next(it)]
        return [_norm({hdrs[i]:(r[i] if i<len(r) else "") for i in range(len(hdrs))})
                for r in it]
    raise ValueError(f"Formato nao suportado: {ext}")

# ── Otimizador round-trip TSP ─────────────────────────────────────────────────

def build_distance_matrix(all_ids: list, progress_cb=None) -> dict:
    """Monta matriz completa (cache SQLite + ESI paralelo)."""
    matrix = {}
    needed = []
    for i in range(len(all_ids)):
        for j in range(i+1, len(all_ids)):
            a,b = all_ids[i], all_ids[j]
            lo,hi = (a,b) if a<b else (b,a)
            c = _jdb_get(lo,hi)
            if c is not None: matrix[(lo,hi)] = c
            else: needed.append((lo,hi))

    if progress_cb:
        progress_cb(f"Cache: {len(matrix)} pares | ESI: {len(needed)} pares novos...")

    if needed:
        lock = threading.Lock()
        done = [0]
        def _fetch(pair):
            a,b = pair
            r = fetch_route_ids(a,b)
            j = max(0,len(r)-1) if r else 999
            _jdb_set(a,b,j)
            with lock:
                matrix[(a,b)] = j
                done[0] += 1
                if progress_cb and done[0] % 8 == 0:
                    pct = int(done[0]/len(needed)*100)
                    progress_cb(f"ESI: {done[0]}/{len(needed)} ({pct}%)...")
        with concurrent.futures.ThreadPoolExecutor(max_workers=8) as pool:
            list(pool.map(_fetch, needed))

    return matrix

def mg(matrix, a, b):
    if a==b: return 0
    lo,hi = (a,b) if a<b else (b,a)
    return matrix.get((lo,hi), 999)

def round_trip_cost(route: list, base_id: int, matrix: dict) -> int:
    """Custo total: base→r[0]→r[1]→...→r[-1]→base."""
    if not route: return 0
    full = [base_id] + route + [base_id]
    return sum(mg(matrix, full[i], full[i+1]) for i in range(len(full)-1))

def nearest_neighbor_rt(base_id: int, stops: list, matrix: dict,
                          progress_cb=None) -> list:
    """NN a partir da base, inclui retorno à base no critério de escolha."""
    remaining = list(stops)
    ordered   = []
    current   = base_id
    total     = len(stops)

    while remaining:
        if progress_cb:
            progress_cb(f"Nearest-Neighbor: {len(ordered)}/{total}...")
        nearest = min(remaining, key=lambda s: mg(matrix, current, s))
        ordered.append(nearest)
        remaining.remove(nearest)
        current = nearest

    return ordered

def two_opt_rt(route: list, base_id: int, matrix: dict,
                progress_cb=None) -> tuple:
    """
    2-opt para rota com retorno à base.
    A rota completa é: [base, r0, r1, ..., rN, base].
    Testa reversão de cada segmento [i..j] e aceita se reduz o custo total.
    """
    if len(route) < 3:
        return route, 0

    full  = [base_id] + list(route) + [base_id]
    n     = len(full)

    def cost(r):
        return sum(mg(matrix, r[i], r[i+1]) for i in range(len(r)-1))

    best      = list(full)
    best_cost = cost(best)
    improved  = True
    saved     = 0
    passes    = 0

    while improved:
        improved = False
        passes  += 1
        if progress_cb:
            progress_cb(f"2-opt passo {passes}: {best_cost} pulos...")
        for i in range(1, n-2):
            for j in range(i+1, n-1):
                cb = mg(matrix,best[i-1],best[i]) + mg(matrix,best[j],best[j+1])
                ca = mg(matrix,best[i-1],best[j]) + mg(matrix,best[i],best[j+1])
                if ca < cb:
                    best[i:j+1] = best[i:j+1][::-1]
                    s = cb-ca; best_cost -= s; saved += s; improved = True

    return best[1:-1], saved   # remove base do início e do fim

def optimise_shopping_route(base_id: int, stop_ids: list,
                             progress_cb=None) -> tuple:
    """
    Pipeline completo: matriz → NN → 2-opt.
    Retorna (rota_ordenada, stats_dict).
    """
    if not stop_ids: return [], {}

    all_ids = list(set([base_id] + stop_ids))

    before = len(_jump_mem)
    if progress_cb: progress_cb("Construindo matriz de distâncias...")
    matrix  = build_distance_matrix(all_ids, progress_cb)
    after   = len(_jump_mem)

    n_pairs  = len(all_ids)*(len(all_ids)-1)//2
    cache_h  = after - before
    esi_c    = max(0, n_pairs - cache_h)

    if progress_cb: progress_cb("Rodando Nearest-Neighbor...")
    nn_route = nearest_neighbor_rt(base_id, stop_ids, matrix, progress_cb)
    nn_cost  = round_trip_cost(nn_route, base_id, matrix)

    if progress_cb: progress_cb("Rodando melhoria 2-opt...")
    opt_route, saved = two_opt_rt(nn_route, base_id, matrix, progress_cb)
    final_cost = round_trip_cost(opt_route, base_id, matrix)

    return opt_route, {
        "nn_cost":     nn_cost,
        "final_cost":  final_cost,
        "saved":       saved,
        "cache_hits":  cache_h,
        "esi_calls":   esi_c,
        "n_pairs":     n_pairs,
    }

def build_shopping_plan(rows: list, base_station: str,
                         sell_station: str, sys_name_map: dict,
                         opt_route: list, matrix: dict) -> dict:
    """
    Agrupa os itens por sistema de compra na ordem da rota otimizada.
    """
    def num(v):
        try: return float(str(v or 0).replace(",","").replace(" ",""))
        except: return 0.0

    base_sys  = sys_name_from_station(base_station)
    base_id   = sys_name_map.get(base_sys, 0)
    sell_sys  = sys_name_from_station(sell_station)

    # Agrupa por sistema de compra
    stops_map: dict = {}
    for r in rows:
        if not r.get("Origem") or not r.get("Item"): continue
        sn = sys_name_from_station(r["Origem"])
        if not sn: continue
        if sn not in stops_map:
            stops_map[sn] = {
                "system_name": sn,
                "system_id":   sys_name_map.get(sn, 0),
                "stations":    set(),
                "items":       [],
                "total_vol":   0.0,
                "total_capital": 0.0,
                "total_profit":  0.0,
            }
        stops_map[sn]["stations"].add(r["Origem"])
        stops_map[sn]["items"].append(r)
        stops_map[sn]["total_vol"]     += num(r.get("Vol"))
        stops_map[sn]["total_capital"] += num(r.get("Capital"))
        stops_map[sn]["total_profit"]  += num(r.get("Lucro"))

    for v in stops_map.values():
        v["stations"] = sorted(v["stations"])

    # Ordena paradas conforme rota otimizada
    id_to_sn = {v["system_id"]: k for k, v in stops_map.items()}
    ordered_stops = []
    prev_id = base_id
    cum_j   = 0

    for sid in opt_route:
        sn = id_to_sn.get(sid)
        if not sn: continue
        info  = stops_map[sn]
        j_prev = mg(matrix, prev_id, sid)
        j_base = mg(matrix, base_id, sid)
        cum_j += j_prev
        ordered_stops.append({
            **info,
            "seq":               len(ordered_stops)+1,
            "jumps_from_prev":   j_prev,
            "jumps_from_base":   j_base,
            "cumulative_jumps":  cum_j,
            "n_items":           len(info["items"]),
        })
        prev_id = sid

    # Pulos de volta
    return_jumps = mg(matrix, prev_id, base_id) if opt_route else 0
    cum_j += return_jumps

    return {
        "base_station":  base_station,
        "base_sys":      base_sys,
        "base_id":       base_id,
        "sell_station":  sell_station,
        "sell_sys":      sell_sys,
        "stops":         ordered_stops,
        "return_jumps":  return_jumps,
        "total_jumps":   cum_j,
        "total_items":   sum(len(s["items"]) for s in ordered_stops),
        "total_vol":     sum(s["total_vol"] for s in ordered_stops),
        "total_capital": sum(s["total_capital"] for s in ordered_stops),
        "total_profit":  sum(s["total_profit"] for s in ordered_stops),
    }

# ── Diálogo de configuração ───────────────────────────────────────────────────

class SetupDialog(tk.Toplevel):
    """Janela que pede o arquivo, base de partida e ponto de venda."""

    def __init__(self, parent, default_file=""):
        super().__init__(parent)
        self.title("EVE Shopping Route — Configuração")
        self.configure(bg=COLORS["bg"])
        self.resizable(False, False)
        self.grab_set()
        self.result = None

        pad = {"padx":10,"pady":5}

        tk.Label(self, text="  EVE SHOPPING ROUTE  ",
                 bg=COLORS["bg"], fg=COLORS["header"],
                 font=("Courier",14,"bold")).pack(pady=(16,4))
        tk.Label(self, text="Planejador de rota de compras com retorno à base",
                 bg=COLORS["bg"], fg=COLORS["subtext"],
                 font=("Courier",9)).pack(pady=(0,10))

        # Arquivo
        frm = tk.Frame(self, bg=COLORS["panel"],
                       highlightbackground=COLORS["border"], highlightthickness=1)
        frm.pack(fill="x", padx=16, pady=6)
        tk.Label(frm, text=" Lista de compras (XLSX / CSV / JSON):",
                 bg=COLORS["panel"], fg=COLORS["header"],
                 font=("Courier",10,"bold")).pack(anchor="w", padx=8, pady=(8,2))
        fr2 = tk.Frame(frm, bg=COLORS["panel"]); fr2.pack(fill="x", padx=8, pady=(0,8))
        self._file_var = tk.StringVar(value=default_file)
        tk.Entry(fr2, textvariable=self._file_var, width=48,
                 bg=COLORS["border"], fg=COLORS["text"],
                 insertbackground=COLORS["text"],
                 font=("Courier",10), relief="flat", bd=2).pack(side="left")
        tk.Button(fr2, text=" ... ", command=self._browse,
                  bg=COLORS["accent"], fg="#000",
                  font=("Courier",10,"bold"), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)

        # Base de partida (onde VENDE / onde começa)
        frm2 = tk.Frame(self, bg=COLORS["panel"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        frm2.pack(fill="x", padx=16, pady=6)
        tk.Label(frm2,
                 text=" Base de operações (partida e chegada para venda):",
                 bg=COLORS["panel"], fg=COLORS["header"],
                 font=("Courier",10,"bold")).pack(anchor="w", padx=8, pady=(8,2))
        tk.Label(frm2,
                 text=" Normalmente Jita ou a estacao detectada no arquivo.",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier",9)).pack(anchor="w", padx=8)
        self._base_var = tk.StringVar(value="Jita IV - Moon 4 - Caldari Navy Assembly Plant")
        tk.Entry(frm2, textvariable=self._base_var, width=56,
                 bg=COLORS["border"], fg=COLORS["text"],
                 insertbackground=COLORS["text"],
                 font=("Courier",10), relief="flat", bd=2).pack(padx=8, pady=(4,8), fill="x")

        # Opcao: detectar base automaticamente do arquivo
        self._auto_base = tk.BooleanVar(value=True)
        tk.Checkbutton(frm2,
                       text=" Detectar automaticamente do arquivo (coluna Destino)",
                       variable=self._auto_base,
                       command=self._toggle_base,
                       bg=COLORS["panel"], fg=COLORS["text"],
                       selectcolor=COLORS["border"],
                       activebackground=COLORS["panel"],
                       font=("Courier",9)).pack(anchor="w", padx=8, pady=(0,8))

        # Capacidade
        frm3 = tk.Frame(self, bg=COLORS["panel"],
                        highlightbackground=COLORS["border"], highlightthickness=1)
        frm3.pack(fill="x", padx=16, pady=6)
        tk.Label(frm3, text=" Parâmetros da viagem:",
                 bg=COLORS["panel"], fg=COLORS["header"],
                 font=("Courier",10,"bold")).pack(anchor="w", padx=8, pady=(8,2))
        prow = tk.Frame(frm3, bg=COLORS["panel"]); prow.pack(fill="x", padx=8, pady=(0,8))

        tk.Label(prow, text="Capacidade da nave (m³):",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier",10)).pack(side="left")
        self._cap_var = tk.StringVar(value="500000")
        tk.Entry(prow, textvariable=self._cap_var, width=12,
                 bg=COLORS["border"], fg=COLORS["text"],
                 insertbackground=COLORS["text"],
                 font=("Courier",10), relief="flat", bd=2).pack(side="left", padx=8)

        tk.Label(prow, text="Seg/pulo (para ISK/hora):",
                 bg=COLORS["panel"], fg=COLORS["subtext"],
                 font=("Courier",10)).pack(side="left", padx=(16,0))
        self._spj_var = tk.StringVar(value="30")
        tk.Entry(prow, textvariable=self._spj_var, width=6,
                 bg=COLORS["border"], fg=COLORS["text"],
                 insertbackground=COLORS["text"],
                 font=("Courier",10), relief="flat", bd=2).pack(side="left", padx=4)

        # Botoes
        br = tk.Frame(self, bg=COLORS["bg"])
        br.pack(pady=14)
        tk.Button(br, text="  CALCULAR ROTA  ",
                  command=self._ok,
                  bg=COLORS["green"], fg="#000",
                  font=("Courier",12,"bold"), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=8)
        tk.Button(br, text=" Cancelar ",
                  command=self.destroy,
                  bg=COLORS["border"], fg=COLORS["text"],
                  font=("Courier",11), relief="flat", cursor="hand2"
                  ).pack(side="left", padx=4)

        self._toggle_base()
        self.update_idletasks()
        self.geometry(f"+{parent.winfo_rootx()+80}+{parent.winfo_rooty()+80}")
        self.wait_window()

    def _browse(self):
        p = filedialog.askopenfilename(
            title="Abrir lista de compras",
            filetypes=[("Todos","*.xlsx *.csv *.json"),
                       ("Excel","*.xlsx"),("CSV","*.csv"),("JSON","*.json")])
        if p:
            self._file_var.set(p)
            self._try_detect_base(p)

    def _try_detect_base(self, path):
        """Tenta detectar a base do arquivo (coluna Destino mais frequente)."""
        try:
            rows = load_file(path)
            counts = {}
            for r in rows:
                d = r.get("Destino","")
                if d: counts[d] = counts.get(d,0)+1
            if counts:
                best = max(counts, key=counts.get)
                self._base_var.set(best)
        except Exception:
            pass

    def _toggle_base(self):
        # Nao desabilita o campo, apenas oferece deteccao automatica
        if self._auto_base.get():
            f = self._file_var.get()
            if f and os.path.exists(f):
                self._try_detect_base(f)

    def _ok(self):
        f = self._file_var.get().strip()
        if not f or not os.path.exists(f):
            messagebox.showerror("Arquivo não encontrado",
                                 f"Arquivo não encontrado:\n{f}")
            return
        try:
            cap = float(self._cap_var.get().replace(",","") or "999999")
            spj = float(self._spj_var.get() or "30")
        except ValueError:
            messagebox.showerror("Parâmetro inválido",
                                 "Capacidade e seg/pulo devem ser números.")
            return
        self.result = {
            "file":        f,
            "base":        self._base_var.get().strip(),
            "auto_base":   self._auto_base.get(),
            "capacity_m3": cap,
            "sec_per_jump":spj,
        }
        self.destroy()

# ── Aplicação principal ───────────────────────────────────────────────────────

class EVEShoppingRoute(tk.Tk):
    def __init__(self, config: dict):
        super().__init__()
        self.title("EVE Shopping Route")
        self.configure(bg=COLORS["bg"])
        self.geometry("1560x920")
        self.minsize(1100, 650)

        self._config   = config
        self._rows     = []
        self._plan     = None
        self._all_sys  = {}
        self._norm_pos = {}
        self._gate_edges = []
        self._nodes    = {}
        self._sel_sys  = 0
        self._sel_stop = -1
        self._loading  = False

        self._scale    = 1.0
        self._off_x    = 0.0
        self._off_y    = 0.0
        self._drag     = None

        self._build_ui()
        self.after(200, self._start_load)

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        top = tk.Frame(self, bg=COLORS["panel"],
                       highlightbackground=COLORS["border"], highlightthickness=1)
        top.pack(fill="x")

        tk.Label(top, text="  EVE SHOPPING ROUTE  ",
                 bg=COLORS["panel"], fg=COLORS["header"],
                 font=("Courier",13,"bold")).pack(side="left", padx=6, pady=6)

        for lbl, cmd, fg in [
            (" NOVA ROTA ",    self._new_route,   COLORS["accent"]),
            (" CENTRALIZAR ",  self._center,       COLORS["text"]),
            (" EXPORTAR PLANO ", self._export_plan, COLORS["yellow"]),
        ]:
            tk.Button(top, text=lbl, command=cmd,
                      bg=COLORS["border"], fg=fg,
                      font=("Courier",10), relief="flat", cursor="hand2"
                      ).pack(side="left", padx=4, pady=6)

        # Legenda
        leg = tk.Frame(top, bg=COLORS["panel"]); leg.pack(side="left", padx=14)
        for col,lbl in [(COLORS["sys_base"],"Base"),(COLORS["sys_stop"],"Compra"),
                         (COLORS["route_out"],"Ida"),(COLORS["route_ret"],"Volta")]:
            tk.Label(leg, text="O", fg=col, bg=COLORS["panel"],
                     font=("Courier",10)).pack(side="left", padx=1)
            tk.Label(leg, text=lbl, fg=COLORS["subtext"],
                     bg=COLORS["panel"], font=("Courier",8)).pack(side="left", padx=(0,6))

        self._spin = ttk.Progressbar(top, mode="indeterminate", length=80)
        self._spin.pack(side="left", padx=8)
        self._status = tk.Label(top, text="Iniciando...",
                                bg=COLORS["panel"], fg=COLORS["subtext"],
                                font=("Courier",10))
        self._status.pack(side="left", padx=4)

        body = tk.Frame(self, bg=COLORS["bg"])
        body.pack(fill="both", expand=True)

        # Canvas
        self._canvas = tk.Canvas(body, bg=COLORS["canvas_bg"],
                                  highlightthickness=0, cursor="crosshair")
        self._canvas.pack(side="left", fill="both", expand=True)

        # Painel direito
        right = tk.Frame(body, bg=COLORS["panel"], width=400,
                         highlightbackground=COLORS["border"], highlightthickness=1)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)

        nb = ttk.Notebook(right)
        nb.pack(fill="both", expand=True)
        self._nb = nb

        t1 = tk.Frame(nb, bg=COLORS["panel"]); nb.add(t1, text="  Resumo  ")
        self._build_tab_resumo(t1)

        t2 = tk.Frame(nb, bg=COLORS["panel"]); nb.add(t2, text="  Rota de Compras  ")
        self._build_tab_rota(t2)

        t3 = tk.Frame(nb, bg=COLORS["panel"]); nb.add(t3, text="  Lista por Parada  ")
        self._build_tab_lista(t3)

        self._statusbar = tk.Label(
            self,
            text="Scroll=zoom  |  Arraste=mover  |  Clique no sistema=info  |  Dbl-clique=selecionar parada",
            bg=COLORS["border"], fg=COLORS["subtext"],
            font=("Courier",9), anchor="w")
        self._statusbar.pack(fill="x", side="bottom")

        self._canvas.bind("<ButtonPress-1>",  self._on_dn)
        self._canvas.bind("<B1-Motion>",       self._on_drag)
        self._canvas.bind("<ButtonRelease-1>", self._on_up)
        self._canvas.bind("<MouseWheel>",       self._on_scroll)
        self._canvas.bind("<Button-4>",         self._on_scroll)
        self._canvas.bind("<Button-5>",         self._on_scroll)
        self._canvas.bind("<Configure>",        lambda e: self._redraw())

    def _build_tab_resumo(self, p):
        def sep(): tk.Frame(p, bg=COLORS["border"], height=1).pack(fill="x")
        def hdr(t): tk.Label(p, text=f"  {t}  ", bg=COLORS["panel"],
                              fg=COLORS["header"], font=("Courier",10,"bold")).pack(fill="x",pady=(6,2))
        hdr("RESUMO DA VIAGEM")
        sep()
        self._lbl_resumo = tk.Label(p, text="Calculando...",
                                     bg=COLORS["panel"], fg=COLORS["subtext"],
                                     font=("Courier",10), justify="left",
                                     anchor="nw", wraplength=380)
        self._lbl_resumo.pack(fill="x", padx=8, pady=6, anchor="nw")
        sep()
        hdr("SISTEMA SELECIONADO")
        self._lbl_sys = tk.Label(p, text="Clique num sistema.",
                                  bg=COLORS["panel"], fg=COLORS["subtext"],
                                  font=("Courier",10), justify="left",
                                  anchor="nw", wraplength=380)
        self._lbl_sys.pack(fill="x", padx=8, pady=4, anchor="nw")
        sep()
        hdr("OTIMIZAÇÃO")
        self._lbl_opt = tk.Label(p, text="—",
                                  bg=COLORS["panel"], fg=COLORS["subtext"],
                                  font=("Courier",10), justify="left",
                                  anchor="nw", wraplength=380)
        self._lbl_opt.pack(fill="x", padx=8, pady=4, anchor="nw")

    def _build_tab_rota(self, p):
        tk.Label(p, text="  ORDEM DE VISITA  ", bg=COLORS["panel"],
                 fg=COLORS["header"], font=("Courier",10,"bold")).pack(fill="x",pady=(8,2))
        tk.Frame(p, bg=COLORS["border"], height=1).pack(fill="x")
        lf = tk.Frame(p, bg=COLORS["panel"]); lf.pack(fill="both", expand=True, padx=4, pady=4)
        self._stop_list = tk.Listbox(lf, bg=COLORS["border"], fg=COLORS["text"],
                                      font=("Courier",9), relief="flat", bd=0,
                                      selectbackground=COLORS["sys_stop"],
                                      selectforeground="#fff", activestyle="none")
        vsb = ttk.Scrollbar(lf, orient="vertical", command=self._stop_list.yview)
        self._stop_list.configure(yscrollcommand=vsb.set)
        self._stop_list.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._stop_list.bind("<<ListboxSelect>>", self._on_stop_sel)

    def _build_tab_lista(self, p):
        tk.Label(p, text="  ITENS A COMPRAR NA PARADA  ", bg=COLORS["panel"],
                 fg=COLORS["header"], font=("Courier",10,"bold")).pack(fill="x",pady=(8,2))
        tk.Frame(p, bg=COLORS["border"], height=1).pack(fill="x")
        self._lbl_stop_hdr = tk.Label(p, text="Selecione uma parada na aba Rota.",
                                       bg=COLORS["panel"], fg=COLORS["subtext"],
                                       font=("Courier",10), justify="left",
                                       anchor="nw", wraplength=380)
        self._lbl_stop_hdr.pack(fill="x", padx=8, pady=6, anchor="nw")
        tk.Frame(p, bg=COLORS["border"], height=1).pack(fill="x")
        lf2 = tk.Frame(p, bg=COLORS["panel"]); lf2.pack(fill="both", expand=True, padx=4, pady=4)
        self._item_list = tk.Listbox(lf2, bg=COLORS["border"], fg=COLORS["text"],
                                      font=("Courier",9), relief="flat", bd=0,
                                      selectbackground=COLORS["accent"],
                                      selectforeground="#000", activestyle="none")
        vsb2 = ttk.Scrollbar(lf2, orient="vertical", command=self._item_list.yview)
        self._item_list.configure(yscrollcommand=vsb2.set)
        self._item_list.pack(side="left", fill="both", expand=True)
        vsb2.pack(side="right", fill="y")

    # ── Carregamento e otimização ─────────────────────────────────────────────

    def _new_route(self):
        dlg = SetupDialog(self)
        if dlg.result:
            self._config  = dlg.result
            self._plan    = None
            self._all_sys = {}
            self._norm_pos= {}
            self._gate_edges = []
            self.after(100, self._start_load)

    def _start_load(self):
        if self._loading: return
        self._loading = True
        self._spin.start(12)
        self._status.config(text="Carregando...", fg=COLORS["yellow"])
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            cfg = self._config

            # 1. Lê arquivo
            self.after(0, lambda: self._status.config(
                text="Lendo arquivo...", fg=COLORS["yellow"]))
            rows = load_file(cfg["file"])
            rows = [r for r in rows if r.get("Origem") and r.get("Item")]
            if not rows:
                self.after(0, lambda: messagebox.showerror(
                    "Arquivo vazio", "Nenhuma linha com Origem válida encontrada."))
                return
            self._rows = rows

            # 2. Detecta base
            base_station = cfg["base"]
            if not base_station:
                counts = {}
                for r in rows:
                    d = r.get("Destino","")
                    if d: counts[d] = counts.get(d,0)+1
                base_station = max(counts, key=counts.get) if counts else "Jita"

            # 3. Resolve sistemas
            all_stations = set([base_station])
            for r in rows:
                all_stations.add(r.get("Origem",""))

            self.after(0, lambda n=len(all_stations): self._status.config(
                text=f"Resolvendo {n} estações via ESI...", fg=COLORS["yellow"]))

            sys_map: dict = {}
            failed = []
            for st in all_stations:
                sn  = sys_name_from_station(st)
                if not sn or sn in sys_map: continue
                sid = resolve_name(st)
                if sid: sys_map[sn] = sid
                else:   failed.append(sn)

            if not sys_map:
                self.after(0, lambda: messagebox.showerror(
                    "Erro ESI", "Não foi possível resolver nenhum sistema."))
                return

            base_sys = sys_name_from_station(base_station)
            base_id  = sys_map.get(base_sys, 0)
            if not base_id:
                self.after(0, lambda: messagebox.showerror(
                    "Base não resolvida",
                    f"Sistema '{base_sys}' não resolvido.\nVerifique o nome da base."))
                return

            # 4. Sistemas de compra únicos
            buy_sys_ids = []
            seen = set()
            for r in rows:
                sn  = sys_name_from_station(r.get("Origem",""))
                sid = sys_map.get(sn, 0)
                if sid and sid != base_id and sid not in seen:
                    buy_sys_ids.append(sid)
                    seen.add(sid)

            if not buy_sys_ids:
                self.after(0, lambda: messagebox.showerror(
                    "Sem paradas", "Nenhum sistema de compra encontrado."))
                return

            # 5. Otimização
            def prog(msg):
                self.after(0, lambda m=msg: self._status.config(
                    text=m, fg=COLORS["yellow"]))

            prog(f"Otimizando rota: {len(buy_sys_ids)} sistemas de compra...")
            opt_route, opt_stats = optimise_shopping_route(base_id, buy_sys_ids, prog)

            # 6. Carrega info dos sistemas
            all_ids = list(set([base_id] + buy_sys_ids))
            for i, sid in enumerate(all_ids):
                prog(f"Carregando sistemas {i+1}/{len(all_ids)}...")
                fetch_system(sid)

            all_sys = {s: _sys_cache[s] for s in all_ids if s in _sys_cache}

            # 7. Stargates
            prog("Carregando stargates...")
            edges = []
            for sid, info in all_sys.items():
                for gid in info.get("stargates",[]):
                    dst = fetch_gate_dest(gid)
                    if dst and dst in all_sys:
                        e = tuple(sorted((sid,dst)))
                        if e not in edges: edges.append(e)

            # 8. Constrói plano
            matrix = {}
            for sid in all_ids:
                for sid2 in all_ids:
                    if sid < sid2:
                        j = fetch_jumps(sid, sid2)
                        matrix[(sid,sid2)] = j

            plan = build_shopping_plan(rows, base_station,
                                       rows[0].get("Destino","Jita"),
                                       sys_map, opt_route, matrix)
            plan["opt_stats"]  = opt_stats
            plan["base_id"]    = base_id
            plan["sys_map"]    = sys_map
            plan["matrix"]     = matrix
            plan["capacity"]   = cfg["capacity_m3"]
            plan["sec_per_jump"] = cfg["sec_per_jump"]

            self._plan       = plan
            self._all_sys    = all_sys
            self._gate_edges = edges

            self.after(0, self._on_done)

        except Exception as exc:
            err = str(exc)
            self.after(0, lambda e=err: (
                self._status.config(text=f"Erro: {e[:60]}", fg=COLORS["red"]),
                messagebox.showerror("Erro", e)))
        finally:
            self._loading = False
            self.after(0, self._spin.stop)

    def _on_done(self):
        p   = self._plan
        tj  = p["total_jumps"]
        n   = len(p["stops"])
        tp  = p["total_profit"]
        tc  = p["total_capital"]
        tv  = p["total_vol"]
        roi = tp/tc*100 if tc else 0
        spj = p.get("sec_per_jump", 30)
        time_h = tj * spj / 3600
        isk_h  = tp / time_h if time_h > 0 else 0
        cap    = p.get("capacity", 999999)
        viagens= math.ceil(tv / cap) if cap > 0 else 1

        stats = p.get("opt_stats",{})
        nn    = stats.get("nn_cost", tj)
        saved = stats.get("saved", 0)
        hits  = stats.get("cache_hits",0)
        npairs= stats.get("n_pairs",0)

        self._status.config(
            text=(f"✓ {n} paradas | {tj} pulos (ida+volta) | "
                  f"Lucro: {self._isk(tp)} | ROI: {roi:.1f}%"),
            fg=COLORS["green"])

        self._lbl_resumo.config(fg=COLORS["text"], text=(
            f"Base:    {p['base_sys']} → {p['sell_sys']}\n"
            f"Paradas: {n} sistemas de compra\n"
            f"Pulos:   {tj} (ida + volta)\n"
            f"Itens:   {p['total_items']}\n"
            f"Volume:  {tv:,.1f} m³\n"
            f"{'─'*32}\n"
            f"Capital: {self._isk(tc)}\n"
            f"Lucro:   {self._isk(tp)}\n"
            f"ROI:     {roi:.1f}%\n"
            f"Viagens: {viagens} "
            f"({'1 viagem' if viagens==1 else f'{viagens} viagens p/ {cap:,.0f} m³'})\n"
            f"Tempo:   {time_h:.1f}h est. ({spj}s/pulo)\n"
            f"ISK/h:   {self._isk(isk_h)}\n"))

        self._lbl_opt.config(fg=COLORS["text"], text=(
            f"NN:     {nn} pulos\n"
            f"2-opt:  {tj} pulos  (-{saved} pulos)\n"
            f"Cache:  {hits}/{npairs} pares "
            f"({int(hits/max(npairs,1)*100)}% hit rate)\n"))

        # Lista de paradas
        self._stop_list.delete(0,"end")
        self._stop_list.insert("end",
            f"  {'PARTIDA':>2}  {p['base_sys']:<18}  base de operações")
        for s in p["stops"]:
            cj = s["cumulative_jumps"]
            jp = s["jumps_from_prev"]
            self._stop_list.insert("end",
                f"  {s['seq']:>2}.  {s['system_name']:<18}"
                f"  +{jp:>2}j ({cj:>3}j)  {s['n_items']}it  {s['total_vol']:.0f}m³")
        self._stop_list.insert("end",
            f"  {'RETORNO':>2}  {p['base_sys']:<18}"
            f"  +{p['return_jumps']:>2}j ({tj:>3}j)  — venda")

        self._compute_layout()
        self._center()

    # ── Painel lista por parada ───────────────────────────────────────────────

    def _on_stop_sel(self, event):
        sel = self._stop_list.curselection()
        if not sel or not self._plan: return
        idx = sel[0] - 1   # -1 porque linha 0 é "PARTIDA"
        stops = self._plan["stops"]
        if idx < 0 or idx >= len(stops):
            return
        stop = stops[idx]
        self._sel_stop = idx
        self._sel_sys  = stop["system_id"]

        # Atualiza aba Lista
        self._lbl_stop_hdr.config(
            text=(f"Parada #{stop['seq']}: {stop['system_name']}\n"
                  f"Pulos da base: {stop['jumps_from_base']} | "
                  f"Do anterior:   {stop['jumps_from_prev']}\n"
                  f"Acumulado:     {stop['cumulative_jumps']}\n"
                  f"Vol. total:    {stop['total_vol']:.1f} m³\n"
                  f"Capital:       {self._isk(stop['total_capital'])}\n"
                  f"Lucro est.:    {self._isk(stop['total_profit'])}\n"
                  f"\nEstações nesta parada:\n" +
                  "\n".join(f"  {s[:42]}" for s in stop["stations"])),
            fg=COLORS["text"])

        self._item_list.delete(0,"end")
        for it in stop["items"]:
            item = (it.get("Item") or "?")
            qty  = it.get("Qtd","?")
            prec = self._isk(self._num(it.get("Preco")))
            cap  = self._isk(self._num(it.get("Capital")))
            vol  = self._num(it.get("Vol"))
            self._item_list.insert("end",
                f"  {item[:28]}")
            self._item_list.insert("end",
                f"    Qtd:{qty}  Preco:{prec}  Capital:{cap}  Vol:{vol:.2f}m³")

        self._nb.select(2)
        self._redraw()

    # ── Layout e renderização ─────────────────────────────────────────────────

    def _compute_layout(self):
        if not self._all_sys: return
        pts = {sid:(i["x"],i["z"]) for sid,i in self._all_sys.items()}
        xs  = [p[0] for p in pts.values()]
        zs  = [p[1] for p in pts.values()]
        sx  = (max(xs)-min(xs)) or 1
        sz  = (max(zs)-min(zs)) or 1
        self._norm_pos = {
            sid: ((x-min(xs))/sx*2-1, (z-min(zs))/sz*2-1)
            for sid,(x,z) in pts.items()}

    def _w2c(self, nx, nz):
        w  = self._canvas.winfo_width()  or 1100
        h  = self._canvas.winfo_height() or 700
        mg = 70
        uw = (w-mg*2)*self._scale
        uh = (h-mg*2)*self._scale
        return (mg+(nx+1)/2*uw+self._off_x,
                mg+(nz+1)/2*uh+self._off_y)

    def _redraw(self):
        c = self._canvas
        c.delete("all")
        if not self._all_sys or not self._norm_pos or not self._plan:
            c.create_text(
                (self._canvas.winfo_width() or 1000)//2,
                (self._canvas.winfo_height() or 600)//2,
                text="Calculando rota...",
                fill=COLORS["subtext"], font=("Courier",13), justify="center")
            return

        p      = self._plan
        base_id = p["base_id"]
        stops  = p["stops"]
        stop_ids = [s["system_id"] for s in stops]
        full_seq = [base_id] + stop_ids + [base_id]   # rota completa: ida e volta

        # Stargates (fundo)
        stop_set = set(full_seq)
        for a,b in self._gate_edges:
            if a not in self._norm_pos or b not in self._norm_pos: continue
            ax,ay = self._w2c(*self._norm_pos[a])
            bx,by = self._w2c(*self._norm_pos[b])
            in_r  = a in stop_set and b in stop_set
            c.create_line(ax,ay,bx,by,
                          fill=COLORS["route_out"] if in_r else COLORS["gate"],
                          width=1.8 if in_r else 0.6,
                          dash=() if in_r else (3,8))

        # Linhas da rota (ida = laranja, volta = azul)
        mid = len(full_seq)-1   # índice do passo de retorno
        for i in range(len(full_seq)-1):
            a = full_seq[i]; b = full_seq[i+1]
            if a not in self._norm_pos or b not in self._norm_pos: continue
            ax,ay = self._w2c(*self._norm_pos[a])
            bx,by = self._w2c(*self._norm_pos[b])
            color  = COLORS["route_ret"] if i == len(full_seq)-2 else COLORS["route_out"]
            c.create_line(ax,ay,bx,by, fill=color, width=2.2,
                          dash=(5,3), arrow="last", arrowshape=(10,12,4))

        # Nós
        nr = max(5, int(8*self._scale))
        self._nodes = {}
        for sid, info in self._all_sys.items():
            if sid not in self._norm_pos: continue
            cx,cy = self._w2c(*self._norm_pos[sid])
            self._nodes[sid] = (cx,cy)

            is_base = sid == base_id
            si      = next((i for i,s in enumerate(stops) if s["system_id"]==sid), -1)
            is_sel  = sid == self._sel_sys

            if is_sel:
                fill,r,outl = COLORS["sys_hi"],  nr+4,"#ffffff"
            elif is_base:
                fill,r,outl = COLORS["sys_base"],nr+3,COLORS["yellow"]
            elif si >= 0:
                fill,r,outl = COLORS["sys_stop"],nr+2,COLORS["border"]
            else:
                fill,r,outl = sec_color(info["sec"]),nr,COLORS["gate"]

            c.create_oval(cx-r,cy-r,cx+r,cy+r, fill=fill, outline=outl,
                          width=1.5, tags=("node",f"sys_{sid}"))

            if si >= 0:
                c.create_text(cx,cy, text=str(si+1),
                              fill="#fff" if not is_sel else "#000",
                              font=("Courier",max(6,int(7*self._scale)),"bold"))

            if is_base or si >= 0 or is_sel or self._scale > 1.4:
                lc = (COLORS["yellow"] if is_sel else
                      COLORS["green"]  if is_base else COLORS["text"])
                extra = f" ({stops[si]['n_items']}it)" if si >= 0 else ""
                c.create_text(cx, cy+r+9,
                              text=f"{info['name']}{extra}",
                              fill=lc, font=("Courier",max(7,int(8*self._scale))))

        c.tag_bind("node","<Button-1>",   self._on_node_click)
        c.tag_bind("node","<Double-1>",   self._on_node_dbl)
        c.tag_bind("node","<Enter>",      self._on_node_enter)
        c.tag_bind("node","<Leave>",      self._on_node_leave)

    # ── Interação ─────────────────────────────────────────────────────────────

    def _on_dn(self,e):  self._drag=(e.x,e.y)
    def _on_up(self,e):  self._drag=None

    def _on_drag(self,e):
        if self._drag:
            self._off_x += e.x-self._drag[0]
            self._off_y += e.y-self._drag[1]
            self._drag   = (e.x,e.y)
            self._redraw()

    def _on_scroll(self,e):
        up = e.num==4 or (hasattr(e,"delta") and e.delta>0)
        self._scale = max(0.1,min(12.0,self._scale*(1.15 if up else 1/1.15)))
        self._redraw()

    def _on_node_click(self,e):
        for t in self._canvas.gettags("current"):
            if t.startswith("sys_"):
                sid = int(t[4:])
                self._sel_sys = sid
                self._show_sys_info(sid)
                self._redraw()
                break

    def _on_node_dbl(self,e):
        """Double-click: seleciona a parada correspondente na lista."""
        for t in self._canvas.gettags("current"):
            if t.startswith("sys_") and self._plan:
                sid = int(t[4:])
                for i,s in enumerate(self._plan["stops"]):
                    if s["system_id"] == sid:
                        self._stop_list.selection_clear(0,"end")
                        self._stop_list.selection_set(i+1)   # +1 pela linha PARTIDA
                        self._stop_list.see(i+1)
                        self._on_stop_sel(None)
                        break
                break

    def _on_node_enter(self,e):
        for t in self._canvas.gettags("current"):
            if t.startswith("sys_"):
                info = self._all_sys.get(int(t[4:]),{})
                self._statusbar.config(
                    text=f"  {info.get('name','?')}  |  "
                         f"Seg: {info.get('sec',0):.1f}  |  "
                         f"Planetas: {info.get('planets',0)}  |  ID: {t[4:]}")
                break

    def _on_node_leave(self,e):
        self._statusbar.config(
            text="Scroll=zoom  |  Arraste=mover  |  Clique=info  |  Dbl-clique=selecionar parada")

    def _center(self):
        self._scale=1.0; self._off_x=0.0; self._off_y=0.0
        self._redraw()

    # ── Info de sistema ───────────────────────────────────────────────────────

    def _show_sys_info(self, sid):
        info = self._all_sys.get(sid,{})
        if not info: return
        plan = self._plan
        if not plan: return
        sec  = info.get("sec",0)
        sl   = "HiSec" if sec>=0.5 else "LowSec" if sec>=0.1 else "Null"

        stop = next((s for s in plan["stops"] if s["system_id"]==sid), None)
        if sid == plan["base_id"]:
            role = "BASE (partida/chegada)"
            detail = f"Itens totais a vender: {plan['total_items']}\nLucro est.: {self._isk(plan['total_profit'])}"
        elif stop:
            role   = f"Parada #{stop['seq']} — compra"
            detail = (f"Pulos da base: {stop['jumps_from_base']}\n"
                      f"Itens: {stop['n_items']}\n"
                      f"Vol: {stop['total_vol']:.1f} m³\n"
                      f"Capital: {self._isk(stop['total_capital'])}\n"
                      f"Lucro: {self._isk(stop['total_profit'])}")
        else:
            role   = "sistema intermediário"
            detail = ""

        self._lbl_sys.config(
            text=(f"Sistema: {info['name']}\n"
                  f"Função:  {role}\n"
                  f"Seg.:    {sec:.2f} ({sl})\n"
                  f"Planetas:{info.get('planets',0)}\n"
                  f"ID ESI:  {sid}\n"
                  f"{'─'*26}\n{detail}"),
            fg=COLORS["text"])

    # ── Exportar plano ────────────────────────────────────────────────────────

    def _export_plan(self):
        if not self._plan:
            messagebox.showinfo("Sem dados","Calcule uma rota primeiro."); return
        p    = self._plan
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Texto","*.txt"),("Todos","*.*")],
            initialfile="shopping_route.txt")
        if not path: return

        lines = [
            "═"*62,
            "  EVE SHOPPING ROUTE — Plano de Compras",
            "═"*62,
            f"  Base:    {p['base_station']}",
            f"  Venda:   {p['sell_station']}",
            f"  Pulos:   {p['total_jumps']} (round-trip)",
            f"  Paradas: {len(p['stops'])} sistemas",
            f"  Itens:   {p['total_items']}",
            f"  Capital: {self._isk(p['total_capital'])}",
            f"  Lucro:   {self._isk(p['total_profit'])}",
            f"  Volume:  {p['total_vol']:,.1f} m³",
            "═"*62, "",
        ]
        lines.append(f"  PARTIDA — {p['base_sys']}  (base de operações)")
        lines.append("")
        for s in p["stops"]:
            lines.append(f"  {'─'*58}")
            lines.append(f"  Parada {s['seq']:>2}: {s['system_name']}"
                         f"   +{s['jumps_from_prev']}j  "
                         f"({s['cumulative_jumps']}j acumulados)")
            for st in s["stations"]:
                lines.append(f"    Estação: {st}")
            lines.append(f"    Volume total: {s['total_vol']:.1f} m³  "
                         f"Capital: {self._isk(s['total_capital'])}")
            lines.append("")
            for it in s["items"]:
                item = (it.get("Item") or "?")
                qty  = it.get("Qtd","?")
                prec = self._isk(self._num(it.get("Preco")))
                vol  = self._num(it.get("Vol"))
                lines.append(f"    [ ] {item}")
                lines.append(f"        Qtd: {qty}  Preço: {prec}  Vol: {vol:.2f} m³")
            lines.append("")
        lines.append(f"  {'─'*58}")
        lines.append(f"  RETORNO — {p['base_sys']}"
                     f"   +{p['return_jumps']}j  ({p['total_jumps']}j total)")
        lines.append(f"  VENDA em: {p['sell_station']}")
        lines.append("")
        lines.append("═"*62)

        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        messagebox.showinfo("Exportado", f"Plano salvo em:\n{path}")

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _num(self, v):
        try: return float(str(v or 0).replace(",","").replace(" ",""))
        except: return 0.0

    def _isk(self, v):
        v = self._num(v)
        if v >= 1e9: return f"{v/1e9:.2f} B"
        if v >= 1e6: return f"{v/1e6:.2f} M"
        return f"{v:,.0f}"

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if not HAS_REQUESTS:
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Dependência ausente","pip install requests")
        return

    root = tk.Tk(); root.withdraw()

    # Arquivo padrão se passado como argumento
    default_file = sys.argv[1] if len(sys.argv)>1 and os.path.exists(sys.argv[1]) else ""

    dlg = SetupDialog(root, default_file)
    root.destroy()

    if not dlg.result:
        return   # Usuário cancelou

    app = EVEShoppingRoute(dlg.result)
    app.mainloop()


if __name__ == "__main__":
    main()
