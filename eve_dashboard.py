"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  EVE MARKET TERMINAL  ·  Bloomberg-style dashboard for EVE Online trading   ║
║  Importa dados diretamente do eve_trade.py — execute na mesma pasta.        ║
╚══════════════════════════════════════════════════════════════════════════════╝

Uso:   python3 eve_dashboard.py
Req.:  eve_trade.py na mesma pasta  |  pip install requests
"""

import sys, os, json, time, threading, concurrent.futures
from datetime import datetime, timezone, timedelta

# ── Resolve o path do eve_trade.py ────────────────────────────────────────────
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

try:
    import eve_trade as et
except ImportError as _e:
    print(f"[ERRO] Não foi possível importar eve_trade.py: {_e}")
    print("       Coloque eve_dashboard.py na mesma pasta que eve_trade.py")
    sys.exit(1)

import tkinter as tk
from tkinter import ttk, font as tkfont

# ── Paleta Bloomberg Terminal ──────────────────────────────────────────────────
T = {
    "bg":        "#0a0a0a",    # fundo quase preto
    "panel":     "#0e1117",    # painéis
    "border":    "#1a2332",    # bordas sutis
    "header_bg": "#0f1824",    # cabeçalho de painel
    "title":     "#ff7b00",    # laranja Bloomberg
    "header":    "#00c8ff",    # azul ciano
    "text":      "#c8d8e8",    # texto principal
    "dim":       "#4a6080",    # texto secundário
    "green":     "#00e676",    # positivo
    "red":       "#ff1744",    # negativo
    "yellow":    "#ffd600",    # alerta
    "cyan":      "#00e5ff",    # destaque
    "purple":    "#e040fb",    # accent secundário
    "white":     "#ffffff",    # máximo destaque
    "grid":      "#111820",    # linhas de grade
    "hot":       "#ff6d00",    # item quente
    "neutral":   "#607d8b",    # neutro
}

MONO  = ("Courier New", 9)
MONO_B= ("Courier New", 9, "bold")
MONO_S= ("Courier New", 8)
MONO_L= ("Courier New", 11, "bold")
MONO_XL=("Courier New", 13, "bold")

# ── Itens de referência para o ticker ─────────────────────────────────────────
TICKER_ITEMS = [
    (29668, "PLEX",          10000002),
    (40520, "INJECTOR",      10000002),
    (34,    "TRITANIUM",     10000002),
    (35,    "PYERITE",       10000002),
    (38,    "NOCXIUM",       10000002),
    (39,    "ZYDRINE",       10000002),
    (40,    "MEGACYTE",      10000002),
    (16273, "STRONTIUM",     10000002),
    (4051,  "HE-ISOTOPES",   10000002),
    (17887, "N-ISOTOPES",    10000002),
    (22,    "ARKONOR",       10000002),
    (11396, "MERCOXIT",      10000002),
    (2411,  "ROBOTICS",      10000002),
    (9830,  "CONS.ELECTR.",  10000002),
]

# ── Ciclos: grupos e itens representativos ────────────────────────────────────
CYCLE_WATCH = {
    "Minerais": [34, 35, 36, 37, 38, 39, 40],
    "PLEX/Inj.": [29668, 40520],
    "Ice":       [16262, 16263, 16264, 16265],
    "Moon R4":   [45491, 45490, 45492, 45493],
    "PI P2":     [2317, 2321, 9830, 2411],
    "Fragatas":  [587, 588, 582, 583],
    "Naves T2":  [11176, 11178, 12003, 12005],
}

HUB_RID = dict(et.HUB_REGIONS)   # {name: rid}

# ─────────────────────────────────────────────────────────────────────────────
#  Utilitários
# ─────────────────────────────────────────────────────────────────────────────

def _fmt_isk(v, abbrev=True):
    if v is None: return "--"
    if abbrev:
        if abs(v) >= 1e12: return f"{v/1e12:.2f}T"
        if abs(v) >= 1e9:  return f"{v/1e9:.2f}B"
        if abs(v) >= 1e6:  return f"{v/1e6:.1f}M"
        if abs(v) >= 1e3:  return f"{v/1e3:.0f}K"
    return f"{v:,.0f}"

def _fmt_pct(v):
    if v is None: return "--"
    return f"{v:+.2f}%"

def _color_pct(v):
    if v is None: return T["dim"]
    if v > 1:   return T["green"]
    if v < -1:  return T["red"]
    return T["yellow"]

def _color_score(s):
    if s >= 65: return T["green"]
    if s >= 40: return T["yellow"]
    return T["red"]

def _now_str():
    return datetime.now(timezone.utc).strftime("%H:%M:%S UTC")

# ─────────────────────────────────────────────────────────────────────────────
#  Widget auxiliares reutilizáveis
# ─────────────────────────────────────────────────────────────────────────────

class PanelFrame(tk.Frame):
    """Painel com borda, título e área de conteúdo — estilo Bloomberg."""
    def __init__(self, master, title, badge=None, **kw):
        super().__init__(master, bg=T["panel"],
                         highlightbackground=T["border"],
                         highlightthickness=1, **kw)
        # Header
        hdr = tk.Frame(self, bg=T["header_bg"], height=22)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        tk.Label(hdr, text=f"  {title}",
                 bg=T["header_bg"], fg=T["title"],
                 font=MONO_B, anchor="w").pack(side="left")

        if badge:
            self._badge = tk.Label(hdr, text=badge,
                                   bg=T["border"], fg=T["dim"],
                                   font=MONO_S, padx=4)
            self._badge.pack(side="right", padx=4)

        # Área de conteúdo
        self.body = tk.Frame(self, bg=T["panel"])
        self.body.pack(fill="both", expand=True)

    def set_badge(self, text, color=None):
        if hasattr(self, '_badge'):
            self._badge.config(text=text,
                                fg=color or T["dim"])


class BlinkLabel(tk.Label):
    """Label que pisca brevemente ao atualizar — feedback visual de refresh."""
    def __init__(self, *a, **kw):
        super().__init__(*a, **kw)
        self._orig_fg = kw.get("fg", T["text"])

    def flash(self, color=T["cyan"]):
        self.config(fg=color)
        self.after(300, lambda: self.config(fg=self._orig_fg))


class ScrollableFrame(tk.Frame):
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        vsb = tk.Scrollbar(self, orient="vertical",
                            bg=T["panel"], troughcolor=T["bg"],
                            width=8)
        vsb.pack(side="right", fill="y")
        self._canvas = tk.Canvas(self, bg=T["panel"],
                                 highlightthickness=0,
                                 yscrollcommand=vsb.set)
        self._canvas.pack(side="left", fill="both", expand=True)
        vsb.config(command=self._canvas.yview)
        self.inner = tk.Frame(self._canvas, bg=T["panel"])
        self._win  = self._canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>",
                        lambda e: self._canvas.configure(
                            scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>",
                          lambda e: self._canvas.itemconfig(self._win, width=e.width))
        self.inner.bind_all("<MouseWheel>",
                            lambda e: self._canvas.yview_scroll(
                                int(-1*(e.delta/120)), "units"))


# ─────────────────────────────────────────────────────────────────────────────
#  Ticker horizontal animado (topo)
# ─────────────────────────────────────────────────────────────────────────────

class TickerBar(tk.Frame):
    """Barra de ticker Bloomberg com preços rolantes."""
    _REFRESH_INTERVAL = 120_000   # 2 min

    def __init__(self, master, **kw):
        super().__init__(master, bg=T["bg"], height=28, **kw)
        self.pack_propagate(False)

        # Faixa laranja esquerda (logo)
        logo = tk.Frame(self, bg=T["title"], width=120)
        logo.pack(side="left", fill="y")
        logo.pack_propagate(False)
        tk.Label(logo, text="⬡ EVE MARKET",
                 bg=T["title"], fg="#000",
                 font=("Courier New", 9, "bold")).place(relx=.5, rely=.5, anchor="center")

        # Separador
        tk.Frame(self, bg=T["border"], width=1).pack(side="left", fill="y")

        # Área do ticker (canvas para scroll)
        self._canvas = tk.Canvas(self, bg=T["bg"], height=28,
                                 highlightthickness=0)
        self._canvas.pack(side="left", fill="both", expand=True)

        # Timestamp direito
        self._ts_lbl = tk.Label(self, text="--:--:-- UTC",
                                 bg=T["bg"], fg=T["dim"],
                                 font=MONO_S, padx=8)
        self._ts_lbl.pack(side="right")

        self._items: list = []
        self._x     = 0
        self._speed = 1.5
        self._text_items = []
        self._loading = False

        self._canvas.bind("<Configure>", self._on_resize)
        self._after_id_scroll = None
        self._after_id_refresh = None

        self.refresh_data()

    def _on_resize(self, _e=None):
        self._x = self._canvas.winfo_width()
        self._rerender()

    def refresh_data(self):
        if self._loading: return
        self._loading = True
        self._ts_lbl.config(text="  LOADING...  ", fg=T["yellow"])
        threading.Thread(target=self._fetch, daemon=True).start()

    def _fetch(self):
        results = {}
        def _get(tid, rid):
            try:
                d = et.fetch_live_price(rid, tid)
                if d: results[tid] = d
            except Exception:
                pass
        with concurrent.futures.ThreadPoolExecutor(max_workers=6) as pool:
            list(pool.map(lambda t: _get(t[0], t[2]), TICKER_ITEMS))

        self._items = []
        ic = et.load_item_cache()
        for tid, label, rid in TICKER_ITEMS:
            d = results.get(tid)
            if d:
                self._items.append({
                    "label":    label,
                    "sell":     d.get("sell_min", 0),
                    "buy":      d.get("buy_max", 0),
                    "spread":   d.get("spread_pct", 0),
                    "net":      d.get("net_pct", 0),
                })
        self._loading = False
        self.after(0, self._rerender)
        self.after(0, lambda: self._ts_lbl.config(
            text=f"  {_now_str()}  ", fg=T["dim"]))
        if self._after_id_refresh:
            self.after_cancel(self._after_id_refresh)
        self._after_id_refresh = self.after(self._REFRESH_INTERVAL, self.refresh_data)

    def _rerender(self):
        self._canvas.delete("all")
        self._text_items = []
        if not self._items: return

        # Monta a string do ticker
        parts = []
        for it in self._items:
            sell = _fmt_isk(it["sell"])
            net  = f"{it['net']:+.1f}%" if it.get("net") else ""
            parts.append((it["label"], sell, it.get("net", 0), net))

        # Renderiza no canvas
        x = self._x if self._x > 0 else self._canvas.winfo_width()
        y = 14
        self._text_items = []
        for label, sell, net_val, net_str in parts:
            col = T["green"] if net_val > 0.5 else (T["red"] if net_val < -0.5 else T["yellow"])
            # Nome
            id1 = self._canvas.create_text(x, y, text=f" {label} ",
                                            font=MONO_B, fill=T["dim"],
                                            anchor="w")
            x += len(f" {label} ") * 7
            # Preço
            id2 = self._canvas.create_text(x, y, text=sell,
                                            font=MONO_B, fill=T["white"],
                                            anchor="w")
            x += len(sell) * 7 + 2
            # Net %
            if net_str:
                id3 = self._canvas.create_text(x, y, text=f"({net_str})",
                                                font=MONO_S, fill=col, anchor="w")
                x += len(f"({net_str})") * 6 + 2
            # Separador
            id4 = self._canvas.create_text(x, y, text=" ·",
                                            font=MONO_S, fill=T["border"], anchor="w")
            x += 14
            self._text_items.extend([id1, id2, id4])

        self._total_w = x
        self._x = self._canvas.winfo_width()
        self._scroll()

    def _scroll(self):
        if self._after_id_scroll:
            try: self.after_cancel(self._after_id_scroll)
            except Exception: pass
        self._do_scroll()

    def _do_scroll(self):
        if not self._text_items: return
        for item in self._text_items:
            self._canvas.move(item, -self._speed, 0)
        # Verifica se todos saíram da tela → reinicia
        coords = self._canvas.coords(self._text_items[0])
        if coords and coords[0] < -self._total_w:
            # Reposiciona tudo
            delta = self._canvas.winfo_width() - coords[0]
            for item in self._text_items:
                self._canvas.move(item, delta, 0)
        self._after_id_scroll = self.after(16, self._do_scroll)  # ~60fps


# ─────────────────────────────────────────────────────────────────────────────
#  Painel: Banco de Dados
# ─────────────────────────────────────────────────────────────────────────────

class DBStatusPanel(PanelFrame):
    _REFRESH = 30_000

    def __init__(self, master, **kw):
        super().__init__(master, "DATABASE STATUS", **kw)
        b = self.body
        self._rows = {}
        hubs = list(HUB_RID.keys())
        for i, hub in enumerate(hubs):
            bg = T["panel"] if i % 2 == 0 else T["grid"]
            row_fr = tk.Frame(b, bg=bg)
            row_fr.pack(fill="x", padx=2, pady=0)
            tk.Label(row_fr, text=f"  {hub:<8}",
                     bg=bg, fg=T["cyan"], font=MONO_B, width=10).pack(side="left")
            lbl_items = tk.Label(row_fr, text="--",
                                  bg=bg, fg=T["dim"], font=MONO_S, width=8)
            lbl_items.pack(side="left")
            lbl_dates = tk.Label(row_fr, text="--",
                                  bg=bg, fg=T["dim"], font=MONO_S, width=26)
            lbl_dates.pack(side="left")
            lbl_status = tk.Label(row_fr, text="",
                                   bg=bg, fg=T["green"], font=MONO_S, width=6)
            lbl_status.pack(side="right", padx=4)
            self._rows[hub] = (lbl_items, lbl_dates, lbl_status)

        self.refresh()

    def refresh(self):
        threading.Thread(target=self._fetch, daemon=True).start()
        self.after(self._REFRESH, self.refresh)

    def _fetch(self):
        for hub, rid in HUB_RID.items():
            try:
                st = et._db_stats(rid)
                n  = st["n_items"]
                d1 = st["min_date"]
                d2 = st["max_date"]
                # Verifica se está atualizado
                yesterday = (datetime.now(timezone.utc) - timedelta(days=1)).strftime("%Y-%m-%d")
                up_to_date = d2 >= yesterday if d2 != "--" else False
                self.after(0, lambda h=hub, n=n, d1=d1, d2=d2, u=up_to_date:
                           self._update_row(h, n, d1, d2, u))
            except Exception:
                pass

    def _update_row(self, hub, n, d1, d2, up_to_date):
        lbl_i, lbl_d, lbl_s = self._rows[hub]
        lbl_i.config(text=f"{n:,} itens", fg=T["text"] if n > 0 else T["red"])
        lbl_d.config(text=f"{d1} → {d2}",
                     fg=T["dim"] if d2 == "--" else T["text"])
        lbl_s.config(text="● OK" if up_to_date else "● OLD",
                     fg=T["green"] if up_to_date else T["yellow"])
        self.set_badge(f"atualizado {_now_str()}", T["dim"])


# ─────────────────────────────────────────────────────────────────────────────
#  Painel: Watchlist com preços live
# ─────────────────────────────────────────────────────────────────────────────

class WatchlistPanel(PanelFrame):
    _REFRESH = 60_000

    def __init__(self, master, **kw):
        super().__init__(master, "WATCHLIST  ·  LIVE PRICES", **kw)
        # Header row
        hdr = tk.Frame(self.body, bg=T["header_bg"])
        hdr.pack(fill="x", padx=2, pady=(2, 0))
        for txt, w in [("ITEM", 18), ("SELL MIN", 12), ("BUY MAX", 12),
                        ("SPREAD%", 8), ("NET%", 8), ("STATUS", 8)]:
            tk.Label(hdr, text=txt, bg=T["header_bg"], fg=T["dim"],
                     font=MONO_S, width=w, anchor="e").pack(side="left")

        self._sf = ScrollableFrame(self.body, bg=T["panel"])
        self._sf.pack(fill="both", expand=True, padx=2)
        self._row_widgets = {}
        self._wl = []
        self._loading = False

        self.refresh()

    def refresh(self):
        if self._loading: return
        self._loading = True
        self._wl = et.load_watchlist()
        threading.Thread(target=self._fetch, daemon=True).start()
        self.after(self._REFRESH, self.refresh)

    def _fetch(self):
        ic = et.load_item_cache()
        results = {}
        def _get(w):
            tid = w.get("type_id")
            if tid:
                d = et.fetch_live_price(10000002, tid)
                if d: results[tid] = d
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
            list(pool.map(_get, self._wl[:20]))
        self._loading = False
        self.after(0, lambda: self._render(results, ic))

    def _render(self, results, ic):
        # Limpa e reconstrói
        for w in self._sf.inner.winfo_children():
            w.destroy()

        n = 0
        for i, item in enumerate(self._wl[:20]):
            tid   = item.get("type_id")
            name  = item.get("name", str(tid))
            alert_above = item.get("alert_above", 0)
            alert_below = item.get("alert_below", 0)
            d = results.get(tid)
            bg = T["panel"] if i % 2 == 0 else T["grid"]
            row = tk.Frame(self._sf.inner, bg=bg)
            row.pack(fill="x", padx=0, pady=0)

            # Nome
            tk.Label(row, text=f" {name[:17]:<17}",
                     bg=bg, fg=T["text"], font=MONO_S, width=18,
                     anchor="w").pack(side="left")

            if d:
                sell = d.get("sell_min", 0)
                buy  = d.get("buy_max", 0)
                sp   = d.get("spread_pct", 0)
                net  = d.get("net_pct", 0)

                # Status de alerta
                status, sc = "  —", T["dim"]
                if alert_above and sell >= alert_above:
                    status, sc = " ▲ HI", T["green"]
                elif alert_below and alert_below > 0 and sell <= alert_below:
                    status, sc = " ▼ LO", T["red"]

                for val, w, col in [
                    (_fmt_isk(sell), 12, T["white"]),
                    (_fmt_isk(buy),  12, T["cyan"]),
                    (f"{sp:.2f}%",   8,  T["yellow"] if sp > 3 else T["dim"]),
                    (f"{net:+.2f}%", 8,  T["green"] if net > 0 else T["red"]),
                    (status,         8,  sc),
                ]:
                    tk.Label(row, text=val, bg=bg, fg=col,
                             font=MONO_S, width=w, anchor="e").pack(side="left")
            else:
                tk.Label(row, text="  sem dados live",
                         bg=bg, fg=T["dim"], font=MONO_S).pack(side="left")
            n += 1

        self.set_badge(f"{n} itens  ·  {_now_str()}", T["dim"])


# ─────────────────────────────────────────────────────────────────────────────
#  Painel: Paper Trading P&L
# ─────────────────────────────────────────────────────────────────────────────

class PaperTradingPanel(PanelFrame):
    _REFRESH = 30_000

    def __init__(self, master, **kw):
        super().__init__(master, "PAPER TRADING  ·  P&L", **kw)
        b = self.body

        # Stats topo
        stats_fr = tk.Frame(b, bg=T["header_bg"])
        stats_fr.pack(fill="x", padx=2, pady=(2, 1))
        self._lbl_balance  = self._stat(stats_fr, "SALDO",    "--")
        self._lbl_pnl      = self._stat(stats_fr, "P&L TOTAL","--")
        self._lbl_winrate  = self._stat(stats_fr, "WIN RATE", "--")
        self._lbl_trades   = self._stat(stats_fr, "TRADES",   "--")
        self._lbl_open     = self._stat(stats_fr, "POSIÇÕES", "--")

        # Header da tabela
        hdr = tk.Frame(b, bg=T["header_bg"])
        hdr.pack(fill="x", padx=2, pady=0)
        for txt, w in [("ITEM", 16), ("COMPRA", 12), ("CUSTO TOTAL", 12),
                        ("ABERTO", 10), ("QTD", 6)]:
            tk.Label(hdr, text=txt, bg=T["header_bg"], fg=T["dim"],
                     font=MONO_S, width=w, anchor="e").pack(side="left")

        self._sf = ScrollableFrame(b, bg=T["panel"])
        self._sf.pack(fill="both", expand=True, padx=2)

        self.refresh()

    def _stat(self, parent, label, value):
        fr = tk.Frame(parent, bg=T["header_bg"], padx=6)
        fr.pack(side="left", pady=2)
        tk.Label(fr, text=label, bg=T["header_bg"],
                 fg=T["dim"], font=MONO_S).pack()
        lbl = tk.Label(fr, text=value, bg=T["header_bg"],
                        fg=T["white"], font=MONO_B)
        lbl.pack()
        return lbl

    def refresh(self):
        threading.Thread(target=self._fetch, daemon=True).start()
        self.after(self._REFRESH, self.refresh)

    def _fetch(self):
        data = et.load_paper_trades()
        self.after(0, lambda: self._render(data))

    def _render(self, data):
        balance   = data.get("balance", 0)
        positions = data.get("positions", [])
        history   = data.get("history", [])

        # P&L do histórico
        pnl_total = sum(t.get("pnl", 0) for t in history)
        wins      = sum(1 for t in history if t.get("pnl", 0) > 0)
        winrate   = wins / len(history) * 100 if history else 0

        pnl_col = T["green"] if pnl_total >= 0 else T["red"]
        self._lbl_balance.config(text=_fmt_isk(balance), fg=T["cyan"])
        self._lbl_pnl.config(text=f"{_fmt_isk(pnl_total)}",
                               fg=pnl_col)
        self._lbl_winrate.config(
            text=f"{winrate:.0f}%",
            fg=T["green"] if winrate >= 50 else T["yellow"])
        self._lbl_trades.config(text=str(len(history)), fg=T["text"])
        self._lbl_open.config(text=str(len(positions)), fg=T["yellow"])

        # Limpa e recostrói tabela de posições abertas
        for w in self._sf.inner.winfo_children():
            w.destroy()

        if not positions:
            tk.Label(self._sf.inner, text="  Sem posições abertas.",
                     bg=T["panel"], fg=T["dim"], font=MONO_S).pack(anchor="w")
        else:
            for i, pos in enumerate(positions[:10]):
                bg = T["panel"] if i % 2 == 0 else T["grid"]
                row = tk.Frame(self._sf.inner, bg=bg)
                row.pack(fill="x")
                name = pos.get("name", "?")[:16]
                buy  = pos.get("buy_price", 0)
                cost = pos.get("total_cost", 0)
                qty  = pos.get("qty", 0)
                opened = pos.get("opened_at", "")[:10]
                for val, w, col in [
                    (f" {name:<15}", 16, T["text"]),
                    (_fmt_isk(buy),  12, T["dim"]),
                    (_fmt_isk(cost), 12, T["yellow"]),
                    (opened,         10, T["dim"]),
                    (str(qty),        6, T["white"]),
                ]:
                    tk.Label(row, text=val, bg=bg, fg=col,
                             font=MONO_S, width=w, anchor="e").pack(side="left")

        self.set_badge(f"atualizado {_now_str()}", T["dim"])


# ─────────────────────────────────────────────────────────────────────────────
#  Painel: Top Oportunidades (da análise de Max Margem no CSV local)
# ─────────────────────────────────────────────────────────────────────────────

class OpportunitiesPanel(PanelFrame):
    _REFRESH = 300_000   # 5 min

    def __init__(self, master, **kw):
        super().__init__(master, "TOP OPORTUNIDADES  ·  JITA", **kw)
        b = self.body

        # Header
        hdr = tk.Frame(b, bg=T["header_bg"])
        hdr.pack(fill="x", padx=2, pady=(2, 0))
        for txt, w in [("ITEM", 20), ("SELL", 11), ("BUY", 11),
                        ("MARG.LÍQ%", 9), ("EDP", 10), ("VOL/DIA", 8)]:
            tk.Label(hdr, text=txt, bg=T["header_bg"], fg=T["dim"],
                     font=MONO_S, width=w, anchor="e").pack(side="left")

        self._sf = ScrollableFrame(b, bg=T["panel"])
        self._sf.pack(fill="both", expand=True, padx=2)

        self._loading = False
        self.set_badge("clique REFRESH ou aguarde 5min", T["dim"])

        # Botão refresh
        tk.Button(b, text="  ↻ REFRESH  ",
                  command=self.refresh,
                  bg=T["border"], fg=T["title"],
                  font=MONO_S, relief="flat", cursor="hand2",
                  activebackground=T["title"], activeforeground="#000"
                  ).pack(side="bottom", pady=2)

        self.after(3000, self.refresh)  # load automático após 3s

    def refresh(self):
        if self._loading: return
        self._loading = True
        self.set_badge("⟳ calculando...", T["yellow"])
        threading.Thread(target=self._fetch, daemon=True).start()

    def _fetch(self):
        try:
            cfg = et.load_config()
            br  = int(cfg.get("g_brel", 0))
            ac  = int(cfg.get("g_acc",  0))
            orders = et.load_orders_from_csv("The Forge (Jita)")
            if not orders:
                self.after(0, lambda: self.set_badge("sem CSV — baixe em DADOS & CACHE", T["red"]))
                self._loading = False
                return
            ic = et.load_item_cache()
            results = et.analyze_max_margin(
                orders, "The Forge (Jita)", ic,
                broker_rel=br, accounting=ac,
                min_margin=5.0, min_vol=5, top_n=25)
            self._loading = False
            self.after(0, lambda r=results: self._render(r))
        except Exception as e:
            self._loading = False
            self.after(0, lambda e=str(e): self.set_badge(f"erro: {e[:40]}", T["red"]))

    def _render(self, results):
        for w in self._sf.inner.winfo_children():
            w.destroy()

        for i, r in enumerate(results[:20]):
            hot = r.get("net_margin", 0) >= 15
            bg  = T["hot"] if hot else (T["panel"] if i % 2 == 0 else T["grid"])
            fg_row = "#000" if hot else T["text"]
            row = tk.Frame(self._sf.inner, bg=bg)
            row.pack(fill="x")

            edp = r.get("edp") or 0
            vol = r.get("avg_vol_90") or 0

            for val, w, col in [
                (f" {r['item'][:19]:<19}", 20, fg_row),
                (_fmt_isk(r.get("ps", 0)),  11, fg_row),
                (_fmt_isk(r.get("pb", 0)),  11, T["dim"] if not hot else "#000"),
                (f"{r.get('net_margin',0):.1f}%", 9,
                 T["green"] if not hot else "#000"),
                (_fmt_isk(edp),             10,
                 T["yellow"] if not hot else "#000"),
                (_fmt_isk(vol, abbrev=True), 8, T["dim"] if not hot else "#000"),
            ]:
                tk.Label(row, text=val, bg=bg, fg=col,
                         font=MONO_S if not hot else MONO_B,
                         width=w, anchor="e").pack(side="left")

        n = len(results)
        self.set_badge(f"{n} oportunidades  ·  {_now_str()}", T["green"])
        self.after(self._REFRESH, self.refresh)


# ─────────────────────────────────────────────────────────────────────────────
#  Painel: Sinais de Ciclo de Mercado
# ─────────────────────────────────────────────────────────────────────────────

class CycleSignalPanel(PanelFrame):
    _REFRESH = 180_000   # 3 min
    _PHASE_COLOR = {
        "Acumulação": T["green"],
        "Alta":       T["cyan"],
        "Distribuição": T["yellow"],
        "Queda":      T["red"],
        "Neutro":     T["dim"],
        "Inconclusivo": T["neutral"],
    }

    def __init__(self, master, **kw):
        super().__init__(master, "CICLOS DE MERCADO  ·  JITA", **kw)
        b = self.body

        hdr = tk.Frame(b, bg=T["header_bg"])
        hdr.pack(fill="x", padx=2, pady=(2, 0))
        for txt, w in [("GRUPO", 16), ("FASE", 12), ("MOM7D", 7),
                        ("TREND", 7), ("SINAL", 10)]:
            tk.Label(hdr, text=txt, bg=T["header_bg"], fg=T["dim"],
                     font=MONO_S, width=w, anchor="e").pack(side="left")

        self._rows_frame = tk.Frame(b, bg=T["panel"])
        self._rows_frame.pack(fill="both", expand=True, padx=2)
        self._loading = False

        self.after(5000, self.refresh)  # delay inicial

    def refresh(self):
        if self._loading: return
        self._loading = True
        self.set_badge("⟳ analisando...", T["yellow"])
        threading.Thread(target=self._fetch, daemon=True).start()

    def _fetch(self):
        rid     = HUB_RID["Jita"]
        results = {}
        try:
            et._warm_hist_cache(rid)
        except Exception:
            pass

        def _calc(group_name, tids):
            sigs = []
            for tid in tids[:5]:   # max 5 itens por grupo para performance
                try:
                    hist = et._get_hist_cached(rid, tid)
                    if hist and len(hist) >= 14:
                        sig = et.detect_cycle_item(hist)
                        if sig: sigs.append(sig)
                except Exception:
                    pass
            if not sigs:
                return group_name, {"phase": "Sem dados", "momentum_7d": 0,
                                     "trend": 0, "anomaly": False}
            # Voto de fase
            from collections import Counter
            phases = Counter(s.get("phase", "Neutro") for s in sigs)
            phase  = phases.most_common(1)[0][0]
            avg_mom = sum(s.get("momentum_7d", 0) for s in sigs) / len(sigs)
            avg_tr  = sum(s.get("trend", 0) for s in sigs) / len(sigs)
            return group_name, {
                "phase":       phase,
                "momentum_7d": avg_mom,
                "trend":       avg_tr,
                "n":           len(sigs),
            }

        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
            futures = {pool.submit(_calc, g, t): g
                       for g, t in CYCLE_WATCH.items()}
            for fut in concurrent.futures.as_completed(futures):
                g, sig = fut.result()
                results[g] = sig

        self._loading = False
        self.after(0, lambda r=results: self._render(r))
        self.after(self._REFRESH, self.refresh)

    def _render(self, results):
        for w in self._rows_frame.winfo_children():
            w.destroy()

        for i, (group, sig) in enumerate(results.items()):
            phase   = sig.get("phase", "?")
            mom     = sig.get("momentum_7d", 0)
            trend   = sig.get("trend", 0)
            bg      = T["panel"] if i % 2 == 0 else T["grid"]
            ph_col  = self._PHASE_COLOR.get(phase, T["dim"])

            # Sinal visual
            if "Alta" in phase or "Acumul" in phase:
                sinal, sc = "▲ COMPRAR", T["green"]
            elif "Queda" in phase or "Distrib" in phase:
                sinal, sc = "▼ VENDER", T["red"]
            else:
                sinal, sc = "◇ AGUARDAR", T["yellow"]

            row = tk.Frame(self._rows_frame, bg=bg)
            row.pack(fill="x")
            for val, w, col in [
                (f" {group:<15}", 16, T["text"]),
                (f"{phase:<11}", 12, ph_col),
                (f"{mom:+.3f}", 7, T["green"] if mom > 0 else T["red"]),
                (f"{trend:+.3f}", 7, T["green"] if trend > 0 else T["red"]),
                (f"{sinal:<10}", 10, sc),
            ]:
                tk.Label(row, text=val, bg=bg, fg=col,
                         font=MONO_S, width=w, anchor="e").pack(side="left")

        self.set_badge(f"atualizado {_now_str()}", T["dim"])


# ─────────────────────────────────────────────────────────────────────────────
#  Painel: Hub Comparison (5 hubs × métricas chave)
# ─────────────────────────────────────────────────────────────────────────────

class HubComparePanel(PanelFrame):
    """Compara itens de referência nos 5 hubs principais."""
    _ITEMS = [
        (29668, "PLEX"),
        (40520, "INJECTOR"),
        (34,    "TRITANIUM"),
        (16273, "STRONTIUM"),
    ]
    _REFRESH = 180_000

    def __init__(self, master, **kw):
        super().__init__(master, "HUB COMPARISON  ·  PREÇOS LIVE", **kw)
        b = self.body

        hubs = list(HUB_RID.keys())
        # Cabeçalho de hubs
        hdr = tk.Frame(b, bg=T["header_bg"])
        hdr.pack(fill="x", padx=2, pady=(2, 0))
        tk.Label(hdr, text="ITEM", bg=T["header_bg"], fg=T["dim"],
                 font=MONO_S, width=12, anchor="w").pack(side="left")
        for hub in hubs:
            tk.Label(hdr, text=hub, bg=T["header_bg"], fg=T["dim"],
                     font=MONO_S, width=13, anchor="e").pack(side="left")

        self._body_fr = tk.Frame(b, bg=T["panel"])
        self._body_fr.pack(fill="both", expand=True, padx=2)

        self._loading = False
        self.after(8000, self.refresh)

    def refresh(self):
        if self._loading: return
        self._loading = True
        self.set_badge("⟳ buscando...", T["yellow"])
        threading.Thread(target=self._fetch, daemon=True).start()

    def _fetch(self):
        results = {}
        def _get(args):
            tid, hub, rid = args
            try:
                d = et.fetch_live_price(rid, tid)
                results[(tid, hub)] = d
            except Exception:
                results[(tid, hub)] = None
        work = [(tid, hub, rid) for tid, _ in self._ITEMS for hub, rid in HUB_RID.items()]
        with concurrent.futures.ThreadPoolExecutor(max_workers=8) as pool:
            list(pool.map(_get, work))
        self._loading = False
        self.after(0, lambda r=results: self._render(r))
        self.after(self._REFRESH, self.refresh)

    def _render(self, results):
        for w in self._body_fr.winfo_children():
            w.destroy()
        hubs = list(HUB_RID.keys())

        for row_i, (tid, label) in enumerate(self._ITEMS):
            bg = T["panel"] if row_i % 2 == 0 else T["grid"]
            row = tk.Frame(self._body_fr, bg=bg)
            row.pack(fill="x")
            tk.Label(row, text=f" {label:<11}", bg=bg, fg=T["cyan"],
                     font=MONO_B, width=12, anchor="w").pack(side="left")

            # Coleta preços para encontrar o mínimo (melhor hub)
            prices = {}
            for hub in hubs:
                d = results.get((tid, hub))
                if d: prices[hub] = d.get("sell_min", 0)

            best_hub = min(prices, key=prices.get) if prices else None

            for hub in hubs:
                d = results.get((tid, hub))
                if d:
                    sell = d.get("sell_min", 0)
                    is_best = hub == best_hub and len(prices) > 1
                    col = T["green"] if is_best else T["text"]
                    fnt = MONO_B if is_best else MONO_S
                    txt = _fmt_isk(sell)
                else:
                    txt, col, fnt = "--", T["dim"], MONO_S
                tk.Label(row, text=txt, bg=bg, fg=col,
                         font=fnt, width=13, anchor="e").pack(side="left")

        self.set_badge(f"verde = menor preço  ·  {_now_str()}", T["dim"])


# ─────────────────────────────────────────────────────────────────────────────
#  Painel: Métricas do Sistema
# ─────────────────────────────────────────────────────────────────────────────

class SystemMetricsPanel(PanelFrame):
    _REFRESH = 15_000

    def __init__(self, master, **kw):
        super().__init__(master, "SYSTEM  ·  METRICS", **kw)
        b = self.body

        self._lbls = {}
        metrics = [
            ("watchlist",  "Watchlist"),
            ("positions",  "Posições Abertas"),
            ("history",    "Trades Fechados"),
            ("balance",    "Saldo Paper"),
            ("db_jita",    "DB Jita (itens)"),
            ("db_amarr",   "DB Amarr (itens)"),
            ("config",     "Config (brel/acc)"),
            ("cache",      "Item Cache"),
            ("csv_jita",   "CSV Jita (ordens)"),
        ]
        for i, (key, label) in enumerate(metrics):
            bg = T["panel"] if i % 2 == 0 else T["grid"]
            row = tk.Frame(b, bg=bg)
            row.pack(fill="x", padx=2)
            tk.Label(row, text=f"  {label:<22}",
                     bg=bg, fg=T["dim"], font=MONO_S, width=26, anchor="w").pack(side="left")
            lbl = tk.Label(row, text="--",
                            bg=bg, fg=T["text"], font=MONO_B)
            lbl.pack(side="right", padx=6)
            self._lbls[key] = (lbl, bg)

        self.refresh()

    def refresh(self):
        threading.Thread(target=self._fetch, daemon=True).start()
        self.after(self._REFRESH, self.refresh)

    def _fetch(self):
        try:
            wl   = et.load_watchlist()
            pt   = et.load_paper_trades()
            cfg  = et.load_config()
            ic   = et.load_item_cache()
            n_jita  = et._db_item_count(HUB_RID["Jita"])
            n_amarr = et._db_item_count(HUB_RID["Amarr"])
            # CSV de ordens
            try:
                orders = et.load_orders_from_csv("The Forge (Jita)")
                csv_n  = len(orders)
            except Exception:
                csv_n  = 0

            data = {
                "watchlist": (str(len(wl)), T["cyan"]),
                "positions": (str(len(pt.get("positions", []))), T["yellow"]),
                "history":   (str(len(pt.get("history", []))),   T["text"]),
                "balance":   (_fmt_isk(pt.get("balance", 0)),     T["green"]),
                "db_jita":   (f"{n_jita:,}", T["green"] if n_jita > 1000 else T["red"]),
                "db_amarr":  (f"{n_amarr:,}", T["green"] if n_amarr > 1000 else T["red"]),
                "config":    (f"brel={cfg.get('g_brel',0)} / acc={cfg.get('g_acc',0)}", T["dim"]),
                "cache":     (f"{len(ic):,} itens", T["text"]),
                "csv_jita":  (f"{csv_n:,} ordens", T["green"] if csv_n > 100 else T["red"]),
            }
            self.after(0, lambda d=data: self._update(d))
        except Exception as e:
            self.after(0, lambda e=str(e): self.set_badge(f"erro: {e[:40]}", T["red"]))

    def _update(self, data):
        for key, (val, col) in data.items():
            if key in self._lbls:
                lbl, bg = self._lbls[key]
                lbl.config(text=val, fg=col)
        self.set_badge(f"atualizado {_now_str()}", T["dim"])


# ─────────────────────────────────────────────────────────────────────────────
#  Barra de status inferior
# ─────────────────────────────────────────────────────────────────────────────

class StatusBar(tk.Frame):
    def __init__(self, master, **kw):
        super().__init__(master, bg=T["header_bg"], height=22, **kw)
        self.pack_propagate(False)

        self._lbl = tk.Label(self, text="",
                              bg=T["header_bg"], fg=T["dim"],
                              font=MONO_S, anchor="w")
        self._lbl.pack(side="left", padx=8)

        tk.Label(self, text="EVE MARKET TERMINAL  v14  ·  dados do eve_trade.py",
                 bg=T["header_bg"], fg=T["dim"], font=MONO_S).pack(side="right", padx=8)

        self._tick()

    def _tick(self):
        self._lbl.config(text=f"  ● {_now_str()}  |  "
                               f"python {sys.version.split()[0]}  |  "
                               f"db: {et.HIST_DB[-40:]}")
        self.after(1000, self._tick)


# ─────────────────────────────────────────────────────────────────────────────
#  Janela principal do Dashboard
# ─────────────────────────────────────────────────────────────────────────────

class EVETerminal(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("EVE MARKET TERMINAL  ·  Bloomberg Style")
        self.configure(bg=T["bg"])
        self.geometry("1600x950")
        self.minsize(1280, 720)

        # Ícone ASCII no title bar (tkinter não suporta .ico no Linux)
        # self.iconbitmap(...)  — omitido por portabilidade

        self._build()

    def _build(self):
        # ── Ticker (topo) ──────────────────────────────────────────────────────
        self._ticker = TickerBar(self)
        self._ticker.pack(fill="x", side="top")

        # ── Separador laranja ─────────────────────────────────────────────────
        tk.Frame(self, bg=T["title"], height=2).pack(fill="x", side="top")

        # ── Status bar (fundo) ────────────────────────────────────────────────
        self._statusbar = StatusBar(self)
        self._statusbar.pack(fill="x", side="bottom")
        tk.Frame(self, bg=T["border"], height=1).pack(fill="x", side="bottom")

        # ── Área principal: 3 colunas ─────────────────────────────────────────
        main = tk.Frame(self, bg=T["bg"])
        main.pack(fill="both", expand=True, padx=4, pady=4)

        # ── Coluna esquerda ────────────────────────────────────────────────────
        col_l = tk.Frame(main, bg=T["bg"])
        col_l.pack(side="left", fill="both", expand=True, padx=(0, 3))

        OpportunitiesPanel(col_l).pack(fill="both", expand=True, pady=(0, 3))
        WatchlistPanel(col_l).pack(fill="both", expand=True)

        # ── Coluna central ─────────────────────────────────────────────────────
        col_m = tk.Frame(main, bg=T["bg"])
        col_m.pack(side="left", fill="both", expand=True, padx=3)

        HubComparePanel(col_m).pack(fill="x", pady=(0, 3))
        CycleSignalPanel(col_m).pack(fill="both", expand=True, pady=(0, 3))
        PaperTradingPanel(col_m).pack(fill="both", expand=True)

        # ── Coluna direita ─────────────────────────────────────────────────────
        col_r = tk.Frame(main, bg=T["bg"])
        col_r.pack(side="left", fill="y", padx=(3, 0))

        DBStatusPanel(col_r).pack(fill="x", pady=(0, 3))
        SystemMetricsPanel(col_r).pack(fill="x", pady=(0, 3))

        # Legenda de cores
        legend = PanelFrame(col_r, "LEGENDA")
        legend.pack(fill="x")
        items = [
            ("●", T["green"],  "Dado positivo / OK"),
            ("●", T["yellow"], "Alerta / abaixo do ideal"),
            ("●", T["red"],    "Negativo / ação necessária"),
            ("●", T["cyan"],   "Info / preço de referência"),
            ("▲", T["green"],  "Sinal de COMPRA (ciclo)"),
            ("▼", T["red"],    "Sinal de VENDA (ciclo)"),
            ("◇", T["yellow"], "Neutro / aguardar"),
        ]
        for sym, col, desc in items:
            fr = tk.Frame(legend.body, bg=T["panel"])
            fr.pack(fill="x", padx=4, pady=0)
            tk.Label(fr, text=f" {sym}",
                     bg=T["panel"], fg=col,
                     font=MONO_B, width=2).pack(side="left")
            tk.Label(fr, text=desc,
                     bg=T["panel"], fg=T["dim"],
                     font=MONO_S).pack(side="left", padx=4)

        # Atalhos
        info = PanelFrame(col_r, "ATALHOS")
        info.pack(fill="x", pady=(3, 0))
        for key, action in [("F5",    "Abrir EVE Trade Finder"),
                              ("R",     "Refresh todos os painéis"),
                              ("ESC",   "Fechar terminal")]:
            fr = tk.Frame(info.body, bg=T["panel"])
            fr.pack(fill="x", padx=4)
            tk.Label(fr, text=f" [{key}]",
                     bg=T["panel"], fg=T["title"],
                     font=MONO_B, width=6).pack(side="left")
            tk.Label(fr, text=action,
                     bg=T["panel"], fg=T["dim"],
                     font=MONO_S).pack(side="left", padx=4)

        # ── Atalhos de teclado ─────────────────────────────────────────────────
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<F5>",     lambda e: self._open_eve_trade())
        self.bind("<r>",      lambda e: self._refresh_all())
        self.bind("<R>",      lambda e: self._refresh_all())

    def _open_eve_trade(self):
        """Abre o EVE Trade Finder principal em processo separado."""
        import subprocess
        eve_path = os.path.join(_HERE, "eve_trade.py")
        if os.path.exists(eve_path):
            subprocess.Popen([sys.executable, eve_path])

    def _refresh_all(self):
        """Força refresh em todos os painéis com dado."""
        self._ticker.refresh_data()


# ─────────────────────────────────────────────────────────────────────────────
#  Entry point
# ─────────────────────────────────────────────────────────────────────────────

def main():
    app = EVETerminal()

    # Splash de boas-vindas
    splash = tk.Toplevel(app)
    splash.overrideredirect(True)
    splash.configure(bg=T["bg"])
    w, h = 480, 200
    sw, sh = app.winfo_screenwidth(), app.winfo_screenheight()
    splash.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    tk.Frame(splash, bg=T["title"], height=3).pack(fill="x")
    tk.Label(splash, text="\n  ⬡  EVE MARKET TERMINAL",
             bg=T["bg"], fg=T["title"],
             font=("Courier New", 16, "bold")).pack(fill="x")
    tk.Label(splash, text="  Bloomberg-style dashboard  ·  v14",
             bg=T["bg"], fg=T["dim"],
             font=("Courier New", 10)).pack(fill="x")
    tk.Label(splash,
             text="\n  Carregando dados do eve_trade.py...\n"
                  "  Preços serão buscados da ESI em background.",
             bg=T["bg"], fg=T["text"],
             font=("Courier New", 9), justify="left").pack(fill="x")
    pb = ttk.Progressbar(splash, mode="indeterminate", length=440)
    pb.pack(padx=20, pady=8)
    pb.start(8)
    tk.Frame(splash, bg=T["title"], height=2).pack(fill="x", side="bottom")

    def _close_splash():
        pb.stop()
        splash.destroy()

    splash.after(2500, _close_splash)
    app.mainloop()


if __name__ == "__main__":
    main()
