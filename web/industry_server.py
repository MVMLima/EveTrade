#!/usr/bin/env python3
"""Servidor web do módulo industrial do EveTrade.

Interface mínima, só leitura, sobre `evetrade/core/industry.py` e as funções de
dados de `eve_trade.py`: calculadora de custo de manufatura com materiais,
preços de mercado e taxa de instalação.

  python web/industry_server.py --port 8009

Rotas:
  GET  /                       página (HTML + JS)
  GET  /api/meta               regiões e hubs disponíveis
  GET  /api/blueprint?tid=..   materiais do blueprint/produto + preços
  POST /api/calc               custo total a partir dos materiais retornados
"""
import argparse
import json
import os
import sys
import threading
import time
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)

import eve_trade as et  # noqa: E402

HUB_SISTEMAS = [
    ("Jita 4-4", 30000142),
    ("Amarr VIII", 30002187),
    ("Dodixie IX", 30002659),
    ("Rens VI", 30002510),
    ("Hek VIII", 30002053),
    ("Perimeter", 30000144),
]

# Construção e venda sempre em Jita 4-4: preço de região mistura citadel e
# transforma margem de 17% em 45%. A estação é única e explícita.
REGIAO_MERCADO = 10000002
ESTACAO_MERCADO = 60003760  # Jita 4-4
ALPHA_CLONE_TAX = 0.0025    # 0,25% do EIV, cobrado em conta Alpha

_lock = threading.Lock()
_cache: dict = {}
_item_cache: dict = {}

DIR_DADOS = os.path.expanduser("~/.local/share/evetrade")
ARQ_USO = os.path.join(DIR_DADOS, "industria_uso.json")
_lock_uso = threading.RLock()


def _uso_ler() -> dict:
    with _lock_uso:
        try:
            with open(ARQ_USO, encoding="utf-8") as f:
                dados = json.load(f)
        except (OSError, ValueError):
            dados = {}
        dados.setdefault("uso", {})
        dados.setdefault("nomes", {})
        dados.setdefault("favoritos", [])
        return dados


def _uso_gravar(dados: dict) -> None:
    os.makedirs(DIR_DADOS, exist_ok=True)
    tmp = ARQ_USO + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=1)
    os.replace(tmp, ARQ_USO)


def _registrar_uso(tid: int, nome: str) -> None:
    with _lock_uso:
        dados = _uso_ler()
        chave = str(int(tid))
        dados["uso"][chave] = int(dados["uso"].get(chave, 0)) + 1
        if nome:
            dados["nomes"][chave] = nome
        _uso_gravar(dados)


def _alternar_favorito(tid: int, nome: str | None) -> dict:
    with _lock_uso:
        dados = _uso_ler()
        chave = str(int(tid))
        if chave in dados["favoritos"]:
            dados["favoritos"].remove(chave)
        else:
            dados["favoritos"].append(chave)
        dados["nomes"].setdefault(chave, nome or chave)
        if nome:
            dados["nomes"][chave] = nome
        _uso_gravar(dados)
        return dados


def _listar_uso(limite: int = 10) -> dict:
    dados = _uso_ler()
    nomes = dados["nomes"]

    def item(chave: str) -> dict:
        return {"tid": int(chave), "nome": nomes.get(chave, chave),
                "usos": int(dados["uso"].get(chave, 0))}

    favoritos = [item(c) for c in dados["favoritos"] if c.isdigit()]
    vistos = {i["tid"] for i in favoritos}
    ordenado = sorted(
        (c for c in dados["uso"] if c.isdigit()),
        key=lambda c: (-int(dados["uso"][c]), c),
    )
    top = [item(c) for c in ordenado if int(c) not in vistos][:limite]
    return {"favoritos": favoritos, "top": top}


def _cached(chave: str, ttl: float, fn):
    """Cache simples com TTL por chave (processo único, trancado)."""
    agora = time.time()
    with _lock:
        item = _cache.get(chave)
        if item and agora - item[0] < ttl:
            return item[1]
    valor = fn()
    with _lock:
        _cache[chave] = (agora, valor)
    return valor


def _indices(ttl: float = 3600) -> dict:
    return _cached("indices", ttl, et.fetch_industry_indices) or {}


def _adj_prices(ttl: float = 3600) -> dict:
    return _cached("adj", ttl, et.fetch_adjusted_prices) or {}


def _itens() -> dict:
    global _item_cache
    if not _item_cache:
        _item_cache = et.load_item_cache() or {}
    return _item_cache


def _nome(tid: int) -> str:
    ic = _itens()
    et._ensure_names([tid], ic)
    return ic.get(tid, {}).get("name", str(tid))


def _blueprint(tid: int, region_id: int, contar_uso: bool = False) -> dict:
    bp = et.fetch_blueprint_materials(tid)
    if not bp.get("materials"):
        raise ValueError(
            f"Sem materiais para o type_id {tid}. Use o ID do blueprint ou do "
            f"produto (ex.: 691 = Rifter Blueprint, 587 = Rifter)."
        )
    adj = _adj_prices()
    precos = et.fetch_material_prices_bulk([m["type_id"] for m in bp["materials"]],
                                           region_id, station_id=ESTACAO_MERCADO)
    materiais = []
    for m in bp["materials"]:
        tid_mat = m["type_id"]
        preco = precos.get(tid_mat, {}).get("sell_min", 0.0)
        materiais.append({
            "type_id": tid_mat,
            "name": _nome(tid_mat),
            "quantity": m["quantity"],
            "unit_price": round(preco, 2),
            "adjusted_price": round(adj.get(tid_mat, 0.0), 2),
        })
    bp_tid = bp.get("bp_tid") or tid
    prod_tid = et._get_product_tid(bp_tid, bp, _itens()) or 0
    bp_nome = _nome(bp_tid)
    if contar_uso:
        _registrar_uso(tid, bp_nome)
    return {
        "bp_tid": bp_tid,
        "prod_tid": prod_tid,
        "bp_name": bp_nome,
        "product_name": _nome(prod_tid) if prod_tid else "",
        "qty_per_run": bp.get("qty_per_run", 1),
        "source": bp.get("source", ""),
        "materials": materiais,
    }


def _calc(payload: dict) -> dict:
    materiais = [
        {
            "type_id": int(m["type_id"]),
            "quantity": int(m["quantity"]),
            "unit_price": float(m.get("unit_price") or 0.0),
        }
        for m in payload.get("materials", [])
    ]
    if not materiais:
        raise ValueError("Nenhum material informado.")

    indices = _indices()
    sci = indices.get(int(payload.get("system_id") or 30000142), {}).get("manufacturing")
    if sci is None:
        raise ValueError("Sistema sem índice de manufatura na ESI.")
    if payload.get("system_cost_index") not in (None, ""):
        sci = float(payload["system_cost_index"]) / 100.0

    res = et.calculate_industry_cost(
        materials=materiais,
        product_qty=int(payload.get("product_qty") or 1),
        facility_tax=float(payload.get("facility_tax") or 0) / 100.0,
        system_cost_index=sci,
        adjusted_prices=_adj_prices(),
        me_level=int(payload.get("me_level") or 0),
        runs=max(1, int(payload.get("runs") or 1)),
        structure_mat_bonus=float(payload.get("structure_mat_bonus") or 0) / 100.0,
        structure_job_bonus=float(payload.get("structure_job_bonus") or 0) / 100.0,
        alpha_clone_tax=ALPHA_CLONE_TAX if payload.get("alpha_clone") else 0.0,
    )

    prod_tid = int(payload.get("prod_tid") or 0)
    region_id = int(payload.get("region_id") or REGIAO_MERCADO)
    if prod_tid:
        mercado = et.fetch_live_price(region_id, prod_tid,
                                      station_id=ESTACAO_MERCADO) or {}
        res["market"] = {
            "sell_min": mercado.get("sell_min", 0),
            "buy_max": mercado.get("buy_max", 0),
            "profit_unit": round(mercado.get("buy_max", 0) - res["cost_per_unit"], 2),
            "roi_pct": round(
                (mercado.get("buy_max", 0) - res["cost_per_unit"])
                / res["cost_per_unit"] * 100, 2
            ) if res["cost_per_unit"] else None,
        }
        res["market"]["profit_total"] = round(
            res["market"]["profit_unit"] * res.get("product_qty_total", 0), 2
        )
    for linha in res["materials"]:
        linha["name"] = _nome(linha["type_id"])
    return res


class Handler(BaseHTTPRequestHandler):
    server_version = "EveTradeIndustry/1.0"

    def log_message(self, format, *args):  # log enxuto
        sys.stderr.write(f"{self.address_string()} - {format % args}\n")

    def _send(self, code: int, body: bytes, ctype: str):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code: int, dados: dict):
        self._send(code, json.dumps(dados, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8")

    def do_GET(self):
        url = urlparse(self.path)
        try:
            if url.path in ("/", "/index.html"):
                with open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       "industry.html"), "rb") as f:
                    self._send(200, f.read(), "text/html; charset=utf-8")
            elif url.path == "/api/meta":
                self._json(200, {
                    "regions": [{"id": REGIAO_MERCADO, "name": "Jita 4-4 · The Forge"}],
                    "systems": [{"name": n, "id": i} for n, i in HUB_SISTEMAS],
                    "station": {"id": ESTACAO_MERCADO, "name": "Jita 4-4"},
                })
            elif url.path == "/api/blueprint":
                q = parse_qs(url.query)
                tid = int(q.get("tid", ["0"])[0])
                region_id = int(q.get("region", ["10000002"])[0])
                contar = q.get("uso", ["0"])[0] == "1"
                if tid <= 0:
                    raise ValueError("Informe o type_id do blueprint ou do produto.")
                self._json(200, _blueprint(tid, region_id, contar))
            elif url.path == "/api/uso":
                self._json(200, _listar_uso())
            else:
                self._json(404, {"erro": "rota não encontrada"})
        except ValueError as exc:
            self._json(400, {"erro": str(exc)})
        except Exception as exc:
            traceback.print_exc()
            self._json(500, {"erro": f"falha interna: {exc}"})

    def do_POST(self):
        url = urlparse(self.path)
        if url.path not in ("/api/calc", "/api/favorito"):
            self._json(404, {"erro": "rota não encontrada"})
            return
        try:
            tamanho = int(self.headers.get("Content-Length") or 0)
            payload = json.loads(self.rfile.read(tamanho) or b"{}")
            if url.path == "/api/favorito":
                tid = int(payload.get("tid") or 0)
                if tid <= 0:
                    raise ValueError("Informe o type_id para favoritar.")
                nome = str(payload.get("nome") or "").strip() or None
                dados = _alternar_favorito(tid, nome)
                self._json(200, {
                    "favorito": str(tid) in dados["favoritos"],
                    **_listar_uso(),
                })
            else:
                self._json(200, _calc(payload))
        except ValueError as exc:
            self._json(400, {"erro": str(exc)})
        except Exception as exc:
            traceback.print_exc()
            self._json(500, {"erro": f"falha interna: {exc}"})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8009)
    ap.add_argument("--host", default="127.0.0.1")
    args = ap.parse_args()
    srv = ThreadingHTTPServer((args.host, args.port), Handler)
    print(f"servindo em http://{args.host}:{args.port}", flush=True)
    srv.serve_forever()


if __name__ == "__main__":
    main()
