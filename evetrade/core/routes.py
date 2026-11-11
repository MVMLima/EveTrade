"""Rotas entre regiões e arbitragem com casamento ordem a ordem.

Módulo puro: sem Tkinter, sem rede, sem estado global. Os nomes dos itens
devem estar em item_cache antes da chamada (eve_trade resolve via ESI/SDE).
"""

from .orders import build_order_index, order_min_volume

# location_id acima disto é estrutura de jogador (citadel)
MAX_INT32 = 2_147_483_647


def find_routes(
    src_orders,
    dst_orders,
    src_name,
    dst_name,
    min_roi=10,
    max_budget=1e18,
    min_budget=0,
    max_volume=1e18,
    min_lucro_total=0,
    station_cache=None,
    item_cache=None,
    sales_tax_pct=7.5,
):
    """Rotas pelo topo do livro + profundidade executável (ver execution_depth).

    sales_tax_pct: imposto de venda usado na profundidade (7,5% = Accounting 0,
    o pior caso). Lucro/un, ROI e Lucro Total continuam brutos, como antes.
    """
    src_idx = build_order_index(src_orders)
    dst_idx = build_order_index(dst_orders)
    sc, ic = station_cache or {}, item_cache or {}
    routes = []

    for tid in set(src_idx) & set(dst_idx):
        sell_list = src_idx[tid]["sell"]
        buy_list = dst_idx[tid]["buy"]
        if not sell_list or not buy_list:
            continue
        # 2F — listas ja pre-ordenadas por build_order_index; [0] e suficiente
        sell_at_src = sell_list[0]
        buy_at_dst = buy_list[0]

        buy_price = float(sell_at_src["price"])
        sell_price = float(buy_at_dst["price"])
        if sell_price <= buy_price:
            continue

        profit_unit = sell_price - buy_price
        roi = (profit_unit / buy_price) * 100
        if roi < min_roi:
            continue

        volume_avail = int(sell_at_src.get("volume_remain", 0))
        dst_vol_demand = int(buy_at_dst.get("volume_remain", 0))
        vol_realizavel = min(volume_avail, dst_vol_demand)
        # Ordem de compra com min_volume maior que o lote disponivel nao aceita a venda
        if vol_realizavel < order_min_volume(buy_at_dst):
            continue
        lucro_total = profit_unit * vol_realizavel
        capital = buy_price * volume_avail

        if lucro_total < min_lucro_total:
            continue
        if min_budget > 0 and capital < min_budget:
            continue
        if max_budget < 1e18 and capital > max_budget:
            continue

        item_info = ic.get(tid, {})
        item_name = item_info.get("name", str(tid))
        item_volume = float(item_info.get("volume", 0))
        total_volume = item_volume * vol_realizavel

        loc_id = int(sell_at_src.get("location_id", 0))
        dst_loc_id = int(buy_at_dst.get("location_id", 0))
        # Vol max (m3) limita a carga executável, não descarta a rota:
        # um item volumoso continua valendo se couber no porão.
        # 0 (campo vazio) significa "sem limite", como em _fread_limit.
        limite_m3 = max_volume if 0 < max_volume < 1e18 else None
        depth = execution_depth(sell_list, buy_list, item_volume, sales_tax_pct, limite_m3)
        if limite_m3 is not None and depth["unid_executavel"] <= 0:
            continue

        routes.append(
            {
                "type_id": tid,
                "name": item_name,
                "station": sc.get(loc_id, f"Loc {loc_id}"),
                "dst_station": sc.get(dst_loc_id, f"Loc {dst_loc_id}"),
                "dst_vol_demand": dst_vol_demand,
                "vol_realizavel": vol_realizavel,
                "lucro_total": lucro_total,
                "buy_price": buy_price,
                "sell_price": sell_price,
                "profit_unit": profit_unit,
                "roi": roi,
                "volume_avail": volume_avail,
                "capital": capital,
                "item_volume": item_volume,
                "total_volume": total_volume,
                "competition": len(dst_idx[tid]["sell"]),
                "src": src_name,
                "dst": dst_name,
                "lucro_dia": None,
                "giro_dia": None,
                **depth,
                "reposicoes_dia": None,
            }
        )

    routes.sort(key=lambda r: r["roi"], reverse=True)
    return routes


def execution_depth(sells: list, buys: list, item_volume: float = 0.0,
                    sales_tax_pct: float = 7.5, max_volume_m3: float | None = None) -> dict:
    """Quanto da rota dá para executar AGORA, caminhando o livro até o break-even.

    Compra das vendas mais baratas na origem e vende para as compras mais altas
    no destino (respeitando min_volume) enquanto o lucro por unidade, já
    descontado o imposto de venda, for positivo. Comprar de ordem de venda e
    vender para ordem de compra não paga broker fee; frete não entra.
    As ordens podem estar em estações diferentes da região.

    max_volume_m3: capacidade do porão. Quando informada (e item_volume > 0), as
    unidades são limitadas ao que cabe — mantendo primeiro as de maior lucro por
    unidade, já que todas ocupam o mesmo volume.

    Retorna unid_executavel, lucro_garantido, capital_executavel,
    preco_venda_medio e m3_executavel.
    """
    trades = match_orders_granular(sells, buys, item_volume, 0.0, 0.0, sales_tax_pct, 0.0)

    if max_volume_m3 and item_volume > 0:
        limite = int(max_volume_m3 // item_volume)
        if sum(t["qty"] for t in trades) > limite:
            trades = sorted(trades, key=lambda t: t["lucro_liq_unit"], reverse=True)
            cortados, resto = [], limite
            for t in trades:
                if resto <= 0:
                    break
                q = min(t["qty"], resto)
                cortados.append({**t, "qty": q,
                                 "lucro_liq_total": t["lucro_liq_unit"] * q,
                                 "custo_total": t["custo_unit"] * q})
                resto -= q
            trades = cortados

    qty = sum(t["qty"] for t in trades)
    return {
        "unid_executavel": qty,
        "lucro_garantido": sum(t["lucro_liq_total"] for t in trades),
        "capital_executavel": sum(t["custo_total"] for t in trades),
        "m3_executavel": qty * item_volume,
        "preco_venda_medio": (sum(t["price_dest"] * t["qty"] for t in trades) / qty
                              if qty else 0.0),
    }


def match_orders_granular(
    sells_raw, buys_raw, volume_m3, tax_broker_buy, tax_broker_sell, tax_sales, custo_frete_m3
):
    # Listas chegam pre-ordenadas pelo build_order_index.
    # 3F — alias removido; volume_remain ja e int (normalizado no CSV loader).
    tax_buy_mult = 1.0 + (tax_broker_buy / 100.0)
    tax_sell_mult = 1.0 - ((tax_broker_sell + tax_sales) / 100.0)
    frete_unit = volume_m3 * custo_frete_m3

    # Compras cujo min_volume excede todo o volume a venda nunca aceitariam o lote
    total_sell = sum(int(o["volume_remain"]) for o in sells_raw)
    buys_raw = [b for b in buys_raw if order_min_volume(b) <= total_sell]

    s_vols = [o["volume_remain"] for o in sells_raw]
    b_vols = [o["volume_remain"] for o in buys_raw]

    trades = []
    s_idx = b_idx = 0

    while s_idx < len(sells_raw) and b_idx < len(buys_raw):
        p_orig = float(sells_raw[s_idx]["price"])
        p_dest = float(buys_raw[b_idx]["price"])

        custo_real = p_orig * tax_buy_mult + frete_unit
        receita_real = p_dest * tax_sell_mult
        lucro_unit = receita_real - custo_real

        if lucro_unit <= 0:
            break

        qty = min(s_vols[s_idx], b_vols[b_idx])
        if qty <= 0:
            if s_vols[s_idx] == 0:
                s_idx += 1
            if b_vols[b_idx] == 0:
                b_idx += 1
            continue

        trades.append(
            {
                "qty": qty,
                "price_orig": p_orig,
                "price_dest": p_dest,
                "loc_orig_id": int(sells_raw[s_idx].get("location_id", 0)),
                "loc_dest_id": int(buys_raw[b_idx].get("location_id", 0)),
                "lucro_liq_unit": lucro_unit,
                "lucro_liq_total": lucro_unit * qty,
                "custo_unit": custo_real,
                "custo_total": custo_real * qty,
            }
        )

        s_vols[s_idx] -= qty
        b_vols[b_idx] -= qty
        if s_vols[s_idx] == 0:
            s_idx += 1
        if b_vols[b_idx] == 0:
            b_idx += 1

    return trades


def arbitrage_opportunities(
    src_orders,
    dst_orders,
    src_name,
    dst_name,
    tax_broker_buy,
    tax_broker_sell,
    tax_sales,
    custo_frete_m3,
    station_cache,
    item_cache,
    min_margem=0.0,
    min_lucro_total=0.0,
    progress_cb=None,
):
    src_idx = build_order_index(src_orders)
    dst_idx = build_order_index(dst_orders)
    sc, ic = station_cache or {}, item_cache or {}

    candidates = list(set(src_idx) & set(dst_idx))
    if progress_cb:
        progress_cb(f"{len(candidates)} itens candidatos...")

    opportunities = []

    total_c = len(candidates)
    for i, tid in enumerate(candidates):
        if progress_cb and i % 200 == 0:
            progress_cb((i, total_c, f"Calculando... {i}/{total_c}"))

        sells = src_idx[tid]["sell"]
        buys = dst_idx[tid]["buy"]
        if not sells or not buys:
            continue

        info = ic.get(tid, {})
        item_name = info.get("name", str(tid))
        vol_m3 = float(info.get("volume", 0)) or 0.01

        trades = match_orders_granular(
            sells, buys, vol_m3, tax_broker_buy, tax_broker_sell, tax_sales, custo_frete_m3
        )

        if not trades:
            continue

        grouped = {}
        for t in trades:
            key = (t["loc_orig_id"], t["loc_dest_id"])
            if key not in grouped:
                grouped[key] = {
                    "qty": 0,
                    "lucro": 0.0,
                    "custo": 0.0,
                    "sum_p_o": 0.0,
                    "sum_p_d": 0.0,
                }
            g = grouped[key]
            g["qty"] += t["qty"]
            g["lucro"] += t["lucro_liq_total"]
            g["custo"] += t["custo_total"]
            g["sum_p_o"] += t["price_orig"] * t["qty"]
            g["sum_p_d"] += t["price_dest"] * t["qty"]

        for (lid_o, lid_d), dat in grouped.items():
            if dat["qty"] <= 0:
                continue
            avg_p_o = dat["sum_p_o"] / dat["qty"]
            avg_p_d = dat["sum_p_d"] / dat["qty"]
            lucro_unit = dat["lucro"] / dat["qty"]
            custo_unit = dat["custo"] / dat["qty"]
            margem = (lucro_unit / custo_unit * 100) if custo_unit > 0 else 0
            spread = ((avg_p_d - avg_p_o) / avg_p_o * 100) if avg_p_o > 0 else 0

            # Filtros aplicados antes de criar o objeto
            if margem < min_margem:
                continue
            if dat["lucro"] < min_lucro_total:
                continue

            name_o = sc.get(
                lid_o, f"Estrutura Privada {lid_o}" if lid_o > MAX_INT32 else f"Loc {lid_o}"
            )
            name_d = sc.get(
                lid_d, f"Estrutura Privada {lid_d}" if lid_d > MAX_INT32 else f"Loc {lid_d}"
            )

            opportunities.append(
                {
                    "type_id": tid,
                    "item": item_name,
                    "volume_m3": vol_m3,
                    "preco_origem": round(avg_p_o, 2),
                    "loc_origem": name_o,
                    "qtd_origem": dat["qty"],
                    "preco_destino": round(avg_p_d, 2),
                    "loc_destino": name_d,
                    "qtd_destino": dat["qty"],
                    "spread_pct": round(spread, 2),
                    "lucro_bruto": round(avg_p_d - avg_p_o, 2),
                    "lucro_liq_unit": round(lucro_unit, 2),
                    "margem_pct": round(margem, 2),
                    "roi_total": round(dat["lucro"], 2),
                    "custo_total": round(dat["custo"], 2),
                    "vol_total_m3": round(vol_m3 * dat["qty"], 2),
                    # Concorrencia no destino: vendedores ja posicionados abaixo do
                    # preco medio que planejamos vender (avg_p_d)
                    "concorr_dst": sum(
                        1 for o in dst_idx[tid]["sell"] if float(o["price"]) < avg_p_d
                    ),
                    "src": src_name,
                    "dst": dst_name,
                    "lucro_dia": None,
                    "giro_dia": None,
                }
            )

    opportunities.sort(key=lambda x: x["roi_total"], reverse=True)
    if progress_cb:
        progress_cb(f"{len(opportunities)} oportunidades encontradas.")
    return opportunities
