"""Estatísticas do histórico diário de mercado (ESI /history).

Módulo puro: sem Tkinter, sem rede, sem estado global.
"""

from datetime import datetime, timedelta, timezone


def history_stats(hist, days_week=7, days_month=30):
    """
    Retorna dict com avg_vol_90, avg_p_week, avg_p_month, avg_v_week, avg_v_month.

    Melhorias:
      - now_str removido (variavel morta da v6)
      - cutoffs calculados UMA vez (antes: 4x datetime.now(timezone.utc) por chamada)
      - lista percorrida UMA vez O(n) em vez de 3x O(n) (h90, h30, h7)
    """
    if not hist:
        return None
    now = datetime.now(timezone.utc)
    cutoff90 = (now - timedelta(days=90)).strftime("%Y-%m-%d")
    cutoff30 = (now - timedelta(days=days_month)).strftime("%Y-%m-%d")
    cutoff7 = (now - timedelta(days=days_week)).strftime("%Y-%m-%d")

    # Uma unica passagem classifica cada entrada nos tres buckets
    sum_vol90 = cnt90 = 0
    sum_p7 = sum_v7 = cnt7 = 0
    sum_p30 = sum_v30 = cnt30 = 0

    for d in hist:
        date = d.get("date", "")
        if date < cutoff90:
            continue
        vol = d.get("volume", 0)
        avg = d.get("average", 0)
        sum_vol90 += vol
        cnt90 += 1
        if date >= cutoff30:
            sum_p30 += avg
            sum_v30 += vol
            cnt30 += 1
        if date >= cutoff7:
            sum_p7 += avg
            sum_v7 += vol
            cnt7 += 1

    if cnt90 == 0:
        return None

    return {
        "avg_vol_90": sum_vol90 / cnt90,
        "avg_p_week": sum_p7 / cnt7 if cnt7 else 0,
        "avg_p_month": sum_p30 / cnt30 if cnt30 else 0,
        "avg_v_week": sum_v7 / cnt7 if cnt7 else 0,
        "avg_v_month": sum_v30 / cnt30 if cnt30 else 0,
    }


def calc_lucro_dia(history_dst, lucro_liq_unit, dias=90):
    """
    Estima ISK/dia usando o giro medio diario no destino dos ultimos `dias` dias.
    Retorna (lucro_dia, giro_dia_medio) ou (None, None) se dados insuficientes.
    """
    if not history_dst:
        return None, None
    cutoff = (datetime.now(timezone.utc) - timedelta(days=dias)).strftime("%Y-%m-%d")
    total_vol = cnt = 0
    for d in history_dst:
        if d.get("date", "") >= cutoff:
            total_vol += d.get("volume", 0)
            cnt += 1
    if cnt == 0:
        return None, None
    giro_dia = total_vol / cnt
    return round(lucro_liq_unit * giro_dia, 2), round(giro_dia, 1)
