"""Custo de manufatura pela fórmula da CCP.

Módulo puro: sem Tkinter, sem rede, sem estado global.
"""

import math


def calculate_industry_cost(
    materials: list,  # [{type_id, quantity, unit_price}]
    product_qty: int,  # unidades produzidas por run
    facility_tax: float,  # taxa do dono da estrutura (0.0–1.0, ex: 0.05 = 5%)
    system_cost_index: float,  # da fetch_industry_indices (ex: 0.0343 = 3.43%)
    adjusted_prices: dict,  # da fetch_adjusted_prices {type_id: adj_price}
    me_level: int = 0,  # Material Efficiency do blueprint (0–10)
    runs: int = 1,
    structure_mat_bonus: float = 0.0,  # bônus da estrutura nos MATERIAIS (0–1, ex: 0.01 = 1%)
    structure_job_bonus: float = 0.0,  # bônus da estrutura no JOB COST  (0–1, ex: 0.03 = 3%)
) -> dict:
    """
    Calcula custo total de fabricação para 'runs' execuções.

    Fórmula oficial da CCP (verificada nas imagens do jogo):

    1. Materiais:
       me_combined = ME_blueprint% + structure_mat_bonus%
       qty_final   = ceil(qty_base × (1 - me_combined) × runs)  [mínimo: runs]

    2. EIV (Estimated Item Value):
       EIV = Σ(adjusted_price × qty_final)

    3. Job Gross Cost:
       job_gross = EIV × system_cost_index × (1 - structure_job_bonus)

    4. SCC Surcharge (taxa fixa da CCP, sempre 4% do EIV):
       scc = EIV × 0.04

    5. Facility Tax (taxa do dono da estrutura, % do job_gross):
       facility_cost = job_gross × facility_tax

    6. Total:
       total_job  = job_gross + scc + facility_cost
       total_cost = mat_cost + total_job
       cost/unit  = total_cost / (product_qty × runs)
    """
    SCC_RATE = 0.04  # 4% do EIV — taxa fixa CCP (SCC Surcharge), não configurável

    # ME combinado: blueprint ME + bônus da estrutura nos materiais
    me_combined = me_level / 100.0 + structure_mat_bonus
    me_factor = max(0.0, 1.0 - me_combined)

    mat_lines = []
    total_mat_cost = 0.0
    eiv = 0.0  # base para job cost e SCC

    for mat in materials:
        tid = mat["type_id"]
        qty = mat["quantity"]  # quantidade base (sem ME) por run
        uprice = mat.get("unit_price", 0.0)

        # Quantidade final: aplica ME+estrutura, mínimo de 1 por run
        qty_final = max(runs, math.ceil(qty * me_factor * runs))

        adj = adjusted_prices.get(tid, uprice)
        eiv += adj * qty_final

        subtotal = qty_final * uprice
        total_mat_cost += subtotal
        mat_lines.append(
            {
                "type_id": tid,
                "quantity": qty_final,
                "unit_price": round(uprice, 2),
                "subtotal": round(subtotal, 2),
                "adj_price": round(adj, 2),
            }
        )

    # Job Gross Cost com bônus de estrutura
    job_gross_before = eiv * system_cost_index
    job_struct_saving = job_gross_before * structure_job_bonus
    job_gross = job_gross_before - job_struct_saving

    # SCC Surcharge: taxa fixa da CCP (4% do EIV, sempre cobrado)
    scc_surcharge = eiv * SCC_RATE

    # Facility Tax: taxa do dono da estrutura (% do job_gross)
    facility_cost = job_gross * facility_tax

    total_job = job_gross + scc_surcharge + facility_cost
    total_cost = total_mat_cost + total_job
    cost_per_unit = total_cost / max(product_qty * runs, 1)

    return {
        "materials": mat_lines,
        "total_mat_cost": round(total_mat_cost, 2),
        "eiv": round(eiv, 2),
        "system_cost_index": system_cost_index,
        "structure_job_bonus": round(structure_job_bonus * 100, 1),
        "structure_mat_bonus": round(structure_mat_bonus * 100, 1),
        "job_gross_before": round(job_gross_before, 2),
        "job_struct_saving": round(job_struct_saving, 2),
        "job_cost": round(job_gross, 2),
        "scc_surcharge": round(scc_surcharge, 2),
        "facility_tax_pct": round(facility_tax * 100, 2),
        "facility_cost": round(facility_cost, 2),
        "total_job_cost": round(total_job, 2),
        "total_cost": round(total_cost, 2),
        "product_qty_total": product_qty * runs,
        "cost_per_unit": round(cost_per_unit, 2),
        "me_level": me_level,
        "runs": runs,
    }
