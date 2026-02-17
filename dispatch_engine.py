"""
KGJ Dispatch Engine
Core calculation and optimization logic.
"""

import pandas as pd
import numpy as np
import pulp
from dataclasses import dataclass
from typing import Optional


# ──────────────────────────────────────────────
# TECHNOLOGY PARAMETERS (dataclass)
# ──────────────────────────────────────────────

@dataclass
class TechParams:
    kgj_heat_output: float = 1.09
    kgj_el_output: float = 0.999
    kgj_heat_eff: float = 0.46
    kgj_service: float = 12.0
    kgj_min_load: float = 0.55

    boiler_eff: float = 0.95
    boiler_max_heat: float = 3.91

    eboiler_eff: float = 0.98
    eboiler_max_heat: float = 0.60564

    ee_dist_cost: float = 33.0
    heat_min_cover: float = 0.99
    min_up: int = 4
    min_down: int = 4
    initial_state: int = 0

    @property
    def kgj_gas_input(self):
        return self.kgj_heat_output / self.kgj_heat_eff

    @property
    def kgj_el_per_heat(self):
        return self.kgj_el_output / self.kgj_heat_output

    @property
    def kgj_gas_per_heat(self):
        return self.kgj_gas_input / self.kgj_heat_output


# ──────────────────────────────────────────────
# MARGINAL COST / MARGIN CALCULATIONS
# ──────────────────────────────────────────────

def compute_margins(ee: float, gas: float, heat_p: float, p: TechParams) -> dict:
    """Compute marginal costs and margins for all 4 heat sources."""

    cost1 = gas / p.boiler_eff

    cost2 = (gas * p.kgj_gas_per_heat
             + (p.kgj_service / p.kgj_heat_output)
             - ee * p.kgj_el_per_heat)

    cost3 = (ee + p.ee_dist_cost) / p.eboiler_eff

    total_heat_agg = p.kgj_heat_output + p.eboiler_max_heat
    ee_surplus = p.kgj_el_output - (p.eboiler_max_heat / p.eboiler_eff)
    cost4 = (gas * p.kgj_gas_input + p.kgj_service - ee * ee_surplus) / total_heat_agg

    trigger_ee_only = (gas * p.kgj_gas_input + p.kgj_service) / p.kgj_el_output
    trigger_full    = (gas * p.kgj_gas_input + p.kgj_service - heat_p * p.kgj_heat_output) / p.kgj_el_output

    return {
        "cost1": cost1,
        "cost2": cost2,
        "cost3": cost3,
        "cost4": cost4,
        "m1": heat_p - cost1,
        "m2": heat_p - cost2,
        "m3": heat_p - cost3,
        "m4": heat_p - cost4,
        "trigger_ee_only": trigger_ee_only,
        "trigger_full": trigger_full,
        "kgj_margin_ee": ee - trigger_ee_only,
    }


def best_source(m1, m2, m3, m4) -> dict:
    """Return the most profitable heat source."""
    options = [
        {"name": "Plynový kotel",      "short": "Kotel",    "id": 1, "m": m1, "color": "#5B8DEE"},
        {"name": "KGJ + spot prodej",  "short": "KGJ+Spot", "id": 2, "m": m2, "color": "#F0A500"},
        {"name": "Elektrokotel (síť)", "short": "EKotel",   "id": 3, "m": m3, "color": "#9B59B6"},
        {"name": "KGJ + elektrokotel", "short": "KGJ+EK",  "id": 4, "m": m4, "color": "#00D4AA"},
    ]
    positive = [o for o in options if o["m"] > 0]
    if not positive:
        return {"name": "Žádný zdroj", "short": "—", "id": 0, "m": 0.0, "color": "#E05555"}
    return sorted(positive, key=lambda x: x["m"], reverse=True)[0]


# ──────────────────────────────────────────────
# FULL LP DISPATCH OPTIMIZATION
# ──────────────────────────────────────────────

def run_dispatch(df: pd.DataFrame, p: TechParams) -> Optional[pd.DataFrame]:
    """
    Solve the full MIP dispatch problem using PuLP/CBC.
    Returns a results DataFrame or None on failure.
    """
    T = len(df)
    BYPASS_TOL = 0.001

    model = pulp.LpProblem("KGJ_Integrated_Dispatch", pulp.LpMaximize)

    # Variables
    q_kgj     = pulp.LpVariable.dicts("q_KGJ",    range(T), 0, p.kgj_heat_output)
    q_boiler  = pulp.LpVariable.dicts("q_boiler",  range(T), 0, p.boiler_max_heat)
    q_eboiler = pulp.LpVariable.dicts("q_eboiler", range(T), 0, p.eboiler_max_heat)

    ee_from_kgj          = pulp.LpVariable.dicts("ee_from_kgj",          range(T), 0)
    ee_sold_spot         = pulp.LpVariable.dicts("ee_sold_spot",         range(T), 0)
    ee_to_eboiler_int    = pulp.LpVariable.dicts("ee_to_eboiler_int",    range(T), 0)
    ee_to_eboiler_grid   = pulp.LpVariable.dicts("ee_to_eboiler_grid",   range(T), 0)

    kgj_on    = pulp.LpVariable.dicts("KGJ_on",    range(T), 0, 1, cat="Binary")
    kgj_start = pulp.LpVariable.dicts("KGJ_start", range(T), 0, 1, cat="Binary")
    kgj_stop  = pulp.LpVariable.dicts("KGJ_stop",  range(T), 0, 1, cat="Binary")
    heat_def  = pulp.LpVariable.dicts("heat_def",  range(T), 0)

    # Constraints
    for t in range(T):
        demand     = df.loc[t, "heat_demand"]
        h_required = p.heat_min_cover * demand

        model += q_kgj[t] <= p.kgj_heat_output * kgj_on[t]
        model += q_kgj[t] >= p.kgj_min_load * p.kgj_heat_output * kgj_on[t]

        if demand > 0:
            model += q_kgj[t] + q_boiler[t] + q_eboiler[t] >= h_required
        else:
            model += q_boiler[t] == 0
            model += q_eboiler[t] == 0

        model += heat_def[t] >= h_required - q_kgj[t]
        model += q_boiler[t] + q_eboiler[t] <= heat_def[t]

        model += ee_from_kgj[t]       == q_kgj[t] * p.kgj_el_per_heat
        model += ee_sold_spot[t] + ee_to_eboiler_int[t] == ee_from_kgj[t]
        model += q_eboiler[t]         == p.eboiler_eff * (ee_to_eboiler_int[t] + ee_to_eboiler_grid[t])

        if t > 0:
            model += kgj_on[t] - kgj_on[t-1] == kgj_start[t] - kgj_stop[t]
        else:
            model += kgj_on[t] - p.initial_state == kgj_start[t] - kgj_stop[t]

    for t in range(T - p.min_up):
        model += pulp.lpSum(kgj_on[t+i] for i in range(p.min_up)) >= p.min_up * kgj_start[t]

    for t in range(T - p.min_down):
        model += pulp.lpSum(1 - kgj_on[t+i] for i in range(p.min_down)) >= p.min_down * kgj_stop[t]

    # Objective
    profit_terms = []
    for t in range(T):
        ee_p      = df.loc[t, "ee_price"]
        gas_p     = df.loc[t, "gas_price"]
        heat_p    = df.loc[t, "heat_price"]
        delivered = p.heat_min_cover * df.loc[t, "heat_demand"]
        profit_terms.append(
            heat_p * delivered
            + ee_p * ee_sold_spot[t]
            - gas_p * (q_kgj[t] * p.kgj_gas_per_heat + q_boiler[t] / p.boiler_eff)
            - (ee_p + p.ee_dist_cost) * ee_to_eboiler_grid[t]
            - p.kgj_service * kgj_on[t]
        )
    model += pulp.lpSum(profit_terms)

    # Solve
    solver = pulp.PULP_CBC_CMD(msg=False, timeLimit=120)
    status = model.solve(solver)

    if pulp.LpStatus[status] not in ("Optimal", "Feasible"):
        return None

    # Build output DataFrame
    rows = []
    for t in range(T):
        demand    = df.loc[t, "heat_demand"]
        ee_p      = df.loc[t, "ee_price"]
        gas_p     = df.loc[t, "gas_price"]
        heat_p    = df.loc[t, "heat_price"]

        kgj_total = q_kgj[t].varValue or 0.0
        bypass    = max(kgj_total - p.heat_min_cover * demand, 0)
        if bypass < BYPASS_TOL:
            bypass = 0.0

        r = compute_margins(ee_p, gas_p, heat_p, p)

        rows.append({
            "datetime":                          df.loc[t, "datetime"],
            "EE_price_EUR_MWh":                  ee_p,
            "Gas_price_EUR_MWh":                 gas_p,
            "Heat_price_EUR_MWh":                heat_p,
            "Heat_demand_MWh":                   demand,
            "Bypass_heat_MWh":                   bypass,
            "KGJ_heat_MWh":                      kgj_total - bypass,
            "KGJ_load_pct":                      100 * kgj_total / p.kgj_heat_output,
            "KGJ_on":                            int(round(kgj_on[t].varValue or 0)),
            "KGJ_start":                         int(round(kgj_start[t].varValue or 0)),
            "KGJ_stop":                          int(round(kgj_stop[t].varValue or 0)),
            "Gas_boiler_heat_MWh":               q_boiler[t].varValue or 0.0,
            "Gas_boiler_load_pct":               100 * (q_boiler[t].varValue or 0) / p.boiler_max_heat,
            "Electric_boiler_heat_MWh":          q_eboiler[t].varValue or 0.0,
            "Electric_boiler_load_pct":          100 * (q_eboiler[t].varValue or 0) / p.eboiler_max_heat,
            "KGJ_Electricity_MWh":               ee_from_kgj[t].varValue or 0.0,
            "EE_Sold_Spot_MWh":                  ee_sold_spot[t].varValue or 0.0,
            "EE_to_EBoiler_Internal_MWh":        ee_to_eboiler_int[t].varValue or 0.0,
            "EE_to_EBoiler_Grid_MWh":            ee_to_eboiler_grid[t].varValue or 0.0,
            "Total_profit_EUR":                  profit_terms[t].value() or 0.0,
            "KGJ_Power_Trigger_EE_only":         r["trigger_ee_only"],
            "Cost_1_Boiler_EUR_per_MWh":         r["cost1"],
            "Cost_2_KGJ_Spot_EUR_per_MWh":       r["cost2"],
            "Cost_3_EBoiler_Grid_EUR_per_MWh":   r["cost3"],
            "Cost_4_KGJ_EBoiler_EUR_per_MWh":    r["cost4"],
            "KGJ_margin_EE_only_EUR_per_MWh":    r["kgj_margin_ee"],
            "Margin_1_Boiler_EUR_per_MWh":       r["m1"],
            "Margin_2_KGJ_Spot_EUR_per_MWh":     r["m2"],
            "Margin_3_EBoiler_Grid_EUR_per_MWh": r["m3"],
            "Margin_4_KGJ_EBoiler_EUR_per_MWh":  r["m4"],
        })

    return pd.DataFrame(rows)
