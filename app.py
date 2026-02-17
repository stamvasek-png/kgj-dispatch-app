import streamlit as st
import pandas as pd
import pulp
import tempfile

st.set_page_config(layout="wide")
st.title("KGJ Dispatch Optimalizace")

uploaded_file = st.file_uploader("Nahraj vstupní Excel (aki11.xlsx)", type="xlsx")

if uploaded_file:

    df = pd.read_excel(uploaded_file)
    df.columns = ["datetime", "ee_price", "gas_price", "heat_price", "heat_demand"]
    df = df.reset_index(drop=True)
    T = len(df)

    # =====================
    # PARAMETRY
    # =====================
    kgj_heat_output = 1.09
    kgj_el_output = 0.999
    kgj_heat_eff = 0.46
    kgj_service = 12.0
    kgj_gas_input = kgj_heat_output / kgj_heat_eff
    kgj_el_per_heat = kgj_el_output / kgj_heat_output
    kgj_gas_per_heat = kgj_gas_input / kgj_heat_output
    KGJ_MIN_LOAD = 0.55
    boiler_eff = 0.95
    boiler_max_heat = 3.91
    eboiler_eff = 0.98
    eboiler_max_heat = 0.60564
    EE_DISTRIBUTION_COST = 33.0
    HEAT_MIN_COVER = 0.99

    model = pulp.LpProblem("Dispatch", pulp.LpMaximize)

    q_kgj = pulp.LpVariable.dicts("q_KGJ", range(T), 0, kgj_heat_output)
    q_boiler = pulp.LpVariable.dicts("q_boiler", range(T), 0, boiler_max_heat)
    q_eboiler = pulp.LpVariable.dicts("q_eboiler", range(T), 0, eboiler_max_heat)

    ee_sold_spot = pulp.LpVariable.dicts("ee_sold_spot", range(T), 0)
    kgj_on = pulp.LpVariable.dicts("KGJ_on", range(T), 0, 1, cat="Binary")

    # =====================
    # CONSTRAINTS
    # =====================
    for t in range(T):
        demand = df.loc[t, "heat_demand"]
        h_required = HEAT_MIN_COVER * demand

        model += q_kgj[t] <= kgj_heat_output * kgj_on[t]
        model += q_kgj[t] >= KGJ_MIN_LOAD * kgj_heat_output * kgj_on[t]
        model += q_kgj[t] + q_boiler[t] + q_eboiler[t] >= h_required

        model += ee_sold_spot[t] == q_kgj[t] * kgj_el_per_heat

    # =====================
    # OBJECTIVE
    # =====================
    profit_terms = []

    for t in range(T):
        ee = df.loc[t, "ee_price"]
        gas = df.loc[t, "gas_price"]
        heat_p = df.loc[t, "heat_price"]
        delivered_heat = HEAT_MIN_COVER * df.loc[t, "heat_demand"]

        profit_terms.append(
            heat_p * delivered_heat
            + ee * ee_sold_spot[t]
            - gas * (q_kgj[t] * kgj_gas_per_heat + q_boiler[t] / boiler_eff)
            - kgj_service * kgj_on[t]
        )

    model += pulp.lpSum(profit_terms)

    model.solve(pulp.PULP_CBC_CMD(msg=False))

    # =====================
    # OUTPUT
    # =====================
    rows = []

    for t in range(T):
        rows.append({
            "datetime": df.loc[t, "datetime"],
            "KGJ_heat": q_kgj[t].value(),
            "Boiler_heat": q_boiler[t].value(),
            "EBoiler_heat": q_eboiler[t].value(),
            "Profit": profit_terms[t].value()
        })

    out_df = pd.DataFrame(rows)

    st.success("Optimalizace dokončena")

    st.dataframe(out_df)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    out_df.to_excel(tmp.name, index=False)

    with open(tmp.name, "rb") as f:
        st.download_button(
            label="Stáhnout Excel výstup",
            data=f,
            file_name="dispatch_output.xlsx"
        )
