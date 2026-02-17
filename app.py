"""
KGJ Dispatch Kalkulátor
Streamlit frontend — full dispatch optimization & margin analytics
"""

import io
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

from dispatch_engine import TechParams, compute_margins, best_source, run_dispatch
import chart_helpers as ch

# ──────────────────────────────────────────────
# PAGE CONFIG
# ──────────────────────────────────────────────

st.set_page_config(
    page_title="KGJ Dispatch",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────
# CUSTOM CSS
# ──────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@300;400;500;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

/* ── HEADER BANNER ── */
.kgj-header {
    background: linear-gradient(135deg, #141720 0%, #0D0F14 60%, #1a1f2e 100%);
    border-bottom: 1px solid #2A3050;
    padding: 18px 28px 14px 28px;
    margin: -1rem -1rem 1.5rem -1rem;
    display: flex;
    align-items: center;
    gap: 16px;
}
.kgj-header-hex {
    width: 38px; height: 38px;
    background: #F0A500;
    clip-path: polygon(50% 0%, 100% 25%, 100% 75%, 50% 100%, 0% 75%, 0% 25%);
    display: flex; align-items: center; justify-content: center;
    font-size: 18px;
    flex-shrink: 0;
}
.kgj-header-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 15px;
    font-weight: 600;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: #E8EAF0;
}
.kgj-header-title span { color: #F0A500; }
.kgj-header-sub {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    color: #6B7490;
    letter-spacing: 0.12em;
    margin-top: 3px;
}
.kgj-tag {
    margin-left: auto;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    color: #6B7490;
    border: 1px solid #2A3050;
    padding: 4px 12px;
    border-radius: 2px;
    letter-spacing: 0.1em;
}

/* ── KPI CARDS ── */
.kpi-card {
    background: #141720;
    border: 1px solid #2A3050;
    border-radius: 3px;
    padding: 14px 16px;
    position: relative;
    overflow: hidden;
}
.kpi-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0;
    width: 3px; height: 100%;
    border-radius: 3px 0 0 3px;
}
.kpi-positive::before { background: #4CAF88; }
.kpi-negative::before { background: #E05555; }
.kpi-neutral::before  { background: #F0A500; }
.kpi-info::before     { background: #5B8DEE; }
.kpi-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 9px;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: #6B7490;
    margin-bottom: 6px;
}
.kpi-value {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 24px;
    font-weight: 600;
    line-height: 1;
}
.kpi-pos { color: #4CAF88; }
.kpi-neg { color: #E05555; }
.kpi-acc { color: #F0A500; }
.kpi-sub {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 9px;
    color: #6B7490;
    margin-top: 5px;
}

/* ── DECISION PANEL ── */
.decision-on {
    display: inline-flex; align-items: center; gap: 6px;
    background: rgba(76,175,136,0.12);
    border: 1px solid rgba(76,175,136,0.35);
    color: #4CAF88;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    padding: 4px 12px; border-radius: 2px;
    font-weight: 600;
}
.decision-off {
    display: inline-flex; align-items: center; gap: 6px;
    background: rgba(224,85,85,0.10);
    border: 1px solid rgba(224,85,85,0.28);
    color: #E05555;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    padding: 4px 12px; border-radius: 2px;
    font-weight: 600;
}
.decision-maybe {
    display: inline-flex; align-items: center; gap: 6px;
    background: rgba(240,165,0,0.10);
    border: 1px solid rgba(240,165,0,0.28);
    color: #F0A500;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    padding: 4px 12px; border-radius: 2px;
    font-weight: 600;
}

/* ── SECTION HEADERS ── */
.section-hd {
    display: flex; align-items: center; gap: 8px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: #6B7490;
    border-bottom: 1px solid #2A3050;
    padding-bottom: 8px;
    margin-bottom: 12px;
}
.section-hd .dot {
    width: 6px; height: 6px;
    border-radius: 50%;
    background: #F0A500;
    box-shadow: 0 0 6px rgba(240,165,0,0.5);
}

/* ── TRIGGER TABLE ── */
.trigger-table {
    width: 100%;
    border-collapse: collapse;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
}
.trigger-table th {
    background: #1C2030;
    padding: 8px 12px;
    text-align: left;
    font-size: 9px;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    color: #6B7490;
    border-bottom: 1px solid #2A3050;
}
.trigger-table td {
    padding: 8px 12px;
    border-bottom: 1px solid #1C2030;
    color: #E8EAF0;
}
.trigger-table tr:hover td { background: rgba(255,255,255,0.02); }
.t-acc { color: #F0A500; font-weight: 600; }
.t-pos { color: #4CAF88; }
.t-neg { color: #E05555; }
.t-muted { color: #6B7490; font-size: 10px; }

/* ── INFO BOX ── */
.info-box {
    background: rgba(91,141,238,0.06);
    border-left: 3px solid #5B8DEE;
    border: 1px solid rgba(91,141,238,0.2);
    border-left-width: 3px;
    padding: 12px 16px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    color: #6B7490;
    line-height: 1.65;
    margin-bottom: 12px;
}

/* ── SIDEBAR ── */
[data-testid="stSidebar"] {
    background: #141720 !important;
    border-right: 1px solid #2A3050 !important;
}
[data-testid="stSidebar"] .stMarkdown p {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    color: #6B7490;
}

/* ── STATUS CHIP ── */
.status-chip-ok {
    display: inline-flex; align-items: center; gap: 6px;
    background: rgba(76,175,136,0.1);
    border: 1px solid rgba(76,175,136,0.3);
    color: #4CAF88;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    padding: 3px 10px; border-radius: 2px;
}
.status-chip-err {
    display: inline-flex; align-items: center; gap: 6px;
    background: rgba(224,85,85,0.1);
    border: 1px solid rgba(224,85,85,0.3);
    color: #E05555;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    padding: 3px 10px; border-radius: 2px;
}

/* Plotly charts */
.stPlotlyChart { border: 1px solid #2A3050; border-radius: 2px; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# HEADER
# ──────────────────────────────────────────────

st.markdown("""
<div class="kgj-header">
  <div class="kgj-header-hex">⚡</div>
  <div>
    <div class="kgj-header-title">KGJ <span>DISPATCH</span> KALKULÁTOR</div>
    <div class="kgj-header-sub">Kogenerační jednotka · Dispatch Optimalizace · Analýza marží</div>
  </div>
  <div class="kgj-tag">v3.0 · STREAMLIT</div>
</div>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────
# SIDEBAR — PARAMETERS
# ──────────────────────────────────────────────

with st.sidebar:
    st.markdown('<div class="section-hd"><div class="dot"></div> Parametry systému</div>', unsafe_allow_html=True)

    with st.expander("⚡ KGJ (kogenerace)", expanded=True):
        kgj_heat_output = st.number_input("Tep. výkon (MWh)", value=1.09, step=0.01, format="%.3f")
        kgj_el_output   = st.number_input("El. výkon (MWh)",  value=0.999, step=0.001, format="%.3f")
        kgj_heat_eff    = st.number_input("Tepelná účinnost", value=0.46, step=0.01, min_value=0.1, max_value=1.0, format="%.2f")
        kgj_service     = st.number_input("Servis (EUR/hod)",  value=12.0, step=0.5, format="%.1f")
        kgj_min_load    = st.slider("Min. zatížení", 0.10, 1.00, 0.55, 0.01)

    with st.expander("🔥 Plynový kotel"):
        boiler_eff      = st.number_input("Účinnost kotle", value=0.95, step=0.01, min_value=0.5, max_value=1.0, format="%.2f")
        boiler_max_heat = st.number_input("Max. výkon (MW)",  value=3.91, step=0.01, format="%.2f")

    with st.expander("🌡 Elektrokotel"):
        eboiler_eff      = st.number_input("Účinnost EKotlu", value=0.98, step=0.01, min_value=0.5, max_value=1.0, format="%.2f")
        eboiler_max_heat = st.number_input("Max. výkon (MW)",  value=0.60564, step=0.001, format="%.5f")

    with st.expander("⚙ Systémové"):
        ee_dist_cost   = st.number_input("Distribuce EE (EUR/MWh)", value=33.0, step=0.5, format="%.1f")
        heat_min_cover = st.slider("Min. pokrytí tepla", 0.80, 1.00, 0.99, 0.01)
        min_up         = st.number_input("Min. doba běhu (hod)",  value=4, step=1, min_value=1, max_value=24)
        min_down       = st.number_input("Min. doba stání (hod)", value=4, step=1, min_value=1, max_value=24)
        initial_state  = st.selectbox("Počáteční stav KGJ", [0, 1], format_func=lambda x: "Vypnuto" if x == 0 else "Zapnuto")

    params = TechParams(
        kgj_heat_output=kgj_heat_output,
        kgj_el_output=kgj_el_output,
        kgj_heat_eff=kgj_heat_eff,
        kgj_service=kgj_service,
        kgj_min_load=kgj_min_load,
        boiler_eff=boiler_eff,
        boiler_max_heat=boiler_max_heat,
        eboiler_eff=eboiler_eff,
        eboiler_max_heat=eboiler_max_heat,
        ee_dist_cost=ee_dist_cost,
        heat_min_cover=heat_min_cover,
        min_up=int(min_up),
        min_down=int(min_down),
        initial_state=initial_state,
    )

    st.divider()
    st.caption("KGJ Dispatch v3.0 · PuLP/CBC solver")


# ──────────────────────────────────────────────
# MAIN TABS
# ──────────────────────────────────────────────

tab_dashboard, tab_dispatch, tab_data, tab_info = st.tabs([
    "📊 Dashboard",
    "🔧 LP Dispatch",
    "📂 Data & Export",
    "📖 Metodika",
])


# ═══════════════════════════════════════════════
# TAB 1: DASHBOARD (single-point margin analysis)
# ═══════════════════════════════════════════════

with tab_dashboard:

    st.markdown('<div class="section-hd"><div class="dot"></div> Aktuální ceny — single-point analýza</div>', unsafe_allow_html=True)

    col_ee, col_gas, col_heat, col_dem = st.columns(4)
    with col_ee:
        ee_price   = st.number_input("⚡ Cena EE (EUR/MWh)",   value=80.0,  step=1.0, format="%.1f", key="sp_ee")
    with col_gas:
        gas_price  = st.number_input("🔥 Cena plynu (EUR/MWh)", value=35.0, step=1.0, format="%.1f", key="sp_gas")
    with col_heat:
        heat_price = st.number_input("🌡 Cena tepla (EUR/MWh)", value=55.0, step=1.0, format="%.1f", key="sp_heat")
    with col_dem:
        heat_dem   = st.number_input("📦 Poptávka tepla (MWh)", value=0.90, step=0.05, format="%.2f", key="sp_dem")

    r = compute_margins(ee_price, gas_price, heat_price, params)
    best = best_source(r["m1"], r["m2"], r["m3"], r["m4"])

    # ── KPI CARDS ──
    st.markdown('<div class="section-hd" style="margin-top:18px;"><div class="dot"></div> Marže zdrojů (EUR/MWh tepla)</div>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    def kpi_card(col, label, val, card_class):
        sign = "+" if val >= 0 else ""
        val_class = "kpi-pos" if val >= 0 else "kpi-neg"
        col.markdown(f"""
        <div class="kpi-card {card_class}">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value {val_class}">{sign}{val:.2f}</div>
          <div class="kpi-sub">EUR / MWh tepla</div>
        </div>""", unsafe_allow_html=True)

    kpi_card(c1, "1 — Plynový kotel",     r["m1"], "kpi-info")
    kpi_card(c2, "2 — KGJ + spot",        r["m2"], "kpi-neutral")
    kpi_card(c3, "3 — EKotel (síť)",      r["m3"], "kpi-info")
    kpi_card(c4, "4 — KGJ + elektrokotel", r["m4"], "kpi-positive" if r["m4"] >= 0 else "kpi-negative")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── DECISION + TRIGGERS ──
    col_dec, col_trg = st.columns([1, 1.5])

    with col_dec:
        st.markdown('<div class="section-hd"><div class="dot"></div> Doporučení</div>', unsafe_allow_html=True)

        # KGJ decision
        kgj_on = r["m2"] > 0 or r["m4"] > 0
        kgj_possible = heat_dem > 0 and (params.kgj_heat_output >= heat_dem * params.heat_min_cover)
        ee_trigger_hit = ee_price > r["trigger_ee_only"]

        if kgj_on and kgj_possible:
            badge = '<span class="decision-on">▲ KGJ ZAPNOUT</span>'
            msg   = "KGJ je ziskový při aktuálních cenách."
        elif ee_trigger_hit and not kgj_on:
            badge = '<span class="decision-maybe">~ KGJ ZVÁŽIT</span>'
            msg   = "Elektřina triggeruje, ale marže z tepla záporná."
        else:
            badge = '<span class="decision-off">▼ KGJ VYPNOUT</span>'
            msg   = "KGJ není ziskový při aktuálních cenách."

        st.markdown(f"""
        {badge}
        <p style="font-family:'IBM Plex Mono',monospace;font-size:11px;color:#6B7490;margin-top:8px;">{msg}</p>
        """, unsafe_allow_html=True)

        st.markdown(f"""
        <p style="font-family:'IBM Plex Mono',monospace;font-size:11px;color:#6B7490;margin-top:14px;line-height:1.8;">
        <span style="color:#6B7490;">Optimální zdroj:</span><br>
        <span style="color:{best['color']};font-size:14px;font-weight:600;">{best['name']}</span><br>
        <span style="color:{'#4CAF88' if best['m']>=0 else '#E05555'};">{'+' if best['m']>=0 else ''}{best['m']:.2f} EUR/MWh</span>
        </p>""", unsafe_allow_html=True)

    with col_trg:
        st.markdown('<div class="section-hd"><div class="dot"></div> Trigger analýza</div>', unsafe_allow_html=True)
        trg1 = r["trigger_ee_only"]
        trg2 = r["trigger_full"]
        dist1 = ee_price - trg1
        dist2 = ee_price - trg2

        st.markdown(f"""
        <table class="trigger-table">
          <thead>
            <tr><th>Trigger</th><th>Hodnota</th><th>Aktuální</th><th>Vzdálenost</th><th>Stav</th></tr>
          </thead>
          <tbody>
            <tr>
              <td>KGJ trigger — EE only</td>
              <td class="t-acc">{trg1:.2f}</td>
              <td>{ee_price:.1f}</td>
              <td class="{'t-pos' if dist1>=0 else 't-neg'}">{'+' if dist1>=0 else ''}{dist1:.2f}</td>
              <td class="t-muted">{'✓ EE nad triggerem' if dist1>=0 else '✗ EE pod triggerem'}</td>
            </tr>
            <tr>
              <td>KGJ trigger — full (tep.+el.)</td>
              <td class="t-acc">{trg2:.2f}</td>
              <td>{ee_price:.1f}</td>
              <td class="{'t-pos' if dist2>=0 else 't-neg'}">{'+' if dist2>=0 else ''}{dist2:.2f}</td>
              <td class="t-muted">{'✓ Ziskový i bez tepla' if dist2>=0 else '✗ Nutný příjem z tepla'}</td>
            </tr>
            <tr>
              <td>Kotel breakeven (teplo)</td>
              <td class="t-acc">{r['cost1']:.2f}</td>
              <td>{heat_price:.1f}</td>
              <td class="{'t-pos' if heat_price>=r['cost1'] else 't-neg'}">{'+' if heat_price>=r['cost1'] else ''}{heat_price-r['cost1']:.2f}</td>
              <td class="t-muted">{'✓ Ziskový' if heat_price>=r['cost1'] else '✗ Ztrátový'}</td>
            </tr>
            <tr>
              <td>EKotel breakeven (teplo)</td>
              <td class="t-acc">{r['cost3']:.2f}</td>
              <td>{heat_price:.1f}</td>
              <td class="{'t-pos' if heat_price>=r['cost3'] else 't-neg'}">{'+' if heat_price>=r['cost3'] else ''}{heat_price-r['cost3']:.2f}</td>
              <td class="t-muted">{'✓ Ziskový' if heat_price>=r['cost3'] else '✗ Ztrátový'}</td>
            </tr>
          </tbody>
        </table>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── CHARTS ──
    col_chart1, col_chart2 = st.columns(2)
    with col_chart1:
        st.plotly_chart(ch.margin_bar_chart(r, heat_price), use_container_width=True)
    with col_chart2:
        fig_sens = ch.sensitivity_chart(gas_price, heat_price, params)
        # Add current EE vertical line
        fig_sens.add_vline(
            x=ee_price, line_color="#F0A500", line_width=1.5, line_dash="dash",
            annotation_text=f"Aktuální EE: {ee_price:.0f}",
            annotation_font=dict(color="#F0A500", size=10),
        )
        st.plotly_chart(fig_sens, use_container_width=True)

    # ── DETAIL TABLE ──
    st.markdown('<div class="section-hd"><div class="dot"></div> Detailní rozpad — všechny zdroje</div>', unsafe_allow_html=True)

    detail_df = pd.DataFrame({
        "Zdroj": ["Plynový kotel", "KGJ + spot prodej", "Elektrokotel (síť)", "KGJ + elektrokotel"],
        "Náklad (EUR/MWh)": [r["cost1"], r["cost2"], r["cost3"], r["cost4"]],
        "Cena tepla (EUR/MWh)": [heat_price]*4,
        "Marže (EUR/MWh)": [r["m1"], r["m2"], r["m3"], r["m4"]],
        "Status": [
            "✅ Ziskový" if r["m1"] >= 0 else "❌ Ztrátový",
            "✅ Ziskový" if r["m2"] >= 0 else "❌ Ztrátový",
            "✅ Ziskový" if r["m3"] >= 0 else "❌ Ztrátový",
            "✅ Ziskový" if r["m4"] >= 0 else "❌ Ztrátový",
        ]
    })

    def color_margin(val):
        if isinstance(val, float):
            return f"color: {'#4CAF88' if val >= 0 else '#E05555'}"
        return ""

    st.dataframe(
        detail_df.style
            .format({"Náklad (EUR/MWh)": "{:.2f}", "Cena tepla (EUR/MWh)": "{:.2f}", "Marže (EUR/MWh)": "{:+.2f}"})
            .map(color_margin, subset=["Marže (EUR/MWh)", "Náklad (EUR/MWh)"])
            .set_properties(**{
                "font-family": "IBM Plex Mono, monospace",
                "font-size": "12px",
            }),
        use_container_width=True,
        hide_index=True,
    )


# ═══════════════════════════════════════════════
# TAB 2: FULL LP DISPATCH
# ═══════════════════════════════════════════════

with tab_dispatch:

    st.markdown('<div class="section-hd"><div class="dot"></div> LP Dispatch Optimalizace (PuLP / CBC)</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="info-box">
    Nahrajte soubor <strong>aki11.xlsx</strong> (sloupce: datetime, ee_price, gas_price, heat_price, heat_demand) nebo vygenerujte vzorová data.
    CBC solver řeší celý MIP model s min-up/min-down constrainty.
    </div>
    """, unsafe_allow_html=True)

    col_up, col_sample = st.columns([2, 1])
    with col_up:
        uploaded = st.file_uploader("📂 Nahrát Excel / CSV", type=["xlsx", "xls", "csv"], key="dispatch_upload")
    with col_sample:
        st.markdown("<br>", unsafe_allow_html=True)
        gen_sample = st.button("📋 Vzorová data (24h)", use_container_width=True)

    df_input = None

    if uploaded is not None:
        try:
            if uploaded.name.endswith(".csv"):
                df_input = pd.read_csv(uploaded)
            else:
                df_input = pd.read_excel(uploaded)
            # Normalize columns
            df_input.columns = [c.strip().lower().replace(" ", "_") for c in df_input.columns]
            if list(df_input.columns[:5]) != ["datetime", "ee_price", "gas_price", "heat_price", "heat_demand"]:
                df_input.columns = ["datetime", "ee_price", "gas_price", "heat_price", "heat_demand"] + list(df_input.columns[5:])
            df_input = df_input.reset_index(drop=True)
            st.markdown(f'<span class="status-chip-ok">✓ Načteno {len(df_input)} řádků</span>', unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Chyba při načítání: {e}")

    if gen_sample:
        np.random.seed(42)
        hours = pd.date_range("2024-01-15 00:00", periods=24, freq="h")
        ee_base = 75 + 25 * np.sin(np.linspace(0, 2*np.pi, 24)) + np.random.randn(24)*8
        gas_base = np.random.uniform(32, 38, 24)
        heat_base = np.random.uniform(50, 60, 24)
        demand_base = 0.3 + 0.8 * np.clip(np.sin(np.linspace(-np.pi/2, 3*np.pi/2, 24)), 0, 1) + np.random.rand(24)*0.15
        df_input = pd.DataFrame({
            "datetime":    hours,
            "ee_price":    np.round(ee_base, 2),
            "gas_price":   np.round(gas_base, 2),
            "heat_price":  np.round(heat_base, 2),
            "heat_demand": np.round(np.clip(demand_base, 0.1, 1.2), 3),
        })
        st.markdown(f'<span class="status-chip-ok">✓ Vygenerováno {len(df_input)} vzorových hodin</span>', unsafe_allow_html=True)

    if df_input is not None:
        # Show input preview
        with st.expander("Náhled vstupních dat", expanded=False):
            st.dataframe(df_input, use_container_width=True, height=200)
            st.plotly_chart(ch.prices_chart(df_input), use_container_width=True)

        # Solve button
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("▶ SPUSTIT DISPATCH OPTIMALIZACI", type="primary", use_container_width=True):
            with st.spinner("⚙ CBC solver pracuje… (může trvat desítky sekund)"):
                result_df = run_dispatch(df_input, params)

            if result_df is None:
                st.error("❌ Solver nenašel optimální řešení. Zkontrolujte vstupní data.")
            else:
                st.session_state["result_df"] = result_df
                st.markdown('<span class="status-chip-ok">✓ Optimalizace dokončena</span>', unsafe_allow_html=True)

    # Show results if available
    if "result_df" in st.session_state:
        result_df = st.session_state["result_df"]

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="section-hd"><div class="dot"></div> Výsledky dispatch plánu</div>', unsafe_allow_html=True)

        # Summary KPIs
        total_profit     = result_df["Total_profit_EUR"].sum()
        kgj_hours        = result_df["KGJ_on"].sum()
        kgj_starts_count = result_df["KGJ_start"].sum()
        avg_kgj_load     = result_df.loc[result_df["KGJ_on"]==1, "KGJ_load_pct"].mean() if kgj_hours > 0 else 0
        total_ee_sold    = result_df["EE_Sold_Spot_MWh"].sum()
        total_heat_del   = (result_df["KGJ_heat_MWh"] + result_df["Gas_boiler_heat_MWh"] + result_df["Electric_boiler_heat_MWh"]).sum()

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.markdown(f"""<div class="kpi-card {'kpi-positive' if total_profit>=0 else 'kpi-negative'}">
            <div class="kpi-label">Celkový zisk</div>
            <div class="kpi-value {'kpi-pos' if total_profit>=0 else 'kpi-neg'}">{'+' if total_profit>=0 else ''}{total_profit:,.0f}</div>
            <div class="kpi-sub">EUR za period</div></div>""", unsafe_allow_html=True)
        k2.markdown(f"""<div class="kpi-card kpi-neutral">
            <div class="kpi-label">KGJ hodiny</div>
            <div class="kpi-value kpi-acc">{int(kgj_hours)}</div>
            <div class="kpi-sub">z {len(result_df)} hodin</div></div>""", unsafe_allow_html=True)
        k3.markdown(f"""<div class="kpi-card kpi-info">
            <div class="kpi-label">Průměrné zatížení KGJ</div>
            <div class="kpi-value">{avg_kgj_load:.1f}%</div>
            <div class="kpi-sub">při běhu</div></div>""", unsafe_allow_html=True)
        k4.markdown(f"""<div class="kpi-card kpi-neutral">
            <div class="kpi-label">EE prodáno spot</div>
            <div class="kpi-value kpi-acc">{total_ee_sold:.2f}</div>
            <div class="kpi-sub">MWh celkem</div></div>""", unsafe_allow_html=True)
        k5.markdown(f"""<div class="kpi-card kpi-info">
            <div class="kpi-label">Teplo dodáno</div>
            <div class="kpi-value">{total_heat_del:.2f}</div>
            <div class="kpi-sub">MWh celkem</div></div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Charts grid
        st.plotly_chart(ch.dispatch_area_chart(result_df), use_container_width=True)

        col_el, col_cum = st.columns(2)
        with col_el:
            st.plotly_chart(ch.electricity_flow_chart(result_df), use_container_width=True)
        with col_cum:
            st.plotly_chart(ch.cumulative_profit_chart(result_df), use_container_width=True)

        st.plotly_chart(ch.kgj_status_chart(result_df), use_container_width=True)
        st.plotly_chart(ch.margin_heatmap(result_df), use_container_width=True)


# ═══════════════════════════════════════════════
# TAB 3: DATA & EXPORT
# ═══════════════════════════════════════════════

with tab_data:

    st.markdown('<div class="section-hd"><div class="dot"></div> Výsledková tabulka & Export</div>', unsafe_allow_html=True)

    if "result_df" not in st.session_state:
        st.info("Spusťte nejprve LP Dispatch Optimalizaci v záložce 🔧 LP Dispatch.")
    else:
        result_df = st.session_state["result_df"]

        # Column selector
        all_cols = list(result_df.columns)
        default_cols = [
            "datetime", "EE_price_EUR_MWh", "Gas_price_EUR_MWh", "Heat_price_EUR_MWh",
            "Heat_demand_MWh", "KGJ_on", "KGJ_load_pct",
            "KGJ_heat_MWh", "Gas_boiler_heat_MWh", "Electric_boiler_heat_MWh",
            "EE_Sold_Spot_MWh", "Total_profit_EUR",
            "Margin_2_KGJ_Spot_EUR_per_MWh", "Margin_4_KGJ_EBoiler_EUR_per_MWh",
        ]
        sel_cols = st.multiselect(
            "Zobrazit sloupce",
            all_cols,
            default=[c for c in default_cols if c in all_cols],
        )

        if sel_cols:
            show_df = result_df[sel_cols].copy()

            def highlight_profit(val):
                if isinstance(val, (int, float)):
                    if "profit" in str(val).lower() or "margin" in str(val).lower():
                        return f"color: {'#4CAF88' if val >= 0 else '#E05555'}"
                return ""

            st.dataframe(
                show_df.style.format(
                    {c: "{:.3f}" for c in show_df.select_dtypes("float").columns}
                ).set_properties(**{
                    "font-family": "IBM Plex Mono, monospace",
                    "font-size": "11px",
                }),
                use_container_width=True,
                height=420,
            )

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="section-hd"><div class="dot"></div> Export</div>', unsafe_allow_html=True)

        col_xlsx, col_csv = st.columns(2)

        with col_xlsx:
            buf_xlsx = io.BytesIO()
            with pd.ExcelWriter(buf_xlsx, engine="xlsxwriter") as writer:
                result_df.to_excel(writer, index=False, sheet_name="Dispatch_Results")
                # Also write summary
                summary = pd.DataFrame({
                    "Metrika": ["Celkový zisk (EUR)", "KGJ hodiny", "EE prodáno (MWh)", "Teplo dodáno (MWh)"],
                    "Hodnota": [
                        result_df["Total_profit_EUR"].sum(),
                        int(result_df["KGJ_on"].sum()),
                        result_df["EE_Sold_Spot_MWh"].sum(),
                        (result_df["KGJ_heat_MWh"] + result_df["Gas_boiler_heat_MWh"] + result_df["Electric_boiler_heat_MWh"]).sum(),
                    ]
                })
                summary.to_excel(writer, index=False, sheet_name="Souhrn")
            st.download_button(
                "⬇ Stáhnout Excel (.xlsx)",
                data=buf_xlsx.getvalue(),
                file_name="kgj_dispatch_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

        with col_csv:
            csv_data = result_df.to_csv(index=False, float_format="%.4f").encode("utf-8")
            st.download_button(
                "⬇ Stáhnout CSV",
                data=csv_data,
                file_name="kgj_dispatch_results.csv",
                mime="text/csv",
                use_container_width=True,
            )


# ═══════════════════════════════════════════════
# TAB 4: METODIKA
# ═══════════════════════════════════════════════

with tab_info:

    st.markdown('<div class="section-hd"><div class="dot"></div> Metodika výpočtu</div>', unsafe_allow_html=True)

    col_m1, col_m2 = st.columns(2)

    with col_m1:
        st.markdown("""
        <div class="info-box">
        <strong style="color:#F0A500;">Zdroj 1 — Plynový kotel</strong><br><br>
        Náklad = Cena_plynu / η_kotel
        <br><br>
        Nejjednodušší zdroj tepla. Ziskový pokud cena tepla > náklad.
        </div>

        <div class="info-box">
        <strong style="color:#F0A500;">Zdroj 2 — KGJ + prodej EE na spot</strong><br><br>
        Náklad = Cena_plynu × (Q_plyn / Q_tep)<br>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;+ Servis / Q_tep<br>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;− Cena_EE × (Q_el / Q_tep)
        <br><br>
        KGJ generuje elektřinu → prodej na spot kompenzuje část nákladů.
        </div>

        <div class="info-box">
        <strong style="color:#F0A500;">Zdroj 3 — Elektrokotel ze sítě</strong><br><br>
        Náklad = (Cena_EE + Distribuce_EE) / η_ekotel
        <br><br>
        Přímá konverze nakoupené EE na teplo. Drahé při vysoké ceně EE.
        </div>

        <div class="info-box">
        <strong style="color:#F0A500;">Zdroj 4 — KGJ + elektrokotel (agregát)</strong><br><br>
        Q_agg = Q_tep_KGJ + Q_tep_ekotel<br>
        EE_surplus = Q_el_KGJ − (Q_ekotel / η_ekotel)<br>
        Náklad = (Cena_plynu × Q_plyn + Servis − Cena_EE × EE_surplus) / Q_agg
        <br><br>
        KGJ napájí elektrokotel interně + přebytek jde na spot.
        </div>
        """, unsafe_allow_html=True)

    with col_m2:
        st.markdown("""
        <div class="info-box">
        <strong style="color:#00D4AA;">Trigger — KGJ el. only (EE threshold)</strong><br><br>
        Trigger_EE = (Cena_plynu × Q_plyn + Servis) / Q_el
        <br><br>
        Pokud cena EE > Trigger_EE → KGJ generuje zisk pouze z elektřiny,
        bez ohledu na příjem z tepla.
        </div>

        <div class="info-box">
        <strong style="color:#00D4AA;">Trigger — KGJ full (EE + teplo)</strong><br><br>
        Trigger_full = (Cena_plynu × Q_plyn + Servis − Cena_tep × Q_tep) / Q_el
        <br><br>
        Pokud cena EE > Trigger_full → KGJ je ziskový i bez nutnosti prodeje tepla.
        </div>

        <div class="info-box">
        <strong style="color:#5B8DEE;">LP Model (záložka Dispatch)</strong><br><br>
        Solver: PuLP / CBC (MILP)<br>
        Proměnné: q_kgj, q_boiler, q_eboiler, ee_sold_spot, ee_to_eboiler_int/grid<br>
        Binární: kgj_on, kgj_start, kgj_stop<br><br>
        Constrainty:<br>
        · Min. pokrytí tepelné poptávky (heat_min_cover)<br>
        · Min. zatížení KGJ (KGJ_MIN_LOAD)<br>
        · Min. doba běhu (MIN_UP) / stání (MIN_DOWN)<br>
        · Bilance elektřiny a tepla<br><br>
        Účelová funkce: MAX Σ (příjem teplo + příjem EE − náklad plyn − náklad EE distribuce − servis)
        </div>

        <div class="info-box" style="border-left-color:#E05555;background:rgba(224,85,85,0.05);">
        <strong style="color:#E05555;">Poznámky k interpretaci</strong><br><br>
        · Dashboard (single-point) počítá marginální analýzu bez LP constraintů.<br>
        · LP Dispatch respektuje min-up/min-down a Min. zatížení — výsledky se mohou lišit.<br>
        · Bypass tepla nastává, když KGJ při min. zatížení přesahuje poptávku.<br>
        · Distribuce EE se platí pouze při nákupu EE ze sítě pro elektrokotel.
        </div>
        """, unsafe_allow_html=True)

    st.markdown("""
    <div class="info-box" style="margin-top:8px;">
    <strong>Technické parametry (výchozí hodnoty)</strong><br><br>
    kgj_heat_output = 1.09 MWh · kgj_el_output = 0.999 MWh · kgj_heat_eff = 0.46<br>
    kgj_service = 12 EUR/h · KGJ_MIN_LOAD = 55% · boiler_eff = 0.95 · boiler_max = 3.91 MW<br>
    eboiler_eff = 0.98 · eboiler_max = 0.60564 MW · EE_DIST = 33 EUR/MWh · HEAT_MIN_COVER = 99%<br>
    MIN_UP = MIN_DOWN = 4 h
    </div>
    """, unsafe_allow_html=True)
