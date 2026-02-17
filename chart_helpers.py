"""
KGJ Chart Helpers — Plotly visualizations
"""

import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import pandas as pd
import numpy as np

# ──────────────────────────────────────────────
# PALETTE
# ──────────────────────────────────────────────
COLORS = {
    "bg":       "#0D0F14",
    "surface":  "#141720",
    "surface2": "#1C2030",
    "border":   "#2A3050",
    "accent":   "#F0A500",
    "accent2":  "#00D4AA",
    "blue":     "#5B8DEE",
    "purple":   "#9B59B6",
    "green":    "#4CAF88",
    "red":      "#E05555",
    "text":     "#E8EAF0",
    "muted":    "#6B7490",
}

SRC_COLORS = {
    "Plynový kotel":      COLORS["blue"],
    "KGJ + spot prodej":  COLORS["accent"],
    "Elektrokotel (síť)": COLORS["purple"],
    "KGJ + elektrokotel": COLORS["accent2"],
    "Žádný zdroj":        COLORS["red"],
}

LAYOUT_DEFAULTS = dict(
    paper_bgcolor=COLORS["surface"],
    plot_bgcolor=COLORS["surface2"],
    font=dict(family="IBM Plex Mono, monospace", color=COLORS["text"], size=11),
    margin=dict(l=16, r=16, t=40, b=16),
    legend=dict(
        bgcolor=COLORS["surface"],
        bordercolor=COLORS["border"],
        borderwidth=1,
        font=dict(size=10),
    ),
    xaxis=dict(
        gridcolor=COLORS["border"],
        linecolor=COLORS["border"],
        tickfont=dict(size=10),
        zerolinecolor=COLORS["border"],
    ),
    yaxis=dict(
        gridcolor=COLORS["border"],
        linecolor=COLORS["border"],
        tickfont=dict(size=10),
        zerolinecolor=COLORS["border"],
    ),
)


def apply_layout(fig, title="", height=380):
    fig.update_layout(
        **LAYOUT_DEFAULTS,
        title=dict(
            text=title,
            font=dict(size=12, color=COLORS["muted"]),
            x=0,
            xref="paper",
        ),
        height=height,
    )
    return fig


# ──────────────────────────────────────────────
# 1. MARGIN BAR CHART (single point)
# ──────────────────────────────────────────────

def margin_bar_chart(costs: dict, heat_price: float) -> go.Figure:
    labels = ["Plynový kotel", "KGJ + spot prodej", "Elektrokotel (síť)", "KGJ + elektrokotel"]
    cost_vals = [costs["cost1"], costs["cost2"], costs["cost3"], costs["cost4"]]
    margin_vals = [costs["m1"], costs["m2"], costs["m3"], costs["m4"]]
    colors = [COLORS["blue"], COLORS["accent"], COLORS["purple"], COLORS["accent2"]]

    fig = make_subplots(
        rows=1, cols=2,
        subplot_titles=["Náklady (EUR/MWh)", "Marže (EUR/MWh)"],
        horizontal_spacing=0.08,
    )

    # Cost bars
    fig.add_trace(go.Bar(
        x=labels, y=cost_vals,
        marker_color=colors,
        marker_line_width=0,
        text=[f"{v:.2f}" for v in cost_vals],
        textposition="outside",
        textfont=dict(size=10),
        name="Náklad",
        showlegend=False,
    ), row=1, col=1)

    # Heat price reference line
    fig.add_hline(
        y=heat_price, line_dash="dash",
        line_color=COLORS["green"], line_width=1.5,
        annotation_text=f"Cena tepla: {heat_price:.1f}",
        annotation_font=dict(size=10, color=COLORS["green"]),
        row=1, col=1,
    )

    # Margin bars
    bar_colors = [COLORS["green"] if m >= 0 else COLORS["red"] for m in margin_vals]
    fig.add_trace(go.Bar(
        x=labels, y=margin_vals,
        marker_color=bar_colors,
        marker_line_width=0,
        text=[f"{'+' if m>=0 else ''}{m:.2f}" for m in margin_vals],
        textposition="outside",
        textfont=dict(size=10),
        name="Marže",
        showlegend=False,
    ), row=1, col=2)

    fig.add_hline(y=0, line_color=COLORS["border"], line_width=1, row=1, col=2)

    apply_layout(fig, height=360)
    fig.update_layout(
        annotations=[
            dict(text="Náklady (EUR/MWh)", x=0.2, y=1.08, xref="paper", yref="paper",
                 showarrow=False, font=dict(size=11, color=COLORS["muted"])),
            dict(text="Marže (EUR/MWh)", x=0.8, y=1.08, xref="paper", yref="paper",
                 showarrow=False, font=dict(size=11, color=COLORS["muted"])),
        ]
    )
    fig.update_xaxes(tickangle=-15, tickfont=dict(size=9))
    return fig


# ──────────────────────────────────────────────
# 2. SENSITIVITY – margin vs EE price
# ──────────────────────────────────────────────

def sensitivity_chart(gas: float, heat_p: float, params) -> go.Figure:
    from dispatch_engine import compute_margins

    ee_range = np.arange(10, 250, 5)
    traces = {
        "Plynový kotel":      ([], COLORS["blue"]),
        "KGJ + spot prodej":  ([], COLORS["accent"]),
        "Elektrokotel (síť)": ([], COLORS["purple"]),
        "KGJ + elektrokotel": ([], COLORS["accent2"]),
    }
    for ee in ee_range:
        r = compute_margins(ee, gas, heat_p, params)
        traces["Plynový kotel"][0].append(r["m1"])
        traces["KGJ + spot prodej"][0].append(r["m2"])
        traces["Elektrokotel (síť)"][0].append(r["m3"])
        traces["KGJ + elektrokotel"][0].append(r["m4"])

    fig = go.Figure()
    for name, (vals, color) in traces.items():
        fig.add_trace(go.Scatter(
            x=ee_range, y=vals,
            name=name,
            line=dict(color=color, width=2),
            mode="lines",
        ))

    # Zero line
    fig.add_hline(y=0, line_color=COLORS["border"], line_width=1, line_dash="dot")

    # Current EE marker — will be added externally
    apply_layout(fig, "Citlivostní analýza — Marže zdrojů vs. Cena EE", height=350)
    fig.update_xaxes(title_text="Cena EE (EUR/MWh)", title_font=dict(size=10))
    fig.update_yaxes(title_text="Marže (EUR/MWh)", title_font=dict(size=10))
    return fig


# ──────────────────────────────────────────────
# 3. DISPATCH TIMESERIES – stacked area
# ──────────────────────────────────────────────

def dispatch_area_chart(result_df: pd.DataFrame) -> go.Figure:
    x = result_df["datetime"].astype(str)

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=x, y=result_df["KGJ_heat_MWh"],
        name="KGJ teplo", stackgroup="heat",
        fillcolor="rgba(240,165,0,0.45)",
        line=dict(color=COLORS["accent"], width=0.5),
        mode="lines",
    ))
    fig.add_trace(go.Scatter(
        x=x, y=result_df["Gas_boiler_heat_MWh"],
        name="Kotel teplo", stackgroup="heat",
        fillcolor="rgba(91,141,238,0.45)",
        line=dict(color=COLORS["blue"], width=0.5),
        mode="lines",
    ))
    fig.add_trace(go.Scatter(
        x=x, y=result_df["Electric_boiler_heat_MWh"],
        name="EKotel teplo", stackgroup="heat",
        fillcolor="rgba(155,89,182,0.45)",
        line=dict(color=COLORS["purple"], width=0.5),
        mode="lines",
    ))

    # Demand line
    fig.add_trace(go.Scatter(
        x=x, y=result_df["Heat_demand_MWh"],
        name="Poptávka", mode="lines",
        line=dict(color=COLORS["green"], width=2, dash="dash"),
    ))

    apply_layout(fig, "Dispatch plán — Výroba tepla (MWh)", height=340)
    fig.update_xaxes(title_text="Čas", title_font=dict(size=10), tickangle=-45)
    fig.update_yaxes(title_text="MWh", title_font=dict(size=10))
    return fig


# ──────────────────────────────────────────────
# 4. ELECTRICITY FLOW
# ──────────────────────────────────────────────

def electricity_flow_chart(result_df: pd.DataFrame) -> go.Figure:
    x = result_df["datetime"].astype(str)

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=x, y=result_df["EE_Sold_Spot_MWh"],
        name="EE prodej spot", stackgroup="ee",
        fillcolor="rgba(240,165,0,0.45)",
        line=dict(color=COLORS["accent"], width=0.5),
        mode="lines",
    ))
    fig.add_trace(go.Scatter(
        x=x, y=result_df["EE_to_EBoiler_Internal_MWh"],
        name="EE → EKotel (intern.)", stackgroup="ee",
        fillcolor="rgba(0,212,170,0.35)",
        line=dict(color=COLORS["accent2"], width=0.5),
        mode="lines",
    ))
    fig.add_trace(go.Scatter(
        x=x, y=result_df["KGJ_Electricity_MWh"],
        name="KGJ el. celkem", mode="lines",
        line=dict(color=COLORS["red"], width=1.5, dash="dot"),
    ))

    apply_layout(fig, "Dispatch plán — Elektřina z KGJ (MWh)", height=300)
    fig.update_xaxes(tickangle=-45, tickfont=dict(size=9))
    fig.update_yaxes(title_text="MWh", title_font=dict(size=10))
    return fig


# ──────────────────────────────────────────────
# 5. KGJ ON/OFF GANTT
# ──────────────────────────────────────────────

def kgj_status_chart(result_df: pd.DataFrame) -> go.Figure:
    x = result_df["datetime"].astype(str)

    fig = make_subplots(rows=2, cols=1, shared_xaxes=True,
                        row_heights=[0.6, 0.4], vertical_spacing=0.06)

    # KGJ Load
    fig.add_trace(go.Bar(
        x=x, y=result_df["KGJ_load_pct"],
        name="KGJ zatížení (%)",
        marker_color=[COLORS["accent"] if v > 0 else COLORS["surface2"]
                      for v in result_df["KGJ_load_pct"]],
        marker_line_width=0,
    ), row=1, col=1)

    # Starts / stops
    starts = result_df[result_df["KGJ_start"] == 1]
    stops  = result_df[result_df["KGJ_stop"]  == 1]

    fig.add_trace(go.Scatter(
        x=starts["datetime"].astype(str), y=[105]*len(starts),
        mode="markers", marker=dict(symbol="triangle-up", color=COLORS["green"], size=10),
        name="Start",
    ), row=1, col=1)
    fig.add_trace(go.Scatter(
        x=stops["datetime"].astype(str),  y=[105]*len(stops),
        mode="markers", marker=dict(symbol="triangle-down", color=COLORS["red"], size=10),
        name="Stop",
    ), row=1, col=1)

    # Profit per hour
    colors_profit = [COLORS["green"] if v >= 0 else COLORS["red"]
                     for v in result_df["Total_profit_EUR"]]
    fig.add_trace(go.Bar(
        x=x, y=result_df["Total_profit_EUR"],
        name="Zisk (EUR)",
        marker_color=colors_profit,
        marker_line_width=0,
    ), row=2, col=1)
    fig.add_hline(y=0, line_color=COLORS["border"], line_width=1, row=2, col=1)

    apply_layout(fig, "KGJ Provozní stav & Zisk po hodinách", height=380)
    fig.update_yaxes(title_text="Zatížení (%)", row=1, col=1, title_font=dict(size=10))
    fig.update_yaxes(title_text="EUR/h", row=2, col=1, title_font=dict(size=10))
    fig.update_xaxes(tickangle=-45, tickfont=dict(size=9), row=2, col=1)
    return fig


# ──────────────────────────────────────────────
# 6. PRICES OVERVIEW
# ──────────────────────────────────────────────

def prices_chart(df: pd.DataFrame) -> go.Figure:
    """Works with both raw input df (ee_price, ...) and result df (EE_price_EUR_MWh, ...)."""
    cols = df.columns.str.lower()

    def find_col(prefix):
        matches = [c for c in df.columns if c.lower().startswith(prefix)]
        return matches[0] if matches else None

    col_ee     = find_col("ee_price")
    col_gas    = find_col("gas_price")
    col_heat_p = find_col("heat_price")
    col_demand = find_col("heat_demand")

    x = df["datetime"].astype(str)

    fig = make_subplots(
        rows=2, cols=1, shared_xaxes=True,
        row_heights=[0.55, 0.45],
        vertical_spacing=0.07,
    )

    if col_ee:
        fig.add_trace(go.Scatter(
            x=x, y=df[col_ee],
            name="EE", line=dict(color=COLORS["accent"], width=2), mode="lines",
        ), row=1, col=1)
    if col_gas:
        fig.add_trace(go.Scatter(
            x=x, y=df[col_gas],
            name="Plyn", line=dict(color=COLORS["blue"], width=2), mode="lines",
        ), row=1, col=1)
    if col_heat_p:
        fig.add_trace(go.Scatter(
            x=x, y=df[col_heat_p],
            name="Teplo", line=dict(color=COLORS["green"], width=2), mode="lines",
        ), row=1, col=1)
    if col_demand:
        fig.add_trace(go.Scatter(
            x=x, y=df[col_demand],
            name="Poptávka tepla",
            line=dict(color=COLORS["accent2"], width=2), mode="lines",
            fill="tozeroy", fillcolor="rgba(0,212,170,0.12)",
        ), row=2, col=1)

    apply_layout(fig, height=380)
    fig.update_xaxes(tickangle=-45, tickfont=dict(size=9), row=2, col=1)
    fig.update_layout(
        annotations=[
            dict(text="Ceny energií (EUR/MWh)", x=0, y=1.04, xref="paper", yref="paper",
                 showarrow=False, font=dict(size=11, color=COLORS["muted"])),
            dict(text="Poptávka tepla (MWh)", x=0, y=0.42, xref="paper", yref="paper",
                 showarrow=False, font=dict(size=11, color=COLORS["muted"])),
        ]
    )
    return fig


# ──────────────────────────────────────────────
# 7. CUMULATIVE PROFIT
# ──────────────────────────────────────────────

def cumulative_profit_chart(result_df: pd.DataFrame) -> go.Figure:
    x = result_df["datetime"].astype(str)
    cum = result_df["Total_profit_EUR"].cumsum()

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=x, y=cum,
        mode="lines",
        line=dict(color=COLORS["accent2"], width=2.5),
        fill="tozeroy",
        fillcolor="rgba(0,212,170,0.08)",
        name="Kumulativní zisk",
    ))
    fig.add_hline(y=0, line_color=COLORS["border"], line_width=1)

    apply_layout(fig, "Kumulativní zisk (EUR)", height=300)
    fig.update_xaxes(tickangle=-45, tickfont=dict(size=9))
    fig.update_yaxes(title_text="EUR", title_font=dict(size=10))
    return fig


# ──────────────────────────────────────────────
# 8. MARGIN HEATMAP (timeseries)
# ──────────────────────────────────────────────

def margin_heatmap(result_df: pd.DataFrame) -> go.Figure:
    cols = [
        "Margin_1_Boiler_EUR_per_MWh",
        "Margin_2_KGJ_Spot_EUR_per_MWh",
        "Margin_3_EBoiler_Grid_EUR_per_MWh",
        "Margin_4_KGJ_EBoiler_EUR_per_MWh",
    ]
    labels = ["Kotel", "KGJ+Spot", "EKotel", "KGJ+EK"]

    valid_cols   = [c for c in cols   if c in result_df.columns]
    valid_labels = [labels[i] for i, c in enumerate(cols) if c in result_df.columns]

    if not valid_cols:
        return go.Figure()

    z = result_df[valid_cols].T.values.tolist()
    x = result_df["datetime"].astype(str).tolist()

    heatmap = go.Heatmap(
        z=z,
        x=x,
        y=valid_labels,
        colorscale=[
            [0.0, COLORS["red"]],
            [0.5, COLORS["surface2"]],
            [1.0, COLORS["green"]],
        ],
        zmid=0,
        showscale=True,
    )

    fig = go.Figure(heatmap)

    fig.update_traces(
        selector=dict(type="heatmap"),
        colorbar_tickfont_size=9,
        colorbar_thickness=12,
        colorbar_len=0.8,
    )

    apply_layout(fig, "Marže zdrojů — heatmapa (EUR/MWh)", height=260)
    fig.update_xaxes(tickangle=-45, tickfont=dict(size=9))
    return fig
