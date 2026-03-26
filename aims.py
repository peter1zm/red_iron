import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="US Forest Service Vendors - AIMS Dashboard",
    page_icon="🌲",
    layout="wide",
)
st.markdown("""
<style>
    .dashboard-header {
        background: linear-gradient(135deg, #b91c1c, #7f1d1d);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .dashboard-header h1 { margin: 0; font-size: 2rem; font-weight: 700; }
    .dashboard-header p  { margin: 0.25rem 0 0; opacity: 0.85; font-size: 1rem; }

    .kpi-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: 1rem 1.25rem;
        text-align: center;
        box-shadow: 0 1px 4px rgba(0,0,0,.06);
    }
    .kpi-label { font-size: 0.78rem; color: #6b7280; font-weight: 600;
                 text-transform: uppercase; letter-spacing: .05em; margin-bottom: .2rem; }
    .kpi-value { font-size: 1.7rem; font-weight: 700; color: #111827; line-height: 1.1; }
    .kpi-sub   { font-size: 0.78rem; color: #6b7280; margin-top: .15rem; }

    .section-header {
        font-size: 1.1rem; font-weight: 700; color: #1f2937;
        border-left: 4px solid #b91c1c;
        padding-left: 0.6rem;
        margin: 1.5rem 0 0.75rem;
    }
</style>
""", unsafe_allow_html=True)


# ── Data loading ───────────────────────────────────────────────────────────────
DATA_PATH = "Red Iron - AIMS Data.xlsx"

RED_PALETTE = [
    "#b91c1c", "#ef4444", "#f97316", "#fb923c",
    "#fca5a5", "#7f1d1d", "#dc2626", "#fecaca",
]

# Extended palette for many dispatch centers
REGION_COLORS = {
    "Region 1": "#b91c1c", "Region 2": "#f97316", "Region 3": "#eab308",
    "Region 4": "#22c55e", "Region 5 North": "#3b82f6", "Region 5 South": "#8b5cf6",
    "Region 6": "#06b6d4", "Region 8": "#ec4899",
    "1": "#b91c1c", "2": "#f97316", "3": "#eab308",
    "4": "#22c55e", "5 North Ops": "#3b82f6", "5 South Ops": "#8b5cf6",
    "6": "#06b6d4", "8": "#ec4899",
}

PLOTLY_TEMPLATE = "plotly_white"

def safe_mean(series):
    return pd.to_numeric(series, errors="coerce").dropna().mean()

def safe_sum(series):
    return pd.to_numeric(series, errors="coerce").dropna().sum()

@st.cache_data
def load_data():
    # ── Toilets ──────────────────────────────────────────────────────────────
    df_t = pd.read_excel(DATA_PATH, sheet_name="Toilets and Handwash")
    df_t.columns = [
        "Rank", "Agreement Number", "Company Name", "Email", "Phone",
        "City", "State", "Region", "Dispatch Center",
        "Rate Toilet", "Rate Wheelchair", "Rate Handwash",
        "Mileage Rate", "Delivery/Pickup Rate", "Service Call Rate", "Relocation Fee Rate",
    ]
    df_t["Region"] = df_t["Region"].apply(
        lambda v: f"Region {int(v)}" if isinstance(v, (int, float)) and str(v) not in ("nan", "") else str(v).strip()
    )

    piv_t = pd.read_excel(DATA_PATH, sheet_name="Toilets Pivot", header=None)
    active_t = set(piv_t[piv_t[3].notna()][2].dropna().astype(str).str.strip())
    df_t["Active"] = df_t["Agreement Number"].astype(str).str.strip().isin(active_t)

    # ── UTVs ─────────────────────────────────────────────────────────────────
    df_u = pd.read_excel(DATA_PATH, sheet_name="UTVs")
    df_u.columns = [
        "Rank", "Agreement Number", "Vendor", "POC Name", "Email",
        "Region", "Dispatch Center", "Dispatch Center Code",
        "UTV Type", "Make", "Model", "Quantity", "Equipment Location",
        "Daily Rate", "Weekly Rate", "Monthly Rate",
        "Delivery Single", "Delivery Multiple", "MSRP",
    ]
    df_u["Region"] = df_u["Region"].apply(
        lambda v: f"Region {int(v)}" if isinstance(v, (int, float)) and str(v) not in ("nan", "") else str(v).strip()
    )
    df_u["Make"] = df_u["Make"].fillna("Unknown").str.strip()
    # Normalise "Honda " → "Honda"
    df_u["Make"] = df_u["Make"].str.strip()

    # Parse state from Equipment Location (e.g. "Whitehall, MT" → "MT")
    df_u["Eq State"] = df_u["Equipment Location"].str.extract(r",\s*([A-Z]{2})$")

    piv_u = pd.read_excel(DATA_PATH, sheet_name="UTVs Pivot", header=None)
    active_u = set(piv_u[piv_u[3].notna()][2].dropna().astype(str).str.strip())
    df_u["Active"] = df_u["Agreement Number"].astype(str).str.strip().isin(active_u)

    return df_t, df_u

df_toilets_raw, df_utvs_raw = load_data()

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="dashboard-header">
  <h1>🌲 US Forest Service Vendors – AIMS Dashboard</h1>
  <p>Vendor rate comparison, coverage, and rankings for Toilets & Handwash and UTVs</p>
</div>
""", unsafe_allow_html=True)

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab_toilets, tab_utvs = st.tabs(["🚽  Toilets & Handwash", "🚗  UTVs"])


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 1 – TOILETS
# ══════════════════════════════════════════════════════════════════════════════
with tab_toilets:

    with st.sidebar:
        st.title("Filters")
        with st.expander("🚽 Toilets Filters"):
            active_only_t = st.toggle("Active agreements only", value=True, key="active_t")

            regions_t = sorted(df_toilets_raw["Region"].unique())
            sel_regions_t = st.multiselect("Region", regions_t, default=regions_t, key="reg_t")

            states_t = sorted(
                df_toilets_raw[df_toilets_raw["Region"].isin(sel_regions_t)]["State"]
                .dropna().unique()
            )
            sel_states_t = st.multiselect("State", states_t, default=states_t, key="state_t")

            dcs_t = sorted(
                df_toilets_raw[df_toilets_raw["Region"].isin(sel_regions_t)]["Dispatch Center"]
                .dropna().unique()
            )
            sel_dcs_t = st.multiselect("Dispatch Center", dcs_t, default=dcs_t, key="dc_t")

            top_n_t = st.slider("# vendors in bar chart", 10, 50, 20, key="topn_t")

    # ── Filter ────────────────────────────────────────────────────────────────
    df_t = df_toilets_raw.copy()
    if active_only_t:
        df_t = df_t[df_t["Active"]]
    df_t = df_t[
        df_t["Region"].isin(sel_regions_t) &
        df_t["State"].isin(sel_states_t) &
        df_t["Dispatch Center"].isin(sel_dcs_t)
    ]
    for _c in ["Rate Toilet", "Rate Wheelchair", "Rate Handwash",
               "Mileage Rate", "Delivery/Pickup Rate", "Service Call Rate", "Relocation Fee Rate"]:
        df_t[_c] = pd.to_numeric(df_t[_c], errors="coerce")

    # ── KPIs ──────────────────────────────────────────────────────────────────
    k1, k2, k3, k4, k5 = st.columns(5)

    def kpi(col, label, value, sub=""):
        col.markdown(
            f'<div class="kpi-card"><div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div>'
            f'<div class="kpi-sub">{sub}</div></div>',
            unsafe_allow_html=True,
        )

    kpi(k1, "Vendors", f"{df_t['Company Name'].nunique():,}", f"of {df_toilets_raw['Company Name'].nunique()} total")
    kpi(k2, "States Covered", df_t["State"].nunique(), "unique states")
    kpi(k3, "Dispatch Centers", df_t["Dispatch Center"].nunique(), "unique centers")
    kpi(k4, "Avg Toilet Rate", f"${safe_mean(df_t['Rate Toilet']):.2f}", "per day")
    kpi(k5, "Avg Handwash Rate", f"${safe_mean(df_t['Rate Handwash']):.2f}", "per day")
    st.markdown("<br>", unsafe_allow_html=True)

    # ══ SECTION 1: Geographic Coverage ═══════════════════════════════════════
    st.markdown('<div class="section-header">📍 Geographic Coverage – What states are represented?</div>', unsafe_allow_html=True)

    col_map, col_state_bar = st.columns([3, 2])

    with col_map:
        st.subheader("Vendor Count by State")
        state_counts_t = (
            df_t.groupby("State")["Company Name"]
            .nunique()
            .reset_index()
            .rename(columns={"Company Name": "Vendors"})
        )
        fig_statemap_toilets = px.choropleth(
            state_counts_t,
            locations="State",
            locationmode="USA-states",
            color="Vendors",
            scope="usa",
            color_continuous_scale=["#fef2f2", "#b91c1c"],
            hover_data={"Vendors": True},
            template=PLOTLY_TEMPLATE,
        )
        fig_statemap_toilets.update_layout(
            height=380, margin=dict(l=0, r=0, t=10, b=10),
            coloraxis_colorbar=dict(title="# Vendors"),
        )
        st.plotly_chart(fig_statemap_toilets, use_container_width=True)

    with col_state_bar:
        st.subheader("Vendors per State (ranked)")
        state_bar_df = state_counts_t.sort_values("Vendors", ascending=True)
        fig_state_bar_toilets = px.bar(
            state_bar_df, x="Vendors", y="State", orientation="h",
            color="Vendors", color_continuous_scale=["#fecaca", "#b91c1c"],
            template=PLOTLY_TEMPLATE, text_auto=True,
        )
        fig_state_bar_toilets.update_layout(
            height=380, coloraxis_showscale=False,
            margin=dict(l=0, r=20, t=10, b=10),
            yaxis=dict(title=""), xaxis=dict(title="# Vendors"),
        )
        fig_state_bar_toilets.update_traces(textposition="outside")
        st.plotly_chart(fig_state_bar_toilets, use_container_width=True)

    # ══ SECTION 2: Vendor Rankings ════════════════════════════════════════════
    st.markdown('<div class="section-header">🏆 Vendor Rankings by Region & Dispatch Center</div>', unsafe_allow_html=True)

    rank_df = (
        df_t[[
            "Rank", "Company Name", "Region", "Dispatch Center", "State",
            "Rate Toilet", "Rate Wheelchair", "Rate Handwash",
            "Mileage Rate", "Delivery/Pickup Rate", "Service Call Rate",
            "Relocation Fee Rate", "Active",
        ]]
        .copy()
        .sort_values(["Region", "Dispatch Center", "Rank"])
        .reset_index(drop=True)
    )
    rank_df["Active"] = rank_df["Active"].map({True: "✅ Active", False: "❌ Inactive"})

    dollar_cols_t = [
        "Rate Toilet", "Rate Wheelchair", "Rate Handwash",
        "Mileage Rate", "Delivery/Pickup Rate", "Service Call Rate", "Relocation Fee Rate",
    ]
    st.caption("⬆️ Click any column header to sort. Default sorted by Region → Dispatch Center → Rank.")
    st.dataframe(
        rank_df.style.format(
            {c: "${:.2f}" for c in dollar_cols_t}, na_rep="—"
        ).background_gradient(subset=["Rate Toilet"], cmap="RdYlGn_r"),
        use_container_width=True,
        height=400,
    )
    st.download_button(
        "⬇️ Download Rankings",
        rank_df.to_csv(index=False),
        file_name="toilet_rankings_by_region_dc.csv",
        mime="text/csv",
    )

    # ══ SECTION 3: Rates by Dispatch Center ══════════════════════════════════
    st.markdown('<div class="section-header">💲 Daily Rates by Dispatch Center – One chart per rate type</div>', unsafe_allow_html=True)

    rate_types = {
        "Rate Toilet": "🚽 Daily Toilet Rate",
        "Rate Wheelchair": "♿ Daily Wheelchair Rate",
        "Rate Handwash": "🚿 Daily Handwash Rate",
    }

    # Add region color column for chart
    region_color_map = {r: REGION_COLORS.get(r, "#6b7280") for r in df_t["Region"].unique()}

    dc_rate_df = (
        df_t.groupby(["Region", "Dispatch Center"])[list(rate_types.keys())]
        .mean()
        .round(2)
        .reset_index()
    )
    dc_rate_df["Region_Color"] = dc_rate_df["Region"].map(region_color_map)

    for rate_col, rate_label in rate_types.items():
        st.subheader(f"{rate_label} – Average by Dispatch Center")
        chart_dc = (
            dc_rate_df[["Region", "Dispatch Center", rate_col]]
            .dropna(subset=[rate_col])
            .sort_values([rate_col], ascending=True)
        )
        fig_dc_toilets = px.bar(
            chart_dc,
            x=rate_col,
            y="Dispatch Center",
            color="Region",
            orientation="h",
            template=PLOTLY_TEMPLATE,
            text_auto=".2f",
            color_discrete_map=region_color_map,
            hover_data={"Region": True, "Dispatch Center": True, rate_col: ":.2f"},
        )
        fig_dc_toilets.update_layout(
            height=max(350, len(chart_dc) * 22),
            margin=dict(l=0, r=60, t=10, b=10),
            xaxis=dict(title="Avg Daily Rate ($)"),
            yaxis=dict(title=""),
            legend=dict(orientation="v", x=1.02, y=1),
        )
        fig_dc_toilets.update_traces(textposition="outside")
    st.plotly_chart(fig_dc_toilets, use_container_width=True)

    # ══ SECTION 4: Rate Distribution ══════════════════════════════════════════
    st.markdown('<div class="section-header">📊 Rate Distributions & Regional Averages</div>', unsafe_allow_html=True)

    col_box, col_region = st.columns([2, 3])

    with col_box:
        st.subheader("Rate Distribution by Type")
        melt_df = df_t[["Rate Toilet", "Rate Wheelchair", "Rate Handwash"]].melt(
            var_name="Type", value_name="Rate"
        ).dropna()
        melt_df["Type"] = melt_df["Type"].map({
            "Rate Toilet": "Toilet", "Rate Wheelchair": "Wheelchair", "Rate Handwash": "Handwash"
        })
        fig_box_toilets = px.box(
            melt_df, x="Type", y="Rate",
            color="Type",
            color_discrete_sequence=["#b91c1c", "#ef4444", "#f97316"],
            template=PLOTLY_TEMPLATE,
            points="outliers",
        )
        fig_box_toilets.update_layout(
            height=380, showlegend=False,
            margin=dict(l=0, r=0, t=10, b=10),
            yaxis=dict(title="Daily Rate ($)"),
        )
        st.plotly_chart(fig_box_toilets, use_container_width=True)

    with col_region:
        st.subheader("Average Rates by Region")
        rate_cols_t = ["Rate Toilet", "Rate Wheelchair", "Rate Handwash",
                       "Mileage Rate", "Delivery/Pickup Rate", "Service Call Rate"]
        region_df = df_t.groupby("Region")[rate_cols_t].mean().round(2).reset_index()
        nice_names = {
            "Rate Toilet": "Toilet ($/day)",
            "Rate Wheelchair": "Wheelchair ($/day)",
            "Rate Handwash": "Handwash ($/day)",
            "Mileage Rate": "Mileage ($/mi)",
            "Delivery/Pickup Rate": "Delivery ($)",
            "Service Call Rate": "Service Call ($)",
        }
        region_df = region_df.rename(columns=nice_names)
        fig_region_toilets = go.Figure()
        for name in nice_names.values():
            fig_region_toilets.add_trace(go.Bar(
                x=region_df["Region"].astype(str),
                y=region_df[name],
                name=name,
            ))
        fig_region_toilets.update_layout(
            barmode="group",
            template=PLOTLY_TEMPLATE,
            height=380,
            margin=dict(l=0, r=0, t=10, b=10),
            colorway=RED_PALETTE,
            xaxis=dict(title="Region"),
            yaxis=dict(title="Avg Rate ($)"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig_region_toilets, use_container_width=True)

    st.download_button(
        "⬇️ Download filtered Toilets data",
        rank_df.to_csv(index=False),
        file_name="filtered_toilets.csv",
        mime="text/csv",
    )


# ══════════════════════════════════════════════════════════════════════════════
#  TAB 2 – UTVs
# ══════════════════════════════════════════════════════════════════════════════
with tab_utvs:

    with st.sidebar:
        with st.expander("🚗 UTV Filters"):
            active_only_u = st.toggle("Active agreements only", value=True, key="active_u")

            regions_u = sorted(df_utvs_raw["Region"].unique())
            sel_regions_u = st.multiselect("Region", regions_u, default=regions_u, key="reg_u")

            #states_u = sorted(
            #    df_utvs_raw[df_utvs_raw["Region"].isin(sel_regions_t)]["State"]
            #    .dropna().unique()
            #)
            #sel_states_u = st.multiselect("State", states_u, default=states_u, key="state_u")

            utv_types = sorted(df_utvs_raw["UTV Type"].dropna().unique())
            sel_types = st.multiselect("UTV Type", utv_types, default=utv_types, key="type_u")

            makes_raw = sorted(df_utvs_raw["Make"].dropna().unique())
            sel_makes = st.multiselect("Make", makes_raw, default=makes_raw, key="make_u")

            dcs_u = sorted(
                df_utvs_raw[df_utvs_raw["Region"].isin(sel_regions_u)]["Dispatch Center"]
                .dropna().unique()
            )
            sel_dcs_u = st.multiselect("Dispatch Center", dcs_u, default=dcs_u, key="dc_u")

            rate_period = st.radio(
                "Rate period (for KPIs & top vendors)",
                ["Daily Rate", "Weekly Rate", "Monthly Rate"],
                key="rate_u",
            )

            top_n_u = st.slider("# vendors in bar chart", 10, 50, 20, key="topn_u")

    # ── Filter ────────────────────────────────────────────────────────────────
    df_u = df_utvs_raw.copy()
    if active_only_u:
        df_u = df_u[df_u["Active"]]
    df_u = df_u[
        df_u["Region"].isin(sel_regions_u) &
        df_u["UTV Type"].isin(sel_types) &
        df_u["Make"].isin(sel_makes) &
        df_u["Dispatch Center"].isin(sel_dcs_u)
    ]
    for _c in ["Daily Rate", "Weekly Rate", "Monthly Rate",
               "Delivery Single", "Delivery Multiple", "Quantity"]:
        df_u[_c] = pd.to_numeric(df_u[_c], errors="coerce")

    # ── KPIs ──────────────────────────────────────────────────────────────────
    k1, k2, k3, k4, k5 = st.columns(5)
    kpi(k1, "Vendors", f"{df_u['Vendor'].nunique():,}", f"of {df_utvs_raw['Vendor'].nunique()} total")
    kpi(k2, "Total Units", f"{int(safe_sum(df_u['Quantity'])):,}", "available")
    kpi(k3, "Dispatch Centers", df_u["Dispatch Center"].nunique(), "unique centers")
    kpi(k4, "Avg Daily Rate", f"${safe_mean(df_u['Daily Rate']):.0f}", "per day")
    kpi(k5, "Avg Weekly Rate", f"${safe_mean(df_u['Weekly Rate']):,.0f}", "per week")
    st.markdown("<br>", unsafe_allow_html=True)

    # ══ SECTION 1: Fleet Composition ══════════════════════════════════════════
    st.markdown('<div class="section-header">🚗 Fleet Composition – What vehicles are vendors using?</div>', unsafe_allow_html=True)

    col_fleet, col_pie = st.columns([3, 2])

    with col_fleet:
        st.subheader("Fleet Breakdown: Make & Quantity by UTV Type")
        fleet_df = (
            df_u.groupby(["UTV Type", "Make"])["Quantity"]
            .sum()
            .reset_index()
            .sort_values(["UTV Type", "Quantity"], ascending=[True, False])
        )
        # Merge "Honda" and "Honda " entries
        fleet_df["Make"] = fleet_df["Make"].str.strip()
        fleet_df = fleet_df.groupby(["UTV Type", "Make"])["Quantity"].sum().reset_index()

        fig_fleet = px.bar(
            fleet_df,
            x="Make",
            y="Quantity",
            color="UTV Type",
            barmode="group",
            template=PLOTLY_TEMPLATE,
            color_discrete_sequence=["#b91c1c", "#f97316", "#3b82f6"],
            text_auto=True,
        )
        fig_fleet.update_layout(
            height=400,
            margin=dict(l=0, r=0, t=10, b=10),
            xaxis=dict(title="Make", tickangle=-35),
            yaxis=dict(title="Total Units"),
            legend=dict(orientation="h", yanchor="bottom", y=1.01),
        )
        fig_fleet.update_traces(textposition="outside")
        st.plotly_chart(fig_fleet, use_container_width=True)

    with col_pie:
        st.subheader("Fleet Mix by UTV Type")
        type_df = df_u.groupby("UTV Type")["Quantity"].sum().reset_index()
        fig_pie = px.pie(
            type_df, names="UTV Type", values="Quantity",
            color_discrete_sequence=["#b91c1c", "#f97316", "#3b82f6"],
            template=PLOTLY_TEMPLATE,
            hole=0.45,
        )
        fig_pie.update_layout(
            height=400, margin=dict(l=0, r=0, t=10, b=10),
            legend=dict(orientation="h", yanchor="top", y=-0.05),
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    # ── Model breakdown table ──────────────────────────────────────────────────
    with st.expander("🔍 Detailed Model Breakdown by UTV Type"):
        model_df = (
            df_u.groupby(["UTV Type", "Make", "Model"])["Quantity"]
            .sum()
            .reset_index()
            .sort_values(["UTV Type", "Quantity"], ascending=[True, False])
        )
        model_df["Quantity"] = model_df["Quantity"].fillna(0).astype(int)
        st.dataframe(model_df, use_container_width=True, height=350)

    # ══ SECTION 2: Geographic Coverage ════════════════════════════════════════
    st.markdown('<div class="section-header">📍 Geographic Coverage – Where are the vehicles located?</div>', unsafe_allow_html=True)

    col_map2, col_state_bar2 = st.columns([3, 2])

    with col_map2:
        st.subheader("Fleet Units by State (Equipment Location)")
        state_u = (
            df_u.groupby("Eq State")["Quantity"]
            .sum()
            .dropna()
            .reset_index()
            .rename(columns={"Eq State": "State", "Quantity": "Total Units"})
        )
        fig_map2 = px.choropleth(
            state_u,
            locations="State",
            locationmode="USA-states",
            color="Total Units",
            scope="usa",
            color_continuous_scale=["#fef2f2", "#b91c1c"],
            hover_data={"Total Units": True},
            template=PLOTLY_TEMPLATE,
        )
        fig_map2.update_layout(
            height=380, margin=dict(l=0, r=0, t=10, b=10),
            coloraxis_colorbar=dict(title="Units"),
        )
        st.plotly_chart(fig_map2, use_container_width=True)

    with col_state_bar2:
        st.subheader("Vendor Count per State (ranked)")
        vendor_state_df = (
            df_u.groupby("Eq State")["Vendor"]
            .nunique()
            .reset_index()
            .rename(columns={"Eq State": "State", "Vendor": "Vendors"})
            .sort_values("Vendors", ascending=True)
        )
        fig_state_bar2 = px.bar(
            vendor_state_df, x="Vendors", y="State", orientation="h",
            color="Vendors", color_continuous_scale=["#fecaca", "#b91c1c"],
            template=PLOTLY_TEMPLATE, text_auto=True,
        )
        fig_state_bar2.update_layout(
            height=380, coloraxis_showscale=False,
            margin=dict(l=0, r=20, t=10, b=10),
            yaxis=dict(title=""), xaxis=dict(title="# Vendors"),
        )
        fig_state_bar2.update_traces(textposition="outside")
        st.plotly_chart(fig_state_bar2, use_container_width=True)

    # ══ SECTION 3: Vendor Rankings ════════════════════════════════════════════
    st.markdown('<div class="section-header">🏆 Vendor Rankings by Region & Dispatch Center</div>', unsafe_allow_html=True)

    rank_u = (
        df_u[[
            "Rank", "Vendor", "Region", "Dispatch Center", "UTV Type", "Make", "Model",
            "Quantity", "Daily Rate", "Weekly Rate", "Monthly Rate",
            "Delivery Single", "Delivery Multiple", "Active",
        ]]
        .copy()
        .sort_values(["Region", "Dispatch Center", "Rank"])
        .reset_index(drop=True)
    )
    rank_u["Active"] = rank_u["Active"].map({True: "✅ Active", False: "❌ Inactive"})

    dollar_cols_u = ["Daily Rate", "Weekly Rate", "Monthly Rate", "Delivery Single", "Delivery Multiple"]

    st.caption("⬆️ Click any column header to sort. Default sorted by Region → Dispatch Center → Rank.")
    st.dataframe(
        rank_u.style.format(
            {c: "${:,.0f}" for c in dollar_cols_u}, na_rep="—"
        ).background_gradient(subset=["Daily Rate"], cmap="RdYlGn_r"),
        use_container_width=True,
        height=400,
    )
    st.download_button(
        "⬇️ Download UTV Rankings",
        rank_u.to_csv(index=False),
        file_name="utv_rankings_by_region_dc.csv",
        mime="text/csv",
    )

    # ══ SECTION 4: Rates by Dispatch Center ═══════════════════════════════════
    st.markdown('<div class="section-header">💲 Rates by Dispatch Center – One chart per rate period</div>', unsafe_allow_html=True)

    region_color_map_u = {r: REGION_COLORS.get(r, "#6b7280") for r in df_u["Region"].unique()}

    dc_rate_u = (
        df_u.groupby(["Region", "Dispatch Center"])[["Daily Rate", "Weekly Rate", "Monthly Rate"]]
        .mean()
        .round(0)
        .reset_index()
    )

    rate_periods = {
        "Daily Rate": "📅 Avg Daily Rate by Dispatch Center",
        "Weekly Rate": "📅 Avg Weekly Rate by Dispatch Center",
        "Monthly Rate": "📅 Avg Monthly Rate by Dispatch Center",
    }

    for rate_col, chart_title in rate_periods.items():
        st.subheader(chart_title)
        chart_dc_u = (
            dc_rate_u[["Region", "Dispatch Center", rate_col]]
            .dropna(subset=[rate_col])
            .sort_values(rate_col, ascending=True)
        )
        fig_dc_u = px.bar(
            chart_dc_u,
            x=rate_col,
            y="Dispatch Center",
            color="Region",
            orientation="h",
            template=PLOTLY_TEMPLATE,
            text_auto=".0f",
            color_discrete_map=region_color_map_u,
            hover_data={"Region": True, "Dispatch Center": True, rate_col: ":,.0f"},
        )
        fig_dc_u.update_layout(
            height=max(400, len(chart_dc_u) * 20),
            margin=dict(l=0, r=60, t=10, b=10),
            xaxis=dict(title=f"Avg {rate_col} ($)"),
            yaxis=dict(title=""),
            legend=dict(orientation="v", x=1.01, y=1),
        )
        fig_dc_u.update_traces(textposition="outside")
        st.plotly_chart(fig_dc_u, use_container_width=True)

    # ══ SECTION 5: Rate Analysis ══════════════════════════════════════════════
    st.markdown('<div class="section-header">📊 Rate Analysis by Make & Distribution</div>', unsafe_allow_html=True)

    col_make, col_box2 = st.columns([3, 2])

    with col_make:
        st.subheader(f"{rate_period} by Make")
        make_df = (
            df_u.groupby("Make")[rate_period]
            .agg(["mean", "min", "max", "count"])
            .round(0)
            .reset_index()
            .sort_values("mean")
        )
        make_df.columns = ["Make", "Avg", "Min", "Max", "# Listings"]
        fig_make = go.Figure()
        fig_make.add_trace(go.Bar(
            x=make_df["Avg"], y=make_df["Make"], orientation="h",
            error_x=dict(
                type="data", symmetric=False,
                array=make_df["Max"] - make_df["Avg"],
                arrayminus=make_df["Avg"] - make_df["Min"],
            ),
            marker_color="#b91c1c",
            text=make_df["# Listings"].apply(lambda n: f"{n} listings"),
            textposition="outside",
        ))
        fig_make.update_layout(
            template=PLOTLY_TEMPLATE, height=380,
            margin=dict(l=0, r=120, t=10, b=10),
            xaxis=dict(title=f"Avg {rate_period} ($)"),
            yaxis=dict(title=""),
        )
        st.plotly_chart(fig_make, use_container_width=True)

    with col_box2:
        st.subheader("Equivalent Daily Rate by Period")
        rate_melt = df_u[["Daily Rate", "Weekly Rate", "Monthly Rate"]].copy()
        rate_melt["Weekly Rate (÷7)"] = rate_melt["Weekly Rate"] / 7
        rate_melt["Monthly Rate (÷30)"] = rate_melt["Monthly Rate"] / 30
        box_df = rate_melt[["Daily Rate", "Weekly Rate (÷7)", "Monthly Rate (÷30)"]].melt(
            var_name="Period", value_name="Equiv. Daily Rate"
        ).dropna()
        fig_box2 = px.box(
            box_df, x="Period", y="Equiv. Daily Rate",
            color="Period",
            color_discrete_sequence=["#b91c1c", "#f97316", "#3b82f6"],
            template=PLOTLY_TEMPLATE, points="outliers",
        )
        fig_box2.update_layout(
            height=380, showlegend=False,
            margin=dict(l=0, r=0, t=10, b=10),
            yaxis=dict(title="Equiv. Daily Rate ($)"),
        )
        st.plotly_chart(fig_box2, use_container_width=True)

    st.download_button(
        "⬇️ Download filtered UTV data",
        rank_u.to_csv(index=False),
        file_name="filtered_utvs.csv",
        mime="text/csv",
    )
