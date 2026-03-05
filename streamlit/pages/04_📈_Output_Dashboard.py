"""
Page 4 — Output Dashboard
===========================
Interactive exploration of computed NLP output data with charts and export.
"""

import streamlit as st
import sys, os
import pandas as pd
import numpy as np

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.styles import inject_custom_css, hero, step_header
from utils.helpers import create_metric_cards, load_output_data

st.set_page_config(page_title="Output Dashboard", page_icon="📈", layout="wide")
inject_custom_css()
hero("📈 Output Dashboard", "Interactive exploration of computed NLP metrics · Filter · Visualize · Export")

try:
    import plotly.express as px
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

output_df = load_output_data()

if output_df.empty:
    st.warning("⚠️ `Output.xlsx` not found. Please ensure it exists in the project root.")
    st.stop()

# ═══════════════════════════════════════════════════════
# STEP 1 — Overview
# ═══════════════════════════════════════════════════════
step_header(1, "Data Overview")

create_metric_cards([
    (str(len(output_df)), "Total Articles"),
    (str(len(output_df.columns)), "Total Columns"),
    (str(output_df.select_dtypes(include=[np.number]).shape[1]), "Numeric Columns"),
])

# ═══════════════════════════════════════════════════════
# STEP 2 — Full Data Table
# ═══════════════════════════════════════════════════════
step_header(2, "Full Dataset")

# Column filter
all_cols = output_df.columns.tolist()
selected_cols = st.multiselect(
    "Select columns to display:",
    all_cols,
    default=all_cols,
    key="col_filter",
)

if selected_cols:
    display_df = output_df[selected_cols]
else:
    display_df = output_df

st.dataframe(display_df, use_container_width=True, height=400)

# ═══════════════════════════════════════════════════════
# STEP 3 — Summary Statistics
# ═══════════════════════════════════════════════════════
step_header(3, "Summary Statistics")

numeric_df = output_df.select_dtypes(include=[np.number])

if not numeric_df.empty:
    stats = numeric_df.describe().T
    stats["median"] = numeric_df.median()
    stats = stats[["count", "mean", "median", "std", "min", "25%", "50%", "75%", "max"]]
    stats = stats.round(3)
    st.dataframe(stats, use_container_width=True, height=350)
else:
    st.info("No numeric columns found in the output data.")

# ═══════════════════════════════════════════════════════
# STEP 4 — Interactive Visualizations
# ═══════════════════════════════════════════════════════
step_header(4, "Interactive Visualizations")

if HAS_PLOTLY and not numeric_df.empty:
    numeric_cols = numeric_df.columns.tolist()

    viz_tab1, viz_tab2, viz_tab3, viz_tab4, viz_tab5 = st.tabs([
        "📊 Histogram", "🔵 Scatter", "🔥 Heatmap", "📦 Box Plot", "📈 Line"
    ])

    with viz_tab1:
        sel = st.selectbox("Metric:", numeric_cols, key="dash_hist")
        fig = px.histogram(
            output_df, x=sel, nbins=30,
            title=f"Distribution — {sel}",
            color_discrete_sequence=["#14b8a6"],
            template="plotly_dark",
        )
        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(30,30,46,0.8)",
            font_color="#94a3b8",
        )
        st.plotly_chart(fig, use_container_width=True)

    with viz_tab2:
        c1, c2 = st.columns(2)
        x = c1.selectbox("X-axis:", numeric_cols, index=0, key="dash_sx")
        y = c2.selectbox("Y-axis:", numeric_cols, index=min(1, len(numeric_cols) - 1), key="dash_sy")
        fig = px.scatter(
            output_df, x=x, y=y,
            title=f"{x} vs {y}",
            color_discrete_sequence=["#5eead4"],
            template="plotly_dark",
            opacity=0.7,
        )
        fig.update_traces(marker=dict(size=8))
        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(30,30,46,0.8)",
            font_color="#94a3b8",
        )
        st.plotly_chart(fig, use_container_width=True)

    with viz_tab3:
        sel_cols = st.multiselect(
            "Columns:", numeric_cols,
            default=numeric_cols[:min(8, len(numeric_cols))],
            key="dash_corr",
        )
        if len(sel_cols) >= 2:
            corr = output_df[sel_cols].corr()
            fig = px.imshow(
                corr, text_auto=".2f",
                color_continuous_scale="Tealgrn",
                title="Correlation Heatmap",
                template="plotly_dark",
            )
            fig.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(30,30,46,0.8)",
                font_color="#94a3b8",
            )
            st.plotly_chart(fig, use_container_width=True)

    with viz_tab4:
        sel = st.selectbox("Metric:", numeric_cols, key="dash_box")
        fig = px.box(
            output_df, y=sel,
            title=f"Box Plot — {sel}",
            color_discrete_sequence=["#14b8a6"],
            template="plotly_dark",
        )
        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(30,30,46,0.8)",
            font_color="#94a3b8",
        )
        st.plotly_chart(fig, use_container_width=True)

    with viz_tab5:
        sel = st.selectbox("Metric:", numeric_cols, key="dash_line")
        fig = px.line(
            output_df.reset_index(), x="index", y=sel,
            title=f"{sel} across Articles",
            color_discrete_sequence=["#14b8a6"],
            template="plotly_dark",
        )
        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(30,30,46,0.8)",
            font_color="#94a3b8",
            xaxis_title="Article Index",
        )
        st.plotly_chart(fig, use_container_width=True)

elif not HAS_PLOTLY:
    st.warning("📦 Install `plotly` for interactive visualizations: `pip install plotly`")

# ═══════════════════════════════════════════════════════
# STEP 5 — Export
# ═══════════════════════════════════════════════════════
step_header(5, "Export Data")

col1, col2 = st.columns(2)

with col1:
    csv = display_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Download as CSV",
        data=csv,
        file_name="blackcoffer_output.csv",
        mime="text/csv",
        use_container_width=True,
    )

with col2:
    from io import BytesIO
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        display_df.to_excel(writer, index=False, sheet_name="Output")
    st.download_button(
        label="📥 Download as Excel",
        data=buffer.getvalue(),
        file_name="blackcoffer_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── Footer ──
st.divider()
st.markdown(
    "<center style='color:#64748b; font-size:0.85rem;'>"
    "Built with ❤️ using Streamlit · Output Dashboard"
    "</center>",
    unsafe_allow_html=True,
)
