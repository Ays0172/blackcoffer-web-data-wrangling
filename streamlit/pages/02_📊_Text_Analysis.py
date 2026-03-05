"""
Page 2 — Text Analysis Dashboard
==================================
NLP metrics exploration: sentiment, readability, and visualizations.
"""

import streamlit as st
import sys, os
import pandas as pd
import numpy as np

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.styles import inject_custom_css, hero, step_header
from utils.helpers import (
    show_code,
    show_explanation,
    create_metric_cards,
    load_output_data,
    load_stopwords,
    load_sentiment_words,
)

st.set_page_config(page_title="Text Analysis", page_icon="📊", layout="wide")
inject_custom_css()
hero("📊 Text Analysis Engine", "Sentiment scoring · Readability metrics · NLP pipeline deep-dive")

try:
    import plotly.express as px
    import plotly.graph_objects as go
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False

# ═══════════════════════════════════════════════════════
# STEP 1 — Stopwords & Dictionaries
# ═══════════════════════════════════════════════════════
step_header(1, "Stopwords & Sentiment Dictionaries")

col1, col2 = st.columns(2)

with col1:
    st.markdown("**Stopword Categories:**")
    stopword_cats = load_stopwords()
    if stopword_cats:
        sw_data = []
        for cat, words in stopword_cats.items():
            sw_data.append({"Category": cat, "Word Count": len(words), "Sample": ", ".join(sorted(words)[:5]) + "…"})
        sw_df = pd.DataFrame(sw_data)
        st.dataframe(sw_df, use_container_width=True, hide_index=True)

        total_sw = sum(len(v) for v in stopword_cats.values())
        st.markdown(f"**Total unique stopwords:** `{total_sw:,}`")
    else:
        st.warning("StopWords folder not found.")

with col2:
    st.markdown("**Sentiment Dictionaries:**")
    pos_words, neg_words = load_sentiment_words()
    create_metric_cards([
        (f"{len(pos_words):,}", "Positive Words"),
        (f"{len(neg_words):,}", "Negative Words"),
    ])

    with st.expander("🔍 Sample Words"):
        tcol1, tcol2 = st.columns(2)
        with tcol1:
            st.markdown("**✅ Positive:**")
            st.write(", ".join(sorted(pos_words)[:20]))
        with tcol2:
            st.markdown("**❌ Negative:**")
            st.write(", ".join(sorted(neg_words)[:20]))

# ═══════════════════════════════════════════════════════
# STEP 2 — Sentiment Analysis
# ═══════════════════════════════════════════════════════
step_header(2, "Sentiment Analysis — Polarity & Subjectivity")

show_explanation(
    "Each article's cleaned words are matched against sentiment dictionaries.<br>"
    "<strong>Polarity</strong> = (Positive − Negative) / (Positive + Negative + ε)<br>"
    "<strong>Subjectivity</strong> = (Positive + Negative) / (Total Words + ε)<br>"
    "Values near 0 indicate neutral; near ±1 indicate strong sentiment."
)

col1, col2 = st.columns(2)

with col1:
    show_code("""# Sentiment Scoring
pos_score = sum(1 for w in words if w.lower() in positive_words)
neg_score = sum(1 for w in words if w.lower() in negative_words)

polarity = (pos_score - neg_score) / ((pos_score + neg_score) + 0.000001)
subjectivity = (pos_score + neg_score) / (word_count + 0.000001)""")

with col2:
    output_df = load_output_data()
    if not output_df.empty:
        # Find sentiment columns
        pol_col = [c for c in output_df.columns if "POLARITY" in c.upper()]
        sub_col = [c for c in output_df.columns if "SUBJECTIVITY" in c.upper()]

        if pol_col and sub_col:
            avg_pol = output_df[pol_col[0]].mean()
            avg_sub = output_df[sub_col[0]].mean()
            create_metric_cards([
                (f"{avg_pol:.3f}", "Avg Polarity"),
                (f"{avg_sub:.3f}", "Avg Subjectivity"),
            ])

# ═══════════════════════════════════════════════════════
# STEP 3 — Readability Metrics
# ═══════════════════════════════════════════════════════
step_header(3, "Readability Metrics — FOG Index")

show_explanation(
    "The <strong>Gunning FOG Index</strong> estimates the years of education needed to understand a text.<br>"
    "<code>FOG Index = 0.4 × (Avg Sentence Length + % Complex Words × 100)</code><br>"
    "Complex words have <strong>&gt;2 syllables</strong>. Higher FOG → harder to read."
)

col1, col2 = st.columns(2)

with col1:
    show_code("""def count_syllables(word):
    word = word.lower()
    if word.endswith(('es', 'ed')) and len(word) > 2:
        word = word[:-2]
    vowels = 'aeiou'
    count, prev_vowel = 0, False
    for ch in word:
        if ch in vowels:
            if not prev_vowel:
                count += 1
            prev_vowel = True
        else:
            prev_vowel = False
    return max(count, 1)

def is_complex(word):
    return count_syllables(word) > 2

# Readability
avg_sentence_length = word_count / sentence_count
fog_index = 0.4 * (avg_sentence_length + pct_complex * 100)""")

with col2:
    if not output_df.empty:
        fog_col = [c for c in output_df.columns if "FOG" in c.upper()]
        asl_col = [c for c in output_df.columns if "AVG_SENTENCE" in c.upper().replace(" ", "_")]
        cw_col = [c for c in output_df.columns if "COMPLEX_WORD_COUNT" in c.upper().replace(" ", "_")]

        metrics = []
        if fog_col:
            metrics.append((f"{output_df[fog_col[0]].mean():.1f}", "Avg FOG Index"))
        if asl_col:
            metrics.append((f"{output_df[asl_col[0]].mean():.1f}", "Avg Sentence Length"))
        if cw_col:
            metrics.append((f"{output_df[cw_col[0]].mean():.0f}", "Avg Complex Words"))
        if metrics:
            create_metric_cards(metrics)

# ═══════════════════════════════════════════════════════
# STEP 4 — Visualizations
# ═══════════════════════════════════════════════════════
step_header(4, "Metric Visualizations")

if not output_df.empty and HAS_PLOTLY:
    # Identify numeric columns for plotting
    numeric_cols = output_df.select_dtypes(include=[np.number]).columns.tolist()

    viz_tab1, viz_tab2, viz_tab3, viz_tab4 = st.tabs([
        "📊 Distributions", "🔵 Scatter", "🔥 Correlation", "📦 Box Plot"
    ])

    with viz_tab1:
        if numeric_cols:
            sel_col = st.selectbox("Select metric:", numeric_cols, key="dist_col")
            fig = px.histogram(
                output_df, x=sel_col, nbins=25,
                title=f"Distribution of {sel_col}",
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
        if len(numeric_cols) >= 2:
            c1, c2 = st.columns(2)
            x_col = c1.selectbox("X-axis:", numeric_cols, index=0, key="scatter_x")
            y_col = c2.selectbox("Y-axis:", numeric_cols, index=min(1, len(numeric_cols) - 1), key="scatter_y")
            fig = px.scatter(
                output_df, x=x_col, y=y_col,
                title=f"{x_col} vs {y_col}",
                color_discrete_sequence=["#5eead4"],
                template="plotly_dark",
                opacity=0.7,
            )
            fig.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(30,30,46,0.8)",
                font_color="#94a3b8",
            )
            st.plotly_chart(fig, use_container_width=True)

    with viz_tab3:
        if len(numeric_cols) >= 3:
            selected_corr = st.multiselect(
                "Select columns for correlation:",
                numeric_cols,
                default=numeric_cols[:min(8, len(numeric_cols))],
                key="corr_cols",
            )
            if len(selected_corr) >= 2:
                corr_matrix = output_df[selected_corr].corr()
                fig = px.imshow(
                    corr_matrix,
                    text_auto=".2f",
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
        if numeric_cols:
            sel_box = st.selectbox("Select metric:", numeric_cols, key="box_col")
            fig = px.box(
                output_df, y=sel_box,
                title=f"Box Plot — {sel_box}",
                color_discrete_sequence=["#14b8a6"],
                template="plotly_dark",
            )
            fig.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(30,30,46,0.8)",
                font_color="#94a3b8",
            )
            st.plotly_chart(fig, use_container_width=True)

elif not HAS_PLOTLY:
    st.warning("📦 Install `plotly` for interactive visualizations: `pip install plotly`")
else:
    st.info("Load Output.xlsx to see visualizations.")

# ── Footer ──
st.divider()
st.markdown(
    "<center style='color:#64748b; font-size:0.85rem;'>"
    "Built with ❤️ using Streamlit · Text Analysis Engine"
    "</center>",
    unsafe_allow_html=True,
)
