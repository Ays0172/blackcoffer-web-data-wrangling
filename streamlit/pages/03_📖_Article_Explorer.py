"""
Page 3 — Article Explorer
==========================
Browse and read extracted articles with per-article NLP metrics.
"""

import streamlit as st
import sys, os
import pandas as pd
import numpy as np

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.styles import inject_custom_css, hero, step_header
from utils.helpers import (
    show_explanation,
    create_metric_cards,
    load_articles,
    load_output_data,
)

st.set_page_config(page_title="Article Explorer", page_icon="📖", layout="wide")
inject_custom_css()
hero("📖 Article Explorer", "Browse extracted articles · View per-article NLP metrics")

articles = load_articles()
output_df = load_output_data()

if not articles:
    st.warning("⚠️ No Netclan*.txt article files found in the project root.")
    st.stop()

# ── Article count ──
create_metric_cards([
    (str(len(articles)), "Articles Available"),
    (str(sum(len(a["text"].split()) for a in articles.values())), "Total Words"),
])

# ── Article selector ──
st.markdown("### 📋 Select an Article")

article_ids = list(articles.keys())
# Show title preview for each article
article_options = [f"{aid} — {articles[aid]['title'][:70]}…" if len(articles[aid]['title']) > 70
                   else f"{aid} — {articles[aid]['title']}" for aid in article_ids]

selected_idx = st.selectbox(
    "Choose an article:",
    range(len(article_options)),
    format_func=lambda i: article_options[i],
    key="article_selector",
)

selected_id = article_ids[selected_idx]
selected_article = articles[selected_id]

# ── Article display ──
st.divider()

col_main, col_stats = st.columns([3, 1])

with col_main:
    st.markdown(f"""
    <div class="article-card">
        <h4>📄 {selected_article['title']}</h4>
        <p>{selected_article['text'][:3000]}{'…' if len(selected_article['text']) > 3000 else ''}</p>
    </div>
    """, unsafe_allow_html=True)

    if len(selected_article["text"]) > 3000:
        with st.expander("📖 Read full article"):
            st.markdown(selected_article["text"])

with col_stats:
    st.markdown("### 📊 Article Stats")

    text = selected_article["text"]
    words = text.split()
    sentences = [s.strip() for s in text.replace("!", ".").replace("?", ".").split(".") if s.strip()]

    st.metric("Words", f"{len(words):,}")
    st.metric("Sentences", f"{len(sentences):,}")
    st.metric("Avg Word Length", f"{np.mean([len(w) for w in words]):.1f}" if words else "0")
    st.metric("Avg Sentence Len", f"{len(words) / max(len(sentences), 1):.1f}")

    # Try to find this article's row in output data
    if not output_df.empty:
        url_id_col = [c for c in output_df.columns if "URL_ID" in c.upper()]
        if url_id_col:
            match = output_df[output_df[url_id_col[0]].astype(str) == selected_id]
            if not match.empty:
                row = match.iloc[0]
                st.divider()
                st.markdown("**NLP Metrics:**")

                metric_map = {
                    "POSITIVE_SCORE": ("✅ Positive", "{:.0f}"),
                    "NEGATIVE_SCORE": ("❌ Negative", "{:.0f}"),
                    "POLARITY_SCORE": ("🔄 Polarity", "{:.3f}"),
                    "SUBJECTIVITY_SCORE": ("💭 Subjectivity", "{:.3f}"),
                    "FOG_INDEX": ("📚 FOG Index", "{:.1f}"),
                    "WORD_COUNT": ("📝 Word Count", "{:.0f}"),
                }

                for col_key, (label, fmt) in metric_map.items():
                    matching_cols = [c for c in output_df.columns if col_key in c.upper().replace(" ", "_")]
                    if matching_cols:
                        val = row[matching_cols[0]]
                        if pd.notna(val):
                            st.metric(label, fmt.format(val))

# ── Search & filter ──
st.divider()
step_header(0, "Search Articles")
search_query = st.text_input("🔎 Search articles by keyword:", key="article_search", placeholder="Enter keywords…")

if search_query:
    results = []
    for aid, article in articles.items():
        if search_query.lower() in article["title"].lower() or search_query.lower() in article["text"].lower():
            results.append((aid, article["title"]))

    if results:
        st.success(f"Found **{len(results)}** articles matching \"{search_query}\":")
        for aid, title in results:
            st.markdown(f"- **{aid}** — {title}")
    else:
        st.warning(f"No articles found matching \"{search_query}\".")

# ── Footer ──
st.divider()
st.markdown(
    "<center style='color:#64748b; font-size:0.85rem;'>"
    "Built with ❤️ using Streamlit · Article Explorer"
    "</center>",
    unsafe_allow_html=True,
)
