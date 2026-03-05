"""
Blackcoffer Web Data Wrangling & NLP Analysis
==============================================
Main entry point for the Streamlit multi-page application.
"""

import streamlit as st
import sys, os

# ── Make utils importable from any page ──
sys.path.insert(0, os.path.dirname(__file__))

from utils.styles import inject_custom_css, hero
from utils.helpers import (
    create_metric_cards,
    get_article_count,
    load_output_data,
    load_sentiment_words,
    load_stopwords,
)

# ── Page config ──
st.set_page_config(
    page_title="Blackcoffer NLP Pipeline",
    page_icon="🕸️",
    layout="wide",
    initial_sidebar_state="expanded",
)

inject_custom_css()

# ── Sidebar branding ──
with st.sidebar:
    st.markdown("## 🕸️ Blackcoffer NLP")
    st.caption("Web Data Wrangling & Text Analysis")
    st.divider()
    st.markdown(
        "**Pipeline**\n"
        "- 🔍 Data Extraction\n"
        "- 📊 Text Analysis\n"
        "- 📖 Article Explorer\n"
        "- 📈 Output Dashboard\n"
    )
    st.divider()
    st.caption("Use the sidebar pages ↑ to navigate.")

# ── Hero section ──
hero(
    "🕸️ Web Data Wrangling & NLP Analysis",
    "Automated web scraping · Sentiment analysis · Readability metrics · Interactive dashboards",
)

# ── Overview metrics ──
article_count = get_article_count()
output_df = load_output_data()
pos_words, neg_words = load_sentiment_words()
stopword_cats = load_stopwords()
total_stopwords = sum(len(v) for v in stopword_cats.values())

create_metric_cards([
    (str(article_count), "Articles Scraped"),
    (str(len(output_df.columns) - 2) if len(output_df.columns) > 2 else "14", "NLP Metrics"),
    (f"{len(pos_words) + len(neg_words):,}", "Sentiment Words"),
    (str(len(stopword_cats)), "Stopword Files"),
])

# ── Pipeline overview cards ──
st.markdown("### 🗺️ Analysis Pipeline")

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="module-card">
        <h3>🔍 Data Extraction</h3>
        <p>Automated web scraping of 100+ articles from Blackcoffer Insights using BeautifulSoup. Extracts titles and article text with robust error handling and logging.</p>
        <span class="badge">BeautifulSoup · Requests · Pandas</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="module-card">
        <h3>📖 Article Explorer</h3>
        <p>Browse extracted articles with full text display. View per-article NLP metrics including sentiment scores, readability, and word statistics.</p>
        <span class="badge">Interactive Reader · Per-Article Stats</span>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="module-card">
        <h3>📊 Text Analysis</h3>
        <p>Comprehensive NLP pipeline computing sentiment polarity, subjectivity, FOG readability index, complex word ratios, syllable counts, and personal pronoun detection.</p>
        <span class="badge">NLTK · Sentiment · Readability</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="module-card">
        <h3>📈 Output Dashboard</h3>
        <p>Interactive exploration of all computed metrics. Correlation heatmaps, distribution charts, filtering, and CSV/Excel export for downstream analysis.</p>
        <span class="badge">Plotly · Filtering · Export</span>
    </div>
    """, unsafe_allow_html=True)

# ── How It Works ──
st.markdown("### ⚙️ How the Pipeline Works")

st.markdown("""
<div class="module-card">
    <h3>1️⃣ Scrape → 2️⃣ Clean → 3️⃣ Analyze → 4️⃣ Visualize</h3>
    <p>
        <strong>Scrape:</strong> Read URLs from <code>Input.xlsx</code> and extract article content via HTTP requests.<br>
        <strong>Clean:</strong> Strip HTML, normalize whitespace, separate titles from body text.<br>
        <strong>Analyze:</strong> Tokenize, remove stopwords, compute sentiment & readability scores.<br>
        <strong>Visualize:</strong> Generate interactive charts and export results to <code>Output.xlsx</code>.
    </p>
</div>
""", unsafe_allow_html=True)

# ── Footer ──
st.divider()
st.markdown(
    "<center style='color:#64748b; font-size:0.85rem;'>"
    "Built with ❤️ using Streamlit · Blackcoffer Web Data Wrangling & NLP Analysis"
    "</center>",
    unsafe_allow_html=True,
)
