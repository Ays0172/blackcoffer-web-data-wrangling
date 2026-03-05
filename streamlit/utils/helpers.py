"""Shared helper utilities for the Blackcoffer Web Data Wrangling Streamlit app."""

import streamlit as st
import pandas as pd
import os
import glob


# ── Paths (relative to streamlit/ folder) ──
ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
OUTPUT_PATH = os.path.join(ROOT_DIR, "Output.xlsx")
INPUT_PATH = os.path.join(ROOT_DIR, "Input.xlsx")
STOPWORDS_DIR = os.path.join(ROOT_DIR, "StopWords-20250702T160606Z-1-001", "StopWords")
MASTER_DICT_DIR = os.path.join(ROOT_DIR, "MasterDictionary-20250702T160600Z-1-001", "MasterDictionary")
POSITIVE_WORDS_FILE = os.path.join(MASTER_DICT_DIR, "positive-words.txt")
NEGATIVE_WORDS_FILE = os.path.join(MASTER_DICT_DIR, "negative-words.txt")


def show_code(code: str, language: str = "python") -> None:
    """Display a code block with syntax highlighting."""
    st.code(code, language=language)


def show_explanation(text: str) -> None:
    """Render an explanation inside a styled info card."""
    st.markdown(
        f'<div class="info-card">{text}</div>',
        unsafe_allow_html=True,
    )


def create_metric_cards(metrics: list) -> None:
    """Render a row of mini metric cards.  metrics: list of (value, label) tuples."""
    cards_html = "".join(
        f'<div class="metric-mini"><div class="value">{val}</div><div class="label">{lbl}</div></div>'
        for val, lbl in metrics
    )
    st.markdown(f'<div class="metric-row">{cards_html}</div>', unsafe_allow_html=True)


@st.cache_data
def load_output_data() -> pd.DataFrame:
    """Load Output.xlsx and return as DataFrame."""
    if os.path.exists(OUTPUT_PATH):
        return pd.read_excel(OUTPUT_PATH)
    return pd.DataFrame()


@st.cache_data
def load_input_data() -> pd.DataFrame:
    """Load Input.xlsx and return as DataFrame."""
    if os.path.exists(INPUT_PATH):
        return pd.read_excel(INPUT_PATH)
    return pd.DataFrame()


@st.cache_data
def load_articles() -> dict:
    """Load all Netclan*.txt article files into a dict mapping URL_ID → {title, text}."""
    articles = {}
    pattern = os.path.join(ROOT_DIR, "Netclan*.txt")
    for filepath in sorted(glob.glob(pattern)):
        url_id = os.path.splitext(os.path.basename(filepath))[0]
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
            title = lines[0].strip() if lines else "(No title)"
            text = "\n".join(line.strip() for line in lines[1:] if line.strip())
        articles[url_id] = {"title": title, "text": text, "path": filepath}
    return articles


@st.cache_data
def load_stopwords() -> dict:
    """Load stopwords from all files in the StopWords folder. Returns dict of {category: set_of_words}."""
    stopwords_by_category = {}
    if not os.path.isdir(STOPWORDS_DIR):
        return stopwords_by_category
    for fname in sorted(os.listdir(STOPWORDS_DIR)):
        if fname.endswith(".txt"):
            category = fname.replace("StopWords_", "").replace(".txt", "")
            fpath = os.path.join(STOPWORDS_DIR, fname)
            words = set()
            with open(fpath, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    word = line.strip().split("|")[0].strip()
                    if word:
                        words.add(word.lower())
            stopwords_by_category[category] = words
    return stopwords_by_category


@st.cache_data
def load_sentiment_words() -> tuple:
    """Load positive and negative word lists. Returns (positive_set, negative_set)."""
    def _load(filepath):
        words = set()
        if os.path.exists(filepath):
            with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    w = line.strip().lower()
                    if w:
                        words.add(w)
        return words

    return _load(POSITIVE_WORDS_FILE), _load(NEGATIVE_WORDS_FILE)


def get_article_count() -> int:
    """Count Netclan*.txt files in root."""
    return len(glob.glob(os.path.join(ROOT_DIR, "Netclan*.txt")))
