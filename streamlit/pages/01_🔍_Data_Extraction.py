"""
Page 1 — Data Extraction Pipeline
===================================
Walkthrough of the web scraping process with annotated code, input data viewer, and error handling.
"""

import streamlit as st
import sys, os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.styles import inject_custom_css, hero, step_header
from utils.helpers import (
    show_code,
    show_explanation,
    create_metric_cards,
    load_input_data,
    get_article_count,
)

st.set_page_config(page_title="Data Extraction", page_icon="🔍", layout="wide")
inject_custom_css()
hero("🔍 Data Extraction Pipeline", "Automated web scraping of Blackcoffer Insights articles")

# ═══════════════════════════════════════════════════════
# STEP 1 — Input Data
# ═══════════════════════════════════════════════════════
step_header(1, "Input Data — URL Source")

input_df = load_input_data()

if not input_df.empty:
    create_metric_cards([
        (str(len(input_df)), "Total URLs"),
        (str(len(input_df.columns)), "Columns"),
        (str(get_article_count()), "Successfully Extracted"),
    ])

    st.markdown("**Source URLs from `Input.xlsx`:**")
    st.dataframe(input_df, use_container_width=True, height=350)
else:
    st.warning("⚠️ `Input.xlsx` not found. Please ensure it exists in the project root.")

# ═══════════════════════════════════════════════════════
# STEP 2 — Extraction Pipeline
# ═══════════════════════════════════════════════════════
step_header(2, "Extraction Pipeline Code")

show_explanation(
    "The extraction script reads each URL from <code>Input.xlsx</code>, sends an HTTP GET request, "
    "parses the HTML with <strong>BeautifulSoup</strong>, and extracts the title + article body text."
)

tab1, tab2, tab3 = st.tabs(["📥 URL Reading", "🔎 HTML Parsing", "💾 File Saving"])

with tab1:
    st.markdown("**Read URLs from Excel and iterate:**")
    show_code("""import pandas as pd
from tqdm import tqdm

input_df = pd.read_excel('Input.xlsx')

for i, row in tqdm(input_df.iterrows(), total=len(input_df)):
    url_id = str(row['URL_ID'])
    url = row['URL']
    title, text = extract_title_and_text(url)""")

    st.markdown("""
    <div class="code-annotation">
        📌 Uses <strong>tqdm</strong> for progress tracking across all URLs.
        Each article is identified by its <code>URL_ID</code>.
    </div>
    """, unsafe_allow_html=True)

with tab2:
    st.markdown("**Extract title and article text from HTML:**")
    show_code("""def extract_title_and_text(url):
    response = requests.get(url, timeout=15)
    response.raise_for_status()
    soup = BeautifulSoup(response.content, 'html.parser')

    # Get title — try <title>, then <h1>, then empty
    if soup.title and soup.title.string:
        title = soup.title.string.strip()
    elif soup.find('h1'):
        title = soup.find('h1').get_text().strip()
    else:
        title = ''

    # Get article text — prefer <article> tag, fallback to all <p>
    article_tag = soup.find('article')
    if article_tag:
        paragraphs = article_tag.find_all('p')
    else:
        paragraphs = soup.find_all('p')

    article_text = '\\n'.join([p.get_text() for p in paragraphs])
    return title, article_text.strip()""")

    st.markdown("""
    <div class="code-annotation">
        🔎 Uses a <strong>fallback strategy</strong>: tries <code>&lt;article&gt;</code> tag first,
        then falls back to all <code>&lt;p&gt;</code> tags. Title extraction also has multiple fallbacks.
    </div>
    """, unsafe_allow_html=True)

with tab3:
    st.markdown("**Save extracted text to individual files:**")
    show_code("""# Clean and save
clean_title = title.strip()
clean_lines = [line.strip() for line in text.split('\\n')]
clean_text = '\\n'.join([line for line in clean_lines if line])

filename = f"articles/{url_id}.txt"
with open(filename, 'w', encoding='utf-8') as f:
    f.write(clean_title + '\\n' + clean_text)""")

    st.markdown("""
    <div class="code-annotation">
        💾 Each article is saved as <code>&lt;URL_ID&gt;.txt</code> with the title on the first line
        followed by the cleaned body text. Empty lines are removed.
    </div>
    """, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════
# STEP 3 — Error Handling
# ═══════════════════════════════════════════════════════
step_header(3, "Error Handling & Logging")

show_explanation(
    "The pipeline categorizes errors into <strong>Timeout</strong>, <strong>HTTP</strong>, "
    "and <strong>Parsing/Other</strong> types. All errors are logged to <code>errors.txt</code> "
    "with the URL_ID, URL, error type, and full error message."
)

col1, col2 = st.columns(2)

with col1:
    st.markdown("**Error Categories:**")
    show_code("""# Three error categories:
# 1. TimeoutError  — server didn't respond within 15s
# 2. HTTPError     — 404, 500, or other bad status codes
# 3. ParsingOrOther — HTML parsing failures, encoding issues

except requests.exceptions.Timeout as e:
    raise Exception("TimeoutError: " + str(e))
except requests.exceptions.HTTPError as e1:
    raise Exception("HTTPError: " + str(e1))
except Exception as e2:
    raise Exception("ParsingOrOtherError: " + str(e2))""")

with col2:
    st.markdown("**Error Log Format:**")
    show_code("""# errors.txt format:
# URL_ID  URL  Error_Type  Error_Message

log_line = f"{url_id}\\t{url}\\t{error_type}\\t{error_message}\\n"
with open('errors.txt', 'a', encoding='utf-8') as ef:
    ef.write(log_line)""")

    # Show actual errors file if present
    errors_path = os.path.join(os.path.dirname(__file__), "..", "..", "errors.txt")
    if os.path.exists(errors_path):
        with open(errors_path, "r", encoding="utf-8") as f:
            error_content = f.read().strip()
        if error_content and len(error_content.split("\n")) > 1:
            st.warning(f"⚠️ Errors logged during extraction:")
            st.text(error_content)
        else:
            st.success("✅ No errors recorded during extraction!")

# ═══════════════════════════════════════════════════════
# STEP 4 — Results Summary
# ═══════════════════════════════════════════════════════
step_header(4, "Extraction Results")

article_count = get_article_count()
total_urls = len(input_df) if not input_df.empty else 0

if total_urls > 0:
    success_rate = (article_count / total_urls) * 100
    create_metric_cards([
        (str(total_urls), "Total URLs"),
        (str(article_count), "Articles Extracted"),
        (f"{success_rate:.0f}%", "Success Rate"),
    ])
else:
    st.info("Load Input.xlsx to see extraction statistics.")

st.info("👉 Head to the **📖 Article Explorer** page to browse the extracted articles!")

# ── Footer ──
st.divider()
st.markdown(
    "<center style='color:#64748b; font-size:0.85rem;'>"
    "Built with ❤️ using Streamlit · Data Extraction Pipeline"
    "</center>",
    unsafe_allow_html=True,
)
