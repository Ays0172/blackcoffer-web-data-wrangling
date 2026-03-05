# 🕸️ Blackcoffer Web Data Wrangling & NLP Analysis

An end-to-end **web scraping** and **Natural Language Processing** pipeline that extracts 100+ articles from Blackcoffer Insights, performs sentiment analysis and readability scoring, and presents everything through a polished **Streamlit** dashboard.

---

## 🚀 Quick Start

```bash
# 1. Clone the repo
git clone https://github.com/<your-username>/blackcoffer-web-data-wrangling.git
cd blackcoffer-web-data-wrangling

# 2. Install dependencies
pip install -r streamlit/requirements.txt

# 3. Launch the dashboard
streamlit run streamlit/app.py
```

The app opens at **http://localhost:8501** with a sidebar to navigate between pages.

---

## 📂 Project Structure

```
blackcoffer-web-data-wrangling/
│
├── Codes (.py files)/
│   ├── Data_Extraction.py          # Web scraper — extracts articles from URLs
│   └── Data_Analysis.py            # NLP engine — sentiment, readability, metrics
│
├── MasterDictionary-…/
│   └── MasterDictionary/
│       ├── positive-words.txt      # Positive sentiment lexicon
│       └── negative-words.txt      # Negative sentiment lexicon
│
├── StopWords-…/
│   └── StopWords/
│       ├── StopWords_Auditor.txt
│       ├── StopWords_Currencies.txt
│       ├── StopWords_DatesandNumbers.txt
│       ├── StopWords_Generic.txt
│       ├── StopWords_GenericLong.txt
│       ├── StopWords_Geographic.txt
│       └── StopWords_Names.txt
│
├── Netclan*.txt                    # 100+ extracted article text files
├── Input.xlsx                      # Source URLs (URL_ID + URL)
├── Output.xlsx                     # Computed NLP metrics per article
├── Output Data Structure.xlsx      # Output schema reference
│
├── streamlit/                      # 🌟 Interactive Streamlit Dashboard
│   ├── app.py                      # Home — hero section, pipeline overview
│   ├── requirements.txt            # Python dependencies
│   ├── utils/
│   │   ├── styles.py               # Custom CSS theming (teal/cyan dark UI)
│   │   └── helpers.py              # Data loaders, display helpers
│   └── pages/
│       ├── 01_🔍_Data_Extraction.py   # Extraction pipeline walkthrough
│       ├── 02_📊_Text_Analysis.py     # NLP metrics & Plotly visualizations
│       ├── 03_📖_Article_Explorer.py  # Browse articles + per-article stats
│       └── 04_📈_Output_Dashboard.py  # Full data table, charts, export
│
├── Documentation for Blackcoffer Data Extraction & Analysis.docx
├── Text Analysis.docx
├── Objective.docx
├── errors.txt                      # Extraction error log
└── README.md                       # ← You are here
```

---

## 🔬 Pipeline Overview

| Stage | Script | Description |
|-------|--------|-------------|
| **1. Extraction** | `Data_Extraction.py` | Reads URLs from `Input.xlsx`, scrapes each page with `requests` + `BeautifulSoup`, saves article text to individual `.txt` files |
| **2. Analysis** | `Data_Analysis.py` | Loads articles, tokenizes with NLTK, computes 14 NLP metrics (sentiment, readability, complexity), exports to `Output.xlsx` |
| **3. Dashboard** | `streamlit/app.py` | Interactive Streamlit app for exploring the entire pipeline and output data |

---

## 📊 NLP Metrics Computed

| Metric | Description |
|--------|-------------|
| **Positive Score** | Count of words matching positive dictionary |
| **Negative Score** | Count of words matching negative dictionary |
| **Polarity Score** | (Positive − Negative) / (Positive + Negative + ε) |
| **Subjectivity Score** | (Positive + Negative) / (Word Count + ε) |
| **Avg Sentence Length** | Words / Sentences |
| **% Complex Words** | Proportion of words with > 2 syllables |
| **FOG Index** | 0.4 × (Avg Sentence Length + % Complex × 100) |
| **Complex Word Count** | Number of words with > 2 syllables |
| **Word Count** | Total cleaned words (stopwords removed) |
| **Syllable Per Word** | Average syllables across all words |
| **Personal Pronouns** | Count of I, we, my, ours, us (excluding "US") |
| **Avg Word Length** | Average character count per word |

---

## 🌟 Streamlit Dashboard Features

- **Home Page** — Hero banner, pipeline overview, dynamic metrics
- **🔍 Data Extraction** — Step-by-step walkthrough with annotated code blocks
- **📊 Text Analysis** — Stopword/sentiment dictionary explorer, Plotly visualizations (histogram, scatter, correlation heatmap, box plot)
- **📖 Article Explorer** — Browse articles by dropdown, read full text, view per-article NLP stats, search by keyword
- **📈 Output Dashboard** — Full interactive data table, summary statistics, 5 chart types, CSV/Excel export
- **Premium dark UI** — Teal/cyan gradient theme, glassmorphism cards, Inter font

---

## 🛠️ Tech Stack

| Technology | Purpose |
|------------|---------|
| **Python 3.11+** | Core language |
| **Streamlit** | Web app framework |
| **BeautifulSoup** | HTML parsing / web scraping |
| **Requests** | HTTP client for fetching pages |
| **NLTK** | Tokenization (sentence & word) |
| **Pandas** | Data manipulation & Excel I/O |
| **NumPy** | Numerical computations |
| **Plotly** | Interactive visualizations |
| **openpyxl** | Excel file read/write |

---

## 📋 Requirements

```
streamlit>=1.30.0
pandas
numpy
plotly
openpyxl
nltk
```

Install with:
```bash
pip install -r streamlit/requirements.txt
```

---

## 📸 Screenshots

> Launch the app with `streamlit run streamlit/app.py` to see the full experience!

- **Home Page** — Hero section with teal gradient, pipeline cards, and process flow
- **Data Extraction** — Annotated code walkthrough with extraction statistics
- **Text Analysis** — Sentiment dictionaries, readability formulas, and 4-tab Plotly charts
- **Article Explorer** — Full-text reader with per-article NLP metrics sidebar
- **Output Dashboard** — Sortable data table, correlation heatmap, and export buttons

---

## 🤝 Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## 📜 License

This project is open source and available for educational purposes.

---

<p align="center">
  Built with ❤️ for Web Data Wrangling & NLP Analysis
</p>
