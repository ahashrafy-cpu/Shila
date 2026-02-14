# 🍕 InsightForge — QFD Dashboard

**Turn customer feedback metrics and reviews into winning strategies.**

A full-featured Streamlit dashboard for Quality Function Deployment (QFD) analysis with bilingual support (English/Persian), ML models, AI insights, and comprehensive export capabilities.

---

## 🚀 Deploy to Streamlit Cloud (Free, 1-Click)

### Step 1: Push to GitHub

```bash
git init
git add .
git commit -m "Initial commit - InsightForge Dashboard"
git remote add origin https://github.com/YOUR_USERNAME/insightforge.git
git branch -M main
git push -u origin main
```

### Step 2: Deploy on Streamlit Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Click **"New app"**
3. Select your GitHub repo → `insightforge`
4. Set **Main file path** to: `app.py`
5. Click **Deploy!**

### Step 3: Add Secrets (Required)

In Streamlit Cloud → Your App → **Settings** → **Secrets**:

```toml
ANTHROPIC_API_KEY = "sk-ant-your-key-here"
DASHBOARD_PASSWORD = "your-secure-password"
```

> Without `ANTHROPIC_API_KEY`, the Claude AI tab will show rule-based insights only.  
> Without `DASHBOARD_PASSWORD`, the default password is `shila2026`.

---

## 📂 Project Structure

```
insightforge/
├── app.py                  # Main Streamlit application (4000+ lines)
├── config.py               # Configuration, labels, column mappings
├── analyzer.py             # Core QFD analysis engine
├── ml_analyzer.py          # Machine learning models
├── ai_insights.py          # AI insights (rule-based + Claude API)
├── Logo.png                # App logo
├── Logo.svg                # App logo (vector)
├── requirements.txt        # Python dependencies
├── packages.txt            # System packages (fonts)
├── .gitignore              # Git ignore rules
├── .streamlit/
│   ├── config.toml         # Streamlit theme & settings
│   └── secrets.toml.example # Template for secrets
└── data/
    └── uploads/            # Pre-loaded data files (optional)
```

---

## ✨ Features

| Tab | Description |
|-----|-------------|
| 📈 Overview | KPIs, rating distribution, NPS, customer segments |
| 📊 Pareto | Issues ranked by rating damage (80/20 rule) |
| 🎨 Kano Model | Must-Be / Performance / Delighter classification |
| 🏪 Branches | Branch comparison, heatmaps, top/bottom performers |
| 🎭 Aspects | Aspect-based sentiment analysis (food, delivery, price...) |
| 📅 Trends | Daily/hourly trends, month-over-month comparison |
| 🤖 AI Insights | Rule-based + Claude AI analysis |
| 🍔 Products | Product performance and category analysis |
| 📝 Text Mining | Word clouds, n-grams, topic discovery, sentiment |
| 🤖 ML | Detractor prediction, clustering, association rules, anomaly detection, churn |

### Additional Features
- 🌐 **Bilingual**: English + Persian (Farsi) with one-click toggle
- 📤 **Export**: Excel (with charts), PowerPoint, Markdown (for NotebookLM)
- 🔒 **Password Protected**: Manager-only access
- 📁 **Multi-file Upload**: Merge multiple daily CSV/Excel files
- 🛵 **SnappFood Support**: Auto-detects and processes SnappFood export format

---

## 🖥️ Run Locally

```bash
# Clone
git clone https://github.com/YOUR_USERNAME/insightforge.git
cd insightforge

# Install
pip install -r requirements.txt

# Set secrets (optional)
export ANTHROPIC_API_KEY="sk-ant-your-key"
export DASHBOARD_PASSWORD="your-password"

# Run
streamlit run app.py
```

---

## 📝 Notes

- **Data Format**: Supports both original Shila CSV format and SnappFood Excel exports
- **ML Models**: Trained on-the-fly from uploaded data (no pre-trained models needed)
- **Font Support**: `packages.txt` installs system fonts for Persian word clouds on Linux
- **Memory**: Streamlit Cloud free tier has 1GB RAM — works well for files up to ~50K rows

---

*Built with Streamlit, Plotly, scikit-learn, and Claude AI*
