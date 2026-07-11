# TradePulse Nepal

**From customs data to market intelligence.**

TradePulse Nepal is a free public-data dashboard that turns Nepal's Department of Customs monthly foreign trade workbooks into clear, interactive, and report-ready trade intelligence.

Live app: https://tradepulsenepal.streamlit.app

---

## What it does

Nepal publishes detailed customs and trade data every month, but the raw Excel files can be difficult to explore directly. TradePulse Nepal cleans, converts, ranks, visualizes, and explains the data so users can quickly understand how Nepal's trade economy is moving.

The dashboard is designed for students, researchers, journalists, businesses, policymakers, trade analysts, and anyone interested in Nepal's economy.

---

## Key features

- Market snapshot of imports, exports, trade deficit, and total trade
- Product-level import and export rankings
- Country and partner-level trade analysis
- Customs route concentration analysis
- Monthly trend dashboard using files inside `monthly_data/`
- Product movement signals across months
- Opportunity Finder for import-substitution and market-screening ideas
- Automated Insights tab with rule-based trade explanations
- Ask TradePulse tab with free rule-based question answering
- Suggested questions for users who do not know what to ask
- Downloadable trade briefs and Ask TradePulse answers
- Data Status section showing latest file and monthly files loaded
- About, Methodology, Limitations, Feedback, and Developer sections

---

## Data source

Primary source:

**Department of Customs, Government of Nepal**

The app is built around monthly foreign trade statistics workbooks. Depending on the official workbook structure, the dataset may include:

- Imports
- Exports
- Trade deficit
- Partner countries
- Commodities
- HS codes
- Customs offices
- Quantities
- Trade values
- Customs revenue

---

## Units and conversion

Many monetary values in the Department of Customs workbook are reported in **Rs. thousands**.

TradePulse Nepal converts those values into **Rs. billion** for easier reading.

```text
Rs. billion = source value / 1,000,000
```

Users should verify important figures with the original Department of Customs workbook before formal citation or publication.

---

## Ask TradePulse

Ask TradePulse is currently a **free rule-based analyst**, not a paid AI API feature.

It answers common questions using the dashboard's processed data, such as:

- What changed in the latest month?
- What is Nepal importing most?
- What is Nepal exporting most?
- Which country dominates imports?
- Which country is the biggest export destination?
- Which customs office handles the most trade?
- What are the main trade risks?
- Where are the business opportunities?
- Give me a short trade brief.
- Give me a media story idea.

This makes the feature safe, transparent, and free to run.

---

## Opportunity Finder

The Opportunity Finder is a screening tool that ranks imported products using three signals:

- **Market Size Score**: based on import value
- **Revenue Signal Score**: based on customs revenue
- **Duty Signal Score**: based on approximate duty/revenue rate

The final Opportunity Score helps identify products that may deserve deeper research for:

- Import substitution
- Sourcing
- Distribution
- Trade finance
- Domestic production potential
- Policy analysis

The score is not an investment recommendation.

---

## Project structure

```text
tradepulse-nepal/
├── app.py
├── requirements.txt
├── README.md
├── customs.xlsx                 # optional latest/default workbook
├── utsav.png                    # optional developer photo
└── monthly_data/
    ├── 01_Shrawan.xlsx
    ├── 02_Bhadra.xlsx
    ├── 03_Asoj.xlsx
    ├── ...
    └── 11_Jestha.xlsx
```

The `monthly_data/` folder is used for the Trends tab. Keep the numeric prefix so files sort in the correct Nepali month order.

Example:

```text
12_Ashadh.xlsx
```

---

## Installation

Clone the repository:

```bash
git clone https://github.com/utsavhatescoding/tradepulse-nepal.git
cd tradepulse-nepal
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the app locally:

```bash
streamlit run app.py
```

---

## Requirements

```text
streamlit
pandas
openpyxl
plotly
reportlab
```

---

## Monthly update workflow

When a new Department of Customs workbook is released:

1. Download the latest Excel file.
2. Rename it using the next month number and month name.
   - Example: `12_Ashadh.xlsx`
3. Add the file to the `monthly_data/` folder.
4. Replace `customs.xlsx` with the latest workbook if you want the main dashboard to open on the newest month.
5. Commit and push to GitHub.
6. Streamlit Cloud redeploys the app automatically.
7. Open the live app and check the Data Status and Trends tabs.

---

## How to use

1. Open the dashboard.
2. Use the default `customs.xlsx` or upload a Department of Customs workbook.
3. Explore the tabs:
   - Overview
   - Products
   - Opportunity Finder
   - Countries
   - Customs Routes
   - Trends
   - About / Methodology
   - Insights
   - Ask TradePulse
4. Download trade briefs or Ask TradePulse answers when needed.

---

## Limitations

- The dashboard depends on the structure and accuracy of the official customs workbook.
- Sheet names and column names must remain close to the expected Department of Customs format.
- Opportunity Scores are screening indicators only.
- The dashboard does not prove profitability, feasibility, or investment potential.
- Import-substitution potential requires additional data on production capacity, costs, demand, regulation, and technology.
- Automated insights are generated only from available dashboard numbers.
- Users should verify figures with the original workbook before formal citation.

---

## Suggested citation

```text
TradePulse Nepal Dashboard. Based on monthly foreign trade statistics published by the Department of Customs, Government of Nepal.
```

---

## Developer

**Utsav Phuyal**  
Developer & Researcher, TradePulse Nepal

I built TradePulse Nepal to make Nepal's public trade data easier to understand and use through dashboards, product-level analysis, country intelligence, opportunity signals, automated briefs, and simple public-data tools.

Contact:

```text
Email: utsavkphuyal@gmail.com
LinkedIn: linkedin.com/in/utsav-phuyal
GitHub: github.com/utsavhatescoding
```

---

## Status

Current version: **Public MVP 0.5**

TradePulse Nepal is currently a free public-data MVP. Feedback, bug reports, and feature suggestions are welcome.
