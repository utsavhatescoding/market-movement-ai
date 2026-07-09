# TradePulse Nepal

**From customs data to market intelligence.**

TradePulse Nepal is a web-based trade intelligence dashboard that converts Nepal’s monthly Department of Customs data into clear, interactive, and report-ready insights.

Live app: https://tradepulsenepal.streamlit.app

---

## Overview

Nepal publishes detailed import-export data every month, but the raw Excel files are often difficult to interpret directly. TradePulse Nepal makes this public data easier to understand by transforming customs data into dashboards, charts, product rankings, country-level analysis, opportunity signals, policy risks, and downloadable trade briefs.

The dashboard is designed for students, researchers, journalists, businesses, policymakers, trade analysts, and anyone interested in understanding how Nepal’s market is moving.

---

## Key Features

- Market snapshot of imports, exports, trade deficit, and total trade
- Top imported and exported products
- Partner-country trade analysis
- Customs route concentration analysis
- Product-level Opportunity Finder
- Country intelligence dashboard
- Business opportunity and policy risk signals
- Downloadable monthly trade brief in TXT and PDF formats
- About / Methodology section
- Developer profile section
- Excel upload support for Department of Customs workbooks

---

## Data Source

The dashboard uses monthly foreign trade statistics published by the:

**Department of Customs, Government of Nepal**

The source workbook generally includes data on:

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

## Unit Conversion

The original customs workbook reports many monetary values in **Rs. thousands**.

This dashboard converts them into **Rs. billion** for readability.

```text
Rs. billion = source value / 1,000,000
```

---

## Opportunity Finder

The Opportunity Finder is a screening tool that ranks imported products based on:

- Market Size Score
- Revenue Signal Score
- Approximate Duty Signal Score

The final Opportunity Score helps identify products that may deserve deeper research for:

- Import substitution
- Sourcing
- Distribution
- Trade finance
- Domestic production potential
- Policy analysis

The score is not an investment recommendation.

---

## Tech Stack

- Python
- Streamlit
- pandas
- Plotly
- openpyxl
- ReportLab

---

## Project Structure

```text
market-movement-ai/
├── app.py
├── requirements.txt
├── customs.xlsx
├── utsav.jpg
└── README.md
```

---

## Installation

Clone the repository:

```bash
git clone https://github.com/your-username/market-movement-ai.git
cd market-movement-ai
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

The `requirements.txt` file should include:

```text
streamlit
pandas
openpyxl
plotly
reportlab
```

---

## How to Use

1. Open the dashboard.
2. Upload a Department of Customs Excel workbook, or use the default `customs.xlsx`.
3. Explore the dashboard tabs:
   - Overview
   - Products
   - Opportunity Finder
   - Countries
   - Customs Routes
   - About / Methodology
   - Insights
4. Download the monthly trade brief as TXT or PDF.

---

## Limitations

- The dashboard depends on the structure and accuracy of the uploaded customs workbook.
- Sheet names in the workbook must match the expected Department of Customs format.
- Opportunity Scores are screening indicators only.
- The dashboard does not prove business feasibility or profitability.
- Import-substitution potential requires additional data on production capacity, costs, technology, regulation, and demand.
- The tool is for research, business intelligence, and policy discussion, not investment advice.

---

## Developer

**Utsav Phuyal**  
Developer & Researcher, TradePulse Nepal

I am interested in economics, data analytics, public data, trade intelligence, and AI-powered decision tools. TradePulse Nepal was built to make Nepal’s customs and trade data easier to understand and use.

Contact:

```text
Email: your-email@example.com
LinkedIn: linkedin.com/in/your-profile
GitHub: github.com/your-username
```

---

## Suggested Citation

```text
TradePulse Nepal Dashboard. Based on monthly foreign trade statistics published by the Department of Customs, Government of Nepal.
```

---

## License

This project is currently released as an MVP for learning, research, and demonstration purposes.
