# Polycab India DCF Financial Model

> **A fully self-contained Python script that generates a professional, multi-sheet Excel DCF financial model for Polycab India Ltd. (NSE: POLYCAB | BSE: 542652).**

---

## 📋 Table of Contents

1. [Overview](#overview)
2. [What the Script Generates](#what-the-script-generates)
3. [Sheet Descriptions](#sheet-descriptions)
4. [Prerequisites & Installation](#prerequisites--installation)
5. [How to Run](#how-to-run)
6. [Customising Assumptions](#customising-assumptions)
7. [Disclaimer](#disclaimer)

---

## Overview

`generate_dcf_model.py` produces a fully formatted Excel workbook (`Polycab_DCF_Model.xlsx`) containing a complete **Discounted Cash Flow (DCF)** financial model for Polycab India Ltd., covering:

- Historical financials (FY2023–FY2025)
- **FCFF-based DCF valuation** (8-year explicit forecast + terminal value)
- **FCFE-based valuation** (equity intrinsic value)
- **Two sensitivity tables** (WACC × Terminal Growth; WACC × Revenue Growth)

The model is built with **`openpyxl`** and requires no other third-party libraries.

---

## What the Script Generates

| Output | Description |
|---|---|
| `Polycab_DCF_Model.xlsx` | Multi-sheet Excel workbook (see below) |
| Console summary | FCFF & FCFE intrinsic values vs CMP printed to terminal |

---

## Sheet Descriptions

### 1. `Assumptions` (Blue tab)
All key model inputs in a **Parameter → Value → Notes** layout, colour-coded for easy editing:
- Company information (name, ticker, CMP, shares outstanding)
- CAPM / WACC inputs: Risk-Free Rate, Beta, ERP, Cost of Equity, Cost of Debt, Tax Rate, capital structure weights → **WACC**
- Revenue growth rates (Phase 1: 22%, Phase 2: 15%), Terminal Growth Rate (5%)
- Margin & reinvestment assumptions: EBITDA margins, D&A %, Capex %, NWC %
- FCFE-specific inputs (constant Capex, D&A schedule, WC schedule)
- Balance sheet inputs (FY2025 Cash & Debt for bridge)

### 2. `Historical Financials` (Green tab)
Reported data for **FY2023, FY2024, FY2025** across four sections:

| Section | Key Metrics |
|---|---|
| Income Statement | Revenue, EBITDA, D&A, EBIT, Interest, PBT, Tax, PAT |
| Cash Flow | CFO, Capex, Free Cash Flow |
| Balance Sheet | Equity, Debt, Cash, Total Assets, Current Assets/Liabilities, NWC |
| Key Ratios | ROE, ROCE, D/E, Current Ratio |

### 3. `DCF Model (FCFF)` (Orange tab)
8-year **Free Cash Flow to Firm** projection (FY26E–FY33E):

```
Revenue → EBITDA (margin %) → D&A → EBIT → NOPAT (after-tax)
FCFF = NOPAT + D&A – Capex – ΔNWC
```

Valuation bridge:
```
Sum of PV(FCFF) + PV(Terminal Value) = Enterprise Value
Enterprise Value + Cash – Debt = Equity Value
Equity Value ÷ Shares Outstanding = Intrinsic Value per Share
```
Shows upside / downside vs CMP (₹8,400).

### 4. `FCFE Model` (Purple tab)
8-year **Free Cash Flow to Equity** projection (FY26E–FY33E):

```
FCFE = Net Income + D&A – Capex – ΔWC + Net Borrowings
```

Discounted at **Cost of Equity (Ke = 12.65%)** using CAPM.  
Terminal value, equity value, and intrinsic value per share calculated.

### 5. `Sensitivity Analysis` (Red tab)
Two colour-coded sensitivity tables computing the FCFF intrinsic value per share across a matrix of assumptions:

| Table | Rows | Columns |
|---|---|---|
| **Table 1** | WACC: 10% – 15% | Terminal Growth: 3% – 7% |
| **Table 2** | WACC: 10% – 15% | Revenue Growth (Phase 1): 18% – 28% |

Base-case cell highlighted in **green**; base-case row/column highlighted in **yellow**.

---

## Prerequisites & Installation

### Requirements
- Python 3.8 or higher
- `openpyxl >= 3.1.0`

### Install dependencies

```bash
pip install -r requirements.txt
```

Or install directly:

```bash
pip install "openpyxl>=3.1.0"
```

---

## How to Run

```bash
# Clone the repository (if you haven't already)
git clone https://github.com/dhwani16271110/polycab-financial-model.git
cd polycab-financial-model

# Install dependencies
pip install -r requirements.txt

# Run the script
python generate_dcf_model.py
```

**Expected output:**

```
Building Sheet 1: Assumptions …
Building Sheet 2: Historical Financials …
Building Sheet 3: DCF Model (FCFF) …
Building Sheet 4: FCFE Model …
Building Sheet 5: Sensitivity Analysis …

✅  Successfully generated: Polycab_DCF_Model.xlsx
    Sheets: Assumptions, Historical Financials, DCF Model (FCFF), FCFE Model, Sensitivity Analysis

    FCFF Intrinsic Value / Share : ₹X,XXX
    FCFE Intrinsic Value / Share : ₹X,XXX
    Current Market Price (CMP)   : ₹8,400
```

The file `Polycab_DCF_Model.xlsx` will be created in the **current working directory**.

---

## Customising Assumptions

All model assumptions are defined as **constants at the top of `generate_dcf_model.py`** (lines ~30–85). You can modify any of these values without touching the rest of the code:

| Constant | Default | Description |
|---|---|---|
| `CMP` | `8400` | Current Market Price (₹) |
| `SHARES_OUTSTANDING` | `150.5` | Shares in Millions |
| `RISK_FREE_RATE` | `0.0670` | India 10Y G-Sec |
| `BETA` | `0.85` | Blended beta estimate |
| `EQUITY_RISK_PREMIUM` | `0.0700` | Damodaran India ERP |
| `TAX_RATE` | `0.2430` | Effective corporate tax rate |
| `REV_GROWTH_PHASE1` | `0.22` | Revenue growth Years 1–5 |
| `REV_GROWTH_PHASE2` | `0.15` | Revenue growth Years 6–8 |
| `TERMINAL_GROWTH` | `0.050` | Long-run terminal growth rate |
| `EBITDA_MARGIN_PHASE1` | `0.140` | EBITDA margin Years 1–5 |
| `EBITDA_MARGIN_PHASE2` | `0.145` | EBITDA margin Years 6–8 |
| `DA_PCT` | `0.0135` | D&A as % of revenue |
| `CAPEX_PCT` | `0.050` | Capex as % of revenue (FCFF) |
| `NWC_PCT` | `0.120` | ΔNWC as % of ΔRevenue |
| `NET_INCOME_GROWTH_PHASE1` | `0.25` | Net Income growth Years 1–5 (FCFE) |
| `NET_INCOME_GROWTH_PHASE2` | `0.15` | Net Income growth Years 6–8 (FCFE) |
| `FCFE_CAPEX` | `1100` | Annual Capex (₹ Crore) for FCFE model |

After editing, simply re-run `python generate_dcf_model.py` to regenerate the workbook.

---

## Disclaimer

> ⚠️ **This model is for educational and research purposes only.**
> It does not constitute investment advice, a solicitation, or a recommendation to buy or sell any securities. All projections, valuations, and estimates are based on publicly available information and analyst consensus data as of March 2026. Actual results may differ materially from projections.
>
> Past performance is not indicative of future results. Please consult a SEBI-registered investment advisor before making any investment decisions. The authors of this repository assume no liability for any investment decisions made based on this model.
>
> Data sources: Polycab India Ltd. annual reports, NSE/BSE filings, analyst reports (Motilal Oswal, Jefferies, Sharekhan), and publicly available financial data portals.
