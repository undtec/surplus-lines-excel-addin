# Surplus Lines Tax - Excel Add-in Test Template

This document describes the Excel test workbook structure that matches the Google Sheets test template.

## Workbook Structure

### Sheet 1: "Calculator" - Tax Calculations

| Row | A | B | C | D | E |
|-----|---|---|---|---|---|
| 1 | **Surplus Lines Tax Calculator** | | | | |
| 2 | | | | | |
| 3 | **State** | **Premium** | **Total Tax** | | |
| 4 | Texas | 10000 | `=SLTAX.CALCULATE(A4, B4)` | | |
| 5 | California | 25000 | `=SLTAX.CALCULATE(A5, B5)` | | |
| 6 | Florida | 15000 | `=SLTAX.CALCULATE(A6, B6)` | | |
| 7 | New York | 50000 | `=SLTAX.CALCULATE(A7, B7)` | | |
| 8 | Illinois | 7500 | `=SLTAX.CALCULATE(A8, B8)` | | |
| 9 | | | | | |
| 10 | **Detailed Breakdown** | | | | |
| 11 | **State** | **Premium** | **Tax** | **Total Due** | |
| 12 | Texas | 10000 | `=SLTAX.CALCULATE_DETAILS(A12, B12)` *(spills to D12)* | | |
| 13 | | | | | |
| 14 | **Compact View** | | | | |
| 15 | **State** | **Premium** | **Tax** | **Total Due** | |
| 16 | California | `=SLTAX.CALCULATE_WITHPREMIUM("California", 10000)` *(spills to D16)* | | | |

---

### Sheet 2: "Rates" - Tax Rates Lookup

| Row | A | B | C |
|-----|---|---|---|
| 1 | **Tax Rate Lookup** | | |
| 2 | | | |
| 3 | **State** | **Rate (%)** | |
| 4 | Texas | `=SLTAX.RATE("Texas")` | |
| 5 | California | `=SLTAX.RATE("California")` | |
| 6 | Florida | `=SLTAX.RATE("Florida")` | |
| 7 | New York | `=SLTAX.RATE("New York")` | |
| 8 | | | |
| 9 | **All Rates** | | |
| 10 | `=SLTAX.RATES()` *(spills 53 rows × 2 columns)* | | |

---

### Sheet 3: "All States" - Complete State List

| Row | A | B |
|-----|---|---|
| 1 | **All 53 Jurisdictions** | |
| 2 | `=SLTAX.STATES()` *(spills 53 rows)* | |

---

### Sheet 4: "Detailed Rates" - Full Rate Breakdown

| Row | A | B | C | D | E | F | G | H | I | J | K |
|-----|---|---|---|---|---|---|---|---|---|---|---|
| 1 | **Complete Rate Details** | | | | | | | | | | |
| 2 | **State** | **Tax Rate** | **Stamping Fee** | **Filing Fee** | **Service Fee** | **Surcharge** | **Regulatory Fee** | **Fire Marshal** | **SLAS Fee** | **Flat Fee** | **Source** |
| 3 | `=SLTAX.RATES_DETAILS()` *(spills 53 rows × 11 columns)* | | | | | | | | | | |

---

### Sheet 5: "Historical" - Historical Rate Lookup

| Row | A | B | C | D |
|-----|---|---|---|---|
| 1 | **Historical Rate Lookup** | | | |
| 2 | | | | |
| 3 | **State** | **Date** | **Rate (%)** | |
| 4 | Iowa | 2025-06-15 | `=SLTAX.HISTORICALRATE(A4, B4)` | |
| 5 | Texas | 2024-01-01 | `=SLTAX.HISTORICALRATE(A5, B5)` | |
| 6 | California | 2023-07-01 | `=SLTAX.HISTORICALRATE(A6, B6)` | |
| 7 | | | | |
| 8 | **Detailed Historical Info** | | | |
| 9 | State: Texas, Date: 2024-01-01 | | | |
| 10 | `=SLTAX.HISTORICALRATE_DETAILS("Texas", "2024-01-01")` *(spills 15 columns)* | | | |
| 11 | | | | |
| 12 | **Vertical View (multiline=TRUE)** | | | |
| 13 | `=SLTAX.HISTORICALRATE_DETAILS("Texas", "2024-01-01", TRUE)` *(spills 15 rows)* | | | |

---

## Function Reference (Excel vs Google Sheets)

| Excel Function | Google Sheets Function | Description |
|----------------|------------------------|-------------|
| `=SLTAX.CALCULATE(state, premium)` | `=CALCULATE_TAX(state, premium)` | Returns total tax amount |
| `=SLTAX.CALCULATE_DETAILS(state, premium, [multiline])` | `=CALCULATE_TAX_DETAILS(state, premium, [multiline])` | Returns [state, premium, tax, due] |
| `=SLTAX.CALCULATE_WITHPREMIUM(state, premium)` | `=CALCULATE_WITH_PREMIUM(state, premium)` | Returns [premium, tax, due] |
| `=SLTAX.RATE(state)` | `=GET_TAX_RATE(state)` | Returns tax rate percentage |
| `=SLTAX.STATES()` | `=GET_STATES()` | Lists all 53 jurisdictions |
| `=SLTAX.RATES()` | `=GET_RATES()` | Returns [state, rate] × 53 |
| `=SLTAX.RATES_DETAILS()` | `=GET_RATES_DETAILS()` | Returns 11 columns × 53 rows |
| `=SLTAX.HISTORICALRATE(state, date)` | `=GET_HISTORICAL_RATE(state, date)` | Returns historical rate |
| `=SLTAX.HISTORICALRATE_DETAILS(state, date, [multiline])` | `=GET_HISTORICAL_RATE_DETAILS(state, date, [multiline])` | Returns 15 columns of historical data |

---

## Setup Instructions

1. Install the Surplus Lines Tax Excel Add-in
2. Configure your API key in the Settings panel
3. Open this workbook
4. All formulas will automatically calculate

---

**Build:** 1.1.0 | 2026-02-02
