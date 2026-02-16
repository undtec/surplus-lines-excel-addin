# Surplus Lines Tax - Excel Add-in Test Template v2.0

This document describes the Excel test workbook structure for the v2.0 add-in with the unified SLAPI function.

## What's New in v2.0

⚡ **Single SLAPI Function**: All operations now use `=SLTAX.SLAPI(calculationType, effectiveDate, stateCode, premiumAmount)`
⚡ **Automatic Fallback**: Historical queries automatically fall back to current rates when data unavailable
⚡ **Matches Google Sheets**: Same function signature across integrations

---

## Workbook Structure

### Sheet 1: "Tax Calculator" - Tax Calculations

| Row | A | B | C | D | E | F |
|-----|---|---|---|---|---|---|
| 1 | **Surplus Lines Tax Calculator v2.0** | | | | | |
| 2 | | | | | | |
| 3 | **State** | **Premium** | **Formula** | **Base Tax** | **Stamping Fee** | |
| 4 | Texas | 10000 | `=SLTAX.SLAPI("Tax", "", A4, B4)` | *(spills)* | *(spills)* | |
| 5 | California | 25000 | `=SLTAX.SLAPI("Tax", "", A5, B5)` | *(spills)* | *(spills)* | |
| 6 | Florida | 15000 | `=SLTAX.SLAPI("Tax", "", A6, B6)` | *(spills)* | *(spills)* | |
| 7 | New York | 50000 | `=SLTAX.SLAPI("Tax", "", A7, B7)` | *(spills)* | *(spills)* | |
| 8 | Illinois | 7500 | `=SLTAX.SLAPI("Tax", "", A8, B8)` | *(spills)* | *(spills)* | |
| 9 | | | | | | |
| 10 | **Historical Tax Calculation (with Fallback)** | | | | | |
| 11 | **State** | **Premium** | **Date** | **Formula** | **Results** | |
| 12 | Texas | 10000 | 2020-01-01 | `=SLTAX.SLAPI("Tax", C12, A12, B12)` | *(2-4 rows)* | |

**Notes:**
- Tax calculations return minimum 2 rows: Base Tax, Stamping Fee
- If historical data unavailable, adds 2 more rows: Notice, Rates From
- Leave space for 4 rows after each formula to accommodate fallback

---

### Sheet 2: "Rate Lookup" - Current & Historical Rates

| Row | A | B | C | D |
|-----|---|---|---|---|
| 1 | **Tax Rate Lookup v2.0** | | | |
| 2 | | | | |
| 3 | **Current Rates** | | | |
| 4 | **State** | **Formula** | **Results** | |
| 5 | Texas | `=SLTAX.SLAPI("Rate", "", A5)` | *(9 rows)* | |
| 6-14 | | | *(spills)* | |
| 15 | | | | |
| 16 | California | `=SLTAX.SLAPI("Rate", "", A16)` | *(9 rows)* | |
| 17-25 | | | *(spills)* | |
| 26 | | | | |
| 27 | **Historical Rates** | | | |
| 28 | **State** | **Date** | **Formula** | **Results** |
| 29 | Iowa | 2024-06-15 | `=SLTAX.SLAPI("Rate", B29, A29)` | *(9-11 rows)* |
| 30-40 | | | | *(spills)* |
| 41 | | | | |
| 42 | **Historical with Fallback** | | | |
| 43 | Texas | 2010-01-01 | `=SLTAX.SLAPI("Rate", B43, A43)` | *(11 rows - includes fallback notice)* |
| 44-54 | | | | *(spills)* |

**Notes:**
- Rate lookups return 9 rows: tax_rate, stamping_fee, filing_fee, service_fee, surcharge, regulatory_fee, fire_marshal_tax, slas_clearinghouse_fee, flat_fee
- If historical fallback occurs, adds 2 more rows: Notice, Rates From (total 11 rows)
- Always leave at least 11 rows after formula for potential fallback

---

### Sheet 3: "Instructions" - Quick Start Guide

| Row | A |
|-----|---|
| 1 | **Surplus Lines Tax Excel Add-in v2.0** |
| 2 | |
| 3 | **Quick Start** |
| 4 | 1. Configure your API key in Settings (click ribbon button) |
| 5 | 2. Use the SLAPI function for all operations |
| 6 | |
| 7 | **Function Syntax** |
| 8 | `=SLTAX.SLAPI(Calculation_Type, Effective_Date, State_Code, Premium_Amount)` |
| 9 | |
| 10 | **Parameters** |
| 11 | • Calculation_Type: "Tax" or "Rate" |
| 12 | • Effective_Date: YYYY-MM-DD format or "" for current |
| 13 | • State_Code: State name (e.g., "Florida") or code (e.g., "FL") |
| 14 | • Premium_Amount: Dollar amount (required for "Tax", ignored for "Rate") |
| 15 | |
| 16 | **Examples** |
| 17 | Calculate Tax: `=SLTAX.SLAPI("Tax", "", "Texas", 10000)` |
| 18 | Get Current Rate: `=SLTAX.SLAPI("Rate", "", "California")` |
| 19 | Get Historical Rate: `=SLTAX.SLAPI("Rate", "2024-06-15", "Iowa")` |
| 20 | |
| 21 | **Return Values** |
| 22 | Tax: Returns 2-4 rows (Base Tax, Stamping Fee, +fallback if needed) |
| 23 | Rate: Returns 9-11 rows (all fee fields, +fallback if needed) |
| 24 | |
| 25 | **Automatic Fallback** |
| 26 | If historical data is unavailable, the function automatically: |
| 27 | • Returns current rates instead of an error |
| 28 | • Adds a "⚠️ Notice" row explaining the fallback |
| 29 | • Adds a "Rates From" row showing "current" |

---

## Migration from v1.x

### Old v1.x Functions → New v2.0

| Old Function | New v2.0 Equivalent |
|--------------|---------------------|
| `=SLTAX.CALCULATE("Texas", 10000)` | `=SLTAX.SLAPI("Tax", "", "Texas", 10000)` |
| `=SLTAX.CALCULATE_DETAILS("CA", 10000)` | `=SLTAX.SLAPI("Tax", "", "CA", 10000)` |
| `=SLTAX.CALCULATE_WITHPREMIUM("FL", 15000)` | `=SLTAX.SLAPI("Tax", "", "FL", 15000)` |
| `=SLTAX.RATE("California")` | `=SLTAX.SLAPI("Rate", "", "California")` |
| `=SLTAX.STATES()` | ❌ Removed (use static list) |
| `=SLTAX.RATES()` | ❌ Removed (too expensive) |
| `=SLTAX.RATES_DETAILS()` | ❌ Removed (too expensive) |
| `=SLTAX.HISTORICALRATE("Iowa", "2024-06-15")` | `=SLTAX.SLAPI("Rate", "2024-06-15", "Iowa")` |
| `=SLTAX.HISTORICALRATE_DETAILS("TX", "2024-01-01")` | `=SLTAX.SLAPI("Rate", "2024-01-01", "TX")` |

---

## Response Row Structure

### Tax Calculation Response (2-4 rows)

**Normal (no fallback):**
```
Base Tax      | 485
Stamping Fee  | 5
```

**With Fallback:**
```
Base Tax      | 485
Stamping Fee  | 5
⚠️ Notice     | No historical data available for 2020-01-01
Rates From    | current
```

### Rate Lookup Response (9-11 rows)

**Normal (no fallback):**
```
tax_rate                  | 4.85%
stamping_fee             | 0.05%
filing_fee               |
service_fee              |
surcharge                |
regulatory_fee           |
fire_marshal_tax         |
slas_clearinghouse_fee   |
flat_fee                 |
```

**With Fallback:**
```
tax_rate                  | 4.85%
stamping_fee             | 0.05%
filing_fee               |
service_fee              |
surcharge                |
regulatory_fee           |
fire_marshal_tax         |
slas_clearinghouse_fee   |
flat_fee                 |
⚠️ Notice                 | No historical data available for 2020-01-01
Rates From               | current
```

---

## Setup Instructions

1. Install the Surplus Lines Tax Excel Add-in v2.0
2. Configure your API key in the Settings panel (click ribbon button)
3. Open this workbook
4. All formulas will automatically calculate

---

## Get API Key

- Sign up at: https://app.surpluslinesapi.com
- 100 free queries included
- $0.38 per query after that

---

**Build:** 2.0.0 | 2026-02-16
