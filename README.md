# Surplus Lines Tax API - Excel Add-in v2.0

Calculate surplus lines taxes for all 50 US states directly in Excel using a single custom function. Automatic fallback for historical data.

## What's New in v2.0

⚡ **Simplified to ONE function**: `SLAPI()` handles both tax calculations and rate lookups
⚡ **Automatic fallback**: When historical data is unavailable, automatically uses current rates with clear notification
⚡ **Cost savings**: No more accidental bulk queries - only pay for what you need ($0.38 per query)
⚡ **Consistent API**: Same function signature across Excel, Google Sheets, n8n, Zapier, Make, and MCP integrations

## Features

✅ **Single Function** - Use `=SLTAX.SLAPI("Tax", "", "Texas", 10000)` for all operations
✅ **Works like native Excel** - Autocomplete, formula bar support, spilling arrays
✅ **All 53 Jurisdictions** - 50 states + DC + Puerto Rico + Virgin Islands
✅ **Historical Rates** - Look up rates from any date with automatic fallback
✅ **Real-time API** - Always current tax rates

---

## The SLAPI Function

Single unified function for tax calculations and rate lookups.

### Syntax

```
=SLTAX.SLAPI(Calculation_Type, Effective_Date, State_Code, Premium_Amount)
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `Calculation_Type` | Text | Yes | "Tax" or "Rate" |
| `Effective_Date` | Text | No | Date in YYYY-MM-DD format, or "" for current rates |
| `State_Code` | Text | Yes | State name (e.g., "Florida") or code (e.g., "FL") |
| `Premium_Amount` | Number | For Tax | Premium amount in dollars (required for "Tax", ignored for "Rate") |

### Examples

**Calculate Tax (Current Rates)**
```
=SLTAX.SLAPI("Tax", "", "Florida", 10000)
```
Returns (2 rows):
```
Base Tax      | 494
Stamping Fee  | 6
```

**Calculate Tax (Historical with Fallback)**
```
=SLTAX.SLAPI("Tax", "2020-01-01", "Texas", 10000)
```
Returns (if historical data not found - 4 rows):
```
Base Tax      | 485
Stamping Fee  | 5
⚠️ Notice     | No historical data available for 2020-01-01
Rates From    | current
```

**Get Current Rates**
```
=SLTAX.SLAPI("Rate", "", "Florida")
```
Returns (9 rows):
```
tax_rate                  | 4.94%
stamping_fee             |
filing_fee               |
service_fee              | 0.06%
surcharge                |
regulatory_fee           |
fire_marshal_tax         |
slas_clearinghouse_fee   |
flat_fee                 |
```

**Get Historical Rates (with Fallback)**
```
=SLTAX.SLAPI("Rate", "2020-01-01", "Texas")
```
Returns (if historical data not found - 11 rows):
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

## Installation

### Option 1: Office Store (Coming Soon)

1. Open Excel
2. Go to **Insert** → **Add-ins** → **Get Add-ins**
3. Search for "Surplus Lines Tax"
4. Click **Add**

### Option 2: Sideload (Development/Testing)

1. Download the `manifest.xml` file
2. Open Excel
3. Go to **Insert** → **Add-ins** → **My Add-ins**
4. Click **Upload My Add-in**
5. Select the `manifest.xml` file
6. Click **Upload**

---

## Setup

### Step 1: Get Your API Key

1. Go to [app.surpluslinesapi.com](https://app.surpluslinesapi.com)
2. Sign in or create an account (100 free queries included)
3. Copy your API key from the dashboard

### Step 2: Configure the Add-in

1. Click the **Surplus Lines Tax** button in the ribbon (Home tab)
2. Paste your API key in the Settings panel
3. Click **Save API Key**

### Step 3: Start Using!

Type a formula in any cell:

```
=SLTAX.SLAPI("Tax", "", "Texas", 10000)
```

Result (spills to 2 cells):
```
Base Tax      | 485
Stamping Fee  | 5
```

---

## Usage Examples

### Simple Tax Calculation Spreadsheet

| A | B | C | D | E |
|---|---|---|---|---|
| **State** | **Premium** | **Formula** | **Base Tax** | **Stamping Fee** |
| Texas | 10000 | `=SLTAX.SLAPI("Tax", "", A2, B2)` | (spills) | (spills) |
| California | 25000 | `=SLTAX.SLAPI("Tax", "", A3, B3)` | (spills) | (spills) |

### Rate Lookup Table

| A | B | C |
|---|---|---|
| **State** | **Formula** | **Rates** |
| Florida | `=SLTAX.SLAPI("Rate", "", A2)` | (9 rows spill) |
| Texas | `=SLTAX.SLAPI("Rate", "", A3)` | (9 rows spill) |

### Historical Comparison

| A | B | C | D |
|---|---|---|---|
| **State** | **Date** | **Formula** | **Results** |
| Iowa | 2024-06-15 | `=SLTAX.SLAPI("Rate", B2, A2)` | (9-11 rows) |
| Iowa | 2025-06-15 | `=SLTAX.SLAPI("Rate", B3, A3)` | (9-11 rows) |

---

## Breaking Changes from v1.x

### Removed Functions

All 8 previous functions have been consolidated into the single `SLAPI` function:

- ❌ `SLTAX.CALCULATE` → Use `SLAPI("Tax", "", state, premium)`
- ❌ `SLTAX.CALCULATE_DETAILS` → Use `SLAPI("Tax", "", state, premium)`
- ❌ `SLTAX.CALCULATE_WITHPREMIUM` → Use `SLAPI("Tax", "", state, premium)`
- ❌ `SLTAX.RATE` → Use `SLAPI("Rate", "", state)`
- ❌ `SLTAX.STATES` → Removed (free data, use static list)
- ❌ `SLTAX.RATES` → Removed (bulk query too expensive)
- ❌ `SLTAX.HISTORICALRATE` → Use `SLAPI("Rate", date, state)`
- ❌ `SLTAX.HISTORICALRATE_DETAILS` → Use `SLAPI("Rate", date, state)`
- ❌ `SLTAX.RATES_DETAILS` → Removed (bulk query too expensive)

### Migration Examples

```excel
// Old v1.x - Calculate tax
=SLTAX.CALCULATE("Texas", 10000)

// New v2.0
=SLTAX.SLAPI("Tax", "", "Texas", 10000)

// Old v1.x - Get current rate
=SLTAX.RATE("California")

// New v2.0
=SLTAX.SLAPI("Rate", "", "California")

// Old v1.x - Get historical rate
=SLTAX.HISTORICALRATE("Iowa", "2024-06-15")

// New v2.0
=SLTAX.SLAPI("Rate", "2024-06-15", "Iowa")
```

---

## Development & Testing

See [TESTING.md](TESTING.md) for detailed instructions on development and testing workflows.

### Quick Start

```bash
# Install dependencies
npm install

# Build the add-in
npm run build

# Start development server (localhost:3001)
npm run dev-server

# In another terminal, sideload the add-in
npm run sideload
```

### Manifest Files

- **`manifest.xml`** - Production manifest (for Office Store submission)
- **`manifest.dev.xml`** - Local development (`https://localhost:3001`)
- **`manifest.network.xml`** - Network testing (`https://192.168.0.106:3001`)

---

## Support

- **Documentation**: https://surpluslinesapi.com/excel/
- **API Dashboard**: https://app.surpluslinesapi.com
- **Support**: support@undtec.com

---

## Pricing

- **100 free queries** included with new accounts
- **$0.38 per query** after free tier
- No monthly fees, only pay for what you use

Get your API key: [app.surpluslinesapi.com](https://app.surpluslinesapi.com)

---

**Version**: 2.0.0
**Last Updated**: 2026-02-16
**© Underwriters Technologies** - https://undtec.com
