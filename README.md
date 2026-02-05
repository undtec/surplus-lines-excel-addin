# Surplus Lines Tax API - Excel Add-in

Calculate surplus lines taxes for all 50 US states directly in Excel using custom functions.

## Features

✅ **Custom Functions** - Use formulas like `=SLTAX.CALCULATE("Texas", 10000)`
✅ **Works like native Excel** - Autocomplete, formula bar support
✅ **All 53 Jurisdictions** - 50 states + DC + Puerto Rico + Virgin Islands
✅ **Historical Rates** - Look up rates from any date
✅ **Real-time API** - Always current tax rates

---

## Available Functions (9)

| Function | Description | Example | Returns |
|----------|-------------|---------|---------|
| `SLTAX.CALCULATE(state, premium)` | Calculate total tax | `=SLTAX.CALCULATE("Texas", 10000)` | 503 |
| `SLTAX.CALCULATE_DETAILS(state, premium, [multiline])` | Full breakdown | `=SLTAX.CALCULATE_DETAILS("CA", 10000)` | [state, premium, tax, due] |
| `SLTAX.CALCULATE_WITHPREMIUM(state, premium)` | Compact breakdown | `=SLTAX.CALCULATE_WITHPREMIUM("FL", 15000)` | [premium, tax, due] |
| `SLTAX.RATE(state)` | Get tax rate % | `=SLTAX.RATE("California")` | 3 |
| `SLTAX.STATES()` | List all 53 jurisdictions | `=SLTAX.STATES()` | Vertical list |
| `SLTAX.RATES()` | All states with rates | `=SLTAX.RATES()` | [state, rate] × 53 |
| `SLTAX.RATES_DETAILS()` | All rates with full fees | `=SLTAX.RATES_DETAILS()` | 11 columns × 53 rows |
| `SLTAX.HISTORICALRATE(state, date)` | Historical rate lookup | `=SLTAX.HISTORICALRATE("Iowa", "2025-06-15")` | 0.95 |
| `SLTAX.HISTORICALRATE_DETAILS(state, date, [multiline])` | Full historical info | `=SLTAX.HISTORICALRATE_DETAILS("TX", "2024-01-01")` | 15 columns |

**Note:** Functions with `[multiline]` parameter accept an optional TRUE/FALSE. When TRUE, returns data vertically (multiple rows, 1 column). Default is FALSE (horizontal).

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
2. Sign in or create an account
3. Copy your API key from the dashboard

### Step 2: Configure the Add-in

1. Click the **Surplus Lines Tax** button in the ribbon (Home tab)
2. Paste your API key in the Settings panel
3. Click **Save API Key**

### Step 3: Start Using!

Type a formula in any cell:

```
=SLTAX.CALCULATE("Texas", 10000)
```

Result: **503**

---

## Usage Examples

### Simple Tax Calculation

| A | B | C |
|---|---|---|
| **State** | **Premium** | **Tax** |
| Texas | 10000 | `=SLTAX.CALCULATE(A2, B2)` |
| California | 25000 | `=SLTAX.CALCULATE(A3, B3)` |

### Detailed Breakdown

```
=SLTAX.CALCULATE_DETAILS("California", 10000)
```

Returns (spills to 4 cells):
| California | 10000 | 318 | 10318 |
|------------|-------|-----|-------|

### Historical Rate Lookup

```
=SLTAX.HISTORICALRATE("Iowa", "2025-06-15")
```

Returns: **0.95** (Iowa's rate during 2025)

### Get All Rates

```
=SLTAX.RATES()
```

Returns a table of all 53 states with their current tax rates.

---

## Pricing

- **100 free calculations** included when you sign up
- **$0.38 per calculation** after free tier
- **$18/month minimum** for active accounts
- **$50 initial deposit** (credited to your balance)
- Add credits at [app.surpluslinesapi.com](https://app.surpluslinesapi.com)

---

## Development

### Prerequisites

- Node.js 18+
- npm or yarn
- Excel 2016+ or Microsoft 365

### Local Development

```bash
# Clone the repository
git clone https://github.com/undtec/surplus-lines-excel-addin.git
cd surplus-lines-excel-addin

# Install dependencies
npm install

# Start dev server
npm run dev-server

# In another terminal, sideload the add-in
npm run start:desktop
```

### Building for Production

```bash
npm run build
```

Output files are in the `dist/` directory.

### Project Structure

```
excel-addin/
├── manifest.xml              # Add-in manifest
├── package.json              # Dependencies & scripts
├── webpack.config.js         # Build configuration
├── src/
│   ├── functions/
│   │   ├── functions.js      # Custom function implementations
│   │   ├── functions.json    # Function metadata
│   │   └── functions.html    # Functions host page
│   ├── taskpane/
│   │   ├── taskpane.html     # Settings UI
│   │   └── taskpane.js       # Taskpane logic
│   └── commands/
│       └── commands.html     # Ribbon commands host
├── assets/                   # Icons and images
└── dist/                     # Build output
```

---

## Support

- **Documentation:** [surpluslinesapi.com/excel/](https://surpluslinesapi.com/excel/)
- **API Dashboard:** [app.surpluslinesapi.com](https://app.surpluslinesapi.com)
- **Email:** support@undtec.com
- **Free Calculator:** [sltax.undtec.com](https://sltax.undtec.com)

---

## Troubleshooting

### "API key not configured"

Open the Settings panel and enter your API key from [app.surpluslinesapi.com](https://app.surpluslinesapi.com).

### Functions not appearing

1. Make sure the add-in is loaded (check Insert → Add-ins)
2. Try restarting Excel
3. Type `=SLTAX.` and check if autocomplete appears

### "Invalid API key"

Verify your API key is correct and your account is active.

### "#VALUE!" error

Check that:
- State name is valid (use full names like "Texas", not "TX")
- Premium is a positive number
- You have API credits remaining

---

## License

MIT License - See [LICENSE](LICENSE) file for details.

---

**Surplus Lines Tax API** is a product of [Underwriters Technologies](https://undtec.com)

Build: 1.1.0 | 2026-02-02
