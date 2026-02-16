# Surplus Lines Tax Excel Add-in - Testing Guide

Comprehensive guide for developing and testing the Excel Add-in across different environments.

---

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Development Environments](#development-environments)
3. [Manifest Files Overview](#manifest-files-overview)
4. [Local Development (localhost)](#local-development-localhost)
5. [Network Testing (LAN)](#network-testing-lan)
6. [Testing Workflows](#testing-workflows)
7. [Debugging Tips](#debugging-tips)
8. [Common Issues](#common-issues)

---

## Prerequisites

### Required Software

- **Node.js 18+** - [Download](https://nodejs.org/)
- **Excel Desktop** (Windows or Mac) or **Excel Online**
- **Git** - For version control
- **Code Editor** - VS Code recommended

### Excel Versions Supported

- ✅ Excel for Windows (Microsoft 365)
- ✅ Excel for Mac (Microsoft 365)
- ✅ Excel Online (Web)
- ⚠️ Excel 2019/2021 (Limited custom functions support)

### Get Your API Key

1. Go to [app.surpluslinesapi.com](https://app.surpluslinesapi.com)
2. Sign in or create an account
3. Copy your API key from the dashboard

---

## Development Environments

### 1. Local Development (Developer Machine)
- URL: `https://localhost:3001`
- Manifest: `manifest.dev.xml`
- Best for: Day-to-day development

### 2. Network Testing (LAN)
- URL: `https://192.168.0.106:3001`
- Manifest: `manifest.network.xml`
- Best for: Testing on multiple devices (iPhone, iPad, other computers)

### 3. Production (Office Store)
- URL: Production CDN
- Manifest: `manifest.xml`
- Best for: Final release

---

## Manifest Files Overview

### manifest.dev.xml (Local Development)

```xml
<DisplayName DefaultValue="Surplus Lines Tax (Dev)"/>
<IconUrl DefaultValue="https://localhost:3001/assets/icon-32.png"/>
<SourceLocation DefaultValue="https://localhost:3001/taskpane.html"/>
```

**Use when:**
- Developing on your local machine
- Making code changes with hot-reload
- Testing new features before network testing

### manifest.network.xml (Network Testing)

```xml
<DisplayName DefaultValue="Surplus Lines Tax"/>
<IconUrl DefaultValue="https://192.168.0.106:3001/assets/icon-32.png"/>
<SourceLocation DefaultValue="https://192.168.0.106:3001/taskpane.html"/>
```

**Use when:**
- Testing on mobile devices (iPhone, iPad)
- Testing on other computers on the same network
- Sharing with colleagues for testing
- Verifying cross-device compatibility

**Important:** Update the IP address (`192.168.0.106`) to match your development machine's local IP

### manifest.xml (Production)

```xml
<DisplayName DefaultValue="Surplus Lines Tax"/>
<IconUrl DefaultValue="https://cdn.surpluslinesapi.com/excel/icon-32.png"/>
<SourceLocation DefaultValue="https://cdn.surpluslinesapi.com/excel/taskpane.html"/>
```

**Use when:**
- Submitting to Office Store
- Production deployment

---

## Local Development (localhost)

### Setup

1. **Clone and Install**

```bash
cd /path/to/excel-addin
npm install
```

2. **Build the Add-in**

```bash
npm run build
```

3. **Start Development Server**

```bash
npm run dev-server
```

Server starts at `https://localhost:3001`

4. **Sideload the Add-in**

In a **new terminal**:

```bash
npm run sideload
```

This will:
- Open Excel
- Prompt you to trust the `manifest.dev.xml`
- Load the add-in automatically

### Alternative Manual Sideload

**Excel for Windows/Mac:**

1. Open Excel
2. Go to **Insert** → **Add-ins** → **My Add-ins**
3. Click **Upload My Add-in** (top right)
4. Select `manifest.dev.xml`
5. Click **Upload**

**Excel Online:**

1. Open Excel Online
2. Go to **Insert** → **Add-ins** → **My Add-ins**
3. Click **Manage My Add-ins** → **Upload My Add-in**
4. Browse and select `manifest.dev.xml`

### Development Workflow

1. **Make code changes** in `src/functions/functions.js` or `src/taskpane/`
2. **Rebuild**: `npm run build`
3. **Refresh Excel**:
   - Close and reopen the workbook
   - Or restart Excel
4. **Test your changes**

### Hot Reload (Optional)

For faster development, use watch mode:

```bash
# Terminal 1: Watch for changes
npm run watch

# Terminal 2: Development server
npm run dev-server
```

Changes to taskpane will hot-reload. Custom functions require Excel restart.

---

## Network Testing (LAN)

Test the add-in on other devices (iPhone, iPad, other computers) on the same network.

### Setup

1. **Find Your Local IP Address**

**Mac:**
```bash
ifconfig | grep "inet " | grep -v 127.0.0.1
```

**Windows:**
```cmd
ipconfig
```

Look for IPv4 address (e.g., `192.168.0.106`)

2. **Update manifest.network.xml**

Replace all instances of `192.168.0.106` with your actual IP address:

```xml
<IconUrl DefaultValue="https://YOUR_IP:3001/assets/icon-32.png"/>
<SourceLocation DefaultValue="https://YOUR_IP:3001/taskpane.html"/>
```

3. **Start Development Server**

```bash
npm run dev-server
```

The server will be accessible at `https://YOUR_IP:3001`

4. **Allow Firewall Access**

Ensure your firewall allows incoming connections on port 3001.

**Mac:**
- System Preferences → Security & Privacy → Firewall Options
- Allow incoming connections for Node.js

**Windows:**
- Windows Defender Firewall → Allow an app
- Add Node.js

### Sideload on Network Devices

#### iPhone/iPad (Excel Mobile)

1. Email yourself the `manifest.network.xml` file
2. Open Excel on your iPhone/iPad
3. Open a workbook
4. Tap **Insert** → **Add-ins**
5. Tap **My Add-ins** → **Upload My Add-in**
6. Select the manifest file from email attachment

#### Other Computers (Same Network)

1. Copy `manifest.network.xml` to the other computer
2. Open Excel
3. Go to **Insert** → **Add-ins** → **My Add-ins**
4. Click **Upload My Add-in**
5. Select the `manifest.network.xml` file

### Trust the Certificate

On first load, you'll see a certificate warning because we're using a self-signed certificate.

**How to Trust:**

1. In browser, navigate to `https://YOUR_IP:3001`
2. Click "Advanced" → "Proceed to YOUR_IP (unsafe)"
3. Now Excel can load the add-in

**Better approach (Mac):**

Add the certificate to your keychain:

```bash
# Generate a self-signed certificate (one-time)
npm run create-cert

# Trust it in Keychain Access
# Open Keychain Access → Certificates
# Find "localhost" certificate → Double-click
# Trust → "Always Trust"
```

---

## Testing Workflows

### Test Cases

#### 1. Tax Calculation (Current Rates)

```excel
=SLTAX.SLAPI("Tax", "", "Texas", 10000)
```

**Expected Result (2 rows):**
```
Base Tax      | 485
Stamping Fee  | 4
```

**Verify:**
- ✅ Result spills to 2 cells
- ✅ Formula bar shows correct syntax
- ✅ Numbers are calculated correctly
- ✅ API call cost: $0.38

#### 2. Tax Calculation (Historical with Data)

```excel
=SLTAX.SLAPI("Tax", "2024-06-15", "Texas", 10000)
```

**Expected Result (2 rows):**
```
Base Tax      | 485
Stamping Fee  | 4
```

**Verify:**
- ✅ Historical data is used (if available)
- ✅ No fallback notice

#### 3. Tax Calculation (Historical Fallback)

```excel
=SLTAX.SLAPI("Tax", "2020-01-01", "Iowa", 15000)
```

**Expected Result (4 rows):**
```
Base Tax          | 138.75
Stamping Fee      | 0
⚠️ Notice          | No historical data available for 2020-01-01
Rates From        | current
```

**Verify:**
- ✅ Fallback notice is displayed
- ✅ Current rates are used
- ✅ "Rates From" shows "current"

#### 4. Rate Lookup (Current)

```excel
=SLTAX.SLAPI("Rate", "", "Florida")
```

**Expected Result (9 rows):**
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

**Verify:**
- ✅ Rate structure spills correctly
- ✅ All fields are present

#### 5. Rate Lookup (Historical)

```excel
=SLTAX.SLAPI("Rate", "2024-06-15", "Texas")
```

**Verify:**
- ✅ Historical rates returned (if available)
- ✅ Or fallback notice with current rates

#### 6. Error Handling

**Invalid state:**
```excel
=SLTAX.SLAPI("Tax", "", "InvalidState", 10000)
```

**Expected:** Error message

**Missing premium:**
```excel
=SLTAX.SLAPI("Tax", "", "Texas")
```

**Expected:** Error message "Premium required for Tax calculation"

**Invalid date format:**
```excel
=SLTAX.SLAPI("Tax", "01/01/2024", "Texas", 10000)
```

**Expected:** Error message about date format

### Performance Testing

Test with multiple formulas:

| A | B | C | D |
|---|---|---|---|
| Texas | 10000 | `=SLTAX.SLAPI("Tax", "", A2, B2)` | (spills) |
| California | 25000 | `=SLTAX.SLAPI("Tax", "", A3, B3)` | (spills) |
| New York | 50000 | `=SLTAX.SLAPI("Tax", "", A4, B4)` | (spills) |

**Verify:**
- ✅ All formulas calculate correctly
- ✅ No rate limiting errors
- ✅ Results appear within 2-3 seconds

---

## Debugging Tips

### Excel Developer Tools

**Enable Developer Tab:**

1. File → Options → Customize Ribbon
2. Check "Developer"
3. Click OK

**Open Console:**

1. Developer → Office Add-ins → My Add-ins
2. Click on your add-in → "..."→ "Debug"
3. Opens Developer Tools (F12)

### Console Logging

In `functions.js`, add logging:

```javascript
console.log("State:", state);
console.log("Premium:", premium);
console.log("API Response:", response);
```

View logs in Developer Tools Console.

### Common Debugging Commands

```javascript
// Check if API key is set
console.log("API Key:", localStorage.getItem("slapi_api_key"));

// Test API connectivity
fetch("https://n8n.undtec.com/webhook/slapi/v1/calculate", {
  method: "POST",
  headers: {
    "X-API-Key": "your_key_here",
    "Content-Type": "application/json"
  },
  body: JSON.stringify({
    state: "Texas",
    premium: 10000
  })
}).then(r => r.json()).then(console.log);
```

### Network Tab

1. Open Developer Tools (F12)
2. Go to **Network** tab
3. Execute a formula
4. Inspect API requests/responses

---

## Common Issues

### Issue: "Add-in won't load"

**Symptoms:**
- Blank taskpane
- Loading spinner forever

**Solutions:**
1. Check dev server is running: `npm run dev-server`
2. Check browser console for errors
3. Clear Excel cache:
   - **Mac:** `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/`
   - **Windows:** `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
4. Restart Excel

### Issue: "Certificate errors"

**Symptoms:**
- "This site is not secure"
- NET::ERR_CERT_AUTHORITY_INVALID

**Solutions:**
1. Navigate to `https://localhost:3001` in browser
2. Accept the certificate warning
3. Go back to Excel and reload

**Mac (Better solution):**
```bash
# Trust the localhost certificate
npm run create-cert
# Open Keychain Access → Always Trust "localhost"
```

### Issue: "Functions not appearing"

**Symptoms:**
- `=SLTAX.SLAPI` shows #NAME? error
- Autocomplete doesn't work

**Solutions:**
1. Verify `functions.json` is properly formatted
2. Check `functions.js` has `@customfunction` JSDoc
3. Rebuild: `npm run build`
4. Restart Excel completely
5. Clear Excel cache

### Issue: "API calls failing"

**Symptoms:**
- Error: "API request failed"
- Network errors in console

**Solutions:**
1. Check API key is configured correctly
2. Check internet connectivity
3. Verify API endpoint: `https://n8n.undtec.com/webhook/slapi/v1/calculate`
4. Test API directly in Postman/curl
5. Check console for detailed error messages

### Issue: "Network testing not working"

**Symptoms:**
- Can't connect from iPhone/iPad
- Other computers can't load add-in

**Solutions:**
1. Verify IP address in `manifest.network.xml` is correct
2. Check firewall allows port 3001
3. Ensure dev server is running: `npm run dev-server`
4. Ping the IP from the test device
5. Try accessing `https://YOUR_IP:3001` in Safari/Chrome first

### Issue: "Hot reload not working"

**Symptoms:**
- Code changes don't appear
- Need to restart Excel every time

**Solutions:**
1. Taskpane changes: Should hot-reload (refresh taskpane)
2. Custom functions: **Always** require Excel restart
3. Manifest changes: Need to re-upload manifest
4. Run `npm run watch` for auto-rebuild

---

## Advanced Testing

### Testing with Templates

Use the test template to verify all scenarios:

```bash
# Open the test template
open templates/SurplusLinesTax-Test-Template.xlsx
```

The template includes pre-built test cases for:
- Current tax calculations
- Historical tax calculations
- Rate lookups
- Error handling

### Automated Testing (Future)

```bash
# Unit tests (when implemented)
npm test

# Integration tests
npm run test:integration
```

---

## Best Practices

### Development

1. ✅ Always run `npm run build` after code changes
2. ✅ Test in both Excel Desktop and Excel Online
3. ✅ Clear cache when troubleshooting
4. ✅ Use console logging liberally
5. ✅ Test error cases, not just happy paths

### Network Testing

1. ✅ Update IP address in `manifest.network.xml` before each session
2. ✅ Test on actual devices (iPhone, iPad, other computers)
3. ✅ Verify certificate trust on each device
4. ✅ Test on different network conditions

### Before Submission

1. ✅ All test cases pass
2. ✅ No console errors
3. ✅ Error handling works correctly
4. ✅ Performance is acceptable (< 3 seconds)
5. ✅ Works in Excel Desktop + Excel Online
6. ✅ Icons and branding correct
7. ✅ Documentation updated

---

## Getting Help

- **Excel Add-in Documentation**: https://learn.microsoft.com/en-us/office/dev/add-ins/
- **Custom Functions Guide**: https://learn.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-overview
- **API Documentation**: https://surpluslinesapi.com/docs/
- **Support Email**: support@undtec.com

---

**Version**: 2.0.0
**Last Updated**: 2026-02-16
**© Underwriters Technologies** - https://undtec.com
