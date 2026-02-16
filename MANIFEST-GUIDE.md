# Manifest Files Quick Reference

Quick guide to understand and use the three different manifest files for the Surplus Lines Tax Excel Add-in.

---

## Overview

| Manifest | Environment | URL | Display Name | Use Case |
|----------|-------------|-----|--------------|----------|
| `manifest.dev.xml` | Local Development | `https://localhost:3001` | "Surplus Lines Tax (Dev)" | Day-to-day development |
| `manifest.network.xml` | Network Testing | `https://192.168.0.106:3001` | "Surplus Lines Tax" | Testing on LAN devices |
| `manifest.xml` | Production | CDN URL | "Surplus Lines Tax" | Office Store submission |

---

## manifest.dev.xml

**🎯 Purpose:** Local development on your machine

**📍 URL:** `https://localhost:3001`

**👤 Display Name:** "Surplus Lines Tax (Dev)"

**✅ When to Use:**
- Daily development work
- Making code changes
- Testing new features locally
- Debugging with hot-reload

**📝 How to Use:**

```bash
# 1. Start dev server
npm run dev-server

# 2. Sideload (in new terminal)
npm run sideload

# Or manually upload manifest.dev.xml in Excel
```

**🔧 Configuration:**

```xml
<DisplayName DefaultValue="Surplus Lines Tax (Dev)"/>
<IconUrl DefaultValue="https://localhost:3001/assets/icon-32.png"/>
<SourceLocation DefaultValue="https://localhost:3001/taskpane.html"/>

<AppDomains>
  <AppDomain>https://localhost:3001</AppDomain>
  <AppDomain>https://n8n.undtec.com</AppDomain>
</AppDomains>
```

---

## manifest.network.xml

**🎯 Purpose:** Testing on other devices (iPhone, iPad, other computers)

**📍 URL:** `https://192.168.0.106:3001` (your local network IP)

**👤 Display Name:** "Surplus Lines Tax"

**✅ When to Use:**
- Testing on iPhone/iPad with Excel Mobile
- Testing on colleague's computers
- Cross-device compatibility testing
- Demonstrating to stakeholders

**📝 How to Use:**

```bash
# 1. Find your IP address
ifconfig | grep "inet " | grep -v 127.0.0.1
# Example output: inet 192.168.0.106

# 2. Update manifest.network.xml
# Replace all instances of 192.168.0.106 with YOUR IP

# 3. Start dev server
npm run dev-server

# 4. Share manifest.network.xml with test devices
# Email it, AirDrop it, or copy to shared folder

# 5. On test device:
# Excel → Insert → Add-ins → Upload My Add-in
# Select manifest.network.xml
```

**🔧 Configuration:**

```xml
<DisplayName DefaultValue="Surplus Lines Tax"/>
<IconUrl DefaultValue="https://192.168.0.106:3001/assets/icon-32.png"/>
<SourceLocation DefaultValue="https://192.168.0.106:3001/taskpane.html"/>

<AppDomains>
  <AppDomain>https://192.168.0.106:3001</AppDomain>
  <AppDomain>https://n8n.undtec.com</AppDomain>
</AppDomains>
```

**⚠️ Important Notes:**

1. **Update IP Address:** Every time your IP changes (WiFi reconnect, VPN, etc.), update ALL occurrences in `manifest.network.xml`

2. **Firewall:** Allow incoming connections on port 3001

3. **Certificate Trust:** On each test device, first navigate to `https://YOUR_IP:3001` in a browser and accept the certificate warning

4. **Same Network:** All test devices must be on the same WiFi/LAN as your development machine

---

## manifest.xml

**🎯 Purpose:** Production deployment to Office Store

**📍 URL:** `https://cdn.surpluslinesapi.com/excel/` (production CDN)

**👤 Display Name:** "Surplus Lines Tax"

**✅ When to Use:**
- Office Store submission
- Production release
- Public distribution

**📝 How to Use:**

This manifest is used for final production deployment only. Users will install it from the Office Store.

**🔧 Configuration:**

```xml
<DisplayName DefaultValue="Surplus Lines Tax"/>
<IconUrl DefaultValue="https://cdn.surpluslinesapi.com/excel/icon-32.png"/>
<SourceLocation DefaultValue="https://cdn.surpluslinesapi.com/excel/taskpane.html"/>

<AppDomains>
  <AppDomain>https://surpluslinesapi.com</AppDomain>
  <AppDomain>https://n8n.undtec.com</AppDomain>
</AppDomains>
```

---

## Key Differences

### Display Name

| Manifest | Display Name | Ribbon Icon Text |
|----------|--------------|------------------|
| `dev` | "Surplus Lines Tax **(Dev)**" | Shows it's development version |
| `network` | "Surplus Lines Tax" | Production name |
| `production` | "Surplus Lines Tax" | Production name |

**Why?** The "(Dev)" suffix helps distinguish development vs production when both are installed.

### Icon URLs

| Manifest | Icon URL | Purpose |
|----------|----------|---------|
| `dev` | `https://localhost:3001/assets/icon-32.png` | Served from local dev server |
| `network` | `https://192.168.0.106:3001/assets/icon-32.png` | Served from network dev server |
| `production` | `https://cdn.surpluslinesapi.com/excel/icon-32.png` | Served from production CDN |

### Source Locations

| Manifest | Source Location | Description |
|----------|-----------------|-------------|
| `dev` | `https://localhost:3001/taskpane.html` | Local machine |
| `network` | `https://192.168.0.106:3001/taskpane.html` | Network accessible |
| `production` | `https://cdn.surpluslinesapi.com/excel/taskpane.html` | Production CDN |

---

## Common Workflows

### 1. Start Development (Using manifest.dev.xml)

```bash
# Terminal 1: Start dev server
npm run dev-server

# Terminal 2: Sideload add-in
npm run sideload

# Excel will open with add-in loaded
# Make code changes → npm run build → refresh Excel
```

### 2. Test on iPhone/iPad (Using manifest.network.xml)

```bash
# 1. Find your IP
ifconfig | grep "inet " | grep -v 127.0.0.1
# Output: inet 192.168.0.106

# 2. Edit manifest.network.xml
# Replace all 192.168.0.106 with your actual IP

# 3. Start dev server
npm run dev-server

# 4. Email manifest.network.xml to yourself
# Open on iPhone → Open in Excel → Upload add-in

# 5. On iPhone, open browser and go to:
# https://YOUR_IP:3001
# Accept certificate warning

# 6. Open Excel → Use add-in
```

### 3. Test on Colleague's Computer (Using manifest.network.xml)

```bash
# 1. Update manifest.network.xml with your current IP
# 2. Start dev server: npm run dev-server
# 3. Share manifest.network.xml with colleague
# 4. Colleague uploads in Excel
# 5. Colleague trusts certificate by visiting https://YOUR_IP:3001
```

### 4. Prepare for Production (Using manifest.xml)

```bash
# 1. Update version in manifest.xml
<Version>2.0.0.0</Version>

# 2. Build production bundle
npm run build

# 3. Upload dist/ to CDN

# 4. Test manifest.xml before submission
# Upload to Excel Online and verify

# 5. Submit to Office Store with manifest.xml
```

---

## Troubleshooting

### "Cannot load add-in" when using manifest.network.xml

**Cause:** IP address changed or firewall blocking port 3001

**Solution:**
```bash
# 1. Get current IP
ifconfig | grep "inet " | grep -v 127.0.0.1

# 2. Update ALL occurrences in manifest.network.xml

# 3. Check firewall allows port 3001

# 4. Restart dev server
npm run dev-server
```

### "Certificate error" on test devices

**Cause:** Self-signed certificate not trusted

**Solution:**
```bash
# On test device:
# 1. Open browser (Safari/Chrome)
# 2. Go to https://YOUR_IP:3001
# 3. Click "Advanced" → "Proceed anyway"
# 4. Now Excel can load the add-in
```

### "Add-in shows old version"

**Cause:** Excel cache

**Solution:**
```bash
# 1. Close Excel completely
# 2. Clear cache:
#    Mac: ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/
#    Windows: %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
# 3. Restart Excel
# 4. Reload add-in
```

---

## Best Practices

### Version Numbers

Keep versions consistent across all manifests:

```xml
<Version>2.0.0.0</Version>
```

Increment when releasing new features:
- `2.0.0.0` → Initial v2.0 release
- `2.0.1.0` → Bug fix
- `2.1.0.0` → Minor feature
- `3.0.0.0` → Major breaking change

### App ID

**IMPORTANT:** All manifests MUST use the same App ID:

```xml
<Id>a1b2c3d4-e5f6-7890-abcd-ef1234567890</Id>
```

If IDs differ, Excel treats them as separate add-ins.

### App Domains

Include all domains your add-in will communicate with:

```xml
<AppDomains>
  <AppDomain>https://localhost:3001</AppDomain>      <!-- Dev server -->
  <AppDomain>https://192.168.0.106:3001</AppDomain>  <!-- Network -->
  <AppDomain>https://surpluslinesapi.com</AppDomain> <!-- Production -->
  <AppDomain>https://n8n.undtec.com</AppDomain>      <!-- API endpoint -->
</AppDomains>
```

---

## Quick Commands Reference

```bash
# Find local IP (Mac)
ifconfig | grep "inet " | grep -v 127.0.0.1

# Find local IP (Windows)
ipconfig

# Start dev server (all manifests)
npm run dev-server

# Auto-sideload manifest.dev.xml
npm run sideload

# Build for production
npm run build

# Watch mode for development
npm run watch
```

---

## Summary

| Task | Manifest | Command |
|------|----------|---------|
| Daily development | `manifest.dev.xml` | `npm run sideload` |
| Test on iPhone | `manifest.network.xml` | Email manifest + accept cert |
| Test on colleague PC | `manifest.network.xml` | Share manifest + accept cert |
| Office Store | `manifest.xml` | Submit to Partner Center |

---

**Need Help?**
- See [TESTING.md](TESTING.md) for detailed testing workflows
- See [README.md](README.md) for function usage and examples

---

**Version**: 2.0.0
**Last Updated**: 2026-02-16
**© Underwriters Technologies**
