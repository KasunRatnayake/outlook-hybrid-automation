# Outlook 2016 Hybrid Automation POC

This repository demonstrates a cross-platform automation strategy for legacy Win32 applications containing modern WebViews. It bridges **WinAppDriver** (for the Windows shell) and **Playwright** (for the GIPHY Add-in).

## The Challenge
Outlook 2019 poses significant automation hurdles:
- **Legacy UI Rendering:** The Ribbon UI is often invisible to standard Accessibility tools.
- **Context Loss:** Initial launch sessions often fail due to splash screens.
- **Hybrid Architecture:** Modern add-ins run in a sandboxed WebView2 environment.

## The Solution (The "Hybrid Bridge")
1. **WinAppDriver:** Orchestrates the Windows OS, handles window handles (HWND), and forces Outlook focus.
2. **CDP (Chrome DevTools Protocol):** Enables Playwright to "attach" to the running WebView2 process inside Outlook on port `9222`.
3. **Self-Healing Locators:** Implements a prioritized array of XPaths and W3C Action fallbacks (Ctrl+N) for resilient navigation.

## 📋 Prerequisites
- Windows 10/11
- [WinAppDriver](https://github.com/microsoft/WinAppDriver) (v1.2.1+)
- Outlook 2016+
- Node.js & Playwright

## 🔧 Environment Setup
To enable Playwright to see the GIPHY add-in, the WebView2 debugging port must be active. Run the following in PowerShell before launching Outlook:

```powershell
$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS = "--remote-debugging-port=9222"


How to Run
Start WinAppDriver.exe.

Ensure Outlook is closed (the script will launch it).

Run the POC:
