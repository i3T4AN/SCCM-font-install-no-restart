# Install-Font.ps1 — README

## Overview

Installs a custom font from a network share, registers it in Windows, and refreshes the font cache without requiring a reboot. Safe for SCCM deployment and repeat runs.

---

## Requirements

* Windows 10/11, PowerShell 5.1
* Admin privileges
* Access to a shared `.otf` or `.ttf` font file

---

## What It Does

1. Checks admin rights.
2. Copies the font file → `%WINDIR%\Fonts`.
3. Updates registry:
   `HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts`
4. Broadcasts `WM_FONTCHANGE` so apps see the font immediately.

---

## Usage

Run from an elevated console:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Install-Font.ps1
```

### Exit Codes

* `0` Success
* `1` Copy/registry error
* `2` Invalid path or preprocessing failure
* `3` Not admin

---

## Verification

* File exists in `%WINDIR%\Fonts`
* Registry has correct display name → filename mapping
* Font visible in **Settings → Fonts**

---

## SCCM Deployment

**Install command:**

```text
powershell.exe -ExecutionPolicy Bypass -File Install-Font.ps1
```

**Detection options:**

* File exists in `%WINDIR%\Fonts`
* Registry entry present under Fonts key

---

## Notes

* Idempotent: reruns won’t break existing installs.
* To update: replace font file on network share and redeploy.
* To uninstall: remove file and registry value, then rebroadcast font change.
