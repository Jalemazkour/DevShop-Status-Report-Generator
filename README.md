# DevShop Report Studio v2.0 — Electron Desktop App

## What This Is
DevShop Report Studio packaged as a native Windows desktop application.
No browser, no Flask server, no Python required. Just a `.exe` installer.

---

## Folder Structure

```
devshop-electron/
├── main.js              ← Electron main process (replaces Flask + subprocess)
├── package.json         ← App config + build config
├── BUILD.bat            ← Run this once to produce the installer
├── src/
│   ├── index.html       ← Full redesigned UI
│   └── preload.js       ← Secure IPC bridge
└── assets/
    ├── icon.png         ← App icon (replace with yours)
    ├── icon.ico         ← Windows taskbar icon (replace with yours)
    └── ICON_INSTRUCTIONS.txt
```

---

## How to Build (One Time)

**Prerequisites:**
- Node.js v16 or higher — https://nodejs.org
- Windows 10/11 machine

**Steps:**
1. Copy this entire folder to your machine
2. Double-click `BUILD.bat`
3. Wait ~2-3 minutes for npm installs and build
4. Find the installer at `dist/DevShop Report Studio Setup x.x.x.exe`
5. Share that `.exe` with your team

---

## How to Add a Custom Icon (Optional but recommended)

1. Create a 512x512 PNG of your logo/icon
2. Save as `assets/icon.png`
3. Convert to `.ico` at https://convertico.com
4. Save as `assets/icon.ico`
5. Run BUILD.bat

If you skip this, the app will build with the default Electron icon.

---

## What Changed from v1 (Flask)

| v1 Flask                          | v2 Electron                          |
|-----------------------------------|--------------------------------------|
| Python + Flask required           | No Python needed                     |
| Browser tab at localhost:5000     | Native desktop window                |
| subprocess → Node.js for docx     | Direct in-process docx generation    |
| Launch_DevShop_Studio.bat         | Installed .exe with desktop shortcut |
| No custom titlebar                | Custom titlebar with window controls |

---

## How to Update the App

To add features or change the UI:
1. Edit `src/index.html` (UI changes)
2. Edit `main.js` (logic/generation changes)
3. Test locally: `npm start`
4. When ready to push to team: run `BUILD.bat` again, share new `.exe`

---

## Output Location

Generated reports are saved to:
`C:\Users\{username}\Documents\DevShop Studio\output\`

The app creates this folder automatically on first run.
