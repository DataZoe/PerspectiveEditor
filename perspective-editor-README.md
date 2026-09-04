# Perspective Editor — Quick start

**Needs:** Windows + Power BI Desktop + PowerShell 7 (`winget install Microsoft.PowerShell`).

**Files:** `perspective-editor-live.html` and `perspective-server-v3.ps1` — put both in the same folder.

**Steps:**
1. Open your `.pbix` in Power BI Desktop.
2. In **Model view → Model explorer**, copy the server (e.g. `localhost:50030`).
3. Start the proxy (leave the window open): `pwsh -File .\perspective-server-v3.ps1`
4. Double-click `perspective-editor-live.html` to open it in a browser.
5. Paste the server → **Load** → edit → **Apply**.
6. Accept Desktop's "reload" prompt and save the PBIX.

**Table checkbox with ⓘ** = whole-table membership (future objects auto-included).
