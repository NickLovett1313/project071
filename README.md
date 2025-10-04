# SAP vs Open Orders — Client-Only Static Web App

This is a 100% client-side (browser-only) implementation of the 4-field row-by-row comparison:
**PO #, PO Item, Model, Location**. Nothing is uploaded or saved.

- Excel parsing via [SheetJS](https://sheetjs.com/) (CDN).
- Model trimmed at `>>`. SAP location code `2913` → `EDMONTON`.
- Strict row-by-row comparison (row 2 vs row 2, etc.).
- Summary metrics, full combined table, discrepancy table, and CSV download.

## Run locally
Just open `index.html` in a modern browser.

## Deploy to Azure Static Web Apps (Free)
1. Create a new **Static Web App** (Free) in the Azure Portal.
2. Link to your GitHub repo with these files in the root.
3. Build preset: **Custom**, App location: `/` (no build step needed).
4. (Optional but recommended) Auth: Ensure the GitHub Action creates SWA; then add **Entra ID** (AAD) auth.
   - You can also use `staticwebapp.config.json` provided here to require login.
5. Once deployed, visit the SWA URL; you'll be prompted to sign in (if auth is enabled).

## Files
- `index.html` – UI and file inputs
- `styles.css` – minimal styling (dark)
- `app.js` – parsing + comparison logic (client-side)
- `staticwebapp.config.json` – SWA auth config (require login, redirect to AAD)
