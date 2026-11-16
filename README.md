# House Finances

Installable phone app for the House Finances Google Sheet: https://kilfarn.github.io/house-finances/

This repo is only a full-screen wrapper (icon, manifest, service worker). The app itself is a
Google Apps Script web app bound to the sheet, and only Google accounts the sheet is shared with can
see any data. Install from Chrome: menu > Install app.

## App code

`apps-script/` holds the Apps Script code (Code.gs, Index.html, appsscript.json). Every push to `main`
that changes it is deployed by `.github/workflows/deploy-apps-script.yml`, which updates the existing
web app deployment so the `/exec` link stays the same.
