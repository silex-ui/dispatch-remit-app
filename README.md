# Dispatch & Remit App

Delivery dispatching & rider remit web app for the business — bookings, calculator,
transaction history, payment status tracking, and rider remit CSV export.

- `index.html` — the whole app (single-file, vanilla HTML/CSS/JS). This is the file Netlify serves.
- `Code.gs` — the Google Apps Script backend (deployed separately in Google Sheets)
  that the app syncs to for shared, cross-device data.

## Deploying

This repo is connected to Netlify for continuous deployment: pushing to `main`
automatically redeploys the live site. No manual steps needed after the first setup.
