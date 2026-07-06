# AA Scooter Rental — Staff Tools

Internal staff web tools for AA Scooter Rental (Chiang Mai), backed by a Google Sheet
via a Google Apps Script web app (`Code.gs`).

## Pages

| File | Purpose |
|---|---|
| `index.html` | Staff tools home / navigation, password-gated |
| `customers.html` | Add customers, search past rentals |
| `pricing.html` | Rental price calculator |
| `parts.html` | Search a bike, edit parts & oil change record |
| `oilchange.html` | Bikes ranked by how soon they need an oil change |
| `bikes.html` | All bikes, current renter, due-back date |
| `bikephotos.html` | Upload / delete bike photos |
| `available-bikes.html` | See what's free, pick dates, get price + reply message |
| `reply-assistant.html` | Voice-instruction WhatsApp reply drafting assistant |
| `Code.gs` | Google Apps Script backend (deploy as a Web App bound to the Google Sheet) |

## Deploying the backend

1. Open the Google Sheet this is meant to run against.
2. Extensions → Apps Script, and paste in `Code.gs`.
3. Deploy → New deployment → **Web app**, execute as yourself, access "Anyone with the link".
4. Copy the resulting `/exec` URL and set it as the `scriptUrl` constant at the top of each HTML file's `<script>` block.

## Hosting the frontend

These are static HTML files — any static host works (GitHub Pages, Netlify, Vercel, etc.),
or they can simply be opened/served locally.

## ⚠️ Security note

`index.html` gates access behind a client-side password check (stored in plain text in
the file, using `sessionStorage` to remember unlock state). This is **not real security**
— anyone who views the page source can read the password. Treat this repo as **private**
if that matters to you, and don't rely on it to protect sensitive customer data. For real
protection, put these pages behind actual server-side auth or a host-level access control
(e.g. Netlify password protection, Cloudflare Access, etc.).
