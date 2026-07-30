# DBMF Holdings Viewer

A Cloudflare Worker that fetches and displays the DBMF (Dynamic Beta-Managed Futures Strategy Fund) holdings data from an Excel file as a beautiful, responsive HTML table.

## Features

- 📥 Automatically fetches the latest DBMF holdings Excel file
- 📊 Converts Excel data to a clean, responsive HTML table
- 🎨 Modern, gradient-styled UI with mobile support
- ⚡ Fast edge-side rendering via Cloudflare Workers
- 💾 5-minute cache for optimal performance

## Prerequisites

- [Node.js](https://nodejs.org/) (v16 or later)
- [npm](https://www.npmjs.com/) or [yarn](https://yarnpkg.com/)
- A [Cloudflare account](https://dash.cloudflare.com/sign-up)
- [Wrangler CLI](https://developers.cloudflare.com/workers/wrangler/install-and-update/)

## Installation

1. Clone this repository and navigate to the project directory:
   ```bash
   cd dbmfview
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Authenticate with Cloudflare (if you haven't already):
   ```bash
   npx wrangler login
   ```

## Development

To run the worker locally for development:

```bash
npm run dev
```

This will start a local server (usually at `http://localhost:8787`) where you can test the worker.

## Deployment

Deploy the worker to Cloudflare:

```bash
npm run deploy
```

After deployment, Wrangler will provide you with a URL where your worker is accessible (e.g., `https://dbmfview.your-subdomain.workers.dev`).

## Configuration

### Custom Route (Optional)

If you want to use a custom domain or route, edit `wrangler.toml`:

```toml
routes = [
  { pattern = "example.com/dbmf", zone_name = "example.com" }
]
```

### Update Excel URL

If the source URL changes, update the `EXCEL_URL` constant in `src/index.js`:

```javascript
const EXCEL_URL = 'https://imgpfunds.com/wp-content/uploads/pdfs/holdings/DBMF-Holdings.xlsx';
```

### Allocation Change Alerts (Pushover)

The worker compares each new holdings snapshot (hourly iMGP cron refresh, and
incoming dbmfwatch emails) against the previous snapshot from the same source
and sends a [Pushover](https://pushover.net/) notification when the allocation
changes:

- 🔄 Expiry rolls (e.g. `CLU6 → CLV6`), even at unchanged size
- ➕ New positions and ➖ closed positions
- Δ Position resizes, sized in **risk terms**: the raw change in shares per
  dollar of NAV, valued at the current price and scaled by the instrument's
  annualized volatility, giving the annualized risk the trade added or
  removed as a fraction of NAV. Shares/NAV is the right base because weight
  drifts with price moves and share counts scale with fund flows, while
  shares/NAV only changes when the fund actually trades; the vol factor is
  what makes a single threshold meaningful across instruments, since a 10pp
  shift in 2-year notes is nothing like a 10pp shift in crude. Raw
  differences rather than relative ones, since long/short positions can
  cross zero. Alerts fire beyond `ALERT_RISK_THRESHOLD` (default `0.002` =
  20bp of NAV vol, configurable in `wrangler.toml`) and report the old and
  new weights, their percentage-point difference, and the risk change

  Per-instrument volatilities live in `ANNUAL_VOL_BY_ROOT` in
  `src/index.js`. They are rough long-run estimates and only need to be
  right to within a factor of ~1.5 to rank trades sensibly. A root with no
  entry is **never given a guessed default** — a wrong vol silently
  mis-scales every threshold for that instrument and looks identical to a
  right one in the output. Instead its changes are always reported, unsized
  (`size unknown` in place of the bp figure) and unfiltered, since without a
  vol there is no basis for judging materiality. The alert carries a
  `⚠️ No annualized volatility configured for <root>` line naming the root
  to add.

Both sources report the same underlying change, so alerted changes are
deduplicated per holdings date — whichever source lands first sends the alert.

To enable, set your Pushover application token and user key as secrets:

```bash
npx wrangler secret put PUSHOVER_TOKEN
npx wrangler secret put PUSHOVER_USER
```

If the secrets are missing, the worker logs a warning and skips the alert.

### Adjust Cache Duration

To change how long the data is cached, modify the `Cache-Control` header in `src/index.js`:

```javascript
'Cache-Control': 'public, max-age=300', // 300 seconds = 5 minutes
```

## Project Structure

```
dbmfview/
├── src/
│   └── index.js          # Main worker script
├── package.json          # Dependencies and scripts
├── wrangler.toml         # Cloudflare Worker configuration
└── README.md            # This file
```

## How It Works

1. The worker receives an HTTP request
2. It fetches the Excel file from the specified URL
3. The `xlsx` library parses the Excel data
4. The data is converted to an HTML table
5. A styled HTML page is generated and returned to the browser
6. The response is cached for 5 minutes to reduce load on the source server

## Dependencies

- **xlsx** (^0.18.5): Library for parsing and writing Excel files
- **wrangler** (^3.0.0): Cloudflare Workers CLI tool (dev dependency)

## Troubleshooting

### Worker fails to fetch Excel file

- Verify the URL is accessible: `curl -I https://imgpfunds.com/wp-content/uploads/pdfs/holdings/DBMF-Holdings.xlsx`
- Check if the source server is blocking Cloudflare Workers
- Review worker logs: `npx wrangler tail`

### Excel parsing errors

- Ensure the file format is valid (.xlsx)
- Check if the file structure has changed
- Look for error messages in the browser or worker logs

### Deployment issues

- Make sure you're authenticated: `npx wrangler whoami`
- Check your Cloudflare account has Workers enabled
- Verify your `wrangler.toml` configuration is correct

## License

MIT

## Data Source

Data is sourced from: [DBMF Holdings Excel File](https://imgpfunds.com/wp-content/uploads/pdfs/holdings/DBMF-Holdings.xlsx)

