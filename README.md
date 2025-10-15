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

