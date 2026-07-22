import PostalMime from 'postal-mime';

const HOLDINGS_URL = 'https://www.imgp.com/us/fund/US53700T8273/';

// Symbol mapping from DBMF tickers to Barchart root symbols
const BARCHART_SYMBOL_MAP = {
  'CL': 'CL',   // Crude Oil
  'ES': 'ES',   // E-mini S&P 500
  'MES': 'M0',  // MSCI Emerging Markets
  'JY': 'J6',   // Japanese Yen
  'MFS': 'DI',  // MSCI EAFE
  'EC': 'E6',   // Euro
  'GC': 'GC',   // Gold
  'US': 'ZB',   // 30-Year Treasury Bond
  'TY': 'ZN',   // 10-Year Treasury Note
  'TU': 'ZT',   // 2-Year Treasury Note (US 2YR NOTE CBT)
};

// Parse plain numbers like "-3,735,600,000" or "$  -3,846,646,537.54"
function parsePlainNumber(s) {
  const n = parseFloat(s.replace(/[$,\s]/g, ''));
  return isNaN(n) ? null : n;
}

// Parse holdings data from the iMGP fund HTML page
function parseHoldingsFromHTML(html) {
  const rows = [];
  let totalNetAssets = null;

  // Match each <tr class="holding row"> in the holdings table body
  const tableMatch = html.match(/<table[^>]*id="breakdown-holdings-us"[^>]*>([\s\S]*?)<\/table>/);
  if (!tableMatch) return { rows, totalNetAssets };

  const tbodyMatch = tableMatch[1].match(/<tbody>([\s\S]*?)<\/tbody>/);
  if (!tbodyMatch) return { rows, totalNetAssets };

  const rowRegex = /<tr class="holding row">([\s\S]*?)<\/tr>/g;
  let match;
  while ((match = rowRegex.exec(tbodyMatch[1])) !== null) {
    const cells = match[1];
    const getValue = (cls) => {
      const m = cells.match(new RegExp(`<td class="${cls}">(.*?)<\\/td>`, 's'));
      return m ? m[1].trim() : '';
    };

    const ticker = getValue('ticker');
    const securityName = getValue('security_name');
    const weight = parseFloat(getValue('weight'));

    if (securityName === 'TOTAL NET ASSETS') {
      totalNetAssets = parsePlainNumber(getValue('market_value'));
      continue;
    }

    // Skip rows without a real ticker (treasury bills, cash)
    if (!ticker || ticker === '-') continue;

    rows.push({
      'DATE': getValue('value_date'),
      'TICKER': ticker,
      'DESCRIPTION': securityName,
      'SHARES': parsePlainNumber(getValue('shares_qty')),
      'VALUE': parsePlainNumber(getValue('market_value')),
      'WEIGHT': weight,
    });
  }

  return { rows, totalNetAssets };
}

// Parse holdings from a dbmfwatch update email (HTML body).
// The history table has per-contract rows like:
//   CLU6 | WTI CRUDEFUTURE SEP26 | 3.1M | $242.3M | 6.0% | ...older dates...
// The first shares/value/pct triple after the description is the current date.
export function parseHoldingsFromEmail(html) {
  const rows = [];
  const seen = new Set();

  // Report date appears as e.g. "DBMF | 2026-07-15"
  const dateMatch = html.match(/DBMF\s*\|\s*(\d{4}-\d{2}-\d{2})/);
  const date = dateMatch ? dateMatch[1] : '';

  const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/g;
  const cellRegex = /<td[^>]*>([\s\S]*?)<\/td>/g;
  let rowMatch;
  while ((rowMatch = rowRegex.exec(html)) !== null) {
    const cells = [];
    let cellMatch;
    cellRegex.lastIndex = 0;
    while ((cellMatch = cellRegex.exec(rowMatch[1])) !== null) {
      cells.push(
        cellMatch[1]
          .replace(/<[^>]*>/g, '')
          .replace(/&amp;/g, '&')
          .replace(/&nbsp;/g, ' ')
          .trim()
      );
    }
    if (cells.length < 5) continue;

    // Per-contract tickers only (root + month code + year digit), e.g. CLU6, MESU6
    const ticker = cells[0];
    if (!/^[A-Z]{1,4}[FGHJKMNQUVXZ]\d$/.test(ticker) || seen.has(ticker)) continue;

    // Current-date weight, e.g. "6.0%" or "-99.0%"; "-" means not currently held
    const pctMatch = cells[4].match(/^(-?\d+(?:\.\d+)?)%$/);
    if (!pctMatch) continue;

    // Parse abbreviated numbers like "3.1M", "-3.7B", "$242.3M"
    const parseAbbreviated = (s) => {
      const m = s.replace(/[$,]/g, '').match(/^(-?\d+(?:\.\d+)?)([KMB])?$/i);
      if (!m) return null;
      const mult = { K: 1e3, M: 1e6, B: 1e9 }[(m[2] || '').toUpperCase()] || 1;
      return parseFloat(m[1]) * mult;
    };

    seen.add(ticker);
    rows.push({
      'DATE': date,
      'TICKER': ticker,
      'DESCRIPTION': cells[1],
      'SHARES': parseAbbreviated(cells[2]),
      'VALUE': parseAbbreviated(cells[3]),
      'WEIGHT': parseFloat(pctMatch[1]) / 100,
    });
  }

  // The email has no explicit TNA row, but VALUE / WEIGHT approximates it
  // for every row (both rounded); the median is a robust estimate
  const ratios = rows
    .filter(r => r.VALUE != null && r.WEIGHT)
    .map(r => r.VALUE / r.WEIGHT)
    .sort((a, b) => a - b);
  const totalNetAssetsEstimate = ratios.length
    ? Math.round(ratios[Math.floor(ratios.length / 2)])
    : null;

  return { date, rows, totalNetAssetsEstimate };
}

// Fetch the iMGP page, parse holdings, and store the result in KV.
// Returns the stored payload, or null on failure.
async function refreshIMGPHoldings(env) {
  try {
    const response = await fetch(HOLDINGS_URL);
    if (!response.ok) {
      console.error(`iMGP page fetch failed: ${response.status} ${response.statusText}`);
      return null;
    }
    const { rows, totalNetAssets } = parseHoldingsFromHTML(await response.text());
    if (rows.length === 0) {
      console.error('iMGP page parsed to 0 holdings rows');
      return null;
    }
    const payload = {
      source: 'imgp',
      date: rows[0]['DATE'] || '',
      fetchedAt: new Date().toISOString(),
      totalNetAssets,
      rows,
    };
    const prev = await env.DBMF_KV.get('imgp:latest', 'json');
    await env.DBMF_KV.put('imgp:latest', JSON.stringify(payload));
    await alertAllocationChanges(env, 'iMGP', prev, payload);
    return payload;
  } catch (e) {
    console.error('Error refreshing iMGP holdings:', e);
    return null;
  }
}

const MONTH_CODES = 'FGHJKMNQUVXZ';

// Extract the futures root from a per-contract ticker (CLU6 -> CL, MESU6 -> MES)
function futuresRoot(ticker) {
  const m = ticker.match(/^([A-Z]+?)[FGHJKMNQUVXZ]\d$/);
  return m ? m[1] : ticker;
}

// Orderable expiry rank (year digit + month code), for pairing rolls
function expiryOrder(ticker) {
  const m = ticker.match(/([FGHJKMNQUVXZ])(\d)$/);
  return m ? parseInt(m[2]) * 12 + MONTH_CODES.indexOf(m[1]) : 0;
}

// Diff two holdings snapshots into allocation changes. Expiry rolls (same
// root, contract replaced) are always reported, even at unchanged weight;
// weight-only moves must exceed weightThreshold (decimal, e.g. 0.01 = 1pp).
// Each change carries a stable key so the same change seen again (e.g. from
// the other source) can be deduped.
export function computeAllocationChanges(prevRows, newRows, weightThreshold) {
  const changes = [];
  const fmt = w => (w >= 0 ? '+' : '') + (w * 100).toFixed(1) + '%';

  const prevByTicker = new Map(prevRows.map(r => [r.TICKER, r]));
  const newByTicker = new Map(newRows.map(r => [r.TICKER, r]));

  const roots = [...new Set([...prevRows, ...newRows].map(r => futuresRoot(r.TICKER)))].sort();
  for (const root of roots) {
    const byExpiry = (a, b) => expiryOrder(a.TICKER) - expiryOrder(b.TICKER);
    const removed = prevRows
      .filter(r => futuresRoot(r.TICKER) === root && !newByTicker.has(r.TICKER))
      .sort(byExpiry);
    const added = newRows
      .filter(r => futuresRoot(r.TICKER) === root && !prevByTicker.has(r.TICKER))
      .sort(byExpiry);

    // A removed and an added contract of the same root is a roll
    const nRolls = Math.min(removed.length, added.length);
    for (let i = 0; i < nRolls; i++) {
      const from = removed[i], to = added[i];
      const resized = Math.abs(to.WEIGHT - from.WEIGHT) >= weightThreshold;
      changes.push({
        key: `roll:${from.TICKER}>${to.TICKER}`,
        text: `\u{1F504} ${root} rolled ${from.TICKER} → ${to.TICKER} ` +
          (resized ? `(${fmt(from.WEIGHT)} → ${fmt(to.WEIGHT)})` : `(${fmt(to.WEIGHT)})`),
      });
    }
    for (const r of removed.slice(nRolls)) {
      changes.push({ key: `close:${r.TICKER}`, text: `➖ Closed ${r.TICKER} (was ${fmt(r.WEIGHT)})` });
    }
    for (const r of added.slice(nRolls)) {
      changes.push({ key: `open:${r.TICKER}`, text: `➕ New ${r.TICKER} at ${fmt(r.WEIGHT)}` });
    }
    for (const r of newRows) {
      if (futuresRoot(r.TICKER) !== root) continue;
      const p = prevByTicker.get(r.TICKER);
      if (p && Math.abs(r.WEIGHT - p.WEIGHT) >= weightThreshold) {
        changes.push({ key: `resize:${r.TICKER}`, text: `Δ ${r.TICKER}: ${fmt(p.WEIGHT)} → ${fmt(r.WEIGHT)}` });
      }
    }
  }
  return changes;
}

async function sendPushover(env, title, message) {
  if (!env.PUSHOVER_TOKEN || !env.PUSHOVER_USER) {
    console.warn('Pushover secrets not configured; skipping alert:', title);
    return false;
  }
  const resp = await fetch('https://api.pushover.net/1/messages.json', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      token: env.PUSHOVER_TOKEN,
      user: env.PUSHOVER_USER,
      title,
      message,
    }),
  });
  if (!resp.ok) console.error('Pushover send failed:', resp.status, await resp.text());
  return resp.ok;
}

// Compare a source's previous snapshot to its new one and push an alert for
// any allocation changes not already alerted for this holdings date. Both
// sources report the same underlying change, so alerted change keys are
// recorded per date and skipped on the second sighting.
async function alertAllocationChanges(env, source, prevPayload, newPayload) {
  try {
    if (!prevPayload?.rows?.length || !newPayload?.rows?.length) return;
    const threshold = parseFloat(env.ALERT_WEIGHT_THRESHOLD) || 0.01;
    const changes = computeAllocationChanges(prevPayload.rows, newPayload.rows, threshold);
    if (changes.length === 0) return;

    const date = normalizeDate(newPayload.date) || 'unknown';
    const sentKey = `alert:sent:${date}`;
    const sent = new Set(await env.DBMF_KV.get(sentKey, 'json') || []);
    const fresh = changes.filter(c => !sent.has(c.key));
    if (fresh.length === 0) return;

    const delivered = await sendPushover(
      env,
      `DBMF allocation change (${date})`,
      fresh.map(c => c.text).join('\n') + `\n\nvia ${source}`
    );

    // Only record keys as alerted on successful delivery, so a failed send
    // (or missing secrets) is retried on the next sighting
    if (delivered) {
      fresh.forEach(c => sent.add(c.key));
      await env.DBMF_KV.put(sentKey, JSON.stringify([...sent]), { expirationTtl: 7 * 86400 });
    }
  } catch (e) {
    console.error('Error alerting allocation changes:', e);
  }
}

// Normalize "MM/DD/YYYY" (iMGP) or "YYYY-MM-DD" (dbmfwatch) to ISO
function normalizeDate(dateStr) {
  if (!dateStr) return '';
  const us = dateStr.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (us) return `${us[3]}-${us[1]}-${us[2]}`;
  return dateStr;
}

// dbmfwatch abbreviates shares to one decimal of K/M/B (e.g. "-1.0B"), so a
// parsed value is quantized to 0.1 × its unit; half that step is the largest
// error pure display rounding can introduce
function abbreviationHalfStep(v) {
  const a = Math.abs(v);
  const unit = a >= 1e9 ? 1e9 : a >= 1e6 ? 1e6 : a >= 1e3 ? 1e3 : 1;
  return unit / 20;
}

// Cross-check the two sources: ticker sets, weights (absolute tolerance,
// dbmfwatch rounds to 0.1pp) and shares. Shares pass on either a relative
// tolerance or an absolute diff within dbmfwatch's abbreviation rounding —
// a flat relative bound alone can't work, since rounding error is up to 5%
// when the abbreviated mantissa is near 1.0 but ~0.02% near 999.9
function compareSources(imgp, dbmfwatch, weightTolerance, sharesTolerance) {
  const imgpMap = new Map(imgp.rows.map(r => [r.TICKER, r]));
  const dwMap = new Map(dbmfwatch.rows.map(r => [r.TICKER, r]));

  const onlyInImgp = [...imgpMap.keys()].filter(t => !dwMap.has(t));
  const onlyInDbmfwatch = [...dwMap.keys()].filter(t => !imgpMap.has(t));

  const diffs = [];
  for (const [ticker, a] of imgpMap) {
    const b = dwMap.get(ticker);
    if (!b) continue;

    const weightDiff = a.WEIGHT - b.WEIGHT;
    const weightOk = Math.abs(weightDiff) <= weightTolerance;

    let sharesRelDiff = null;
    let sharesOk = null;
    if (a.SHARES != null && b.SHARES != null) {
      const denom = Math.max(Math.abs(a.SHARES), Math.abs(b.SHARES));
      sharesRelDiff = denom === 0 ? 0 : Math.abs(a.SHARES - b.SHARES) / denom;
      sharesRelDiff = Math.round(sharesRelDiff * 1e6) / 1e6;
      sharesOk = sharesRelDiff <= sharesTolerance ||
        Math.abs(a.SHARES - b.SHARES) <= abbreviationHalfStep(b.SHARES);
    }

    diffs.push({
      ticker,
      imgpWeight: a.WEIGHT,
      dbmfwatchWeight: b.WEIGHT,
      weightDiff: Math.round(weightDiff * 1e6) / 1e6,
      weightOk,
      imgpShares: a.SHARES,
      dbmfwatchShares: b.SHARES,
      sharesRelDiff,
      sharesOk,
    });
  }

  const imgpDate = normalizeDate(imgp.date);
  const dbmfwatchDate = normalizeDate(dbmfwatch.date);

  // Informational only (not part of ok): the dbmfwatch figure is a derived
  // estimate, not an independent measurement
  const tnaRelDiff = imgp.totalNetAssets && dbmfwatch.totalNetAssetsEstimate
    ? Math.round(Math.abs(imgp.totalNetAssets - dbmfwatch.totalNetAssetsEstimate)
        / Math.abs(imgp.totalNetAssets) * 1e6) / 1e6
    : null;

  return {
    ok: onlyInImgp.length === 0 && onlyInDbmfwatch.length === 0 &&
        diffs.every(d => d.weightOk && d.sharesOk !== false),
    imgpDate,
    dbmfwatchDate,
    sameDate: imgpDate === dbmfwatchDate,
    weightTolerance,
    sharesTolerance,
    imgpTotalNetAssets: imgp.totalNetAssets ?? null,
    dbmfwatchTnaEstimate: dbmfwatch.totalNetAssetsEstimate ?? null,
    tnaRelDiff,
    onlyInImgp,
    onlyInDbmfwatch,
    diffs,
  };
}

function jsonResponse(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      'Content-Type': 'application/json',
      'Cache-Control': 'no-cache',
      'Access-Control-Allow-Origin': '*'
    }
  });
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    // Handle API endpoint for price updates only
    if (url.pathname === '/api/prices') {
      return handlePriceUpdate(request, env, ctx);
    }

    // Both holdings parses (iMGP page + dbmfwatch email) and their consistency
    if (url.pathname === '/api/holdings') {
      const weightTolerance = parseFloat(url.searchParams.get('wtol')) || 0.002;
      const sharesTolerance = parseFloat(url.searchParams.get('stol')) || 0.02;
      let imgp = await env.DBMF_KV.get('imgp:latest', 'json');
      if (!imgp) imgp = await refreshIMGPHoldings(env);
      const dbmfwatch = await env.DBMF_KV.get('dbmfwatch:latest', 'json');
      const comparison = imgp && dbmfwatch
        ? compareSources(imgp, dbmfwatch, weightTolerance, sharesTolerance)
        : null;
      return jsonResponse({
        consistent: comparison ? comparison.ok : null,
        imgp,
        dbmfwatch,
        comparison,
      }, imgp || dbmfwatch ? 200 : 404);
    }

    // Handle main page
    return handleMainPage(request, env, ctx);
  },

  // Cron: refresh the iMGP holdings snapshot in KV
  async scheduled(event, env, ctx) {
    ctx.waitUntil(refreshIMGPHoldings(env));
  },

  // Receives mail forwarded from Gmail via Email Routing (dbmf@thechiao.com)
  async email(message, env, ctx) {
    const raw = await new Response(message.raw).arrayBuffer();
    const parsed = await PostalMime.parse(raw);
    const subject = parsed.subject || message.headers.get('subject') || '';

    // Keep the most recent email around for debugging and for the Gmail
    // forwarding-verification code
    await env.DBMF_KV.put('inbox:last', JSON.stringify({
      from: message.from,
      subject,
      receivedAt: new Date().toISOString(),
      text: (parsed.text || '').slice(0, 20000),
    }));

    if (!/DBMF update/i.test(subject) || !parsed.html) return;

    const { date, rows, totalNetAssetsEstimate } = parseHoldingsFromEmail(parsed.html);
    if (rows.length === 0) {
      console.error(`dbmfwatch email "${subject}" parsed to 0 holdings rows`);
      return;
    }

    const payload = {
      source: 'dbmfwatch',
      date,
      receivedAt: new Date().toISOString(),
      totalNetAssetsEstimate,
      rows,
    };
    if (date) {
      await env.DBMF_KV.put(`dbmfwatch:${date}`, JSON.stringify(payload));
    }
    // Only advance "latest" — an old email forwarded later must not regress it
    const existing = await env.DBMF_KV.get('dbmfwatch:latest', 'json');
    if (!existing || !existing.date || (date && date >= existing.date)) {
      await env.DBMF_KV.put('dbmfwatch:latest', JSON.stringify(payload));
      await alertAllocationChanges(env, 'dbmfwatch', existing, payload);
    }
  }
};

// API endpoint to fetch only price updates
async function handlePriceUpdate(request, env, ctx) {
  try {
    const url = new URL(request.url);
    const tickersParam = url.searchParams.get('tickers');

    if (!tickersParam) {
      return new Response(JSON.stringify({ error: 'Missing tickers parameter' }), {
        status: 400,
        headers: { 'Content-Type': 'application/json' }
      });
    }

    const tickers = tickersParam.split(',');
    const prices = {};

    // Fetch prices in parallel
    const batchSize = 10;
    for (let i = 0; i < tickers.length; i += batchSize) {
      const batch = tickers.slice(i, i + batchSize);
      const batchPromises = batch.map(ticker => fetchSinglePriceForAPI(ticker));
      const batchResults = await Promise.all(batchPromises);

      batch.forEach((ticker, index) => {
        prices[ticker] = batchResults[index];
      });
    }

    return new Response(JSON.stringify({ prices, timestamp: new Date().toISOString() }), {
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-cache',
        'Access-Control-Allow-Origin': '*'
      }
    });
  } catch (error) {
    return new Response(JSON.stringify({ error: error.message }), {
      status: 500,
      headers: { 'Content-Type': 'application/json' }
    });
  }
}

// Helper for API price fetching
async function fetchSinglePriceForAPI(ticker) {
  try {
    const barchartRoot = getBarchartRootForAPI(ticker);
    return await fetchBarchartPriceForAPI(barchartRoot);
  } catch (error) {
    console.error(`Error fetching price for ${ticker}:`, error);
    return { change: 'N/A', numeric: null };
  }
}

function getBarchartRootForAPI(ticker) {
  const match = ticker.match(/^([A-Z]+?)([A-Z]\d+)$/);
  const commodityPrefix = match ? match[1] : ticker.match(/^([A-Z]+)/)?.[1] || ticker;
  return BARCHART_SYMBOL_MAP[commodityPrefix] || commodityPrefix;
}

async function fetchBarchartPriceForAPI(root) {
  try {
    const url = `https://www.barchart.com/futures/quotes/${root}*0/futures-prices`;
    const response = await fetch(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
        'Accept': 'text/html'
      }
    });

    if (!response.ok) {
      return { change: 'N/A', numeric: null };
    }

    const html = await response.text();
    const match = html.match(/"percentChange":"([^"]+)"/);

    if (match && match[1]) {
      const percentStr = match[1];
      if (percentStr.toLowerCase() === 'unch') {
        return { change: '+0.00%', numeric: 0 };
      }
      const percentValue = parseFloat(percentStr.replace('%', ''));
      if (!isNaN(percentValue)) {
        const formatted = new Intl.NumberFormat('en-US', {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
          signDisplay: 'always'
        }).format(percentValue);
        return { change: formatted + '%', numeric: percentValue };
      }
    }

    return { change: 'N/A', numeric: null };
  } catch (error) {
    return { change: 'N/A', numeric: null };
  }
}

async function handleMainPage(request, env, ctx) {
    try {
      // Primary source: cron-refreshed iMGP snapshot (self-seeds on first load)
      let filteredData = [];
      let sourceLabel = 'iMGP Fund Page';
      let sourceUrl = HOLDINGS_URL;

      let holdingsTimestamp = null;

      let imgp = await env.DBMF_KV.get('imgp:latest', 'json');
      if (!imgp || !imgp.rows || imgp.rows.length === 0) {
        imgp = await refreshIMGPHoldings(env);
      }
      if (imgp && imgp.rows) {
        filteredData = imgp.rows;
        holdingsTimestamp = imgp.fetchedAt;
      }

      // Fallback source: weights parsed from the dbmfwatch email subscription
      if (filteredData.length === 0) {
        const cached = await env.DBMF_KV.get('dbmfwatch:latest', 'json');
        if (cached && cached.rows && cached.rows.length > 0) {
          filteredData = cached.rows;
          sourceLabel = `dbmfwatch email (${cached.date})`;
          sourceUrl = 'https://dbmfwatch.com';
          holdingsTimestamp = cached.receivedAt;
        }
      }

      if (filteredData.length === 0) {
        return new Response('Holdings data not yet available. Please try again shortly.', {
          status: 503,
          headers: { 'Content-Type': 'text/plain', 'Retry-After': '60' }
        });
      }

      // Footer shows when the holdings snapshot was actually fetched
      const holdingsUpdatedLabel = holdingsTimestamp
        ? new Date(holdingsTimestamp).toLocaleDateString('en-US', {
            weekday: 'short', day: 'numeric', month: 'short', year: 'numeric',
            hour: '2-digit', minute: '2-digit', timeZone: 'America/New_York'
          }) + ' ET'
        : new Date().toLocaleDateString('en-US', { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });

      // Fetch prices for all tickers
      const tickers = filteredData.map(row => row['TICKER']);
      const prices = await fetchTickerPrices(tickers);

      // Calculate contributions and prepare data for sorting
      const dataWithContributions = filteredData.map((row, index) => {
        const holdingsPct = row['WEIGHT'];
        const dailyChangeStr = prices[row['TICKER']];
        let dailyChangePct = 0;
        let contribution = 0;
        
        // Extract numeric value from daily change string
        if (dailyChangeStr && dailyChangeStr !== 'N/A') {
          dailyChangePct = parseFloat(dailyChangeStr.replace('%', '')) / 100; // Convert to decimal
          contribution = holdingsPct * dailyChangePct; // Both are decimals now
        }
        
        return {
          original: row,
          dailyChangeStr: dailyChangeStr,
          dailyChangePct: dailyChangePct,
          contribution: contribution
        };
      });
      
      // Sort by contribution (descending - highest positive contributions first)
      dataWithContributions.sort((a, b) => b.contribution - a.contribution);
      
      // Format the data for display
      const formattedData = dataWithContributions.map(item => {
        const row = item.original;
        return {
          'Date': row['DATE'] || '',
          'Ticker': row['TICKER'] || '',
          'Description': row['DESCRIPTION'] || '',
          'Holdings %': formatPercent(row['WEIGHT']),
          'Daily Change': item.dailyChangeStr || 'N/A',
          'Contribution': item.dailyChangeStr !== 'N/A' ? formatChangePercent(item.contribution * 100) : 'N/A'
        };
      });
      
      // Calculate total contribution
      const totalContribution = dataWithContributions.reduce((sum, item) => sum + item.contribution, 0);
      
      // Build HTML table manually with color coding
      const htmlTable = buildColorCodedTable(formattedData, dataWithContributions, totalContribution);
      
      // Function to fetch ticker prices from Yahoo Finance
      async function fetchTickerPrices(tickers) {
        const prices = {};
        
        // Fetch prices in parallel with a limit to avoid overwhelming the API
        const batchSize = 10;
        for (let i = 0; i < tickers.length; i += batchSize) {
          const batch = tickers.slice(i, i + batchSize);
          const batchPromises = batch.map(ticker => fetchSinglePrice(ticker));
          const batchResults = await Promise.all(batchPromises);
          
          batch.forEach((ticker, index) => {
            prices[ticker] = batchResults[index];
          });
        }
        
        return prices;
      }
      
      // Fetch price from Barchart HTML page
      async function fetchBarchartPrice(root) {
        try {
          // Fetch the HTML page which has embedded JSON data
          const url = `https://www.barchart.com/futures/quotes/${root}*0/futures-prices`;
          const response = await fetch(url, {
            headers: {
              'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
              'Accept': 'text/html'
            }
          });
          
          if (!response.ok) {
            console.error(`Barchart HTTP ${response.status} for root ${root}`);
            return 'N/A';
          }
          
          const html = await response.text();
          
          // Extract percentChange from embedded JSON in the HTML
          // Pattern matches: "percentChange":"-0.54%" or "percentChange":"1.23%" or "unch"
          const match = html.match(/"percentChange":"([^"]+)"/);
          
          if (match && match[1]) {
            const percentStr = match[1];
            // Handle "unch" (unchanged) as 0%
            if (percentStr.toLowerCase() === 'unch') {
              return formatChangePercent(0);
            }
            // Parse the percentage string (e.g., "-0.54%" -> -0.54)
            const percentValue = parseFloat(percentStr.replace('%', ''));
            if (!isNaN(percentValue)) {
              return formatChangePercent(percentValue);
            }
          }
          
          return 'N/A';
        } catch (error) {
          console.error(`Error fetching Barchart price for root ${root}:`, error);
          return 'N/A';
        }
      }
      
      // Extract Barchart root symbol from a DBMF ticker (e.g., CLZ5 -> CL, MFSZ5 -> DI)
      function getBarchartRoot(ticker) {
        const match = ticker.match(/^([A-Z]+?)([A-Z]\d+)$/);
        const commodityPrefix = match ? match[1] : ticker.match(/^([A-Z]+)/)?.[1] || ticker;
        return BARCHART_SYMBOL_MAP[commodityPrefix] || commodityPrefix;
      }

      async function fetchSinglePrice(ticker) {
        try {
          const barchartRoot = getBarchartRoot(ticker);
          return await fetchBarchartPrice(barchartRoot);
        } catch (error) {
          console.error(`Error fetching price for ${ticker}:`, error);
          return 'N/A';
        }
      }
      
      // Helper functions for formatting
      function formatNumber(num) {
        if (num === null || num === undefined || num === '') return '';
        return new Intl.NumberFormat('en-US').format(num);
      }
      
      function formatCurrency(num) {
        if (num === null || num === undefined || num === '') return '';
        return new Intl.NumberFormat('en-US', {
          style: 'currency',
          currency: 'USD',
          minimumFractionDigits: 0,
          maximumFractionDigits: 0
        }).format(num);
      }
      
      function formatPercent(num) {
        if (num === null || num === undefined || num === '') return '';
        return new Intl.NumberFormat('en-US', {
          style: 'percent',
          minimumFractionDigits: 2,
          maximumFractionDigits: 2
        }).format(num);
      }
      
      function formatChangePercent(num) {
        if (num === null || num === undefined || num === '') return '';
        const formatted = new Intl.NumberFormat('en-US', {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
          signDisplay: 'always'
        }).format(num);
        return formatted + '%';
      }

      function getBarchartUrl(ticker) {
        return `https://www.barchart.com/futures/quotes/${getBarchartRoot(ticker)}*0/futures-prices`;
      }

      function buildColorCodedTable(formattedRows, dataRows, totalContribution) {
        if (formattedRows.length === 0) {
          return '<p>No holdings with tickers found.</p>';
        }
        
        // Get column headers from the first row
        const headers = Object.keys(formattedRows[0]);
        
        // Build table HTML
        let html = '<table id="holdings-table">\n';
        html += '<thead><tr>\n';
        
        // Add headers
        headers.forEach(header => {
          html += `<th>${header}</th>\n`;
        });
        
        html += '</tr></thead>\n<tbody>\n';
        
        // Add data rows
        formattedRows.forEach((row, index) => {
          const originalPercent = dataRows[index].original['WEIGHT'];
          const ticker = row['Ticker'];
          const rowClass = originalPercent > 0 ? 'positive-holding' :
                          originalPercent < 0 ? 'negative-holding' : '';

          // Add data attributes for JS price updates
          html += `<tr${rowClass ? ` class="${rowClass}"` : ''} data-ticker="${ticker}" data-holdings="${originalPercent}">\n`;

          headers.forEach((header, colIndex) => {
            // Add special class for Daily Change and Contribution columns to color them
            let tdClasses = [];
            let tdDataAttr = '';
            let cellValue = row[header];

            // Add col-text class for text columns (Date, CUSIP, Ticker, Description)
            if (['Date', 'Ticker', 'Description'].includes(header)) {
              tdClasses.push('col-text');
            }

            if ((header === 'Daily Change' || header === 'Contribution') && cellValue && cellValue !== 'N/A') {
              const numericValue = parseFloat(cellValue.replace('%', ''));
              if (numericValue > 0) {
                tdClasses.push('positive-change');
              } else if (numericValue < 0) {
                tdClasses.push('negative-change');
              } else {
                tdClasses.push('neutral-change');
              }
            }

            // Add data attribute for cells that need JS updates
            if (header === 'Daily Change') {
              tdDataAttr = ' data-col="change"';
            } else if (header === 'Contribution') {
              tdDataAttr = ' data-col="contribution"';
            }

            // Wrap ticker in a link to Barchart source page
            if (header === 'Ticker' && cellValue) {
              const barchartUrl = getBarchartUrl(cellValue);
              cellValue = `<a href="${barchartUrl}" target="_blank" class="ticker-link">${cellValue}</a>`;
            }

            const classAttr = tdClasses.length > 0 ? ` class="${tdClasses.join(' ')}"` : '';
            html += `<td${classAttr}${tdDataAttr}>${cellValue}</td>\n`;
          });

          html += '</tr>\n';
        });
        
        // Add total row
        html += '<tr class="total-row">\n';
        headers.forEach((header, index) => {
          if (header === 'Contribution') {
            const totalClass = totalContribution > 0 ? 'positive-change' :
                              totalContribution < 0 ? 'negative-change' : 'neutral-change';
            html += `<td class="${totalClass}" data-col="total-contribution">${formatChangePercent(totalContribution * 100)}</td>\n`;
          } else if (index === 0) {
            html += `<td class="col-text"><strong>TOTAL</strong></td>\n`;
          } else if (['Date', 'Ticker', 'Description'].includes(header)) {
            html += `<td class="col-text"></td>\n`;
          } else {
            html += `<td></td>\n`;
          }
        });
        html += '</tr>\n';
        
        html += '</tbody>\n</table>';
        
        return html;
      }
      
      // Create a complete HTML page with styling
      const html = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>DBMF Holdings</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 700;
        }
        
        .header p {
            font-size: 1.1em;
            opacity: 0.9;
        }
        
        .table-container {
            overflow-x: auto;
            padding: 30px;
        }
        
        #holdings-table {
            width: 100%;
            border-collapse: collapse;
            border-spacing: 0;
            font-size: 14px;
        }
        
        #holdings-table th {
            background: #667eea;
            color: white;
            padding: 15px;
            text-align: left;
            font-weight: 600;
            position: sticky;
            top: 0;
            z-index: 10;
            text-transform: uppercase;
            font-size: 12px;
            letter-spacing: 0.5px;
        }
        
        #holdings-table td {
            padding: 12px 15px;
            border-bottom: 1px solid #e0e0e0;
        }
        
        /* Explicit dark text color for text columns to prevent browser theme overrides */
        #holdings-table td.col-text {
            color: #333;
        }
        
        /* Right-align numeric columns */
        #holdings-table td:nth-child(4),  /* Holdings % */
        #holdings-table td:nth-child(5),  /* Daily Change */
        #holdings-table td:nth-child(6) { /* Contribution */
            text-align: right;
            font-family: 'SF Mono', Monaco, 'Cascadia Code', 'Roboto Mono', Consolas, 'Courier New', monospace;
        }

        /* Right-align headers for numeric columns */
        #holdings-table th:nth-child(4),
        #holdings-table th:nth-child(5),
        #holdings-table th:nth-child(6) {
            text-align: right;
        }

        /* Color coding for Holdings % column */
        .positive-holding td:nth-child(4) {
            background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%);
            color: #2e7d32;
            font-weight: 600;
        }
        
        .negative-holding td:nth-child(4) {
            background: linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%);
            color: #c62828;
            font-weight: 600;
        }
        
        /* Color coding for Daily Change column */
        .positive-change {
            background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%);
            color: #2e7d32;
            font-weight: 600;
        }
        
        .negative-change {
            background: linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%);
            color: #c62828;
            font-weight: 600;
        }
        
        .neutral-change {
            background: linear-gradient(135deg, #f5f5f5 0%, #eeeeee 100%);
            color: #666;
            font-weight: 600;
        }
        
        #holdings-table tr:hover {
            background-color: #f5f5f5;
            transition: background-color 0.2s ease;
        }
        
        .positive-holding:hover td:nth-child(4) {
            background: linear-gradient(135deg, #c8e6c9 0%, #a5d6a7 100%);
        }
        
        .negative-holding:hover td:nth-child(4) {
            background: linear-gradient(135deg, #ffcdd2 0%, #ef9a9a 100%);
        }
        
        #holdings-table tr:hover .positive-change {
            background: linear-gradient(135deg, #c8e6c9 0%, #a5d6a7 100%);
        }
        
        #holdings-table tr:hover .negative-change {
            background: linear-gradient(135deg, #ffcdd2 0%, #ef9a9a 100%);
        }
        
        #holdings-table tr:hover .neutral-change {
            background: linear-gradient(135deg, #eeeeee 0%, #e0e0e0 100%);
        }
        
        #holdings-table tr:nth-child(even) {
            background-color: #fafafa;
        }
        
        /* Total row styling */
        .total-row {
            background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
            font-weight: 700;
            border-top: 3px solid #667eea;
        }
        
        .total-row td {
            padding: 15px;
            font-size: 1.1em;
        }
        
        .total-row:hover {
            background: linear-gradient(135deg, #bbdefb 0%, #90caf9 100%);
        }
        
        #holdings-table tr:nth-child(even):hover {
            background-color: #f5f5f5;
        }
        
        .footer {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 12px;
            flex-wrap: wrap;
            background: #f8f9fa;
            padding: 14px 30px;
            color: #666;
            font-size: 13px;
            border-top: 1px solid #e0e0e0;
        }
        
        .footer-item {
            white-space: nowrap;
        }
        
        .footer-divider {
            width: 1px;
            height: 16px;
            background: #ccc;
        }
        
        .footer a {
            color: #667eea;
            text-decoration: none;
            font-weight: 600;
        }

        .footer a:hover {
            text-decoration: underline;
        }

        .ticker-link {
            color: #667eea;
            text-decoration: none;
            font-weight: 600;
        }

        .ticker-link:hover {
            text-decoration: underline;
            color: #764ba2;
        }

        .countdown {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 600;
        }

        .countdown-number {
            font-family: 'SF Mono', Monaco, 'Cascadia Code', monospace;
            min-width: 18px;
            text-align: center;
        }

        .refresh-indicator {
            display: none;
            align-items: center;
        }

        .refresh-indicator.active {
            display: inline-flex;
        }

        .spinner {
            width: 14px;
            height: 14px;
            border: 2px solid #e0e0e0;
            border-top-color: #667eea;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        @media (max-width: 768px) {
            body {
                padding: 5px;
            }
            
            .container {
                border-radius: 8px;
            }
            
            .header {
                padding: 10px;
            }
            
            .header h1 {
                font-size: 1.1em;
            }
            
            .header p {
                font-size: 0.7em;
            }
            
            .table-container {
                padding: 3px;
                overflow-x: auto;
            }
            
            #holdings-table {
                font-size: 7px;
            }
            
            #holdings-table th,
            #holdings-table td {
                padding: 2px 2px;
                white-space: nowrap;
            }
            
            /* Consistent header sizing - all at 6px */
            #holdings-table th {
                font-size: 6px;
                padding: 3px 2px;
            }
            
            /* Make Date column smaller */
            #holdings-table td:nth-child(1) {
                font-size: 6px;
            }

            /* Make ticker column slightly larger for readability */
            #holdings-table td:nth-child(2) {
                font-size: 7px;
                font-weight: 600;
            }

            /* Make description column wrappable and limit width */
            #holdings-table td:nth-child(3) {
                max-width: 60px;
                white-space: normal;
                font-size: 6px;
                line-height: 1.1;
            }
            
            .footer {
                padding: 10px 8px;
                font-size: 10px;
                gap: 8px;
            }
            
            .footer-divider {
                height: 12px;
            }

            .countdown {
                padding: 3px 6px;
                font-size: 9px;
            }

            .total-row td {
                padding: 3px 2px;
                font-size: 0.95em;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 DBMF Holdings</h1>
        </div>
        <div class="table-container">
            ${htmlTable}
        </div>
        <div class="footer">
            <span class="footer-item">Source: <a href="${sourceUrl}" target="_blank">${sourceLabel}</a></span>
            <span class="footer-divider"></span>
            <span class="footer-item">Holdings: ${holdingsUpdatedLabel}</span>
            <span class="footer-divider"></span>
            <div class="countdown">
                <span class="countdown-number" id="countdown">60</span>
                <span>s</span>
            </div>
            <div class="refresh-indicator" id="refresh-indicator">
                <div class="spinner"></div>
            </div>
            <span class="footer-divider"></span>
            <span class="footer-item" id="last-refresh-time">--</span>
        </div>
    </div>
    <script>
    (function() {
        // Get all tickers from the table
        const tickers = Array.from(document.querySelectorAll('#holdings-table tbody tr[data-ticker]'))
            .map(row => row.dataset.ticker)
            .filter(t => t);

        let lastRefreshMinute = new Date().getMinutes();

        const countdownEl = document.getElementById('countdown');
        const refreshIndicator = document.getElementById('refresh-indicator');
        const lastRefreshEl = document.getElementById('last-refresh-time');

        // Format time in viewer's timezone
        function formatLocalTime(date) {
            return date.toLocaleTimeString(undefined, {
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            });
        }

        // Format percentage with sign
        function formatChangePercent(num) {
            if (num === null || num === undefined) return 'N/A';
            const sign = num >= 0 ? '+' : '';
            return sign + num.toFixed(2) + '%';
        }

        // Update cell styling based on value
        function updateCellStyle(cell, value) {
            cell.classList.remove('positive-change', 'negative-change', 'neutral-change');
            if (value > 0) {
                cell.classList.add('positive-change');
            } else if (value < 0) {
                cell.classList.add('negative-change');
            } else {
                cell.classList.add('neutral-change');
            }
        }

        // Fetch and update prices
        async function refreshPrices() {
            if (tickers.length === 0) return;

            refreshIndicator.classList.add('active');

            try {
                const response = await fetch('/api/prices?tickers=' + tickers.join(','));
                const data = await response.json();

                if (data.prices) {
                    let totalContribution = 0;
                    const rowsWithContribution = [];

                    // Update each row and collect contribution data
                    document.querySelectorAll('#holdings-table tbody tr[data-ticker]').forEach(row => {
                        const ticker = row.dataset.ticker;
                        const holdings = parseFloat(row.dataset.holdings);
                        const priceData = data.prices[ticker];
                        let contribution = 0;

                        if (priceData) {
                            // Update Daily Change cell
                            const changeCell = row.querySelector('td[data-col="change"]');
                            if (changeCell) {
                                changeCell.textContent = priceData.change;
                                if (priceData.numeric !== null) {
                                    updateCellStyle(changeCell, priceData.numeric);
                                } else {
                                    changeCell.classList.remove('positive-change', 'negative-change', 'neutral-change');
                                }
                            }

                            // Update Contribution cell
                            const contribCell = row.querySelector('td[data-col="contribution"]');
                            if (contribCell && priceData.numeric !== null) {
                                contribution = holdings * (priceData.numeric / 100);
                                totalContribution += contribution;
                                contribCell.textContent = formatChangePercent(contribution * 100);
                                updateCellStyle(contribCell, contribution);
                            } else if (contribCell) {
                                contribCell.textContent = 'N/A';
                                contribCell.classList.remove('positive-change', 'negative-change', 'neutral-change');
                            }
                        }

                        rowsWithContribution.push({ row, contribution });
                    });

                    // Sort rows by contribution (descending)
                    rowsWithContribution.sort((a, b) => b.contribution - a.contribution);

                    // Reorder rows in the DOM
                    const tbody = document.querySelector('#holdings-table tbody');
                    const totalRow = tbody.querySelector('.total-row');
                    rowsWithContribution.forEach(item => {
                        tbody.insertBefore(item.row, totalRow);
                    });

                    // Update total contribution
                    const totalCell = document.querySelector('td[data-col="total-contribution"]');
                    if (totalCell) {
                        totalCell.textContent = formatChangePercent(totalContribution * 100);
                        updateCellStyle(totalCell, totalContribution);
                    }

                    // Update last refresh time in viewer's timezone
                    lastRefreshEl.textContent = 'Last: ' + formatLocalTime(new Date());
                }
            } catch (error) {
                console.error('Error refreshing prices:', error);
            } finally {
                refreshIndicator.classList.remove('active');
            }
        }

        // Check if refresh is needed and update countdown display
        function checkAndRefresh() {
            const now = new Date();
            const currentMinute = now.getMinutes();

            // If we're in a new minute, refresh
            if (currentMinute !== lastRefreshMinute) {
                refreshPrices();
                lastRefreshMinute = currentMinute;
            }

            // Countdown is always derived from actual time - never drifts
            countdownEl.textContent = 60 - now.getSeconds();
        }

        // Check every second
        setInterval(checkAndRefresh, 1000);

        // Also check immediately when tab becomes visible
        document.addEventListener('visibilitychange', () => {
            if (document.visibilityState === 'visible') {
                checkAndRefresh();
            }
        });

        // Initialize
        lastRefreshEl.textContent = 'Last: ' + formatLocalTime(new Date());
        checkAndRefresh();
    })();
    </script>
</body>
</html>
      `.trim();
      
      return new Response(html, {
        headers: {
          'Content-Type': 'text/html;charset=UTF-8',
          'Cache-Control': 'public, max-age=300', // Cache for 5 minutes
        }
      });
      
    } catch (error) {
      console.error('Error:', error);
      return new Response(`Error processing holdings data: ${error.message}`, {
        status: 500,
        headers: { 'Content-Type': 'text/plain' }
      });
    }
}

