import * as XLSX from 'xlsx';

const EXCEL_URL = 'https://imgpfunds.com/wp-content/uploads/pdfs/holdings/DBMF-Holdings.xlsx';

export default {
  async fetch(request, env, ctx) {
    try {
      // Fetch the Excel file
      const response = await fetch(EXCEL_URL);
      
      if (!response.ok) {
        return new Response(`Failed to fetch Excel file: ${response.status} ${response.statusText}`, {
          status: 500,
          headers: { 'Content-Type': 'text/plain' }
        });
      }

      // Get the file as ArrayBuffer
      const arrayBuffer = await response.arrayBuffer();
      
      // Parse the Excel file
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      
      // Get the first sheet
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      
      // Based on the actual file structure:
      // Row 0: Title
      // Row 1: Empty
      // Row 2-3: Fund info (NAV, SHARES_OUTSTANDING, etc.)
      // Row 4: Empty
      // Row 5: Table headers (DATE, CUSIP, TICKER, DESCRIPTION, SHARES, BASE_MV, PCT_HOLDINGS)
      // Row 6+: Holdings data
      
      // Parse starting from row 5 (0-indexed)
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { range: 5 });
      
      // Filter to only include rows with a ticker (TICKER column is not empty)
      const filteredData = jsonData.filter(row => {
        const ticker = row['TICKER'];
        return ticker && String(ticker).trim() !== '';
      });
      
      // Fetch prices for all tickers
      const tickers = filteredData.map(row => row['TICKER']);
      const prices = await fetchTickerPrices(tickers);
      
      // Calculate contributions and prepare data for sorting
      const dataWithContributions = filteredData.map((row, index) => {
        const holdingsPct = row['PCT_HOLDINGS'];
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
          'Date': formatDate(row['DATE']),
          'CUSIP': row['CUSIP'] || '',
          'Ticker': row['TICKER'] || '',
          'Description': row['DESCRIPTION'] || '',
          'Holdings %': formatPercent(row['PCT_HOLDINGS']),
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
      
      // Fetch price from Barchart HTML page for MSCI indices
      // root: 'DI' for MSCI EAFE (MFS), 'M0' for MSCI Emerging Markets (MME)
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
          // Pattern matches: "percentChange":"-0.54%" or "percentChange":"1.23%"
          const match = html.match(/"percentChange":"([^"]+)"/);
          
          if (match && match[1]) {
            const percentStr = match[1];
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
      
      async function fetchSinglePrice(ticker) {
        try {
          // Extract commodity prefix from ticker (e.g., CLZ5 -> CL, MFSZ5 -> MFS)
          const match = ticker.match(/^([A-Z]+?)([A-Z]\d+)$/);
          const commodityPrefix = match ? match[1] : ticker.match(/^([A-Z]+)/)?.[1] || ticker;
          
          // Use Barchart for MSCI indices (MFS and MES) since Yahoo Finance no longer supports them
          if (commodityPrefix === 'MFS') {
            return await fetchBarchartPrice('DI');
          }
          if (commodityPrefix === 'MES') {
            return await fetchBarchartPrice('M0');
          }
          
          // Convert futures contract symbol to Yahoo Finance format
          // e.g., CLZ5 -> CL=F, GCZ5 -> GC=F
          const yahooTicker = convertToYahooSymbol(ticker);
          
          // Try Yahoo Finance quote API
          const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(yahooTicker)}`;
          const response = await fetch(url, {
            headers: {
              'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
              'Accept': 'application/json'
            }
          });
          
          if (!response.ok) {
            console.error(`HTTP ${response.status} for ${yahooTicker}`);
            return `N/A`;
          }
          
          const data = await response.json();
          const meta = data?.chart?.result?.[0]?.meta;
          
          if (meta) {
            // Calculate percentage change from previous close
            const currentPrice = meta.regularMarketPrice;
            const previousClose = meta.chartPreviousClose || meta.previousClose;
            
            if (currentPrice && previousClose) {
              const changePercent = ((currentPrice - previousClose) / previousClose) * 100;
              return formatChangePercent(changePercent);
            }
          }
          
          if (data?.chart?.error) {
            console.error(`Yahoo Finance error for ${yahooTicker}:`, data.chart.error);
          }
          
          return `N/A`;
        } catch (error) {
          console.error(`Error fetching price for ${ticker}:`, error);
          return 'N/A';
        }
      }
      
      function convertToYahooSymbol(ticker) {
        // Extract the commodity prefix, excluding the month code
        // e.g., CLZ5 -> CL (Z is month, 5 is year)
        //       GCZ5 -> GC
        //       MFSZ5 -> MFS
        // Pattern: commodity letters + single letter month code + year digits
        const match = ticker.match(/^([A-Z]+?)([A-Z]\d+)$/);
        if (match) {
          const commodityPrefix = match[1];
          return commodityPrefix + '=F';
        }
        // Fallback: if pattern doesn't match, just use all letters
        const simpleMatch = ticker.match(/^([A-Z]+)/);
        if (simpleMatch) {
          const commodityPrefix = simpleMatch[1];
          return commodityPrefix + '=F';
        }
        // If no match, return original ticker
        return ticker;
      }
      
      // Helper functions for formatting
      function formatDate(dateNum) {
        if (!dateNum) return '';
        const dateStr = String(dateNum);
        // Format: YYYYMMDD -> YYYY-MM-DD
        if (dateStr.length === 8) {
          return `${dateStr.slice(0,4)}-${dateStr.slice(4,6)}-${dateStr.slice(6,8)}`;
        }
        return dateStr;
      }
      
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
          const originalPercent = dataRows[index].original['PCT_HOLDINGS'];
          const rowClass = originalPercent > 0 ? 'positive-holding' : 
                          originalPercent < 0 ? 'negative-holding' : '';
          
          html += `<tr${rowClass ? ` class="${rowClass}"` : ''}>\n`;
          
          headers.forEach((header, colIndex) => {
            // Add special class for Daily Change and Contribution columns to color them
            let tdClass = '';
            const cellValue = row[header];
            
            if ((header === 'Daily Change' || header === 'Contribution') && cellValue && cellValue !== 'N/A') {
              const numericValue = parseFloat(cellValue.replace('%', ''));
              if (numericValue > 0) {
                tdClass = ' class="positive-change"';
              } else if (numericValue < 0) {
                tdClass = ' class="negative-change"';
              }
            }
            html += `<td${tdClass}>${cellValue}</td>\n`;
          });
          
          html += '</tr>\n';
        });
        
        // Add total row
        html += '<tr class="total-row">\n';
        headers.forEach((header, index) => {
          if (header === 'Contribution') {
            const totalClass = totalContribution > 0 ? 'positive-change' : 
                              totalContribution < 0 ? 'negative-change' : '';
            html += `<td class="${totalClass}">${formatChangePercent(totalContribution * 100)}</td>\n`;
          } else if (index === 0) {
            html += `<td><strong>TOTAL</strong></td>\n`;
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
        
        /* Right-align numeric columns */
        #holdings-table td:nth-child(5),  /* Holdings % */
        #holdings-table td:nth-child(6),  /* Daily Change */
        #holdings-table td:nth-child(7) { /* Contribution */
            text-align: right;
            font-family: 'SF Mono', Monaco, 'Cascadia Code', 'Roboto Mono', Consolas, 'Courier New', monospace;
        }
        
        /* Right-align headers for numeric columns */
        #holdings-table th:nth-child(5),
        #holdings-table th:nth-child(6),
        #holdings-table th:nth-child(7) {
            text-align: right;
        }
        
        /* Color coding for Holdings % column */
        .positive-holding td:nth-child(5) {
            background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%);
            color: #2e7d32;
            font-weight: 600;
            box-shadow: inset 0 0 0 1px rgba(76, 175, 80, 0.2);
        }
        
        .negative-holding td:nth-child(5) {
            background: linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%);
            color: #c62828;
            font-weight: 600;
            box-shadow: inset 0 0 0 1px rgba(244, 67, 54, 0.2);
        }
        
        /* Color coding for Daily Change column */
        .positive-change {
            background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%);
            color: #2e7d32;
            font-weight: 600;
            box-shadow: inset 0 0 0 1px rgba(76, 175, 80, 0.2);
        }
        
        .negative-change {
            background: linear-gradient(135deg, #ffebee 0%, #ffcdd2 100%);
            color: #c62828;
            font-weight: 600;
            box-shadow: inset 0 0 0 1px rgba(244, 67, 54, 0.2);
        }
        
        #holdings-table tr:hover {
            background-color: #f5f5f5;
            transition: background-color 0.2s ease;
        }
        
        .positive-holding:hover td:nth-child(5) {
            background: linear-gradient(135deg, #c8e6c9 0%, #a5d6a7 100%);
            box-shadow: inset 0 0 0 1px rgba(76, 175, 80, 0.3);
        }
        
        .negative-holding:hover td:nth-child(5) {
            background: linear-gradient(135deg, #ffcdd2 0%, #ef9a9a 100%);
            box-shadow: inset 0 0 0 1px rgba(244, 67, 54, 0.3);
        }
        
        #holdings-table tr:hover .positive-change {
            background: linear-gradient(135deg, #c8e6c9 0%, #a5d6a7 100%);
            box-shadow: inset 0 0 0 1px rgba(76, 175, 80, 0.3);
        }
        
        #holdings-table tr:hover .negative-change {
            background: linear-gradient(135deg, #ffcdd2 0%, #ef9a9a 100%);
            box-shadow: inset 0 0 0 1px rgba(244, 67, 54, 0.3);
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
            background: #f8f9fa;
            padding: 20px 30px;
            text-align: center;
            color: #666;
            font-size: 14px;
            border-top: 1px solid #e0e0e0;
        }
        
        .footer a {
            color: #667eea;
            text-decoration: none;
            font-weight: 600;
        }
        
        .footer a:hover {
            text-decoration: underline;
        }
        
        .timestamp {
            margin-top: 10px;
            font-size: 12px;
            opacity: 0.8;
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
            
            /* Make Date and CUSIP columns smaller */
            #holdings-table td:nth-child(1),
            #holdings-table td:nth-child(2) {
                font-size: 6px;
            }
            
            /* Make ticker column slightly larger for readability */
            #holdings-table td:nth-child(3) {
                font-size: 7px;
                font-weight: 600;
            }
            
            /* Make description column wrappable and limit width */
            #holdings-table td:nth-child(4) {
                max-width: 60px;
                white-space: normal;
                font-size: 6px;
                line-height: 1.1;
            }
            
            .footer {
                padding: 8px;
                font-size: 9px;
            }
            
            .footer .timestamp {
                font-size: 7px;
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
            <p>Data source: <a href="${EXCEL_URL}" target="_blank">DBMF-Holdings.xlsx</a></p>
            <p class="timestamp">Last updated: ${new Date().toUTCString()}</p>
        </div>
    </div>
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
      return new Response(`Error processing Excel file: ${error.message}`, {
        status: 500,
        headers: { 'Content-Type': 'text/plain' }
      });
    }
  }
};

