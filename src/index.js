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
      
      // Format the data for better display
      const formattedData = filteredData.map(row => {
        return {
          'Date': formatDate(row['DATE']),
          'CUSIP': row['CUSIP'] || '',
          'Ticker': row['TICKER'] || '',
          'Description': row['DESCRIPTION'] || '',
          'Shares': formatNumber(row['SHARES']),
          'Market Value': formatCurrency(row['BASE_MV']),
          'Holdings %': formatPercent(row['PCT_HOLDINGS'])
        };
      });
      
      // Build HTML table manually with color coding
      const htmlTable = buildColorCodedTable(formattedData, filteredData);
      
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
      
      function buildColorCodedTable(formattedRows, originalRows) {
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
          const originalPercent = originalRows[index]['PCT_HOLDINGS'];
          const rowClass = originalPercent > 0 ? 'positive-holding' : 
                          originalPercent < 0 ? 'negative-holding' : '';
          
          html += `<tr${rowClass ? ` class="${rowClass}"` : ''}>\n`;
          
          headers.forEach(header => {
            html += `<td>${row[header]}</td>\n`;
          });
          
          html += '</tr>\n';
        });
        
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
        #holdings-table td:nth-child(5),  /* Shares */
        #holdings-table td:nth-child(6),  /* Market Value */
        #holdings-table td:nth-child(7) { /* Holdings % */
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
        .positive-holding td:nth-child(7) {
            background-color: #c3f73a;
            color: #2d5016;
            font-weight: 600;
        }
        
        .negative-holding td:nth-child(7) {
            background-color: #ffb3d9;
            color: #8b0045;
            font-weight: 600;
        }
        
        #holdings-table tr:hover {
            background-color: #f5f5f5;
            transition: background-color 0.2s ease;
        }
        
        .positive-holding:hover td:nth-child(7) {
            background-color: #b4e632;
        }
        
        .negative-holding:hover td:nth-child(7) {
            background-color: #ff99cc;
        }
        
        #holdings-table tr:nth-child(even) {
            background-color: #fafafa;
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
            .header h1 {
                font-size: 1.8em;
            }
            
            .table-container {
                padding: 15px;
            }
            
            #holdings-table {
                font-size: 12px;
            }
            
            #holdings-table th,
            #holdings-table td {
                padding: 8px 10px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 DBMF Holdings</h1>
            <p>Securities with Tickers</p>
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

