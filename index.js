const express = require('express');
const AdmZip = require('adm-zip');
const fetch = require('node-fetch');

const app = express();
const PORT = process.env.PORT || 3000;

// Filter by Firm ID 6651 - covers all name variants:
// Benjamin & Co, Kreit & Chiu CPA LLP, Paris Kreit & Chiu CPA LLP
const FIRM_ID = '6651';

function parseCSVLine(line) {
  const cols = [];
  let inQuote = false;
  let current = '';
  for (const ch of line) {
    if (ch === '"') {
      inQuote = !inQuote;
    } else if (ch === ',' && !inQuote) {
      cols.push(current.trim());
      current = '';
    } else {
      current += ch;
    }
  }
  cols.push(current.trim());
  return cols;
}

function escapeCSV(val) {
  const str = String(val == null ? '' : val);
  return str.includes(',') || str.includes('"') || str.includes('\n')
    ? '"' + str.replace(/"/g, '""') + '"'
    : str;
}

app.get('/health', function(req, res) {
  res.json({ status: 'ok', firm_id: FIRM_ID });
});

app.get('/filter', async function(req, res) {
  try {
    console.log('Downloading ZIP from PCAOB...');
    const response = await fetch('https://pcaobus.org/assets/PCAOBFiles/FirmFilings.zip');
    if (!response.ok) {
      throw new Error('PCAOB download failed: ' + response.status);
    }

    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    console.log('ZIP downloaded. Size: ' + buffer.length + ' bytes');

    const zip = new AdmZip(buffer);
    const entries = zip.getEntries();
    const csvEntry = entries.find(function(e) {
      return e.entryName.toLowerCase().endsWith('.csv');
    });

    if (!csvEntry) {
      throw new Error('No CSV found in ZIP');
    }

    const csvText = csvEntry.getData().toString('utf8');
    console.log('CSV extracted successfully');

    // Handle header that may span two lines due to quoted newline
    const firstNewline = csvText.indexOf('\n');
    const secondNewline = csvText.indexOf('\n', firstNewline + 1);
    const headerCandidate = csvText.substring(0, firstNewline);
    const headerContinuation = csvText.substring(firstNewline + 1, secondNewline);

    const openQuotes = (headerCandidate.match(/"/g) || []).length;
    const headerRaw = openQuotes % 2 !== 0
      ? headerCandidate + ' ' + headerContinuation
      : headerCandidate;

    const dataStart = openQuotes % 2 !== 0 ? secondNewline + 1 : firstNewline + 1;
    const lines = csvText.substring(dataStart).split('\n');
    const headers = parseCSVLine(headerRaw).map(function(h) {
      return h.replace(/^"|"$/g, '').replace(/\s+/g, ' ').trim();
    });

    // Find both Firm ID and Firm Name columns
    const firmIdIndex   = headers.indexOf('Firm ID');
    const firmNameIndex = headers.indexOf('Firm Name');

    if (firmIdIndex === -1) {
      throw new Error('Firm ID column not found in CSV headers');
    }
    if (firmNameIndex === -1) {
      throw new Error('Firm Name column not found in CSV headers');
    }

    console.log('Filtering by Firm ID: ' + FIRM_ID);

    const filteredRows = [];
    for (let i = 1; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) continue;
      const cols = parseCSVLine(line);
      // Primary filter: Firm ID 6651 (catches all name variants)
      if ((cols[firmIdIndex] || '').trim() === FIRM_ID) {
        const row = {};
        headers.forEach(function(h, idx) {
          row[h] = cols[idx] || '';
        });
        filteredRows.push(row);
      }
    }

    console.log('Filtered ' + filteredRows.length + ' rows for Firm ID ' + FIRM_ID);

    // Log distinct firm names found (for audit trail)
    const firmNames = [...new Set(filteredRows.map(r => r['Firm Name']))];
    console.log('Firm name variants found: ' + firmNames.join(' | '));

    if (!filteredRows.length) {
      throw new Error('No rows found for Firm ID ' + FIRM_ID);
    }

    const allColumns = Object.keys(filteredRows[0]);
    const csvLines = [allColumns.map(escapeCSV).join(',')];
    filteredRows.forEach(function(row) {
      csvLines.push(allColumns.map(function(col) {
        return escapeCSV(row[col]);
      }).join(','));
    });

    const today = new Date().toISOString().split('T')[0];
    const fileName = 'KC_FirmFilings_' + today + '.csv';

    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', 'attachment; filename="' + fileName + '"');
    res.setHeader('X-Row-Count', String(filteredRows.length));
    res.setHeader('X-Report-Date', today);
    res.setHeader('X-Firm-Names', firmNames.join(' | '));
    res.send(csvLines.join('\n'));

  } catch (err) {
    console.error('Error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, function() {
  console.log('PCAOB filter service running on port ' + PORT + ' | Firm ID: ' + FIRM_ID);
});
