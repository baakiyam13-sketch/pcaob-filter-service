const express = require('express');
const AdmZip = require('adm-zip');
const fetch = require('node-fetch');

const app = express();
const PORT = process.env.PORT || 3000;

const FIRM_NAMES = [
  'Paris, Kreit & Chiu CPA LLP',
  'Kreit & Chiu CPA LLP'
];

const KEEP_COLUMNS = [
  'Form Filing ID',
  'Firm Name',
  'Issuer Name',
  'Issuer CIK',
  'Audit Report Date',
  'Fiscal Period End Date',
  'Engagement Partner Last Name',
  'Engagement Partner First Name',
  'Firm Issuing City',
  'Firm Issuing State',
  'Signed Last Name',
  'Signed First Name',
  'Signed Date',
  'Filing Date'
];

function parseCSVLine(line) {
  const cols = [];
  let inQuote = false;
  let current = '';
  for (const ch of line) {
    if (ch === '"') { inQuote = !inQuote; }
    else if (ch === ',' && !inQuote) { cols.push(current.trim()); current = ''; }
    else { current += ch; }
  }
  cols.push(current.trim());
  return cols;
}

function escapeCSV(val) {
  const str = String(val ?? '');
  return str.includes(',') || str.includes('"') || str.includes('\n')
    ? `"${str.replace(/"/g, '""')}"`
    : str;
}

app.get('/health', (req, res) => {
  res.json({ status: 'ok' });
});

app.get('/filter', async (req, res) => {
  try {
    console.log('Downloading ZIP from PCAOB...');
    const response = await fetch('https://pcaobus.org/assets/PCAOBFiles/FirmFilings.zip');
    if (!response.ok) throw new Error(`PCAOB download failed: ${response.status}`);

    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    console.log(`ZIP downloaded. Size: ${buffer.length} bytes`);

    // Extract ZIP
    const zip = new AdmZip(buffer);
    const entries = zip.getEntries();
    const csvEntry = entries.find(e => e.entryName.toLowerCase().endsWith('.csv'));
    if (!csvEntry) throw new Error('No CSV found in ZIP');

    const csvText = csvEntry.getData().toString('utf8');
    console.log('CSV extracted successfully');

    // Parse CSV
    const lines = csvText.split('\n');
    const headers = parseCSVLine(lines[0]).map(h => h.replace(/^"|"$/g, ''));

    const firmNameIndex = headers.indexOf('Firm Name');
    if (firmNameIndex === -1) throw new Error('Firm Name column not found');

    // Filter rows
    const filteredRows = [];
    for (let i = 1; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) continue;
      const cols = parseCSVLine(line);
      if (FIRM_NAMES.includes(cols[firmNameIndex])) {
        const row = {};
        headers.forEach((h, idx) => { row[h] = cols[idx] || ''; });
        filteredRows.push(row);
      }
    }

    console.log(`Filtered ${filteredRows.length} rows`);

    // Build output CSV with selected columns only
    const csvLines = [
      KEEP_COLUMNS.map(escapeCSV).join(','),
      ...filteredRows.map(row => KEEP_COLUMNS.map(col => escapeCSV(row[col])).join(','))
    ];

    const today = new Date().toISOString().split('T')[0];
    const fileName = `KC_FirmFilings_${today}.csv`;
