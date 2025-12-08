const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const LOGS_XLSX = path.join(__dirname, 'data', 'logs.xlsx');
const RESPONSES_XLSX = path.join(__dirname, 'data', 'responses.xlsx');
const ONE_YEAR_MS = 365 * 24 * 60 * 60 * 1000;

function cleanupOldLogs() {
  if (!fs.existsSync(LOGS_XLSX)) {
    console.log('[cleanup] logs.xlsx not found; skipping cleanup');
    return;
  }

  try {
    const wb = XLSX.readFile(LOGS_XLSX);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const logs = XLSX.utils.sheet_to_json(ws, { defval: '' });

    const now = Date.now();
    const logsToKeep = logs.filter(log => {
      if (!log.timestamp) return true; // Keep logs without timestamp
      const logTime = new Date(log.timestamp).getTime();
      const age = now - logTime;
      return age < ONE_YEAR_MS;
    });

    const deleted = logs.length - logsToKeep.length;

    if (deleted > 0) {
      // Rewrite the file with only recent logs
      const newWs = XLSX.utils.json_to_sheet(logsToKeep, { skipHeader: false });
      const newWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWb, newWs, 'logs');
      XLSX.writeFile(newWb, LOGS_XLSX);
      console.log(`[cleanup] Deleted ${deleted} log entries older than 1 year`);
    } else {
      console.log('[cleanup] No logs older than 1 year found');
    }
  } catch (err) {
    console.error('[cleanup] Error cleaning up logs:', err);
  }
}

function cleanupOldResponses() {
  if (!fs.existsSync(RESPONSES_XLSX)) {
    console.log('[cleanup] responses.xlsx not found; skipping cleanup');
    return;
  }

  try {
    const wb = XLSX.readFile(RESPONSES_XLSX);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const responses = XLSX.utils.sheet_to_json(ws, { defval: '' });

    const now = Date.now();
    const responsesToKeep = responses.filter(response => {
      if (!response.timestamp) return true; // Keep responses without timestamp
      const responseTime = new Date(response.timestamp).getTime();
      const age = now - responseTime;
      return age < ONE_YEAR_MS;
    });

    const deleted = responses.length - responsesToKeep.length;

    if (deleted > 0) {
      // Rewrite the file with only recent responses
      const newWs = XLSX.utils.json_to_sheet(responsesToKeep, { skipHeader: false });
      const newWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWb, newWs, 'responses');
      XLSX.writeFile(newWb, RESPONSES_XLSX);
      console.log(`[cleanup] Deleted ${deleted} response entries older than 1 year`);
    } else {
      console.log('[cleanup] No responses older than 1 year found');
    }
  } catch (err) {
    console.error('[cleanup] Error cleaning up responses:', err);
  }
}

module.exports = { cleanupOldLogs, cleanupOldResponses };
