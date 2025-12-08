const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { cleanupOldLogs, cleanupOldResponses } = require('./cleanup');
const nodemailer = require('nodemailer');

const app = express();
const PORT = process.env.PORT || 3000;

const DATA_DIR = path.join(__dirname, 'data');
const STUDENTS_XLSX = path.join(DATA_DIR, 'students.xlsx');
const LOGS_XLSX = path.join(DATA_DIR, 'logs.xlsx');

app.use(express.static(path.join(__dirname, 'public')));

// Ensure data directory exists
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

const RESPONSES_XLSX = path.join(DATA_DIR, 'responses.xlsx');

function writeStudentsFile(students) {
  const ws = XLSX.utils.json_to_sheet(students, { skipHeader: false });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'students');
  XLSX.writeFile(wb, STUDENTS_XLSX);
}

function writeLogsFile(logs) {
  const ws = XLSX.utils.json_to_sheet(logs, { skipHeader: false });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'logs');
  XLSX.writeFile(wb, LOGS_XLSX);
}

function appendResponseToFile(response) {
  let responses = [];
  if (fs.existsSync(RESPONSES_XLSX)) {
    const wb = XLSX.readFile(RESPONSES_XLSX);
    const ws = wb.Sheets[wb.SheetNames[0]];
    responses = XLSX.utils.sheet_to_json(ws, { defval: '' });
  }
  responses.push({ ...response, timestamp: new Date().toISOString() });
  const ws = XLSX.utils.json_to_sheet(responses, { skipHeader: false });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'responses');
  XLSX.writeFile(wb, RESPONSES_XLSX);
}

// Note: .xlsx files are expected to exist; server does not auto-generate them.
if (!fs.existsSync(STUDENTS_XLSX)) {
  console.warn(`Warning: students.xlsx not found at ${STUDENTS_XLSX}`);
}
if (!fs.existsSync(LOGS_XLSX)) {
  console.warn(`Warning: logs.xlsx not found at ${LOGS_XLSX}`);
}

// Helper to read students into memory
function readStudents() {
  const wb = XLSX.readFile(STUDENTS_XLSX);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { defval: '' });
  // Ensure loggedIn is boolean and normalize fields (drop any id)
  return data.map(s => ({
    firstName: s.firstName || s.FirstName || '',
    lastName: s.lastName || s.LastName || '',
    grade: String(s.grade || s.Grade || ''),
    loggedIn: s.loggedIn === true || s.loggedIn === 'true' || s.loggedIn === 1 || s.loggedIn === 'TRUE'
  }));
}

function readLogs() {
  const wb = XLSX.readFile(LOGS_XLSX);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { defval: '' });
  return data;
}

// API: get all students
app.get('/api/students', (req, res) => {
  try {
    const students = readStudents();
    res.json(students);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to read students' });
  }
});

// API: toggle login state for a student by name (no id required)
app.use(express.json());
app.post('/api/toggle', (req, res) => {
  try {
    const { firstName, lastName } = req.body || {};
    if (!firstName || !lastName) return res.status(400).json({ error: 'firstName and lastName required' });

    const students = readStudents();
    const idx = students.findIndex(s => String(s.firstName) === String(firstName) && String(s.lastName) === String(lastName));
    if (idx === -1) return res.status(404).json({ error: 'Student not found' });

    const student = students[idx];
    student.loggedIn = !Boolean(student.loggedIn);

    // Write back students
    writeStudentsFile(students);

    // Append to logs
    const logs = readLogs();
    const action = student.loggedIn ? 'login' : 'logout';
    logs.push({ timestamp: new Date().toISOString(), firstName: student.firstName, lastName: student.lastName, grade: student.grade, action });
    writeLogsFile(logs);

    res.json(student);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to toggle student' });
  }
});

// API: submit test response
app.post('/api/submit-test-response', (req, res) => {
  try {
    const response = req.body || {};
    appendResponseToFile(response);
    // Send notification email (best-effort)
    sendNotificationEmail(response).catch(err => console.error('Email send error:', err));

    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to save response' });
  }
});

// Helper: send notification email with response summary and attach responses.xlsx
async function sendNotificationEmail(response) {
  const to = process.env.NOTIFY_EMAIL || process.env.ADMIN_EMAIL || 'jstruman4929@gmail.com';

  // Determine transport
  let transporter;
  if (process.env.SMTP_HOST && process.env.SMTP_USER) {
    transporter = nodemailer.createTransport({
      host: process.env.SMTP_HOST,
      port: parseInt(process.env.SMTP_PORT || '587', 10),
      secure: (process.env.SMTP_SECURE === 'true'),
      auth: {
        user: process.env.SMTP_USER,
        pass: process.env.SMTP_PASS
      }
    });
  } else if (process.env.SENDMAIL === 'true') {
    transporter = nodemailer.createTransport({ sendmail: true });
  } else {
    console.warn('[email] SMTP not configured; skipping email send');
    return;
  }

  // Build plain-text summary
  const lines = [];
  lines.push('New test response submitted:');
  for (const k of Object.keys(response)) {
    lines.push(`${k}: ${response[k]}`);
  }
  lines.push(`timestamp: ${new Date().toISOString()}`);
  const textBody = lines.join('\n');

  const attachments = [];
  if (fs.existsSync(RESPONSES_XLSX)) {
    attachments.push({ filename: 'responses.xlsx', path: RESPONSES_XLSX });
  }

  await transporter.sendMail({
    from: process.env.EMAIL_FROM || (process.env.SMTP_USER || 'no-reply@example.com'),
    to,
    subject: 'New test response submitted',
    text: textBody,
    attachments
  });
}

// API: admin login (password hardcoded as "a" for now, can be env var later)
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'a';
app.post('/api/admin-login', (req, res) => {
  const { password } = req.body || {};
  if (password === ADMIN_PASSWORD) {
    // In a real app, set a secure session cookie here
    res.json({ success: true });
  } else {
    res.status(401).json({ error: 'Unauthorized' });
  }
});

// API: get all logs
app.get('/api/admin/logs', (req, res) => {
  try {
    if (!fs.existsSync(LOGS_XLSX)) {
      return res.json([]);
    }
    const wb = XLSX.readFile(LOGS_XLSX);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const logs = XLSX.utils.sheet_to_json(ws, { defval: '' });
    res.json(logs);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to fetch logs' });
  }
});

// API: get all responses
app.get('/api/admin/responses', (req, res) => {
  try {
    if (!fs.existsSync(RESPONSES_XLSX)) {
      return res.json([]);
    }
    const wb = XLSX.readFile(RESPONSES_XLSX);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const responses = XLSX.utils.sheet_to_json(ws, { defval: '' });
    res.json(responses);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to fetch responses' });
  }
});

// API: add a new student (admin only) - appends to students.xlsx
app.post('/api/admin/add-student', (req, res) => {
  try {
    const { firstName, lastName, grade } = req.body || {};
    if (!firstName || !lastName || !grade) return res.status(400).json({ error: 'firstName,lastName,grade required' });

    let students = [];
    if (fs.existsSync(STUDENTS_XLSX)) {
      const wb = XLSX.readFile(STUDENTS_XLSX);
      const ws = wb.Sheets[wb.SheetNames[0]];
      students = XLSX.utils.sheet_to_json(ws, { defval: '' });
    }

    const newStudent = { firstName, lastName, grade: String(grade), loggedIn: false };
    students.push(newStudent);

    // Write back students file
    const ws = XLSX.utils.json_to_sheet(students, { skipHeader: false });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'students');
    XLSX.writeFile(wb, STUDENTS_XLSX);

    res.json(newStudent);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to add student' });
  }
});

// Schedule daily cleanup of old logs (runs once per day)
function scheduleDailyCleanup() {
  // Run cleanup once on startup
  cleanupOldLogs();
  cleanupOldResponses();

  // Schedule to run every 24 hours
  setInterval(() => {
    console.log('[scheduler] Running daily log and response cleanup...');
    cleanupOldLogs();
    cleanupOldResponses();
  }, 24 * 60 * 60 * 1000);
}

app.listen(PORT, () => {
  console.log(`Server listening on http://localhost:${PORT}`);
  scheduleDailyCleanup();
});
