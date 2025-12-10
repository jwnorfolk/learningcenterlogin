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
  // Map internal student objects to the expected workbook headers so external tools
  // that expect these column names continue to work.
  const rows = students.map(s => ({
    'First_name': s.firstName || '',
    'Last_name': s.lastName || '',
    'Grade_Level': s.grade || '',
    // Use 1/0 for logged state to match common spreadsheet exports
    'U_StudentsUserFields.dsf6': (s.loggedIn ? 1 : 0)
  }));
  const ws = XLSX.utils.json_to_sheet(rows, { skipHeader: false });
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

function readResponses() {
  if (!fs.existsSync(RESPONSES_XLSX)) {
    return [];
  }
  const wb = XLSX.readFile(RESPONSES_XLSX);
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
}

function writeResponsesFile(responses) {
  // Write responses with explicit column order to match requested layout
  const headers = [
    'email',            // A
    'firstName',        // B
    'lastName',         // C
    'grade',            // D
    'testType',         // E
    'subject',          // F
    'teacherName',      // G
    'otherTeachername', // H
    'testName',         // I
    'testDate',         // J
    'period',           // K
    'timestamp',        // L
    'teacherEmail',     // M
    'sent'              // N
  ];
  const ws = XLSX.utils.json_to_sheet(responses, { header: headers });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'responses');
  XLSX.writeFile(wb, RESPONSES_XLSX);
}

function appendResponseToFile(response) {
  const responses = readResponses();
  // Add sent flag (default false) and timestamp
  const newRow = {
    email: response.email || '',
    firstName: response.firstName || '',
    lastName: response.lastName || '',
    grade: response.grade || '',
    testType: response.testType || '',
    subject: response.subject || '',
    teacherName: response.teacherName || '',
    // form uses otherTeacherLastName for the OTHER field; store under otherTeachername
    otherTeachername: response.otherTeacherLastName || response.otherTeachername || '',
    testName: response.testName || '',
    testDate: response.testDate || '',
    period: response.period || '',
    timestamp: new Date().toISOString(),
    teacherEmail: response.teacherEmail || response.teacherEmail || '',
    sent: false
  };
  responses.push(newRow);
  writeResponsesFile(responses);
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
  // Normalize rows: support different header names coming from the XLSX
  const firstNameCandidates = ['firstName','FirstName','First_name','First Name','C','c'];
  const lastNameCandidates = ['lastName','LastName','Last_name','Last Name','D','d'];
  const gradeCandidates = ['grade','Grade','Grade_Level','Grade Level','F','f'];
  const loggedCandidates = ['loggedIn','logged','logged_in','U_StudentsUserFields.dsf6','U_StudentsUserFields.dsf6 ','B','b'];

  function pickField(obj, candidates) {
    for (const k of candidates) {
      if (Object.prototype.hasOwnProperty.call(obj, k)) return obj[k];
    }
    // fallback: try case-insensitive match
    const keys = Object.keys(obj);
    const lower = keys.reduce((acc, key) => { acc[key.toLowerCase()] = key; return acc; }, {});
    for (const k of candidates) {
      const lk = String(k).toLowerCase();
      if (lower[lk]) return obj[lower[lk]];
    }
    return undefined;
  }

  function parseLogged(val) {
    if (val === true || val === 1) return true;
    if (typeof val === 'number') return val === 1;
    if (!val && val !== 0) return false;
    const s = String(val).trim().toLowerCase();
    return ['true','yes','1','logged in','loggedin','y'].includes(s);
  }

  return data.map(row => {
    const firstName = pickField(row, firstNameCandidates) || '';
    const lastName = pickField(row, lastNameCandidates) || '';
    const grade = String(pickField(row, gradeCandidates) || '');
    const loggedRaw = pickField(row, loggedCandidates);
    const loggedIn = parseLogged(loggedRaw);
    return { firstName: String(firstName || ''), lastName: String(lastName || ''), grade, loggedIn };
  });
}

function readLogs() {
  const wb = XLSX.readFile(LOGS_XLSX);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { defval: '' });
  return data;
}

// Read teacher names from data/teachers.xlsx column A rows 4..142, skipping blanks
function readTeachers() {
  const TEACHERS_XLSX = path.join(DATA_DIR, 'teachers.xlsx');
  if (!fs.existsSync(TEACHERS_XLSX)) {
    console.warn(`Warning: teachers.xlsx not found at ${TEACHERS_XLSX}`);
    return [];
  }
  const wb = XLSX.readFile(TEACHERS_XLSX);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const names = [];
  for (let r = 4; r <= 142; r++) {
    const cell = ws[`A${r}`];
    if (!cell) continue;
    const v = (cell.v || '').toString().trim();
    if (!v) continue;
    names.push(v);
  }
  return names;
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

// API: get teacher list
app.get('/api/teachers', (req, res) => {
  try {
    const teachers = readTeachers();
    res.json(teachers);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to read teachers' });
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
    const responses = readResponses();
    res.json(responses);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to fetch responses' });
  }
});

// API: get unsent responses (for Email Notifications tab)
app.get('/api/admin/unsent-responses', (req, res) => {
  try {
    const responses = readResponses();
    // Return unsent responses along with their original index so the client can reference them
    function isSentValue(v) {
      if (v === true) return true;
      if (typeof v === 'number') return v === 1;
      if (!v && v !== 0) return false;
      const s = String(v).trim().toLowerCase();
      return ['true','1','yes','y'].includes(s);
    }
    const unsent = responses
      .map((r, i) => ({ __index: i, ...r }))
      .filter(r => !isSentValue(r.sent));
    res.json(unsent);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to fetch unsent responses' });
  }
});

// Helper: extract teacher email from name
// Assumes format "LastName, FirstName" or "FirstName LastName" or just last name
// Email format: lastnamefirstinitial@dlshs.org (e.g., smithj@dlshs.org)
function getTeacherEmail(teacherName) {
  if (!teacherName) return null;
  const name = String(teacherName).trim();
  let lastName = '', firstName = '';
  
  if (name.includes(',')) {
    // Format: "LastName, FirstName"
    const parts = name.split(',').map(p => p.trim());
    lastName = parts[0];
    firstName = parts[1] || '';
  } else if (name.includes(' ')) {
    // Format: "FirstName LastName"
    const parts = name.trim().split(/\s+/);
    firstName = parts[0];
    lastName = parts.slice(1).join(' ');
  } else {
    // Just a last name
    lastName = name;
  }
  
  if (!lastName) return null;
  const firstInitial = firstName ? firstName.charAt(0).toLowerCase() : '';
  const email = `${lastName.toLowerCase()}${firstInitial}@dlshs.org`;
  return email;
}

// API: send email notification to teacher for a test response
app.post('/api/admin/send-notification', (req, res) => {
  try {
    const { responseId, responseData } = req.body || {};
    // allow responseId === 0, so check for undefined/null instead
    if (responseId === undefined || responseData == null) {
      return res.status(400).json({ error: 'responseId and responseData required' });
    }
    
    // Use teacherEmail from the form if provided, otherwise calculate from teacherName
    let teacherEmail = responseData.teacherEmail;
    if (!teacherEmail) {
      teacherEmail = getTeacherEmail(responseData.teacherName);
    }
    
    if (!teacherEmail) {
      return res.status(400).json({ error: 'Teacher email is required' });
    }
    
    // Attempt to send email to teacher and only mark as sent on success
    (async () => {
      try {
        await sendTeacherNotification(responseData, teacherEmail);

        // Mark as sent in spreadsheet
        const responses = readResponses();
        const idx = responses.findIndex((r, i) => i === parseInt(responseId));
        if (idx >= 0) {
          responses[idx].sent = true;
          writeResponsesFile(responses);
        }

        res.json({ success: true, teacherEmail });
      } catch (err) {
        console.error('Failed to send teacher email:', err);
        res.status(500).json({ error: 'Failed to send teacher email', details: String(err.message || err) });
      }
    })();
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to send notification' });
  }
});

// API: mark a response as sent (admin UI can call this when admin manually sends the email)
app.post('/api/admin/mark-sent', (req, res) => {
  try {
    const { responseId } = req.body || {};
    if (responseId === undefined) {
      return res.status(400).json({ error: 'responseId required' });
    }

    const responses = readResponses();
    const idx = responses.findIndex((r, i) => i === parseInt(responseId));
    if (idx >= 0) {
      responses[idx].sent = true;
      writeResponsesFile(responses);
      return res.json({ success: true });
    }

    return res.status(404).json({ error: 'Response not found' });
  } catch (err) {
    console.error('Failed to mark response sent:', err);
    res.status(500).json({ error: 'Failed to mark response sent' });
  }
});

// Helper: send notification email to teacher
async function sendTeacherNotification(response, teacherEmail) {
  // Determine transport. Throw if not configured so the admin UI receives an error.
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
    throw new Error('SMTP not configured. Set SMTP_HOST/SMTP_USER or SENDMAIL=true');
  }

  // Build plain-text summary including all response fields
  const lines = [];
  lines.push('Test Request Notification');
  lines.push('========================');
  for (const k of Object.keys(response || {})) {
    lines.push(`${k}: ${response[k]}`);
  }
  lines.push(`sentAt: ${new Date().toISOString()}`);
  const textBody = lines.join('\n');

  // Attach responses.xlsx when available so teacher can open workbook if needed
  const attachments = [];
  if (fs.existsSync(RESPONSES_XLSX)) {
    attachments.push({ filename: 'responses.xlsx', path: RESPONSES_XLSX });
  }

  await transporter.sendMail({
    from: process.env.EMAIL_FROM || (process.env.SMTP_USER || 'no-reply@example.com'),
    to: teacherEmail,
    subject: `Test Request: ${response.testName || 'Unknown Test'} - ${response.firstName || ''} ${response.lastName || ''}`,
    text: textBody,
    attachments
  });
}

// API: add a new student (admin only) - appends to students.xlsx
app.post('/api/admin/add-student', (req, res) => {
  try {
    const { firstName, lastName, grade } = req.body || {};
    if (!firstName || !lastName || !grade) return res.status(400).json({ error: 'firstName,lastName,grade required' });

    // Read existing students using the normalized reader, append, and write using expected headers
    let students = [];
    if (fs.existsSync(STUDENTS_XLSX)) {
      students = readStudents();
    }

    const newStudent = { firstName: String(firstName), lastName: String(lastName), grade: String(grade), loggedIn: false };
    students.push(newStudent);

    writeStudentsFile(students);

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
