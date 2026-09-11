const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Memory storage for file uploads (handles serverless/cloud environments cleanly)
const storage = multer.memoryStorage();
const upload = multer({
  storage,
  limits: { fileSize: 25 * 1024 * 1024 } // 25MB max
});

// Health check endpoint for cloud hosting platforms (Render, Railway, etc.)
app.get('/health', (req, res) => {
  res.status(200).json({ status: 'healthy', uptime: process.uptime() });
});

// Campaign State in Memory
let activeCampaign = {
  id: null,
  status: 'idle', // 'idle' | 'running' | 'paused' | 'stopped' | 'completed'
  total: 0,
  sent: 0,
  failed: 0,
  pending: 0,
  currentIndex: 0,
  startTime: null,
  logs: [],
  recipients: [],
  config: null,
  pausePromiseResolver: null
};

// SSE Clients for real-time progress updates
let sseClients = [];

function broadcastProgress(data) {
  const payload = `data: ${JSON.stringify(data)}\n\n`;
  sseClients.forEach((client) => {
    try {
      client.res.write(payload);
    } catch (err) {
      console.error('SSE client write error:', err);
    }
  });
}

function appendLog(type, message, details = {}) {
  const logEntry = {
    id: Date.now() + Math.random().toString(36).substr(2, 5),
    timestamp: new Date().toLocaleTimeString(),
    type, // 'info' | 'success' | 'error' | 'warning'
    message,
    details
  };
  activeCampaign.logs.unshift(logEntry);
  if (activeCampaign.logs.length > 500) {
    activeCampaign.logs.pop();
  }
  broadcastProgress({ type: 'log', log: logEntry, campaign: getCampaignSnapshot() });
}

function getCampaignSnapshot() {
  return {
    id: activeCampaign.id,
    status: activeCampaign.status,
    total: activeCampaign.total,
    sent: activeCampaign.sent,
    failed: activeCampaign.failed,
    pending: activeCampaign.pending,
    currentIndex: activeCampaign.currentIndex,
    startTime: activeCampaign.startTime
  };
}

// ----------------------------------------------------
// Helper: Email Syntax Validation
// ----------------------------------------------------
function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email.trim());
}

// ----------------------------------------------------
// Helper: Merge Variable Replacer (e.g., {{Name}} or {Name})
// ----------------------------------------------------
function replaceVariables(template, dataRow) {
  if (!template) return '';
  let result = template;
  for (const [key, val] of Object.entries(dataRow)) {
    const safeVal = val !== undefined && val !== null ? String(val) : '';
    // Replace {{key}}, {key}, {{ key }}
    const regex = new RegExp(`\\{\\{\\s*${key}\\s*\\}\\}|\\{\\s*${key}\\s*\\}`, 'gi');
    result = result.replace(regex, safeVal);
  }
  return result;
}

// ----------------------------------------------------
// Helper: Create Nodemailer Transporter
// ----------------------------------------------------
function createTransporter(smtpConfig) {
  const port = parseInt(smtpConfig.port, 10) || 587;
  const isSecure = smtpConfig.secure === true || smtpConfig.secure === 'true' || port === 465;

  const transportOptions = {
    host: smtpConfig.host,
    port: port,
    secure: isSecure, // true for 465, false for 587 / other
    auth: {
      user: smtpConfig.user,
      pass: smtpConfig.pass
    },
    tls: {
      rejectUnauthorized: smtpConfig.rejectUnauthorized !== false
    },
    pool: true,
    maxConnections: 5,
    maxMessages: 100
  };

  return nodemailer.createTransport(transportOptions);
}

// ----------------------------------------------------
// ENDPOINT: SSE Stream for Real-time Dashboard Updates
// ----------------------------------------------------
app.get('/api/campaign/events', (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders();

  const clientId = Date.now();
  const newClient = { id: clientId, res };
  sseClients.push(newClient);

  // Send initial snapshot
  res.write(`data: ${JSON.stringify({ type: 'snapshot', campaign: getCampaignSnapshot(), logs: activeCampaign.logs })}\n\n`);

  req.on('close', () => {
    sseClients = sseClients.filter((c) => c.id !== clientId);
  });
});

// ----------------------------------------------------
// ENDPOINT: Parse Excel / CSV File (Direct Memory Parsing)
// ----------------------------------------------------
app.post('/api/parse-excel', upload.single('file'), (req, res) => {
  try {
    if (!req.file || !req.file.buffer) {
      return res.status(400).json({ error: 'No file uploaded or file is empty.' });
    }

    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rawData = xlsx.utils.sheet_to_json(sheet, { defval: '' });

    if (!rawData || rawData.length === 0) {
      return res.status(400).json({ error: 'The uploaded file is empty or invalid.' });
    }

    // Identify columns
    const columns = Object.keys(rawData[0]);

    // Auto-detect email column
    let detectedEmailCol = columns.find((c) => /email|e-mail|mail|recipient/i.test(c.trim()));
    if (!detectedEmailCol) {
      // Check first 5 rows for email values
      for (const col of columns) {
        if (rawData.slice(0, 5).some((row) => isValidEmail(row[col]))) {
          detectedEmailCol = col;
          break;
        }
      }
    }
    if (!detectedEmailCol) {
      detectedEmailCol = columns[0];
    }

    // Auto-detect name column
    const detectedNameCol = columns.find((c) => /name|firstname|first_name|client|contact/i.test(c.trim())) || '';

    // Extract valid and invalid count
    let validCount = 0;
    let invalidCount = 0;
    const sampleRows = [];

    rawData.forEach((row, idx) => {
      const emailVal = String(row[detectedEmailCol] || '').trim();
      const valid = isValidEmail(emailVal);
      if (valid) validCount++;
      else invalidCount++;

      if (idx < 50) {
        sampleRows.push({
          ...row,
          __isValidEmail: valid,
          __rowIndex: idx + 1
        });
      }
    });

    res.json({
      success: true,
      fileName: req.file.originalname,
      totalRows: rawData.length,
      validEmails: validCount,
      invalidEmails: invalidCount,
      columns: columns,
      detectedEmailColumn: detectedEmailCol,
      detectedNameColumn: detectedNameCol,
      sampleRows: sampleRows,
      fullRows: rawData
    });
  } catch (err) {
    console.error('Error parsing Excel:', err);
    res.status(500).json({ error: `Failed to parse Excel file: ${err.message}` });
  }
});

// ----------------------------------------------------
// ENDPOINT: Test SMTP Connection & Send Test Email
// ----------------------------------------------------
app.post('/api/test-smtp', async (req, res) => {
  try {
    const { smtp, testEmail, sampleData } = req.body;

    if (!smtp || !smtp.host || !smtp.user || !smtp.pass) {
      return res.status(400).json({ error: 'Missing required SMTP configuration (Host, User, Password).' });
    }

    const transporter = createTransporter(smtp);

    // 1. Verify connection
    await transporter.verify();

    // 2. If test email is requested, send a sample email
    let testSent = false;
    if (testEmail && isValidEmail(testEmail)) {
      const demoData = sampleData || { Name: 'Demo Client', Company: 'Acme Corp', Email: testEmail };
      const fromName = smtp.fromName ? `"${smtp.fromName}" <${smtp.fromEmail || smtp.user}>` : (smtp.fromEmail || smtp.user);

      await transporter.sendMail({
        from: fromName,
        to: testEmail,
        subject: `[Test] ${replaceVariables(req.body.subject || 'Email Automation Test', demoData)}`,
        html: replaceVariables(req.body.htmlBody || '<p>Hello <strong>{{Name}}</strong>, your email automation setup is working perfectly! 🚀</p>', demoData),
        text: replaceVariables(req.body.textBody || 'Hello {{Name}}, your email automation setup is working perfectly!', demoData)
      });
      testSent = true;
    }

    res.json({
      success: true,
      message: testSent
        ? `SMTP Connected successfully! Test email dispatched to ${testEmail}.`
        : 'SMTP Connection verified successfully!'
    });
  } catch (err) {
    console.error('SMTP test error:', err);
    res.status(500).json({
      error: `SMTP Connection Failed: ${err.message}. Please check your Host, Port, and App Password.`
    });
  }
});

// ----------------------------------------------------
// ENDPOINT: Start Campaign
// ----------------------------------------------------
app.post('/api/campaign/start', upload.array('attachments', 10), async (req, res) => {
  try {
    if (activeCampaign.status === 'running') {
      return res.status(400).json({ error: 'A campaign is already running. Please stop or wait for it to finish.' });
    }

    const {
      smtp: smtpStr,
      emailColumn,
      subject,
      htmlBody,
      textBody,
      recipients: recipientsStr,
      delaySeconds = 1
    } = req.body;

    const smtp = typeof smtpStr === 'string' ? JSON.parse(smtpStr) : smtpStr;
    const recipients = typeof recipientsStr === 'string' ? JSON.parse(recipientsStr) : recipientsStr;

    if (!smtp || !smtp.host || !smtp.user || !smtp.pass) {
      return res.status(400).json({ error: 'Incomplete SMTP settings.' });
    }
    if (!recipients || !Array.isArray(recipients) || recipients.length === 0) {
      return res.status(400).json({ error: 'No recipients provided.' });
    }
    if (!subject) {
      return res.status(400).json({ error: 'Subject is required.' });
    }

    // Process attachments from memory
    const attachments = (req.files || []).map((file) => ({
      filename: file.originalname,
      content: file.buffer
    }));

    // Filter valid recipients
    const validRecipients = recipients.filter((r) => isValidEmail(r[emailColumn]));
    if (validRecipients.length === 0) {
      return res.status(400).json({ error: 'No valid email addresses found in the selected column.' });
    }

    // Reset Campaign State
    activeCampaign = {
      id: `camp_${Date.now()}`,
      status: 'running',
      total: validRecipients.length,
      sent: 0,
      failed: 0,
      pending: validRecipients.length,
      currentIndex: 0,
      startTime: new Date().toISOString(),
      logs: [],
      recipients: validRecipients.map((r, i) => ({
        index: i + 1,
        email: String(r[emailColumn]).trim(),
        data: r,
        status: 'pending',
        error: null,
        sentAt: null
      })),
      config: {
        smtp,
        emailColumn,
        subject,
        htmlBody,
        textBody,
        delayMs: Math.max(200, (parseFloat(delaySeconds) || 1) * 1000),
        attachments
      },
      pausePromiseResolver: null
    };

    appendLog('info', `Campaign started: ${activeCampaign.total} recipients queued.`);
    res.json({ success: true, message: 'Campaign initiated', campaignId: activeCampaign.id });

    // Execute background sending queue
    runCampaignQueue();
  } catch (err) {
    console.error('Error starting campaign:', err);
    res.status(500).json({ error: `Failed to start campaign: ${err.message}` });
  }
});

// Background queue processor
async function runCampaignQueue() {
  const { smtp, subject, htmlBody, textBody, delayMs, attachments } = activeCampaign.config;
  let transporter;

  try {
    transporter = createTransporter(smtp);
  } catch (err) {
    appendLog('error', `Failed to initialize SMTP transporter: ${err.message}`);
    activeCampaign.status = 'failed';
    broadcastProgress({ type: 'status', campaign: getCampaignSnapshot() });
    return;
  }

  const fromName = smtp.fromName ? `"${smtp.fromName}" <${smtp.fromEmail || smtp.user}>` : (smtp.fromEmail || smtp.user);

  for (let i = activeCampaign.currentIndex; i < activeCampaign.recipients.length; i++) {
    // Check if campaign was stopped
    if (activeCampaign.status === 'stopped') {
      appendLog('warning', `Campaign manually stopped at #${i + 1}.`);
      break;
    }

    // Check if campaign was paused
    if (activeCampaign.status === 'paused') {
      appendLog('info', `Campaign paused at #${i + 1}. Waiting for resume...`);
      await new Promise((resolve) => {
        activeCampaign.pausePromiseResolver = resolve;
      });
      if (activeCampaign.status === 'stopped') break;
    }

    activeCampaign.currentIndex = i;
    const recipient = activeCampaign.recipients[i];
    const clientData = recipient.data;
    const recipientEmail = recipient.email;

    try {
      // Personalized subject & body
      const personalizedSubject = replaceVariables(subject, clientData);
      const personalizedHtml = htmlBody ? replaceVariables(htmlBody, clientData) : undefined;
      const personalizedText = textBody ? replaceVariables(textBody, clientData) : undefined;

      await transporter.sendMail({
        from: fromName,
        to: recipientEmail,
        subject: personalizedSubject,
        html: personalizedHtml,
        text: personalizedText,
        attachments: attachments.length > 0 ? attachments : undefined
      });

      recipient.status = 'sent';
      recipient.sentAt = new Date().toLocaleTimeString();
      activeCampaign.sent++;
      activeCampaign.pending--;

      appendLog('success', `[#${i + 1}/${activeCampaign.total}] Sent successfully to ${recipientEmail}`);
    } catch (sendErr) {
      recipient.status = 'failed';
      recipient.error = sendErr.message || 'Unknown SMTP error';
      activeCampaign.failed++;
      activeCampaign.pending--;

      appendLog('error', `[#${i + 1}/${activeCampaign.total}] Failed for ${recipientEmail}: ${recipient.error}`);
    }

    broadcastProgress({
      type: 'progress',
      campaign: getCampaignSnapshot(),
      currentRecipient: recipient
    });

    // Rate Limiting Throttling Delay
    if (i < activeCampaign.recipients.length - 1 && activeCampaign.status === 'running') {
      await new Promise((res) => setTimeout(res, delayMs));
    }
  }

  if (activeCampaign.status === 'running') {
    activeCampaign.status = 'completed';
    appendLog('info', `Campaign completed! Sent: ${activeCampaign.sent}, Failed: ${activeCampaign.failed}`);
  }

  broadcastProgress({ type: 'completed', campaign: getCampaignSnapshot() });
}

// ----------------------------------------------------
// ENDPOINT: Pause Campaign
// ----------------------------------------------------
app.post('/api/campaign/pause', (req, res) => {
  if (activeCampaign.status === 'running') {
    activeCampaign.status = 'paused';
    appendLog('warning', 'Campaign paused by user.');
    broadcastProgress({ type: 'status', campaign: getCampaignSnapshot() });
    return res.json({ success: true, message: 'Campaign paused' });
  }
  res.status(400).json({ error: 'No active running campaign to pause' });
});

// ----------------------------------------------------
// ENDPOINT: Resume Campaign
// ----------------------------------------------------
app.post('/api/campaign/resume', (req, res) => {
  if (activeCampaign.status === 'paused') {
    activeCampaign.status = 'running';
    appendLog('info', 'Campaign resumed by user.');
    if (activeCampaign.pausePromiseResolver) {
      activeCampaign.pausePromiseResolver();
      activeCampaign.pausePromiseResolver = null;
    }
    broadcastProgress({ type: 'status', campaign: getCampaignSnapshot() });
    return res.json({ success: true, message: 'Campaign resumed' });
  }
  res.status(400).json({ error: 'Campaign is not paused' });
});

// ----------------------------------------------------
// ENDPOINT: Stop Campaign
// ----------------------------------------------------
app.post('/api/campaign/stop', (req, res) => {
  if (activeCampaign.status === 'running' || activeCampaign.status === 'paused') {
    activeCampaign.status = 'stopped';
    appendLog('warning', 'Campaign cancelled by user.');
    if (activeCampaign.pausePromiseResolver) {
      activeCampaign.pausePromiseResolver();
      activeCampaign.pausePromiseResolver = null;
    }
    broadcastProgress({ type: 'status', campaign: getCampaignSnapshot() });
    return res.json({ success: true, message: 'Campaign stopped' });
  }
  res.status(400).json({ error: 'No active campaign to stop' });
});

// ----------------------------------------------------
// ENDPOINT: Get Current Campaign Status & Details
// ----------------------------------------------------
app.get('/api/campaign/status', (req, res) => {
  res.json({
    campaign: getCampaignSnapshot(),
    recipients: activeCampaign.recipients || [],
    logs: activeCampaign.logs || []
  });
});

// ----------------------------------------------------
// ENDPOINT: Export Campaign Summary as CSV
// ----------------------------------------------------
app.get('/api/campaign/export-csv', (req, res) => {
  if (!activeCampaign.recipients || activeCampaign.recipients.length === 0) {
    return res.status(400).json({ error: 'No campaign data to export' });
  }

  const exportData = activeCampaign.recipients.map((r) => ({
    Index: r.index,
    Email: r.email,
    Status: r.status.toUpperCase(),
    Error: r.error || '',
    SentAt: r.sentAt || '',
    ...r.data
  }));

  const worksheet = xlsx.utils.json_to_sheet(exportData);
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Campaign_Results');

  const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'csv' });

  res.setHeader('Content-Disposition', 'attachment; filename="campaign_report.csv"');
  res.setHeader('Content-Type', 'text/csv');
  res.send(buffer);
});

// ----------------------------------------------------
// Generate Sample Excel File for user convenience
// ----------------------------------------------------
app.get('/api/download-sample-excel', (req, res) => {
  const sampleData = [
    { Name: 'John Smith', Email: 'john.smith@example.com', Company: 'Apex Innovations', Role: 'Product Manager' },
    { Name: 'Sarah Connor', Email: 'sarah.c@example.com', Company: 'Cyberdyne Systems', Role: 'Tech Lead' },
    { Name: 'Michael Scott', Email: 'michael@dundermifflin.com', Company: 'Dunder Mifflin', Role: 'Regional Manager' },
    { Name: 'Elena Rostova', Email: 'elena.rostova@techglobal.io', Company: 'TechGlobal', Role: 'CEO' },
    { Name: 'David Chen', Email: 'dchen@innovate.co', Company: 'Innovate Co', Role: 'Marketing Director' }
  ];

  const ws = xlsx.utils.json_to_sheet(sampleData);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, 'Clients');

  const buf = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', 'attachment; filename="sample_clients.xlsx"');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// SPA Fallback
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Start Server binding to 0.0.0.0 for Cloud Container support
app.listen(PORT, '0.0.0.0', () => {
  console.log(`=======================================================`);
  console.log(`🚀 Bulk Email Automation Server running on port ${PORT}`);
  console.log(`=======================================================`);
});
