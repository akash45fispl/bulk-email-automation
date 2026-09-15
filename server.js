const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const db = require('./database');

const app = express();
app.set('trust proxy', 1);
const PORT = process.env.PORT || 3000;

// Ensure uploads folder exists
const UPLOADS_DIR = path.join(__dirname, 'uploads');
if (!fs.existsSync(UPLOADS_DIR)) {
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}

// Initialize SQLite database
db.initDb().catch((err) => console.error('Database initialization error:', err));

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Memory storage for file uploads
const storage = multer.memoryStorage();
const upload = multer({
  storage,
  limits: { fileSize: 35 * 1024 * 1024 } // 35MB max
});

// 1x1 Transparent GIF buffer for tracking pixel
const TRACKING_PIXEL_BUFFER = Buffer.from(
  'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7',
  'base64'
);

// Health check endpoint
app.get('/health', (req, res) => {
  res.status(200).json({ status: 'healthy', uptime: process.uptime() });
});

// Campaign State in Memory for active live streaming
let activeCampaign = {
  id: null,
  name: '',
  status: 'idle', // 'idle' | 'running' | 'paused' | 'stopped' | 'completed'
  total: 0,
  sent: 0,
  failed: 0,
  pending: 0,
  opened: 0,
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
    name: activeCampaign.name,
    status: activeCampaign.status,
    total: activeCampaign.total,
    sent: activeCampaign.sent,
    failed: activeCampaign.failed,
    pending: activeCampaign.pending,
    opened: activeCampaign.opened,
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
    const regex = new RegExp(`\\{\\{\\s*${key}\\s*\\}\\}|\\{\\s*${key}\\s*\\}`, 'gi');
    result = result.replace(regex, safeVal);
  }
  return result;
}

// ----------------------------------------------------
// Helper: Create Nodemailer Transporter
// ----------------------------------------------------
function createTransporter(smtpConfig) {
  const host = String(smtpConfig.host || '').toLowerCase().trim();
  const user = String(smtpConfig.user || '').trim();
  const rawPass = String(smtpConfig.pass || '').trim();
  const pass = rawPass.replace(/\s+/g, '');

  if (host.includes('gmail') || user.endsWith('@gmail.com') || user.endsWith('@googlemail.com')) {
    return nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: user,
        pass: pass
      },
      pool: true,
      maxConnections: 5,
      maxMessages: 100
    });
  }

  const port = parseInt(smtpConfig.port, 10) || 587;
  const isSecure = smtpConfig.secure === true || smtpConfig.secure === 'true' || port === 465;

  const transportOptions = {
    host: host,
    port: port,
    secure: isSecure,
    auth: {
      user: user,
      pass: pass
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

  res.write(`data: ${JSON.stringify({ type: 'snapshot', campaign: getCampaignSnapshot(), logs: activeCampaign.logs })}\n\n`);

  req.on('close', () => {
    sseClients = sseClients.filter((c) => c.id !== clientId);
  });
});

// ----------------------------------------------------
// ENDPOINT: Parse Excel / CSV File
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

    const columns = Object.keys(rawData[0]);

    let detectedEmailCol = columns.find((c) => /email|e-mail|mail|recipient/i.test(c.trim()));
    if (!detectedEmailCol) {
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

    const detectedNameCol = columns.find((c) => /name|firstname|first_name|client|contact/i.test(c.trim())) || '';

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
    await transporter.verify();

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
// ENDPOINT: Start Campaign (With SQLite & Excel archiving)
// ----------------------------------------------------
app.post('/api/campaign/start', upload.fields([
  { name: 'attachments', maxCount: 10 },
  { name: 'excelFile', maxCount: 1 }
]), async (req, res) => {
  try {
    if (activeCampaign.status === 'running') {
      return res.status(400).json({ error: 'A campaign is already running. Please stop or wait for it to finish.' });
    }

    const {
      campaignName,
      publicBaseUrl,
      smtp: smtpStr,
      emailColumn,
      subject,
      htmlBody,
      textBody,
      recipients: recipientsStr,
      delaySeconds = 1,
      originalFilename
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

    const files = req.files || {};
    const attachmentFiles = files['attachments'] || [];
    const excelFiles = files['excelFile'] || [];

    const attachments = attachmentFiles.map((file) => ({
      filename: file.originalname,
      content: file.buffer
    }));

    const validRecipients = recipients.filter((r) => isValidEmail(r[emailColumn]));
    if (validRecipients.length === 0) {
      return res.status(400).json({ error: 'No valid email addresses found in the selected column.' });
    }

    const campaignId = `camp_${Date.now()}`;
    const cleanCampaignName = (campaignName || `Dump_${new Date().toISOString().slice(0, 10)}`).trim();

    // Save Excel file to uploads/ for permanent storage
    let savedExcelPath = '';
    const origName = originalFilename || (excelFiles[0] ? excelFiles[0].originalname : 'recipients.xlsx');
    if (excelFiles.length > 0 && excelFiles[0].buffer) {
      savedExcelPath = path.join(UPLOADS_DIR, `${campaignId}_${origName}`);
      fs.writeFileSync(savedExcelPath, excelFiles[0].buffer);
    } else {
      // Create workbook from recipients if raw file wasn't re-attached
      const ws = xlsx.utils.json_to_sheet(recipients);
      const wb = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(wb, ws, 'Recipients');
      savedExcelPath = path.join(UPLOADS_DIR, `${campaignId}_recipients.xlsx`);
      xlsx.writeFile(wb, savedExcelPath);
    }

    // Determine Base URL for tracking pixel
    let baseUrl = (publicBaseUrl || '').trim().replace(/\/+$/, '');
    if (!baseUrl) {
      baseUrl = `${req.protocol}://${req.get('host')}`;
    }

    // Persist Campaign in SQLite
    const fromStr = smtp.fromName ? `"${smtp.fromName}" <${smtp.fromEmail || smtp.user}>` : (smtp.fromEmail || smtp.user);
    await db.createCampaign({
      id: campaignId,
      name: cleanCampaignName,
      originalFilename: origName,
      excelFilePath: savedExcelPath,
      totalRecipients: validRecipients.length,
      subject,
      htmlBody,
      textBody,
      attachmentNames: attachments.map((a) => a.filename),
      smtpFrom: fromStr,
      status: 'running'
    });

    const preparedRecipients = validRecipients.map((r, i) => ({
      index: i + 1,
      email: String(r[emailColumn]).trim(),
      name: r.Name || r.name || r.FirstName || '',
      data: r,
      status: 'pending',
      error: null,
      sentAt: null,
      isOpened: 0,
      openedAt: null
    }));

    await db.addCampaignRecipients(campaignId, preparedRecipients);

    // Reset In-Memory State for live streaming
    activeCampaign = {
      id: campaignId,
      name: cleanCampaignName,
      status: 'running',
      total: validRecipients.length,
      sent: 0,
      failed: 0,
      pending: validRecipients.length,
      opened: 0,
      currentIndex: 0,
      startTime: new Date().toISOString(),
      logs: [],
      recipients: preparedRecipients,
      config: {
        smtp,
        emailColumn,
        subject,
        htmlBody,
        textBody,
        delayMs: Math.max(200, (parseFloat(delaySeconds) || 1) * 1000),
        attachments,
        baseUrl
      },
      pausePromiseResolver: null
    };

    appendLog('info', `Campaign "${cleanCampaignName}" initiated: ${activeCampaign.total} recipients queued.`);
    res.json({ success: true, message: 'Campaign initiated', campaignId });

    // Start background worker
    runCampaignQueue();
  } catch (err) {
    console.error('Error starting campaign:', err);
    res.status(500).json({ error: `Failed to start campaign: ${err.message}` });
  }
});

// Background queue processor
async function runCampaignQueue() {
  const { smtp, subject, htmlBody, textBody, delayMs, attachments, baseUrl } = activeCampaign.config;
  let transporter;

  try {
    transporter = createTransporter(smtp);
  } catch (err) {
    appendLog('error', `Failed to initialize SMTP transporter: ${err.message}`);
    activeCampaign.status = 'failed';
    await db.updateCampaignStatus(activeCampaign.id, 'failed');
    broadcastProgress({ type: 'status', campaign: getCampaignSnapshot() });
    return;
  }

  const fromName = smtp.fromName ? `"${smtp.fromName}" <${smtp.fromEmail || smtp.user}>` : (smtp.fromEmail || smtp.user);

  for (let i = activeCampaign.currentIndex; i < activeCampaign.recipients.length; i++) {
    if (activeCampaign.status === 'stopped') {
      appendLog('warning', `Campaign manually stopped at #${i + 1}.`);
      await db.updateCampaignStatus(activeCampaign.id, 'stopped');
      break;
    }

    if (activeCampaign.status === 'paused') {
      appendLog('info', `Campaign paused at #${i + 1}. Waiting for resume...`);
      await db.updateCampaignStatus(activeCampaign.id, 'paused');
      await new Promise((resolve) => {
        activeCampaign.pausePromiseResolver = resolve;
      });
      if (activeCampaign.status === 'stopped') {
        await db.updateCampaignStatus(activeCampaign.id, 'stopped');
        break;
      }
      await db.updateCampaignStatus(activeCampaign.id, 'running');
    }

    activeCampaign.currentIndex = i;
    const recipient = activeCampaign.recipients[i];
    const clientData = recipient.data;
    const recipientEmail = recipient.email;

    try {
      const personalizedSubject = replaceVariables(subject, clientData);
      let personalizedHtml = htmlBody ? replaceVariables(htmlBody, clientData) : '';
      
      // Inject Open Tracking Pixel into HTML email
      const trackingUrl = `${baseUrl}/api/track/open/${activeCampaign.id}/${recipient.index}`;
      const trackingPixelHtml = `\n<img src="${trackingUrl}" alt="" width="1" height="1" style="display:none !important; min-height:1px !important; min-width:1px !important; border:0 !important; outline:none !important;" />`;
      
      if (personalizedHtml) {
        if (personalizedHtml.includes('</body>')) {
          personalizedHtml = personalizedHtml.replace('</body>', `${trackingPixelHtml}</body>`);
        } else {
          personalizedHtml += trackingPixelHtml;
        }
      } else {
        personalizedHtml = `<p>${personalizedSubject}</p>${trackingPixelHtml}`;
      }

      let personalizedText = textBody ? replaceVariables(textBody, clientData) : undefined;
      if (!personalizedText && personalizedHtml) {
        personalizedText = personalizedHtml
          .replace(/<br\s*[\/]?>/gi, '\n')
          .replace(/<\/p>/gi, '\n\n')
          .replace(/<[^>]+>/gi, '')
          .replace(/&nbsp;/g, ' ')
          .replace(/&amp;/g, '&')
          .trim();
      }

      await transporter.sendMail({
        from: fromName,
        to: recipientEmail,
        replyTo: smtp.fromEmail || smtp.user,
        subject: personalizedSubject,
        html: personalizedHtml,
        text: personalizedText,
        headers: {
          'List-Unsubscribe': `<mailto:${smtp.user}?subject=unsubscribe>`,
          'X-Mailer': 'AutoMailer PRO',
          'X-Campaign-ID': activeCampaign.id,
          'X-Recipient-ID': String(recipient.index)
        },
        attachments: attachments.length > 0 ? attachments : undefined
      });

      recipient.status = 'sent';
      recipient.sentAt = new Date().toLocaleTimeString();
      activeCampaign.sent++;
      activeCampaign.pending--;

      // Update SQLite record
      await db.updateRecipientSendStatus(activeCampaign.id, recipient.index, 'sent');
      appendLog('success', `[#${i + 1}/${activeCampaign.total}] Sent successfully to ${recipientEmail}`);
    } catch (sendErr) {
      recipient.status = 'failed';
      recipient.error = sendErr.message || 'Unknown SMTP error';
      activeCampaign.failed++;
      activeCampaign.pending--;

      // Update SQLite record with error
      await db.updateRecipientSendStatus(activeCampaign.id, recipient.index, 'failed', recipient.error);
      appendLog('error', `[#${i + 1}/${activeCampaign.total}] Failed for ${recipientEmail}: ${recipient.error}`);
    }

    broadcastProgress({
      type: 'progress',
      campaign: getCampaignSnapshot(),
      currentRecipient: recipient
    });

    // Anti-Spam Rate Limiting
    if (i < activeCampaign.recipients.length - 1 && activeCampaign.status === 'running') {
      const jitter = Math.floor(Math.random() * 1000) - 300;
      const actualDelay = Math.max(800, delayMs + jitter);
      await new Promise((res) => setTimeout(res, actualDelay));
    }
  }

  if (activeCampaign.status === 'running') {
    activeCampaign.status = 'completed';
    await db.updateCampaignStatus(activeCampaign.id, 'completed');
    appendLog('info', `Campaign completed! Sent: ${activeCampaign.sent}, Failed: ${activeCampaign.failed}`);
  }

  broadcastProgress({ type: 'completed', campaign: getCampaignSnapshot() });
}

// ----------------------------------------------------
// ENDPOINT: Open Tracking Pixel
// ----------------------------------------------------
app.get('/api/track/open/:campaignId/:recipientId', async (req, res) => {
  const { campaignId, recipientId } = req.params;

  try {
    const result = await db.recordEmailOpen(campaignId, recipientId);
    
    if (result && result.recorded && result.firstOpen) {
      if (activeCampaign.id === campaignId) {
        activeCampaign.opened++;
      }
      
      broadcastProgress({
        type: 'open_event',
        campaignId,
        recipientId,
        openedAt: new Date().toLocaleTimeString(),
        email: result.recipient ? result.recipient.email : ''
      });
    }
  } catch (err) {
    console.error('Tracking pixel recording error:', err.message);
  }

  // Return 1x1 transparent GIF with no-cache headers
  res.setHeader('Content-Type', 'image/gif');
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate, max-age=0');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  res.status(200).send(TRACKING_PIXEL_BUFFER);
});

// ----------------------------------------------------
// ENDPOINT: Pause / Resume / Stop Campaign
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

app.post('/api/campaign/stop', (req, res) => {
  activeCampaign.status = 'stopped';
  appendLog('warning', 'Campaign reset / stopped by user.');
  if (activeCampaign.pausePromiseResolver) {
    activeCampaign.pausePromiseResolver();
    activeCampaign.pausePromiseResolver = null;
  }
  broadcastProgress({ type: 'status', campaign: getCampaignSnapshot() });
  return res.json({ success: true, message: 'Campaign stopped' });
});

app.get('/api/campaign/status', (req, res) => {
  res.json({
    campaign: getCampaignSnapshot(),
    recipients: activeCampaign.recipients || [],
    logs: activeCampaign.logs || []
  });
});

// ----------------------------------------------------
// ENDPOINTS: Campaign History & Analytics
// ----------------------------------------------------
app.get('/api/campaigns', async (req, res) => {
  try {
    const campaigns = await db.getAllCampaigns();
    const stats = await db.getOverallStats();
    res.json({
      success: true,
      stats,
      campaigns
    });
  } catch (err) {
    console.error('Error fetching campaigns:', err);
    res.status(500).json({ error: `Failed to fetch campaigns: ${err.message}` });
  }
});

app.get('/api/analytics/recipients', async (req, res) => {
  try {
    const type = req.query.type || 'all';
    const recipients = await db.getAggregateRecipients(type);
    res.json({ success: true, count: recipients.length, recipients });
  } catch (err) {
    console.error('Error fetching aggregate recipients:', err);
    res.status(500).json({ error: `Failed to fetch recipients: ${err.message}` });
  }
});


app.get('/api/campaigns/:id', async (req, res) => {
  try {
    const campaign = await db.getCampaignById(req.params.id);
    if (!campaign) {
      return res.status(404).json({ error: 'Campaign not found' });
    }
    res.json({ success: true, campaign });
  } catch (err) {
    console.error('Error fetching campaign detail:', err);
    res.status(500).json({ error: `Failed to fetch campaign: ${err.message}` });
  }
});

app.delete('/api/campaigns/:id', async (req, res) => {
  try {
    await db.deleteCampaign(req.params.id);
    res.json({ success: true, message: 'Campaign record deleted successfully' });
  } catch (err) {
    console.error('Error deleting campaign:', err);
    res.status(500).json({ error: `Failed to delete campaign: ${err.message}` });
  }
});

// Download original uploaded Excel sheet
app.get('/api/campaigns/:id/download-original', async (req, res) => {
  try {
    const campaign = await db.getCampaignById(req.params.id);
    if (!campaign) {
      return res.status(404).json({ error: 'Campaign not found' });
    }

    if (campaign.excel_file_path && fs.existsSync(campaign.excel_file_path)) {
      return res.download(campaign.excel_file_path, campaign.original_filename || 'campaign_data.xlsx');
    }

    // If file isn't on disk, regenerate from recipient row data
    const exportRows = (campaign.recipients || []).map((r) => r.rowData || {});
    const ws = xlsx.utils.json_to_sheet(exportRows);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');
    const buffer = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });

    res.setHeader('Content-Disposition', `attachment; filename="${campaign.original_filename || 'campaign_recipients.xlsx'}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
  } catch (err) {
    console.error('Error downloading original Excel:', err);
    res.status(500).json({ error: `Failed to download file: ${err.message}` });
  }
});

// Export Enriched Analytics Excel Report (with Seen/Unseen & Opened At)
app.get('/api/campaigns/:id/export-analytics', async (req, res) => {
  try {
    const campaign = await db.getCampaignById(req.params.id);
    if (!campaign) {
      return res.status(404).json({ error: 'Campaign not found' });
    }

    const rows = (campaign.recipients || []).map((r) => ({
      Index: r.recipient_index,
      Email: r.email,
      Name: r.name || '',
      Send_Status: (r.send_status || '').toUpperCase(),
      Sent_At: r.sent_at || '',
      Seen_Status: r.is_opened ? 'SEEN (OPENED)' : 'UNSEEN (UNOPENED)',
      First_Opened_At: r.opened_at || '',
      Open_Count: r.open_count || 0,
      Error_Log: r.error_message || '',
      ...(r.rowData || {})
    }));

    const ws = xlsx.utils.json_to_sheet(rows);
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'Campaign_Analytics');

    const buffer = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const cleanFilename = `${campaign.name.replace(/[^a-zA-Z0-9_-]/g, '_')}_Analytics_Report.xlsx`;

    res.setHeader('Content-Disposition', `attachment; filename="${cleanFilename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
  } catch (err) {
    console.error('Error exporting analytics report:', err);
    res.status(500).json({ error: `Failed to export analytics: ${err.message}` });
  }
});

// Generate Sample Excel File
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

// Start Server
if (process.env.NODE_ENV !== 'test') {
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`=======================================================`);
    console.log(`🚀 AutoMailer PRO Engine running on port ${PORT}`);
    console.log(`📊 SQLite Campaign & Open-Tracking Engine: ONLINE`);
    console.log(`=======================================================`);
  });
}

module.exports = app;
