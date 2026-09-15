const path = require('path');
const fs = require('fs');
const sqlite3 = require('sqlite3').verbose();

const DB_PATH = path.join(__dirname, 'campaigns.db');
const db = new sqlite3.Database(DB_PATH);

// Helper promise wrapper for db operations
function runQuery(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.run(sql, params, function (err) {
      if (err) reject(err);
      else resolve({ lastID: this.lastID, changes: this.changes });
    });
  });
}

function getQuery(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.get(sql, params, (err, row) => {
      if (err) reject(err);
      else resolve(row);
    });
  });
}

function allQuery(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.all(sql, params, (err, rows) => {
      if (err) reject(err);
      else resolve(rows || []);
    });
  });
}

// Initialize Tables
async function initDb() {
  await runQuery(`
    CREATE TABLE IF NOT EXISTS campaigns (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
      original_filename TEXT,
      excel_file_path TEXT,
      total_recipients INTEGER DEFAULT 0,
      sent_count INTEGER DEFAULT 0,
      failed_count INTEGER DEFAULT 0,
      opened_count INTEGER DEFAULT 0,
      subject TEXT,
      html_body TEXT,
      text_body TEXT,
      attachment_names TEXT,
      smtp_from TEXT,
      status TEXT DEFAULT 'pending'
    )
  `);

  await runQuery(`
    CREATE TABLE IF NOT EXISTS campaign_recipients (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      campaign_id TEXT NOT NULL,
      recipient_index INTEGER,
      email TEXT NOT NULL,
      name TEXT,
      row_data TEXT,
      send_status TEXT DEFAULT 'pending',
      sent_at DATETIME,
      error_message TEXT,
      is_opened INTEGER DEFAULT 0,
      opened_at DATETIME,
      open_count INTEGER DEFAULT 0,
      FOREIGN KEY (campaign_id) REFERENCES campaigns(id) ON DELETE CASCADE
    )
  `);

  await runQuery(`CREATE INDEX IF NOT EXISTS idx_recipients_campaign ON campaign_recipients(campaign_id)`);
  await runQuery(`CREATE INDEX IF NOT EXISTS idx_recipients_email ON campaign_recipients(email)`);
  console.log('✅ SQLite Campaign & Tracking Database initialized at:', DB_PATH);
}

// ----------------------------------------------------
// Campaign CRUD & Stats
// ----------------------------------------------------
async function createCampaign(campaign) {
  const sql = `
    INSERT INTO campaigns (
      id, name, created_at, original_filename, excel_file_path,
      total_recipients, sent_count, failed_count, opened_count,
      subject, html_body, text_body, attachment_names, smtp_from, status
    ) VALUES (?, ?, CURRENT_TIMESTAMP, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `;
  await runQuery(sql, [
    campaign.id,
    campaign.name || `Campaign - ${new Date().toLocaleDateString()}`,
    campaign.originalFilename || '',
    campaign.excelFilePath || '',
    campaign.totalRecipients || 0,
    0,
    0,
    0,
    campaign.subject || '',
    campaign.htmlBody || '',
    campaign.textBody || '',
    JSON.stringify(campaign.attachmentNames || []),
    campaign.smtpFrom || '',
    campaign.status || 'running'
  ]);
  return campaign.id;
}

async function addCampaignRecipients(campaignId, recipients) {
  return new Promise((resolve, reject) => {
    db.serialize(() => {
      db.run('BEGIN TRANSACTION');
      const stmt = db.prepare(`
        INSERT INTO campaign_recipients (
          campaign_id, recipient_index, email, name, row_data, send_status
        ) VALUES (?, ?, ?, ?, ?, 'pending')
      `);

      for (let i = 0; i < recipients.length; i++) {
        const r = recipients[i];
        const email = String(r.email || '').trim();
        const name = r.name || (r.data && (r.data.Name || r.data.name || r.data.FirstName || r.data.client)) || '';
        const rowData = JSON.stringify(r.data || r);
        stmt.run(campaignId, i + 1, email, name, rowData);
      }

      stmt.finalize((err) => {
        if (err) {
          db.run('ROLLBACK');
          return reject(err);
        }
        db.run('COMMIT', (commitErr) => {
          if (commitErr) reject(commitErr);
          else resolve(true);
        });
      });
    });
  });
}

async function updateRecipientSendStatus(campaignId, recipientIndex, status, errorMessage = null) {
  const sentAt = status === 'sent' ? new Date().toISOString() : null;
  await runQuery(
    `UPDATE campaign_recipients 
     SET send_status = ?, sent_at = COALESCE(?, sent_at), error_message = ? 
     WHERE campaign_id = ? AND recipient_index = ?`,
    [status, sentAt, errorMessage, campaignId, recipientIndex]
  );

  // Recalculate campaign sent/failed totals
  const stats = await getQuery(
    `SELECT 
       SUM(CASE WHEN send_status = 'sent' THEN 1 ELSE 0 END) as sent,
       SUM(CASE WHEN send_status = 'failed' THEN 1 ELSE 0 END) as failed
     FROM campaign_recipients WHERE campaign_id = ?`,
    [campaignId]
  );

  if (stats) {
    await runQuery(
      `UPDATE campaigns SET sent_count = ?, failed_count = ? WHERE id = ?`,
      [stats.sent || 0, stats.failed || 0, campaignId]
    );
  }
}

async function updateCampaignStatus(campaignId, status) {
  await runQuery(`UPDATE campaigns SET status = ? WHERE id = ?`, [status, campaignId]);
}

// ----------------------------------------------------
// Seen / Open Tracking
// ----------------------------------------------------
async function recordEmailOpen(campaignId, recipientIdentifier) {
  const index = parseInt(recipientIdentifier, 10);
  const now = new Date().toISOString();

  let targetRecipient;
  if (!isNaN(index)) {
    targetRecipient = await getQuery(
      `SELECT * FROM campaign_recipients WHERE campaign_id = ? AND (recipient_index = ? OR id = ?)`,
      [campaignId, index, index]
    );
  } else {
    targetRecipient = await getQuery(
      `SELECT * FROM campaign_recipients WHERE campaign_id = ? AND email = ?`,
      [campaignId, String(recipientIdentifier).trim()]
    );
  }

  if (!targetRecipient) {
    return { recorded: false, error: 'Recipient not found' };
  }

  const isFirstOpen = targetRecipient.is_opened === 0;

  await runQuery(
    `UPDATE campaign_recipients 
     SET is_opened = 1, 
         opened_at = COALESCE(opened_at, ?), 
         open_count = open_count + 1 
     WHERE id = ?`,
    [now, targetRecipient.id]
  );

  if (isFirstOpen) {
    await runQuery(
      `UPDATE campaigns 
       SET opened_count = (SELECT COUNT(*) FROM campaign_recipients WHERE campaign_id = ? AND is_opened = 1) 
       WHERE id = ?`,
      [campaignId, campaignId]
    );
  }

  return { recorded: true, firstOpen: isFirstOpen, recipient: targetRecipient };
}

// ----------------------------------------------------
// Queries for Dashboard & History
// ----------------------------------------------------
async function getAllCampaigns() {
  const campaigns = await allQuery(`
    SELECT 
      c.*,
      (c.total_recipients - c.opened_count) as unseen_count,
      CASE 
        WHEN c.sent_count > 0 THEN ROUND((CAST(c.opened_count AS FLOAT) / c.sent_count) * 100, 1)
        ELSE 0 
      END as open_rate_pct
    FROM campaigns c
    ORDER BY c.created_at DESC
  `);
  return campaigns;
}

async function getCampaignById(campaignId) {
  const campaign = await getQuery(
    `SELECT 
      c.*,
      (c.total_recipients - c.opened_count) as unseen_count,
      CASE 
        WHEN c.sent_count > 0 THEN ROUND((CAST(c.opened_count AS FLOAT) / c.sent_count) * 100, 1)
        ELSE 0 
      END as open_rate_pct
    FROM campaigns c 
    WHERE c.id = ?`,
    [campaignId]
  );

  if (!campaign) return null;

  const recipients = await allQuery(
    `SELECT 
      id, recipient_index, email, name, row_data, send_status, sent_at, error_message, is_opened, opened_at, open_count
     FROM campaign_recipients 
     WHERE campaign_id = ?
     ORDER BY recipient_index ASC`,
    [campaignId]
  );

  const parsedRecipients = recipients.map((r) => {
    let rowDataObj = {};
    try {
      rowDataObj = r.row_data ? JSON.parse(r.row_data) : {};
    } catch (e) {
      rowDataObj = {};
    }
    return {
      ...r,
      rowData: rowDataObj
    };
  });

  let attachments = [];
  try {
    attachments = campaign.attachment_names ? JSON.parse(campaign.attachment_names) : [];
  } catch (e) {
    attachments = [];
  }

  return {
    ...campaign,
    attachmentNames: attachments,
    recipients: parsedRecipients
  };
}

async function deleteCampaign(campaignId) {
  const campaign = await getQuery(`SELECT * FROM campaigns WHERE id = ?`, [campaignId]);
  if (campaign && campaign.excel_file_path && fs.existsSync(campaign.excel_file_path)) {
    try {
      fs.unlinkSync(campaign.excel_file_path);
    } catch (err) {
      console.warn('Could not delete excel file:', err.message);
    }
  }

  await runQuery(`DELETE FROM campaign_recipients WHERE campaign_id = ?`, [campaignId]);
  await runQuery(`DELETE FROM campaigns WHERE id = ?`, [campaignId]);
  return true;
}

async function getOverallStats() {
  const stats = await getQuery(`
    SELECT 
      COUNT(*) as total_campaigns,
      COALESCE(SUM(total_recipients), 0) as total_recipients,
      COALESCE(SUM(sent_count), 0) as total_sent,
      COALESCE(SUM(failed_count), 0) as total_failed,
      COALESCE(SUM(opened_count), 0) as total_opened,
      COALESCE(SUM(total_recipients - opened_count), 0) as total_unseen,
      CASE 
        WHEN SUM(sent_count) > 0 THEN ROUND((CAST(SUM(opened_count) AS FLOAT) / SUM(sent_count)) * 100, 1)
        ELSE 0 
      END as overall_open_rate
    FROM campaigns
  `);
  return stats;
}

async function getAggregateRecipients(type = 'all') {
  let whereClause = '';
  if (type === 'sent' || type === 'dispatched') {
    whereClause = "WHERE r.send_status = 'sent'";
  } else if (type === 'seen' || type === 'opened') {
    whereClause = "WHERE r.is_opened = 1";
  } else if (type === 'unseen' || type === 'unopened') {
    whereClause = "WHERE r.send_status = 'sent' AND r.is_opened = 0";
  } else if (type === 'failed') {
    whereClause = "WHERE r.send_status = 'failed'";
  }

  const rows = await allQuery(`
    SELECT 
      r.id,
      r.recipient_index,
      r.email,
      r.name,
      r.send_status,
      r.sent_at,
      r.is_opened,
      r.opened_at,
      r.open_count,
      r.error_message,
      r.row_data,
      c.id as campaign_id,
      c.name as campaign_name,
      c.subject as campaign_subject,
      c.created_at as campaign_created_at,
      c.original_filename
    FROM campaign_recipients r
    JOIN campaigns c ON r.campaign_id = c.id
    ${whereClause}
    ORDER BY COALESCE(r.opened_at, r.sent_at, c.created_at) DESC
    LIMIT 2000
  `);

  return rows.map((r) => {
    let rowDataObj = {};
    try {
      rowDataObj = r.row_data ? JSON.parse(r.row_data) : {};
    } catch (e) {}
    return { ...r, rowData: rowDataObj };
  });
}

module.exports = {
  db,
  initDb,
  createCampaign,
  addCampaignRecipients,
  updateRecipientSendStatus,
  updateCampaignStatus,
  recordEmailOpen,
  getAllCampaigns,
  getCampaignById,
  deleteCampaign,
  getOverallStats,
  getAggregateRecipients
};

