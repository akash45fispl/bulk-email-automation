const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const BASE_URL = 'http://localhost:3000';

async function runTests() {
  console.log('====================================================');
  console.log('       AUTOMAILER PRO - COMPREHENSIVE TEST SUITE    ');
  console.log('====================================================\n');

  let passed = 0;
  let failed = 0;

  // Test 1: Health Check
  try {
    console.log('[TEST 1] Testing /health endpoint...');
    const res = await fetch(`${BASE_URL}/health`);
    const data = await res.json();
    if (res.ok && data.status === 'healthy') {
      console.log('✅ PASS: Server is healthy.');
      passed++;
    } else {
      throw new Error(`Invalid health response: ${JSON.stringify(data)}`);
    }
  } catch (err) {
    console.error('❌ FAIL: /health -', err.message);
    failed++;
  }

  // Test 2: Sample Excel Download
  try {
    console.log('\n[TEST 2] Testing /api/download-sample-excel endpoint...');
    const res = await fetch(`${BASE_URL}/api/download-sample-excel`);
    const buffer = await res.arrayBuffer();
    const wb = xlsx.read(Buffer.from(buffer), { type: 'buffer' });
    const sheetName = wb.SheetNames[0];
    const rows = xlsx.utils.sheet_to_json(wb.Sheets[sheetName]);
    if (res.ok && rows.length > 0 && rows[0].Email) {
      console.log(`✅ PASS: Sample Excel generated with ${rows.length} contacts.`);
      passed++;
    } else {
      throw new Error('Failed to parse downloaded sample Excel');
    }
  } catch (err) {
    console.error('❌ FAIL: /api/download-sample-excel -', err.message);
    failed++;
  }

  // Test 3: Parse Excel Upload
  try {
    console.log('\n[TEST 3] Testing /api/parse-excel endpoint...');
    const sampleFilePath = path.join(__dirname, 'sample_clients.xlsx');
    const fileBuf = fs.readFileSync(sampleFilePath);

    const boundary = '----WebKitFormBoundary' + Math.random().toString(36).substring(2);
    let formHeader = `--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="sample_clients.xlsx"\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n`;
    const payload = Buffer.concat([
      Buffer.from(formHeader, 'utf8'),
      fileBuf,
      Buffer.from(`\r\n--${boundary}--\r\n`, 'utf8')
    ]);

    const res = await fetch(`${BASE_URL}/api/parse-excel`, {
      method: 'POST',
      headers: { 'Content-Type': `multipart/form-data; boundary=${boundary}` },
      body: payload
    });
    const data = await res.json();

    if (res.ok && data.success && data.columns.includes('Email') && data.validEmails > 0) {
      console.log(`✅ PASS: Excel parsed successfully (${data.totalRows} rows, detected column: "${data.detectedEmailColumn}").`);
      passed++;
    } else {
      throw new Error(`Parse failed: ${JSON.stringify(data)}`);
    }
  } catch (err) {
    console.error('❌ FAIL: /api/parse-excel -', err.message);
    failed++;
  }

  // Test 4: Database & Campaign Creation & Open Tracking Test
  let testCampaignId = null;
  try {
    console.log('\n[TEST 4] Testing SQLite Campaign Creation & Tracking Pixel...');
    const db = require('./database');
    testCampaignId = `camp_test_${Date.now()}`;
    await db.createCampaign({
      id: testCampaignId,
      name: 'Automated Test Dump 2026',
      originalFilename: 'sample_clients.xlsx',
      excelFilePath: '',
      totalRecipients: 2,
      subject: 'Special Offer for {{Company}}',
      htmlBody: '<h1>Hello {{Name}}</h1>',
      textBody: 'Hello {{Name}}',
      attachmentNames: ['brochure.pdf'],
      smtpFrom: 'AutoMailer <test@example.com>',
      status: 'completed'
    });

    await db.addCampaignRecipients(testCampaignId, [
      { email: 'alex@example.com', name: 'Alex', data: { Name: 'Alex', Company: 'Apex Corp' } },
      { email: 'sarah@example.com', name: 'Sarah', data: { Name: 'Sarah', Company: 'Cyberdyne' } }
    ]);

    await db.updateRecipientSendStatus(testCampaignId, 1, 'sent');
    await db.updateRecipientSendStatus(testCampaignId, 2, 'sent');

    // Trigger Open Tracking for Recipient 1
    const openRes = await fetch(`${BASE_URL}/api/track/open/${testCampaignId}/1`);
    const openBuffer = await openRes.arrayBuffer();

    // Check Campaign Details
    const detailRes = await fetch(`${BASE_URL}/api/campaigns/${testCampaignId}`);
    const detailData = await detailRes.json();

    if (
      openRes.ok &&
      openBuffer.byteLength > 0 &&
      detailData.success &&
      detailData.campaign.opened_count === 1 &&
      detailData.campaign.unseen_count === 1 &&
      detailData.campaign.recipients[0].is_opened === 1
    ) {
      console.log('✅ PASS: Campaign persisted, Open tracking pixel processed (1 Seen, 1 Unseen).');
      passed++;
    } else {
      throw new Error(`Campaign/Open verification failed: ${JSON.stringify(detailData)}`);
    }
  } catch (err) {
    console.error('❌ FAIL: Campaign Creation & Open Tracking -', err.message);
    failed++;
  }

  // Test 5: Campaign History List API
  try {
    console.log('\n[TEST 5] Testing /api/campaigns History & Analytics List...');
    const res = await fetch(`${BASE_URL}/api/campaigns`);
    const data = await res.json();
    if (res.ok && data.success && Array.isArray(data.campaigns) && data.stats) {
      console.log(`✅ PASS: Found ${data.campaigns.length} campaigns in history. Total Sent: ${data.stats.total_sent}, Total Seen: ${data.stats.total_opened}`);
      passed++;
    } else {
      throw new Error(`Failed to list campaigns: ${JSON.stringify(data)}`);
    }
  } catch (err) {
    console.error('❌ FAIL: /api/campaigns -', err.message);
    failed++;
  }

  // Test 6: Export Analytics Report as Excel
  try {
    console.log('\n[TEST 6] Testing /api/campaigns/:id/export-analytics...');
    const res = await fetch(`${BASE_URL}/api/campaigns/${testCampaignId}/export-analytics`);
    const buffer = await res.arrayBuffer();
    const wb = xlsx.read(Buffer.from(buffer), { type: 'buffer' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet);
    if (res.ok && rows.length === 2 && rows[0].Seen_Status.includes('SEEN')) {
      console.log(`✅ PASS: Analytics report exported with ${rows.length} rows including Seen/Unseen status.`);
      passed++;
    } else {
      throw new Error(`Analytics export verification failed`);
    }
  } catch (err) {
    console.error('❌ FAIL: /api/campaigns/:id/export-analytics -', err.message);
    failed++;
  }

  // Test 7: Clean up test campaign
  try {
    if (testCampaignId) {
      const delRes = await fetch(`${BASE_URL}/api/campaigns/${testCampaignId}`, { method: 'DELETE' });
      const delData = await delRes.json();
      if (delRes.ok && delData.success) {
        console.log('✅ PASS: Cleaned up test campaign record.');
        passed++;
      }
    }
  } catch (err) {
    console.error('❌ Clean up error:', err.message);
  }

  console.log('\n====================================================');
  console.log(`TEST RESULTS: ${passed} PASSED, ${failed} FAILED`);
  console.log('====================================================\n');
}

runTests().catch(console.error);
