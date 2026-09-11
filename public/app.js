// =========================================================
// AutoMailer PRO - Frontend Application Logic
// =========================================================

document.addEventListener('DOMContentLoaded', () => {
  // Application State
  const state = {
    currentStep: 1,
    smtp: {
      host: 'smtp.gmail.com',
      port: 587,
      secure: false,
      user: '',
      pass: '',
      fromName: '',
      fromEmail: '',
      verified: false
    },
    excel: {
      file: null,
      fileName: '',
      totalRows: 0,
      validEmails: 0,
      invalidEmails: 0,
      columns: [],
      emailColumn: '',
      nameColumn: '',
      fullRows: [],
      sampleRows: []
    },
    message: {
      subject: 'Important update for {{Name}} at {{Company}}',
      body: '<p>Dear <strong>{{Name}}</strong>,</p>\n\n<p>I hope this email finds you well at <strong>{{Company}}</strong>.</p>\n\n<p>I am reaching out regarding our recent project updates and opportunities tailored for your team.</p>\n\n<p>Please let me know if you would like to connect for a quick 5-minute discussion this week.</p>\n\n<br>\n<p>Warm regards,<br><strong>Akash</strong><br><small style="color: #64748b;">If you no longer wish to receive updates, reply with "unsubscribe".</small></p>',
      isHtml: true,
      delaySeconds: 3.0,
      attachments: []
    },
    campaign: {
      id: null,
      status: 'idle',
      total: 0,
      sent: 0,
      failed: 0,
      pending: 0
    }
  };

  // ----------------------------------------------------
  // DOM Elements
  // ----------------------------------------------------
  const stepBtns = document.querySelectorAll('.step-btn');
  const stepPanes = document.querySelectorAll('.step-pane');
  const presetBtns = document.querySelectorAll('.btn-preset');
  const smtpStatusBadge = document.getElementById('smtpStatusBadge');

  // SMTP Inputs
  const smtpHostInput = document.getElementById('smtpHost');
  const smtpPortInput = document.getElementById('smtpPort');
  const smtpSecureSelect = document.getElementById('smtpSecure');
  const smtpUserInput = document.getElementById('smtpUser');
  const smtpPassInput = document.getElementById('smtpPass');
  const smtpFromNameInput = document.getElementById('smtpFromName');
  const smtpFromEmailInput = document.getElementById('smtpFromEmail');
  const btnTogglePass = document.getElementById('btnTogglePass');
  const btnTestConnection = document.getElementById('btnTestConnection');
  const btnSaveStep1 = document.getElementById('btnSaveStep1');

  // Excel / CSV Inputs
  const dropzone = document.getElementById('dropzone');
  const excelFileInput = document.getElementById('excelFileInput');
  const dataPreviewSection = document.getElementById('dataPreviewSection');
  const previewFileName = document.getElementById('previewFileName');
  const previewValidCount = document.getElementById('previewValidCount');
  const previewInvalidCount = document.getElementById('previewInvalidCount');
  const previewColCount = document.getElementById('previewColCount');
  const selectEmailColumn = document.getElementById('selectEmailColumn');
  const selectNameColumn = document.getElementById('selectNameColumn');
  const previewTableHead = document.getElementById('previewTableHead');
  const previewTableBody = document.getElementById('previewTableBody');
  const tableFilterInput = document.getElementById('tableFilterInput');
  const btnReupload = document.getElementById('btnReupload');
  const btnBackToStep1 = document.getElementById('btnBackToStep1');
  const btnGoToStep3 = document.getElementById('btnGoToStep3');

  // Composer Inputs
  const dynamicTagsList = document.getElementById('dynamicTagsList');
  const emailSubjectInput = document.getElementById('emailSubject');
  const emailBodyInput = document.getElementById('emailBody');
  const btnFmtHtml = document.getElementById('btnFmtHtml');
  const btnFmtText = document.getElementById('btnFmtText');
  const delaySlider = document.getElementById('delaySlider');
  const delayVal = document.getElementById('delayVal');
  const estimatedDuration = document.getElementById('estimatedDuration');
  const emailAttachmentsInput = document.getElementById('emailAttachments');
  const attachmentList = document.getElementById('attachmentList');
  const previewRecipientSelect = document.getElementById('previewRecipientSelect');
  const previewToEmail = document.getElementById('previewToEmail');
  const previewSubject = document.getElementById('previewSubject');
  const previewBodyRendered = document.getElementById('previewBodyRendered');
  const btnSendTestEmail = document.getElementById('btnSendTestEmail');
  const btnBackToStep2 = document.getElementById('btnBackToStep2');
  const btnGoToStep4 = document.getElementById('btnGoToStep4');

  // Campaign Dispatch & Live Monitor
  const btnStartCampaign = document.getElementById('btnStartCampaign');
  const btnPauseResumeCampaign = document.getElementById('btnPauseResumeCampaign');
  const btnStopCampaign = document.getElementById('btnStopCampaign');
  const btnExportCsv = document.getElementById('btnExportCsv');
  const progressCircle = document.getElementById('progressCircle');
  const progressPercent = document.getElementById('progressPercent');
  const campaignStatusBadge = document.getElementById('campaignStatusBadge');
  const counterTotal = document.getElementById('counterTotal');
  const counterSent = document.getElementById('counterSent');
  const counterFailed = document.getElementById('counterFailed');
  const counterPending = document.getElementById('counterPending');
  const summaryDelayText = document.getElementById('summaryDelayText');
  const terminalLogs = document.getElementById('terminalLogs');
  const btnClearLogs = document.getElementById('btnClearLogs');
  const recipientStatusList = document.getElementById('recipientStatusList');

  // Modals
  const testEmailModal = document.getElementById('testEmailModal');
  const btnCloseTestModal = document.getElementById('btnCloseTestModal');
  const btnCancelTestModal = document.getElementById('btnCancelTestModal');
  const testTargetEmail = document.getElementById('testTargetEmail');
  const btnConfirmSendTest = document.getElementById('btnConfirmSendTest');

  const helpModal = document.getElementById('helpModal');
  const btnHelpModal = document.getElementById('btnHelpModal');
  const btnCloseHelpModal = document.getElementById('btnCloseHelpModal');

  // Toast Container
  const toastContainer = document.getElementById('toastContainer');

  // ----------------------------------------------------
  // Progress Ring Constants & Initialization
  // ----------------------------------------------------
  const circleRadius = 66;
  const circumference = 2 * Math.PI * circleRadius;

  function setupProgressRing() {
    if (progressCircle) {
      progressCircle.style.strokeDasharray = `${circumference} ${circumference}`;
      progressCircle.style.strokeDashoffset = `${circumference}`;
    }
  }

  function setProgressRing(percent) {
    if (progressCircle) {
      const offset = circumference - (percent / 100) * circumference;
      progressCircle.style.strokeDashoffset = offset;
    }
  }

  // ----------------------------------------------------
  // Theme Selector & Light/Dark Mode Controller
  // ----------------------------------------------------
  const themeSelector = document.getElementById('themeSelector');
  const btnToggleMode = document.getElementById('btnToggleMode');

  function applyTheme(themeKey) {
    document.body.setAttribute('data-theme', themeKey);
    const isLight = themeKey.includes('light');
    document.body.classList.toggle('light-theme', isLight);
    document.body.classList.toggle('dark-theme', !isLight);
    if (themeSelector) themeSelector.value = themeKey;
    if (btnToggleMode) {
      btnToggleMode.innerHTML = isLight ? '<i class="fa-solid fa-moon"></i>' : '<i class="fa-solid fa-sun"></i>';
      btnToggleMode.title = isLight ? 'Switch to Dark Mode' : 'Switch to Light Mode';
    }
    localStorage.setItem('automailer_theme', themeKey);
  }

  if (themeSelector) {
    themeSelector.addEventListener('change', (e) => {
      applyTheme(e.target.value);
    });
  }

  if (btnToggleMode) {
    btnToggleMode.addEventListener('click', () => {
      const currentTheme = document.body.getAttribute('data-theme') || 'cosmic-aurora';
      const isCurrentlyLight = currentTheme.includes('light') || document.body.classList.contains('light-theme');
      const newTheme = isCurrentlyLight ? 'cosmic-aurora' : 'clean-light';
      applyTheme(newTheme);
      showToast(`Switched to ${isCurrentlyLight ? 'Dark' : 'Light'} background!`, 'info');
    });
  }

  // Load saved theme (default to clean-light if requested or cosmic-aurora)
  const savedTheme = localStorage.getItem('automailer_theme') || 'clean-light';
  applyTheme(savedTheme);

  // ----------------------------------------------------
  // Initial Setup & Local Storage Load
  // ----------------------------------------------------
  loadSavedSettings();
  emailBodyInput.value = state.message.body;
  setupProgressRing();
  setupSSE();

  // ----------------------------------------------------
  // STEPPER NAVIGATION
  // ----------------------------------------------------
  function goToStep(stepNum) {
    if (stepNum < 1 || stepNum > 4) return;
    state.currentStep = stepNum;

    stepBtns.forEach((btn) => {
      const bStep = parseInt(btn.dataset.step, 10);
      btn.classList.toggle('active', bStep === stepNum);
      btn.classList.toggle('completed', bStep < stepNum);
    });

    stepPanes.forEach((pane) => {
      pane.classList.remove('active');
    });

    const targetPane = document.getElementById(`step${stepNum}`);
    if (targetPane) targetPane.classList.add('active');

    if (stepNum === 3) {
      updateDynamicTagsUI();
      updateLivePreview();
      updateDurationEstimate();
    } else if (stepNum === 4) {
      updateCampaignSummaryView();
    }
  }

  stepBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      const step = parseInt(btn.dataset.step, 10);
      goToStep(step);
    });
  });

  btnSaveStep1.addEventListener('click', () => {
    saveSmtpFromInputs();
    goToStep(2);
  });
  btnBackToStep1.addEventListener('click', () => goToStep(1));
  btnGoToStep3.addEventListener('click', () => goToStep(3));
  btnBackToStep2.addEventListener('click', () => goToStep(2));
  btnGoToStep4.addEventListener('click', () => goToStep(4));

  // ----------------------------------------------------
  // STEP 1: SMTP PRESETS & LOGIC
  // ----------------------------------------------------
  const smtpPresets = {
    gmail: { host: 'smtp.gmail.com', port: 587, secure: false },
    outlook: { host: 'smtp.office365.com', port: 587, secure: false },
    brevo: { host: 'smtp-relay.brevo.com', port: 587, secure: false },
    custom: { host: '', port: 587, secure: false }
  };

  presetBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      presetBtns.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      const presetKey = btn.dataset.preset;
      const config = smtpPresets[presetKey];
      if (config) {
        if (config.host) smtpHostInput.value = config.host;
        smtpPortInput.value = config.port;
        smtpSecureSelect.value = String(config.secure);
      }
    });
  });

  btnTogglePass.addEventListener('click', () => {
    const isPass = smtpPassInput.type === 'password';
    smtpPassInput.type = isPass ? 'text' : 'password';
    btnTogglePass.innerHTML = isPass ? '<i class="fa-regular fa-eye-slash"></i>' : '<i class="fa-regular fa-eye"></i>';
  });

  function saveSmtpFromInputs() {
    state.smtp.host = smtpHostInput.value.trim();
    state.smtp.port = parseInt(smtpPortInput.value, 10) || 587;
    state.smtp.secure = smtpSecureSelect.value === 'true';
    state.smtp.user = smtpUserInput.value.trim();
    state.smtp.pass = smtpPassInput.value.trim();
    state.smtp.fromName = smtpFromNameInput.value.trim();
    state.smtp.fromEmail = smtpFromEmailInput.value.trim() || state.smtp.user;

    localStorage.setItem('automailer_smtp', JSON.stringify({
      host: state.smtp.host,
      port: state.smtp.port,
      secure: state.smtp.secure,
      user: state.smtp.user,
      fromName: state.smtp.fromName,
      fromEmail: state.smtp.fromEmail
    }));
  }

  function loadSavedSettings() {
    try {
      const saved = localStorage.getItem('automailer_smtp');
      if (saved) {
        const obj = JSON.parse(saved);
        if (obj.host) smtpHostInput.value = obj.host;
        if (obj.port) smtpPortInput.value = obj.port;
        if (obj.secure !== undefined) smtpSecureSelect.value = String(obj.secure);
        if (obj.user) smtpUserInput.value = obj.user;
        if (obj.fromName) smtpFromNameInput.value = obj.fromName;
        if (obj.fromEmail) smtpFromEmailInput.value = obj.fromEmail;
      }
    } catch (e) {
      console.warn('Could not load saved settings', e);
    }
  }

  btnTestConnection.addEventListener('click', async () => {
    saveSmtpFromInputs();
    if (!state.smtp.host || !state.smtp.user || !state.smtp.pass) {
      showToast('Please enter SMTP Host, Username and Password.', 'error');
      return;
    }

    btnTestConnection.disabled = true;
    btnTestConnection.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Verifying...';

    try {
      const res = await fetch('/api/test-smtp', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ smtp: state.smtp })
      });
      const data = await res.json();
      if (res.ok && data.success) {
        state.smtp.verified = true;
        updateSmtpBadge(true, 'SMTP Connected');
        showToast(data.message, 'success');
      } else {
        state.smtp.verified = false;
        updateSmtpBadge(false, 'Connection Failed');
        showToast(data.error || 'Connection failed', 'error');
      }
    } catch (err) {
      state.smtp.verified = false;
      updateSmtpBadge(false, 'Network Error');
      showToast('Failed to reach server.', 'error');
    } finally {
      btnTestConnection.disabled = false;
      btnTestConnection.innerHTML = '<i class="fa-solid fa-plug-circle-check"></i> Test Connection';
    }
  });

  function updateSmtpBadge(connected, text) {
    const dot = smtpStatusBadge.querySelector('.status-dot');
    const label = smtpStatusBadge.querySelector('.status-text');
    dot.className = `status-dot ${connected ? 'connected' : 'disconnected'}`;
    label.textContent = text;
  }

  // ----------------------------------------------------
  // STEP 2: EXCEL / CSV UPLOAD & TABLE PREVIEW
  // ----------------------------------------------------
  dropzone.addEventListener('click', () => excelFileInput.click());
  btnReupload.addEventListener('click', () => excelFileInput.click());

  ['dragenter', 'dragover'].forEach((eventName) => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach((eventName) => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
    });
  });

  dropzone.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) handleFileUpload(files[0]);
  });

  excelFileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleFileUpload(e.target.files[0]);
  });

  async function handleFileUpload(file) {
    if (!file) return;
    const formData = new FormData();
    formData.append('file', file);

    dropzone.innerHTML = '<div class="dropzone-content"><i class="fa-solid fa-spinner fa-spin fa-2x" style="color:var(--primary);margin-bottom:1rem;"></i><h3>Processing spreadsheet...</h3><p>Extracting recipient records and column variables</p></div>';

    try {
      const res = await fetch('/api/parse-excel', {
        method: 'POST',
        body: formData
      });
      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Failed to parse file');
      }

      state.excel.file = file;
      state.excel.fileName = data.fileName;
      state.excel.totalRows = data.totalRows;
      state.excel.validEmails = data.validEmails;
      state.excel.invalidEmails = data.invalidEmails;
      state.excel.columns = data.columns;
      state.excel.emailColumn = data.detectedEmailColumn;
      state.excel.nameColumn = data.detectedNameColumn;
      state.excel.fullRows = data.fullRows;
      state.excel.sampleRows = data.sampleRows;

      // Update UI
      dropzone.style.display = 'none';
      btnReupload.style.display = 'inline-flex';
      dataPreviewSection.style.display = 'block';

      previewFileName.textContent = data.fileName;
      previewValidCount.textContent = data.validEmails;
      previewInvalidCount.textContent = data.invalidEmails;
      previewColCount.textContent = data.columns.length;

      // Populate Column Selectors
      populateColumnSelectors(data.columns, data.detectedEmailColumn, data.detectedNameColumn);

      // Render Preview Table
      renderTable(data.columns, data.sampleRows, data.detectedEmailColumn);

      showToast(`Successfully loaded ${data.totalRows} contacts!`, 'success');
    } catch (err) {
      showToast(err.message, 'error');
      // Reset dropzone
      dropzone.innerHTML = `
        <input type="file" id="excelFileInput" accept=".xlsx, .xls, .csv" hidden>
        <div class="dropzone-content">
          <div class="dropzone-icon"><i class="fa-solid fa-cloud-arrow-up"></i></div>
          <h3>Drag & Drop your Excel (.xlsx, .xls) or .CSV file here</h3>
          <p>or click to browse files from your computer</p>
        </div>
      `;
    }
  }

  function populateColumnSelectors(columns, emailCol, nameCol) {
    selectEmailColumn.innerHTML = '';
    selectNameColumn.innerHTML = '<option value="">-- None --</option>';

    columns.forEach((col) => {
      const opt1 = document.createElement('option');
      opt1.value = col;
      opt1.textContent = col;
      if (col === emailCol) opt1.selected = true;
      selectEmailColumn.appendChild(opt1);

      const opt2 = document.createElement('option');
      opt2.value = col;
      opt2.textContent = col;
      if (col === nameCol) opt2.selected = true;
      selectNameColumn.appendChild(opt2);
    });
  }

  selectEmailColumn.addEventListener('change', (e) => {
    state.excel.emailColumn = e.target.value;
    renderTable(state.excel.columns, state.excel.sampleRows, e.target.value);
  });

  selectNameColumn.addEventListener('change', (e) => {
    state.excel.nameColumn = e.target.value;
  });

  function renderTable(columns, rows, emailCol, filterText = '') {
    previewTableHead.innerHTML = `
      <tr>
        <th style="width: 50px;">#</th>
        <th>Status</th>
        ${columns.map((c) => `<th class="${c === emailCol ? 'text-primary' : ''}">${c}</th>`).join('')}
      </tr>
    `;

    const filtered = filterText
      ? rows.filter((r) => Object.values(r).some((v) => String(v).toLowerCase().includes(filterText.toLowerCase())))
      : rows;

    previewTableBody.innerHTML = filtered.map((row, idx) => {
      const isValid = row.__isValidEmail;
      return `
        <tr>
          <td>${row.__rowIndex || idx + 1}</td>
          <td>
            <span class="badge-cell ${isValid ? 'valid' : 'invalid'}">
              <i class="fa-solid fa-${isValid ? 'check' : 'xmark'}"></i> ${isValid ? 'Valid' : 'Invalid'}
            </span>
          </td>
          ${columns.map((c) => `<td>${escapeHtml(String(row[c] || ''))}</td>`).join('')}
        </tr>
      `;
    }).join('');
  }

  tableFilterInput.addEventListener('input', (e) => {
    renderTable(state.excel.columns, state.excel.sampleRows, state.excel.emailColumn, e.target.value);
  });

  // ----------------------------------------------------
  // STEP 3: MESSAGE COMPOSER & LIVE PREVIEW
  // ----------------------------------------------------
  function updateDynamicTagsUI() {
    const cols = state.excel.columns.length > 0 ? state.excel.columns : ['Name', 'Email', 'Company'];
    dynamicTagsList.innerHTML = cols.map((col) => {
      return `<span class="tag-chip" data-tag="{{${col}}}">&#123;&#123;${col}&#125;&#125;</span>`;
    }).join('');

    // Attach click listeners to tags
    dynamicTagsList.querySelectorAll('.tag-chip').forEach((chip) => {
      chip.addEventListener('click', () => {
        insertAtCursor(emailBodyInput, chip.dataset.tag);
        updateLivePreview();
      });
    });

    // Populate recipient preview switcher
    previewRecipientSelect.innerHTML = '';
    const sampleList = state.excel.fullRows.slice(0, 20);
    if (sampleList.length === 0) {
      previewRecipientSelect.innerHTML = '<option value="0">Demo Client</option>';
    } else {
      sampleList.forEach((r, idx) => {
        const nameVal = r[state.excel.nameColumn] || r[state.excel.columns[0]] || `Row #${idx + 1}`;
        const opt = document.createElement('option');
        opt.value = idx;
        opt.textContent = `#${idx + 1} - ${nameVal}`;
        previewRecipientSelect.appendChild(opt);
      });
    }
  }

  function insertAtCursor(textarea, textToInsert) {
    const startPos = textarea.selectionStart;
    const endPos = textarea.selectionEnd;
    textarea.value = textarea.value.substring(0, startPos) + textToInsert + textarea.value.substring(endPos);
    textarea.focus();
    textarea.selectionStart = startPos + textToInsert.length;
    textarea.selectionEnd = startPos + textToInsert.length;
  }

  btnFmtHtml.addEventListener('click', () => {
    btnFmtHtml.classList.add('active');
    btnFmtText.classList.remove('active');
    state.message.isHtml = true;
    updateLivePreview();
  });

  btnFmtText.addEventListener('click', () => {
    btnFmtText.classList.add('active');
    btnFmtHtml.classList.remove('active');
    state.message.isHtml = false;
    updateLivePreview();
  });

  emailSubjectInput.addEventListener('input', () => updateLivePreview());
  emailBodyInput.addEventListener('input', () => updateLivePreview());
  previewRecipientSelect.addEventListener('change', () => updateLivePreview());

  function updateLivePreview() {
    const selectedIdx = parseInt(previewRecipientSelect.value, 10) || 0;
    const clientData = state.excel.fullRows[selectedIdx] || {
      Name: 'John Smith',
      Email: 'john.smith@example.com',
      Company: 'Acme Corp'
    };

    const emailCol = state.excel.emailColumn || 'Email';
    previewToEmail.textContent = clientData[emailCol] || 'client@example.com';

    const rawSubj = emailSubjectInput.value || 'No Subject';
    previewSubject.textContent = replaceVariables(rawSubj, clientData);

    const rawBody = emailBodyInput.value || '<p>No content</p>';
    const rendered = replaceVariables(rawBody, clientData);

    if (state.message.isHtml) {
      previewBodyRendered.innerHTML = rendered;
    } else {
      previewBodyRendered.textContent = rendered;
    }
  }

  // Delay & Rate slider
  delaySlider.addEventListener('input', (e) => {
    const val = parseFloat(e.target.value);
    state.message.delaySeconds = val;
    delayVal.textContent = `${val.toFixed(1)}s`;
    summaryDelayText.textContent = `${val.toFixed(1)}s per email`;
    updateDurationEstimate();
  });

  function updateDurationEstimate() {
    const totalCount = state.excel.validEmails || 1000;
    const totalSec = totalCount * state.message.delaySeconds;
    const mins = Math.ceil(totalSec / 60);
    estimatedDuration.textContent = `Estimated time for ${totalCount} emails: ~${mins} min${mins > 1 ? 's' : ''}`;
  }

  // Attachment handling
  emailAttachmentsInput.addEventListener('change', (e) => {
    state.message.attachments = Array.from(e.target.files);
    renderAttachmentChips();
  });

  function renderAttachmentChips() {
    attachmentList.innerHTML = state.message.attachments.map((f, i) => `
      <span class="tag-chip" style="background:#1e293b; color:#38bdf8;">
        <i class="fa-solid fa-paperclip"></i> ${escapeHtml(f.name)} (${(f.size / 1024).toFixed(0)} KB)
      </span>
    `).join('');
  }

  // ----------------------------------------------------
  // TEST EMAIL MODAL
  // ----------------------------------------------------
  btnSendTestEmail.addEventListener('click', () => {
    saveSmtpFromInputs();
    if (!state.smtp.user) {
      showToast('Please configure your sender email in Step 1 first.', 'error');
      goToStep(1);
      return;
    }
    testTargetEmail.value = state.smtp.user;
    testEmailModal.classList.add('active');
  });

  btnCloseTestModal.addEventListener('click', () => testEmailModal.classList.remove('active'));
  btnCancelTestModal.addEventListener('click', () => testEmailModal.classList.remove('active'));

  btnConfirmSendTest.addEventListener('click', async () => {
    const target = testTargetEmail.value.trim();
    if (!target) {
      showToast('Please enter a destination email address.', 'error');
      return;
    }

    btnConfirmSendTest.disabled = true;
    btnConfirmSendTest.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Sending...';

    const selectedIdx = parseInt(previewRecipientSelect.value, 10) || 0;
    const sampleData = state.excel.fullRows[selectedIdx] || { Name: 'Demo Client', Company: 'Acme Corp' };

    try {
      const res = await fetch('/api/test-smtp', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          smtp: state.smtp,
          testEmail: target,
          subject: emailSubjectInput.value,
          htmlBody: state.message.isHtml ? emailBodyInput.value : undefined,
          textBody: !state.message.isHtml ? emailBodyInput.value : undefined,
          sampleData
        })
      });
      const data = await res.json();
      if (res.ok && data.success) {
        showToast(`Test email successfully sent to ${target}! Check your inbox.`, 'success');
        testEmailModal.classList.remove('active');
      } else {
        showToast(data.error || 'Failed to send test email.', 'error');
      }
    } catch (err) {
      showToast('Error communicating with server.', 'error');
    } finally {
      btnConfirmSendTest.disabled = false;
      btnConfirmSendTest.innerHTML = '<i class="fa-solid fa-paper-plane"></i> Send Test Now';
    }
  });

  // Help Modal
  btnHelpModal.addEventListener('click', () => helpModal.classList.add('active'));
  btnCloseHelpModal.addEventListener('click', () => helpModal.classList.remove('active'));

  // ----------------------------------------------------
  // STEP 4: CAMPAIGN DISPATCH & LIVE REAL-TIME MONITOR
  // ----------------------------------------------------
  function updateCampaignSummaryView() {
    const total = state.excel.validEmails || 0;
    counterTotal.textContent = total;
    counterSent.textContent = '0';
    counterFailed.textContent = '0';
    counterPending.textContent = total;

    // Render initial recipient list
    if (state.excel.fullRows.length > 0) {
      const emailCol = state.excel.emailColumn;
      recipientStatusList.innerHTML = state.excel.fullRows.map((r, i) => `
        <div class="recipient-item-row" id="recRow_${i}">
          <div class="recipient-item-email">#${i + 1}. ${escapeHtml(String(r[emailCol] || ''))}</div>
          <div class="recipient-item-status pending">Pending</div>
        </div>
      `).join('');
    }
  }

  btnStartCampaign.addEventListener('click', async () => {
    saveSmtpFromInputs();
    if (!state.smtp.host || !state.smtp.user || !state.smtp.pass) {
      showToast('Please complete SMTP configuration in Step 1.', 'error');
      goToStep(1);
      return;
    }
    if (!state.excel.fullRows || state.excel.fullRows.length === 0) {
      showToast('Please upload an Excel contact sheet in Step 2.', 'error');
      goToStep(2);
      return;
    }

    const formData = new FormData();
    formData.append('smtp', JSON.stringify(state.smtp));
    formData.append('emailColumn', state.excel.emailColumn);
    formData.append('subject', emailSubjectInput.value);
    if (state.message.isHtml) {
      formData.append('htmlBody', emailBodyInput.value);
    } else {
      formData.append('textBody', emailBodyInput.value);
    }
    formData.append('recipients', JSON.stringify(state.excel.fullRows));
    formData.append('delaySeconds', state.message.delaySeconds);

    // Attachments
    state.message.attachments.forEach((file) => {
      formData.append('attachments', file);
    });

    btnStartCampaign.disabled = true;
    btnStartCampaign.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Initializing...';

    try {
      const res = await fetch('/api/campaign/start', {
        method: 'POST',
        body: formData
      });
      const data = await res.json();
      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Failed to start campaign');
      }

      showToast('Campaign successfully initiated!', 'success');
      btnStartCampaign.style.display = 'none';
      btnPauseResumeCampaign.style.display = 'inline-flex';
      btnStopCampaign.style.display = 'inline-flex';
      btnExportCsv.disabled = false;
    } catch (err) {
      showToast(err.message, 'error');
      btnStartCampaign.disabled = false;
      btnStartCampaign.innerHTML = '<i class="fa-solid fa-play"></i> START BULK SENDING';
    }
  });

  btnPauseResumeCampaign.addEventListener('click', async () => {
    if (state.campaign.status === 'running') {
      await fetch('/api/campaign/pause', { method: 'POST' });
    } else if (state.campaign.status === 'paused') {
      await fetch('/api/campaign/resume', { method: 'POST' });
    }
  });

  btnStopCampaign.addEventListener('click', async () => {
    if (confirm('Are you sure you want to cancel the bulk campaign?')) {
      await fetch('/api/campaign/stop', { method: 'POST' });
    }
  });

  btnExportCsv.addEventListener('click', () => {
    window.location.href = '/api/campaign/export-csv';
  });

  btnClearLogs.addEventListener('click', () => {
    terminalLogs.innerHTML = '<div class="log-line info">[System] Logs cleared.</div>';
  });

  // ----------------------------------------------------
  // SSE REAL-TIME CONNECTION
  // ----------------------------------------------------
  function setupSSE() {
    const eventSource = new EventSource('/api/campaign/events');

    eventSource.onmessage = (event) => {
      try {
        const payload = JSON.parse(event.data);
        handleSSEUpdate(payload);
      } catch (err) {
        console.error('Error handling SSE message:', err);
      }
    };

    eventSource.onerror = () => {
      console.warn('SSE connection lost. Reconnecting...');
    };
  }

  function handleSSEUpdate(payload) {
    if (payload.campaign) {
      state.campaign = { ...state.campaign, ...payload.campaign };
      updateCampaignUI(payload.campaign);
    }

    if (payload.type === 'log' && payload.log) {
      appendTerminalLog(payload.log);
    }

    if (payload.type === 'progress' && payload.currentRecipient) {
      updateRecipientRow(payload.currentRecipient);
    }

    if (payload.type === 'completed') {
      btnPauseResumeCampaign.style.display = 'none';
      btnStopCampaign.style.display = 'none';
      btnStartCampaign.style.display = 'inline-flex';
      btnStartCampaign.disabled = false;
      btnStartCampaign.innerHTML = '<i class="fa-solid fa-rotate-right"></i> Start New Campaign';
      showToast('🎉 Campaign finished sending!', 'success');
    }
  }

  function updateCampaignUI(camp) {
    counterTotal.textContent = camp.total || 0;
    counterSent.textContent = camp.sent || 0;
    counterFailed.textContent = camp.failed || 0;
    counterPending.textContent = camp.pending || 0;

    const total = camp.total || 1;
    const completed = (camp.sent || 0) + (camp.failed || 0);
    const pct = Math.min(100, Math.round((completed / total) * 100));

    progressPercent.textContent = `${pct}%`;
    setProgressRing(pct);

    campaignStatusBadge.textContent = camp.status.toUpperCase();
    campaignStatusBadge.className = `progress-status-badge ${camp.status}`;

    if (camp.status === 'running') {
      btnPauseResumeCampaign.innerHTML = '<i class="fa-solid fa-pause"></i> Pause';
      btnPauseResumeCampaign.classList.remove('btn-hero-start');
      btnPauseResumeCampaign.classList.add('btn-hero-pause');
    } else if (camp.status === 'paused') {
      btnPauseResumeCampaign.innerHTML = '<i class="fa-solid fa-play"></i> Resume';
      btnPauseResumeCampaign.classList.remove('btn-hero-pause');
      btnPauseResumeCampaign.classList.add('btn-hero-start');
    }
  }

  function updateRecipientRow(recipient) {
    const rowIdx = recipient.index - 1;
    const rowEl = document.getElementById(`recRow_${rowIdx}`);
    if (rowEl) {
      const badge = rowEl.querySelector('.recipient-item-status');
      if (badge) {
        badge.className = `recipient-item-status ${recipient.status}`;
        badge.textContent = recipient.status;
      }
      rowEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
  }

  function appendTerminalLog(log) {
    const line = document.createElement('div');
    line.className = `log-line ${log.type}`;
    line.textContent = `[${log.timestamp}] ${log.message}`;
    terminalLogs.prepend(line);
  }

  function setProgressRing(percent) {
    const circle = document.getElementById('progressCircle');
    if (!circle) return;
    const radius = circle.r.baseVal.value || 66;
    const circumference = 2 * Math.PI * radius;
    circle.style.strokeDasharray = `${circumference} ${circumference}`;
    const offset = circumference - (percent / 100) * circumference;
    circle.style.strokeDashoffset = offset;
  }

  // ----------------------------------------------------
  // UTILITY HELPERS
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

  function escapeHtml(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    const icon = type === 'success' ? 'circle-check' : type === 'error' ? 'circle-exclamation' : 'circle-info';
    toast.innerHTML = `<i class="fa-solid fa-${icon}"></i> <span>${escapeHtml(message)}</span>`;
    toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateX(100%)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 4500);
  }



  // Initial Progress Ring Setup
  setProgressRing(0);
});
