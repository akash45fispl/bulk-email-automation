// =========================================================
// AutoMailer PRO - Frontend Application Logic
// Enhanced with SQLite Campaign Archiving & Open Tracking
// =========================================================

document.addEventListener('DOMContentLoaded', () => {
  // Application State
  const state = {
    currentView: 'wizardView', // 'wizardView' | 'analyticsView'
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
      name: '',
      status: 'idle',
      total: 0,
      sent: 0,
      failed: 0,
      pending: 0,
      opened: 0
    },
    analytics: {
      campaigns: [],
      stats: null,
      currentDetailCampaign: null,
      currentDetailFilter: 'all',
      searchQuery: ''
    }
  };

  // ----------------------------------------------------
  // DOM Elements
  // ----------------------------------------------------
  // Nav Mode Switcher
  const navBtnWizard = document.getElementById('navBtnWizard');
  const navBtnAnalytics = document.getElementById('navBtnAnalytics');
  const wizardView = document.getElementById('wizardView');
  const analyticsView = document.getElementById('analyticsView');
  const headerCampaignCount = document.getElementById('headerCampaignCount');

  // Wizard Stepper & Panes
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

  // Excel / CSV & Campaign Name Inputs
  const campaignNameInput = document.getElementById('campaignNameInput');
  const publicBaseUrlInput = document.getElementById('publicBaseUrlInput');
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

  // Campaign Monitor Elements
  const progressCircle = document.getElementById('progressCircle');
  const progressPercent = document.getElementById('progressPercent');
  const campaignStatusBadge = document.getElementById('campaignStatusBadge');
  const counterCampaignName = document.getElementById('counterCampaignName');
  const counterTotal = document.getElementById('counterTotal');
  const counterSent = document.getElementById('counterSent');
  const counterOpened = document.getElementById('counterOpened');
  const counterFailed = document.getElementById('counterFailed');
  const counterPending = document.getElementById('counterPending');
  const btnStartCampaign = document.getElementById('btnStartCampaign');
  const btnPauseResumeCampaign = document.getElementById('btnPauseResumeCampaign');
  const btnStopCampaign = document.getElementById('btnStopCampaign');
  const summaryDelayText = document.getElementById('summaryDelayText');
  const terminalLogs = document.getElementById('terminalLogs');
  const btnClearLogs = document.getElementById('btnClearLogs');
  const recipientStatusList = document.getElementById('recipientStatusList');
  const btnExportCsv = document.getElementById('btnExportCsv');
  const btnViewInAnalytics = document.getElementById('btnViewInAnalytics');

  // Analytics View Elements
  const btnRefreshAnalytics = document.getElementById('btnRefreshAnalytics');
  const btnNewCampaignFromHistory = document.getElementById('btnNewCampaignFromHistory');
  const kpiTotalCampaigns = document.getElementById('kpiTotalCampaigns');
  const kpiTotalSent = document.getElementById('kpiTotalSent');
  const kpiTotalOpened = document.getElementById('kpiTotalOpened');
  const kpiOverallOpenRate = document.getElementById('kpiOverallOpenRate');
  const kpiTotalUnseen = document.getElementById('kpiTotalUnseen');
  const campaignSearchInput = document.getElementById('campaignSearchInput');
  const historyTableBody = document.getElementById('historyTableBody');
  const historyFilterBtns = document.querySelectorAll('.filter-group .btn-filter-pill');

  // Campaign Detail Modal Elements
  const campaignDetailModal = document.getElementById('campaignDetailModal');
  const btnCloseDetailModal = document.getElementById('btnCloseDetailModal');
  const btnCloseDetailModalFooter = document.getElementById('btnCloseDetailModalFooter');
  const modalCampaignName = document.getElementById('modalCampaignName');
  const modalCampaignMeta = document.getElementById('modalCampaignMeta');
  const modalMetricTotal = document.getElementById('modalMetricTotal');
  const modalMetricSent = document.getElementById('modalMetricSent');
  const modalMetricOpened = document.getElementById('modalMetricOpened');
  const modalMetricUnseen = document.getElementById('modalMetricUnseen');
  const modalMetricOpenRate = document.getElementById('modalMetricOpenRate');
  const modalMetricFailed = document.getElementById('modalMetricFailed');
  const modalTabs = document.querySelectorAll('.modal-tab-btn');
  const modalTabPanes = document.querySelectorAll('.modal-tab-pane');
  const modalRecipientSearch = document.getElementById('modalRecipientSearch');
  const modalRecipFilterBtns = document.querySelectorAll('.status-filter-pills .btn-filter-pill');
  const modalRecipientsTableBody = document.getElementById('modalRecipientsTableBody');
  const countFilterAll = document.getElementById('countFilterAll');
  const countFilterSeen = document.getElementById('countFilterSeen');
  const countFilterUnseen = document.getElementById('countFilterUnseen');
  const countFilterFailed = document.getElementById('countFilterFailed');
  const modalEmailSubject = document.getElementById('modalEmailSubject');
  const modalEmailSender = document.getElementById('modalEmailSender');
  const modalEmailAttachments = document.getElementById('modalEmailAttachments');
  const modalEmailHtmlBody = document.getElementById('modalEmailHtmlBody');
  const btnDownloadOriginalExcel = document.getElementById('btnDownloadOriginalExcel');
  const btnDownloadAnalyticsReport = document.getElementById('btnDownloadAnalyticsReport');

  // Modals & Helpers
  const testEmailModal = document.getElementById('testEmailModal');
  const btnCloseTestModal = document.getElementById('btnCloseTestModal');
  const btnCancelTestModal = document.getElementById('btnCancelTestModal');
  const btnConfirmSendTest = document.getElementById('btnConfirmSendTest');
  const testTargetEmail = document.getElementById('testTargetEmail');
  const helpModal = document.getElementById('helpModal');
  const btnHelpModal = document.getElementById('btnHelpModal');
  const btnCloseHelpModal = document.getElementById('btnCloseHelpModal');
  const toastContainer = document.getElementById('toastContainer');
  const themeSelector = document.getElementById('themeSelector');
  const btnToggleMode = document.getElementById('btnToggleMode');

  // Circle Progress Radius
  const CIRCLE_RADIUS = 66;
  const CIRCLE_CIRCUMFERENCE = 2 * Math.PI * CIRCLE_RADIUS;
  if (progressCircle) {
    progressCircle.style.strokeDasharray = `${CIRCLE_CIRCUMFERENCE} ${CIRCLE_CIRCUMFERENCE}`;
    progressCircle.style.strokeDashoffset = CIRCLE_CIRCUMFERENCE;
  }

  // Auto-detect base URL for tracking
  if (publicBaseUrlInput) {
    const savedBaseUrl = localStorage.getItem('automailer_base_url');
    publicBaseUrlInput.value = savedBaseUrl || window.location.origin;
    publicBaseUrlInput.addEventListener('change', () => {
      localStorage.setItem('automailer_base_url', publicBaseUrlInput.value.trim());
    });
  }

  // ----------------------------------------------------
  // Toast Helper
  // ----------------------------------------------------
  function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    let icon = 'fa-info-circle';
    if (type === 'success') icon = 'fa-circle-check';
    if (type === 'error') icon = 'fa-circle-exclamation';

    toast.innerHTML = `<i class="fa-solid ${icon}"></i> <span>${message}</span>`;
    toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.animation = 'slideIn 0.3s ease reverse forwards';
      setTimeout(() => toast.remove(), 300);
    }, 4500);
  }

  // ----------------------------------------------------
  // View Switcher (Wizard vs History & Analytics)
  // ----------------------------------------------------
  function switchView(viewName) {
    state.currentView = viewName;
    if (viewName === 'wizardView') {
      wizardView.style.display = 'block';
      analyticsView.style.display = 'none';
      navBtnWizard.classList.add('active');
      navBtnAnalytics.classList.remove('active');
    } else {
      wizardView.style.display = 'none';
      analyticsView.style.display = 'block';
      navBtnWizard.classList.remove('active');
      navBtnAnalytics.classList.add('active');
      loadAnalyticsData();
    }
  }

  navBtnWizard.addEventListener('click', () => switchView('wizardView'));
  navBtnAnalytics.addEventListener('click', () => switchView('analyticsView'));
  if (btnViewInAnalytics) {
    btnViewInAnalytics.addEventListener('click', () => switchView('analyticsView'));
  }
  if (btnNewCampaignFromHistory) {
    btnNewCampaignFromHistory.addEventListener('click', () => {
      switchView('wizardView');
      goToStep(2);
    });
  }

  // ----------------------------------------------------
  // Step Navigation Logic
  // ----------------------------------------------------
  function goToStep(stepNumber) {
    state.currentStep = stepNumber;
    stepBtns.forEach((btn) => {
      const step = parseInt(btn.getAttribute('data-step'), 10);
      btn.classList.toggle('active', step === stepNumber);
      btn.classList.toggle('completed', step < stepNumber);
    });

    stepPanes.forEach((pane) => {
      pane.classList.toggle('active', pane.id === `step${stepNumber}`);
    });

    window.scrollTo({ top: 0, behavior: 'smooth' });

    if (stepNumber === 3) updateLivePreview();
    if (stepNumber === 4) refreshCampaignSummary();
  }

  stepBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      const targetStep = parseInt(btn.getAttribute('data-step'), 10);
      if (targetStep === 2 && !state.smtp.user) {
        showToast('Please enter your SMTP details first.', 'error');
        return;
      }
      if (targetStep >= 3 && state.excel.validEmails === 0) {
        showToast('Please upload an Excel file with valid emails in Step 2.', 'error');
        return;
      }
      goToStep(targetStep);
    });
  });

  // ----------------------------------------------------
  // SMTP Configuration & Presets
  // ----------------------------------------------------
  const SMTP_PRESETS = {
    gmail: { host: 'smtp.gmail.com', port: 587, secure: 'false' },
    outlook: { host: 'smtp-mail.outlook.com', port: 587, secure: 'false' },
    brevo: { host: 'smtp-relay.brevo.com', port: 587, secure: 'false' },
    custom: { host: '', port: 587, secure: 'false' }
  };

  presetBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      presetBtns.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      const preset = SMTP_PRESETS[btn.getAttribute('data-preset')];
      if (preset) {
        smtpHostInput.value = preset.host;
        smtpPortInput.value = preset.port;
        smtpSecureSelect.value = preset.secure;
      }
    });
  });

  btnTogglePass.addEventListener('click', () => {
    const isPass = smtpPassInput.type === 'password';
    smtpPassInput.type = isPass ? 'text' : 'password';
    btnTogglePass.innerHTML = isPass ? '<i class="fa-regular fa-eye-slash"></i>' : '<i class="fa-regular fa-eye"></i>';
  });

  function getSmtpConfig() {
    return {
      host: smtpHostInput.value.trim(),
      port: parseInt(smtpPortInput.value, 10) || 587,
      secure: smtpSecureSelect.value === 'true',
      user: smtpUserInput.value.trim(),
      pass: smtpPassInput.value.trim(),
      fromName: smtpFromNameInput.value.trim(),
      fromEmail: smtpFromEmailInput.value.trim()
    };
  }

  btnTestConnection.addEventListener('click', async () => {
    const config = getSmtpConfig();
    if (!config.host || !config.user || !config.pass) {
      showToast('Please enter SMTP Host, User Email, and Password.', 'error');
      return;
    }

    btnTestConnection.disabled = true;
    btnTestConnection.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin"></i> Testing...';

    try {
      const res = await fetch('/api/test-smtp', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ smtp: config })
      });
      const data = await res.json();

      if (res.ok && data.success) {
        state.smtp = { ...config, verified: true };
        smtpStatusBadge.innerHTML = `<span class="status-dot connected"></span><span class="status-text">${config.user}</span>`;
        showToast('SMTP connected and verified successfully! 🚀', 'success');
      } else {
        throw new Error(data.error || 'Connection failed');
      }
    } catch (err) {
      showToast(err.message, 'error');
      smtpStatusBadge.innerHTML = `<span class="status-dot disconnected"></span><span class="status-text">Connection Failed</span>`;
    } finally {
      btnTestConnection.disabled = false;
      btnTestConnection.innerHTML = '<i class="fa-solid fa-plug-circle-check"></i> Test Connection';
    }
  });

  btnSaveStep1.addEventListener('click', () => {
    state.smtp = getSmtpConfig();
    if (!state.smtp.user || !state.smtp.pass) {
      showToast('Please provide your sender email and password.', 'error');
      return;
    }
    goToStep(2);
  });

  // ----------------------------------------------------
  // Excel Drag & Drop / Parsing
  // ----------------------------------------------------
  dropzone.addEventListener('click', () => excelFileInput.click());
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) {
      handleFileSelected(e.dataTransfer.files[0]);
    }
  });

  excelFileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFileSelected(e.target.files[0]);
    }
  });

  btnReupload.addEventListener('click', () => {
    excelFileInput.value = '';
    excelFileInput.click();
  });

  async function handleFileSelected(file) {
    if (!file) return;
    const allowed = ['.xlsx', '.xls', '.csv'];
    const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
    if (!allowed.includes(ext)) {
      showToast('Invalid file format. Please upload .xlsx, .xls, or .csv', 'error');
      return;
    }

    state.excel.file = file;
    state.excel.fileName = file.name;

    // Auto-generate a clean default campaign dump name if empty
    if (campaignNameInput && (!campaignNameInput.value || campaignNameInput.value.startsWith('Campaign -'))) {
      const cleanBase = file.name.replace(/\.[^/.]+$/, '').replace(/[_-]/g, ' ');
      const dateStr = new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
      campaignNameInput.value = `${cleanBase} (${dateStr})`;
    }

    const formData = new FormData();
    formData.append('file', file);

    dropzone.innerHTML = `<div class="dropzone-content"><i class="fa-solid fa-circle-notch fa-spin dropzone-icon text-accent"></i><h3>Parsing Excel Spreadsheet...</h3><p>Extracting rows, headers, and validating email records</p></div>`;

    try {
      const res = await fetch('/api/parse-excel', { method: 'POST', body: formData });
      const data = await res.json();

      if (res.ok && data.success) {
        state.excel.totalRows = data.totalRows;
        state.excel.validEmails = data.validEmails;
        state.excel.invalidEmails = data.invalidEmails;
        state.excel.columns = data.columns;
        state.excel.emailColumn = data.detectedEmailColumn;
        state.excel.nameColumn = data.detectedNameColumn;
        state.excel.fullRows = data.fullRows;
        state.excel.sampleRows = data.sampleRows;

        renderExcelDataPreview();
        renderDynamicTags();
        showToast(`Loaded ${data.validEmails.toLocaleString()} valid recipients from ${file.name}!`, 'success');
      } else {
        throw new Error(data.error || 'Failed to parse Excel file');
      }
    } catch (err) {
      showToast(err.message, 'error');
      resetDropzone();
    }
  }

  function resetDropzone() {
    dropzone.style.display = 'block';
    dataPreviewSection.style.display = 'none';
    btnReupload.style.display = 'none';
    dropzone.innerHTML = `
      <input type="file" id="excelFileInput" accept=".xlsx, .xls, .csv" hidden>
      <div class="dropzone-content">
        <div class="dropzone-icon"><i class="fa-solid fa-cloud-arrow-up"></i></div>
        <h3>Drag & Drop your Excel (.xlsx, .xls) or .CSV file here</h3>
        <p>or click to browse files from your computer</p>
        <div class="dropzone-hints">
          <span><i class="fa-solid fa-check"></i> Supports 10,000+ rows</span>
          <span><i class="fa-solid fa-check"></i> Auto-detects column headers</span>
          <span><i class="fa-solid fa-check"></i> Saved permanently in SQLite</span>
        </div>
      </div>
    `;
    const newFileInput = document.getElementById('excelFileInput');
    newFileInput.addEventListener('change', (e) => {
      if (e.target.files.length > 0) handleFileSelected(e.target.files[0]);
    });
  }

  function renderExcelDataPreview() {
    dropzone.style.display = 'none';
    dataPreviewSection.style.display = 'block';
    btnReupload.style.display = 'inline-flex';

    previewFileName.textContent = state.excel.fileName;
    previewValidCount.textContent = state.excel.validEmails.toLocaleString();
    previewInvalidCount.textContent = state.excel.invalidEmails.toLocaleString();
    previewColCount.textContent = state.excel.columns.length;

    // Populate Column Selectors
    selectEmailColumn.innerHTML = state.excel.columns
      .map((col) => `<option value="${col}" ${col === state.excel.emailColumn ? 'selected' : ''}>${col}</option>`)
      .join('');

    selectNameColumn.innerHTML = `<option value="">-- None / Use Email --</option>` + state.excel.columns
      .map((col) => `<option value="${col}" ${col === state.excel.nameColumn ? 'selected' : ''}>${col}</option>`)
      .join('');

    selectEmailColumn.addEventListener('change', (e) => {
      state.excel.emailColumn = e.target.value;
      recalculateValidEmails();
    });

    renderPreviewTable(state.excel.sampleRows);
  }

  function recalculateValidEmails() {
    const col = state.excel.emailColumn;
    let valid = 0;
    let invalid = 0;
    const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    state.excel.fullRows.forEach((row) => {
      const email = String(row[col] || '').trim();
      if (re.test(email)) valid++;
      else invalid++;
    });

    state.excel.validEmails = valid;
    state.excel.invalidEmails = invalid;
    previewValidCount.textContent = valid.toLocaleString();
    previewInvalidCount.textContent = invalid.toLocaleString();
    renderPreviewTable(state.excel.sampleRows);
  }

  function renderPreviewTable(rows) {
    const cols = state.excel.columns;
    previewTableHead.innerHTML = `<tr><th>#</th>` + cols.map((c) => `<th>${c} ${c === state.excel.emailColumn ? '📧' : ''}</th>`).join('') + `<th>Status</th></tr>`;

    const filterText = (tableFilterInput ? tableFilterInput.value : '').toLowerCase().trim();
    const filteredRows = filterText
      ? rows.filter((r) => Object.values(r).some((v) => String(v).toLowerCase().includes(filterText)))
      : rows;

    previewTableBody.innerHTML = filteredRows.map((r, i) => {
      const emailVal = String(r[state.excel.emailColumn] || '').trim();
      const isValid = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailVal);
      return `
        <tr>
          <td>${i + 1}</td>
          ${cols.map((c) => `<td>${r[c] !== undefined ? r[c] : ''}</td>`).join('')}
          <td><span class="badge-cell ${isValid ? 'valid' : 'invalid'}">${isValid ? '✓ Valid' : '✗ Invalid'}</span></td>
        </tr>
      `;
    }).join('');
  }

  if (tableFilterInput) {
    tableFilterInput.addEventListener('input', () => renderPreviewTable(state.excel.sampleRows));
  }

  btnBackToStep1.addEventListener('click', () => goToStep(1));
  btnGoToStep3.addEventListener('click', () => {
    if (state.excel.validEmails === 0) {
      showToast('No valid recipients found in selected email column.', 'error');
      return;
    }
    goToStep(3);
  });

  // ----------------------------------------------------
  // Dynamic Tags & Message Composer
  // ----------------------------------------------------
  function renderDynamicTags() {
    dynamicTagsList.innerHTML = state.excel.columns.map((col) => {
      return `<span class="tag-chip" data-tag="{{${col}}}">&#123;&#123;${col}&#125;&#125;</span>`;
    }).join('');

    dynamicTagsList.querySelectorAll('.tag-chip').forEach((chip) => {
      chip.addEventListener('click', () => {
        const tag = chip.getAttribute('data-tag');
        insertTagAtCursor(emailBodyInput, tag);
        updateLivePreview();
      });
    });

    previewRecipientSelect.innerHTML = state.excel.sampleRows.slice(0, 20).map((r, idx) => {
      const name = r[state.excel.nameColumn] || `Row #${idx + 1}`;
      const email = r[state.excel.emailColumn] || 'No email';
      return `<option value="${idx}">Row ${idx + 1}: ${name} (${email})</option>`;
    }).join('');
  }

  function insertTagAtCursor(textarea, tag) {
    const start = textarea.selectionStart || 0;
    const end = textarea.selectionEnd || 0;
    const text = textarea.value;
    textarea.value = text.substring(0, start) + tag + text.substring(end);
    textarea.focus();
    textarea.selectionStart = textarea.selectionEnd = start + tag.length;
  }

  emailSubjectInput.addEventListener('input', () => {
    state.message.subject = emailSubjectInput.value;
    updateLivePreview();
  });

  emailBodyInput.addEventListener('input', () => {
    state.message.body = emailBodyInput.value;
    updateLivePreview();
  });

  btnFmtHtml.addEventListener('click', () => {
    state.message.isHtml = true;
    btnFmtHtml.classList.add('active');
    btnFmtText.classList.remove('active');
    updateLivePreview();
  });

  btnFmtText.addEventListener('click', () => {
    state.message.isHtml = false;
    btnFmtText.classList.add('active');
    btnFmtHtml.classList.remove('active');
    updateLivePreview();
  });

  delaySlider.addEventListener('input', (e) => {
    const val = parseFloat(e.target.value);
    state.message.delaySeconds = val;
    delayVal.textContent = `${val.toFixed(1)}s`;
    if (summaryDelayText) summaryDelayText.textContent = `${val.toFixed(1)}s per email`;
    const mins = Math.round((1000 * val) / 60);
    estimatedDuration.textContent = `Estimated time for 1,000 emails: ~${mins} mins (Natural anti-spam delivery)`;
  });

  // Attachments Handling
  emailAttachmentsInput.addEventListener('change', (e) => {
    const files = Array.from(e.target.files);
    state.message.attachments = files;
    renderAttachmentChips();
  });

  function renderAttachmentChips() {
    attachmentList.innerHTML = state.message.attachments.map((file, i) => `
      <div class="attachment-chip">
        <i class="fa-solid fa-file"></i>
        <span>${file.name} (${(file.size / 1024).toFixed(1)} KB)</span>
        <button type="button" class="btn-remove-att" data-index="${i}">&times;</button>
      </div>
    `).join('');

    attachmentList.querySelectorAll('.btn-remove-att').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        const idx = parseInt(e.target.getAttribute('data-index'), 10);
        state.message.attachments.splice(idx, 1);
        renderAttachmentChips();
      });
    });
  }

  function replaceVars(template, row) {
    if (!template || !row) return '';
    let res = template;
    for (const [k, v] of Object.entries(row)) {
      const regex = new RegExp(`\\{\\{\\s*${k}\\s*\\}\\}|\\{\\s*${k}\\s*\\}`, 'gi');
      res = res.replace(regex, v !== undefined && v !== null ? String(v) : '');
    }
    return res;
  }

  function updateLivePreview() {
    const selectedIdx = parseInt(previewRecipientSelect.value, 10) || 0;
    const clientData = state.excel.sampleRows[selectedIdx] || { Name: 'John Smith', Company: 'Apex Corp', Email: 'john@example.com' };

    previewToEmail.textContent = clientData[state.excel.emailColumn] || 'recipient@example.com';
    previewSubject.textContent = replaceVars(state.message.subject, clientData);

    const renderedBody = replaceVars(state.message.body, clientData);
    if (state.message.isHtml) {
      previewBodyRendered.innerHTML = renderedBody;
    } else {
      previewBodyRendered.innerHTML = `<pre style="font-family: inherit; white-space: pre-wrap;">${renderedBody}</pre>`;
    }
  }

  previewRecipientSelect.addEventListener('change', updateLivePreview);

  btnBackToStep2.addEventListener('click', () => goToStep(2));
  btnGoToStep4.addEventListener('click', () => {
    if (!state.message.subject.trim()) {
      showToast('Please enter an email subject line.', 'error');
      return;
    }
    if (!state.message.body.trim()) {
      showToast('Please compose your message content.', 'error');
      return;
    }
    goToStep(4);
  });

  // ----------------------------------------------------
  // Test Email Modal
  // ----------------------------------------------------
  btnSendTestEmail.addEventListener('click', () => {
    testTargetEmail.value = state.smtp.fromEmail || state.smtp.user || '';
    testEmailModal.classList.add('active');
  });
  btnCloseTestModal.addEventListener('click', () => testEmailModal.classList.remove('active'));
  btnCancelTestModal.addEventListener('click', () => testEmailModal.classList.remove('active'));

  btnConfirmSendTest.addEventListener('click', async () => {
    const email = testTargetEmail.value.trim();
    if (!email) {
      showToast('Please enter your personal email address.', 'error');
      return;
    }

    btnConfirmSendTest.disabled = true;
    btnConfirmSendTest.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin"></i> Sending Test...';

    const sampleData = state.excel.sampleRows[0] || { Name: 'Demo Client', Company: 'Your Company', Email: email };

    try {
      const res = await fetch('/api/test-smtp', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          smtp: state.smtp,
          testEmail: email,
          subject: state.message.subject,
          htmlBody: state.message.isHtml ? state.message.body : undefined,
          textBody: !state.message.isHtml ? state.message.body : undefined,
          sampleData
        })
      });
      const data = await res.json();
      if (res.ok && data.success) {
        showToast(`Test email dispatched to ${email}! Check your inbox.`, 'success');
        testEmailModal.classList.remove('active');
      } else {
        throw new Error(data.error || 'Failed to send test email');
      }
    } catch (err) {
      showToast(err.message, 'error');
    } finally {
      btnConfirmSendTest.disabled = false;
      btnConfirmSendTest.innerHTML = '<i class="fa-solid fa-paper-plane"></i> Send Test Now';
    }
  });

  // ----------------------------------------------------
  // STEP 4: Campaign Dispatch & Execution
  // ----------------------------------------------------
  function refreshCampaignSummary() {
    const defaultName = campaignNameInput ? campaignNameInput.value.trim() : `Dump_${new Date().toLocaleDateString()}`;
    if (counterCampaignName) counterCampaignName.textContent = defaultName || 'New Campaign';
    counterTotal.textContent = state.excel.validEmails.toLocaleString();
    counterSent.textContent = '0';
    counterOpened.textContent = '0';
    counterFailed.textContent = '0';
    counterPending.textContent = state.excel.validEmails.toLocaleString();
    updateProgressRing(0);
    renderRecipientActivityInitial();
  }

  function renderRecipientActivityInitial() {
    const list = state.excel.fullRows.slice(0, 100);
    recipientStatusList.innerHTML = list.map((r, i) => {
      const email = r[state.excel.emailColumn];
      const name = r[state.excel.nameColumn] || `Recipient #${i + 1}`;
      return `
        <div class="recipient-item" id="recip-item-${i + 1}">
          <div class="recipient-item-info">
            <strong>${name}</strong>
            <small>${email}</small>
          </div>
          <span class="recipient-item-status pending" id="recip-status-${i + 1}">Queued</span>
        </div>
      `;
    }).join('');
  }

  function updateProgressRing(percent) {
    if (!progressCircle) return;
    const offset = CIRCLE_CIRCUMFERENCE - (percent / 100) * CIRCLE_CIRCUMFERENCE;
    progressCircle.style.strokeDashoffset = offset;
    progressPercent.textContent = `${Math.round(percent)}%`;
  }

  btnStartCampaign.addEventListener('click', async () => {
    if (state.excel.validEmails === 0) {
      showToast('No recipients queued to send.', 'error');
      return;
    }

    const campaignDumpName = (campaignNameInput ? campaignNameInput.value : '').trim() || `Campaign - ${new Date().toLocaleString()}`;
    const baseUrl = (publicBaseUrlInput ? publicBaseUrlInput.value : '').trim() || window.location.origin;

    btnStartCampaign.disabled = true;
    btnStartCampaign.innerHTML = '<i class="fa-solid fa-circle-notch fa-spin"></i> Initializing...';

    const formData = new FormData();
    formData.append('campaignName', campaignDumpName);
    formData.append('publicBaseUrl', baseUrl);
    formData.append('originalFilename', state.excel.fileName || 'recipients.xlsx');
    formData.append('smtp', JSON.stringify(state.smtp));
    formData.append('emailColumn', state.excel.emailColumn);
    formData.append('subject', state.message.subject);
    formData.append('htmlBody', state.message.isHtml ? state.message.body : '');
    formData.append('textBody', !state.message.isHtml ? state.message.body : '');
    formData.append('delaySeconds', state.message.delaySeconds);
    formData.append('recipients', JSON.stringify(state.excel.fullRows));

    // Append attachments and original Excel file
    state.message.attachments.forEach((file) => formData.append('attachments', file));
    if (state.excel.file) {
      formData.append('excelFile', state.excel.file);
    }

    try {
      const res = await fetch('/api/campaign/start', { method: 'POST', body: formData });
      const data = await res.json();

      if (res.ok && data.success) {
        state.campaign.id = data.campaignId;
        state.campaign.name = campaignDumpName;
        state.campaign.status = 'running';
        
        btnStartCampaign.style.display = 'none';
        btnPauseResumeCampaign.style.display = 'inline-flex';
        btnStopCampaign.style.display = 'inline-flex';
        campaignStatusBadge.textContent = 'RUNNING';
        campaignStatusBadge.className = 'progress-status-badge running';
        if (btnExportCsv) btnExportCsv.disabled = true;

        showToast(`Campaign "${campaignDumpName}" launched! Dispatching emails...`, 'success');
        loadAnalyticsData(); // Update counts in header
      } else {
        throw new Error(data.error || 'Failed to start campaign');
      }
    } catch (err) {
      showToast(err.message, 'error');
      btnStartCampaign.disabled = false;
      btnStartCampaign.innerHTML = '<i class="fa-solid fa-play"></i> START BULK SENDING';
    }
  });

  btnPauseResumeCampaign.addEventListener('click', async () => {
    if (state.campaign.status === 'running') {
      const res = await fetch('/api/campaign/pause', { method: 'POST' });
      if (res.ok) {
        state.campaign.status = 'paused';
        btnPauseResumeCampaign.innerHTML = '<i class="fa-solid fa-play"></i> Resume';
        campaignStatusBadge.textContent = 'PAUSED';
        campaignStatusBadge.className = 'progress-status-badge paused';
      }
    } else if (state.campaign.status === 'paused') {
      const res = await fetch('/api/campaign/resume', { method: 'POST' });
      if (res.ok) {
        state.campaign.status = 'running';
        btnPauseResumeCampaign.innerHTML = '<i class="fa-solid fa-pause"></i> Pause';
        campaignStatusBadge.textContent = 'RUNNING';
        campaignStatusBadge.className = 'progress-status-badge running';
      }
    }
  });

  btnStopCampaign.addEventListener('click', async () => {
    if (confirm('Are you sure you want to stop and reset the campaign?')) {
      await fetch('/api/campaign/stop', { method: 'POST' });
      state.campaign.status = 'stopped';
      campaignStatusBadge.textContent = 'STOPPED';
      campaignStatusBadge.className = 'progress-status-badge failed';
      btnStartCampaign.style.display = 'inline-flex';
      btnStartCampaign.disabled = false;
      btnStartCampaign.innerHTML = '<i class="fa-solid fa-play"></i> RE-START CAMPAIGN';
      btnPauseResumeCampaign.style.display = 'none';
      btnStopCampaign.style.display = 'none';
      if (btnExportCsv) btnExportCsv.disabled = false;
      if (btnViewInAnalytics) btnViewInAnalytics.style.display = 'inline-flex';
      loadAnalyticsData();
    }
  });

  if (btnClearLogs) {
    btnClearLogs.addEventListener('click', () => {
      terminalLogs.innerHTML = '<div class="log-line info">[System] Audit logs cleared.</div>';
    });
  }

  if (btnExportCsv) {
    btnExportCsv.addEventListener('click', () => {
      window.location.href = '/api/campaign/export-csv';
    });
  }

  // ----------------------------------------------------
  // SSE Real-Time Progress & Open Events Listener
  // ----------------------------------------------------
  function initSSE() {
    const eventSource = new EventSource('/api/campaign/events');

    eventSource.onmessage = (e) => {
      try {
        const data = JSON.parse(e.data);

        // Real-Time Open Tracking Event
        if (data.type === 'open_event') {
          if (counterOpened) {
            const current = parseInt(counterOpened.textContent, 10) || 0;
            counterOpened.textContent = (current + 1).toLocaleString();
          }
          appendTerminalLog('info', `👁️ Email opened by ${data.email || 'recipient'} at ${data.openedAt}!`);
          showToast(`👁️ Email opened by ${data.email || 'recipient'}!`, 'info');
          // If modal or analytics view is open, refresh
          if (state.currentView === 'analyticsView') loadAnalyticsData();
          return;
        }

        if (data.type === 'log') {
          appendTerminalLog(data.log.type, `[${data.log.timestamp}] ${data.log.message}`);
        }

        if (data.campaign) {
          updateCampaignCounters(data.campaign);
        }

        if (data.currentRecipient) {
          updateRecipientItem(data.currentRecipient);
        }

        if (data.type === 'completed') {
          campaignStatusBadge.textContent = 'COMPLETED';
          campaignStatusBadge.className = 'progress-status-badge completed';
          btnStartCampaign.style.display = 'inline-flex';
          btnStartCampaign.disabled = false;
          btnStartCampaign.innerHTML = '<i class="fa-solid fa-rotate"></i> Start New Campaign';
          btnPauseResumeCampaign.style.display = 'none';
          btnStopCampaign.style.display = 'none';
          if (btnExportCsv) btnExportCsv.disabled = false;
          if (btnViewInAnalytics) btnViewInAnalytics.style.display = 'inline-flex';
          showToast('Campaign successfully completed! 🎉', 'success');
          loadAnalyticsData();
        }
      } catch (err) {
        console.error('SSE parse error:', err);
      }
    };

    eventSource.onerror = () => {
      console.warn('SSE connection lost. Reconnecting in 3s...');
      setTimeout(initSSE, 3000);
    };
  }

  function appendTerminalLog(type, message) {
    const line = document.createElement('div');
    line.className = `log-line ${type}`;
    line.textContent = message;
    terminalLogs.appendChild(line);
    terminalLogs.scrollTop = terminalLogs.scrollHeight;
  }

  function updateCampaignCounters(c) {
    if (counterTotal) counterTotal.textContent = (c.total || 0).toLocaleString();
    if (counterSent) counterSent.textContent = (c.sent || 0).toLocaleString();
    if (counterFailed) counterFailed.textContent = (c.failed || 0).toLocaleString();
    if (counterPending) counterPending.textContent = (c.pending || 0).toLocaleString();
    if (counterOpened && c.opened !== undefined) counterOpened.textContent = (c.opened || 0).toLocaleString();

    if (c.total > 0) {
      const processed = (c.sent || 0) + (c.failed || 0);
      const pct = (processed / c.total) * 100;
      updateProgressRing(pct);
    }
  }

  function updateRecipientItem(recipient) {
    const item = document.getElementById(`recip-item-${recipient.index}`);
    const badge = document.getElementById(`recip-status-${recipient.index}`);
    if (item && badge) {
      badge.className = `recipient-item-status ${recipient.status}`;
      badge.textContent = recipient.status.toUpperCase();
      if (recipient.status === 'sent') {
        item.style.borderColor = 'rgba(16, 185, 129, 0.4)';
      } else if (recipient.status === 'failed') {
        item.style.borderColor = 'rgba(239, 68, 68, 0.4)';
        item.title = recipient.error || 'Failed';
      }
    }
  }

  // ----------------------------------------------------
  // HISTORY & ANALYTICS DASHBOARD LOGIC
  // ----------------------------------------------------
  async function loadAnalyticsData() {
    try {
      const res = await fetch('/api/campaigns');
      const data = await res.json();

      if (res.ok && data.success) {
        state.analytics.campaigns = data.campaigns || [];
        state.analytics.stats = data.stats || {};

        // Update Header Badge
        if (headerCampaignCount) {
          headerCampaignCount.textContent = state.analytics.campaigns.length;
        }

        renderKpis(data.stats);
        renderHistoricalCampaignsTable();
      }
    } catch (err) {
      console.error('Error fetching analytics:', err);
    }
  }

  function renderKpis(stats) {
    if (!stats) return;
    if (kpiTotalCampaigns) kpiTotalCampaigns.textContent = (stats.total_campaigns || 0).toLocaleString();
    if (kpiTotalSent) kpiTotalSent.textContent = (stats.total_sent || 0).toLocaleString();
    if (kpiTotalOpened) kpiTotalOpened.textContent = (stats.total_opened || 0).toLocaleString();
    if (kpiOverallOpenRate) kpiOverallOpenRate.textContent = `${stats.overall_open_rate || 0}% Overall Open Rate`;
    if (kpiTotalUnseen) kpiTotalUnseen.textContent = (stats.total_unseen || 0).toLocaleString();
  }

  function renderHistoricalCampaignsTable() {
    if (!historyTableBody) return;

    let campaigns = state.analytics.campaigns;
    const filter = document.querySelector('.filter-group .btn-filter-pill.active')?.getAttribute('data-filter') || 'all';
    const search = (campaignSearchInput ? campaignSearchInput.value : '').toLowerCase().trim();

    if (filter !== 'all') {
      campaigns = campaigns.filter((c) => (c.status || '').toLowerCase() === filter);
    }
    if (search) {
      campaigns = campaigns.filter((c) =>
        (c.name || '').toLowerCase().includes(search) ||
        (c.subject || '').toLowerCase().includes(search) ||
        (c.created_at || '').toLowerCase().includes(search)
      );
    }

    if (campaigns.length === 0) {
      historyTableBody.innerHTML = `
        <tr>
          <td colspan="9" class="text-center py-5">
            <div class="empty-state">
              <i class="fa-solid fa-folder-open"></i>
              <p>No campaign records found. Create and dispatch your first email campaign above!</p>
            </div>
          </td>
        </tr>
      `;
      return;
    }

    historyTableBody.innerHTML = campaigns.map((c) => {
      const dateFormatted = new Date(c.created_at).toLocaleDateString('en-US', {
        month: 'short', day: 'numeric', year: 'numeric', hour: '2-digit', minute: '2-digit'
      });
      const openRate = c.open_rate_pct || 0;
      const unseen = c.unseen_count !== undefined ? c.unseen_count : (c.total_recipients - c.opened_count);

      let statusBadgeClass = 'valid';
      if (c.status === 'running') statusBadgeClass = 'text-info';
      if (c.status === 'paused') statusBadgeClass = 'text-warning';
      if (c.status === 'failed') statusBadgeClass = 'invalid';

      return `
        <tr>
          <td>
            <strong>${c.name}</strong>
            <br><small class="text-muted">${c.original_filename || 'Excel Dump'}</small>
          </td>
          <td><small>${dateFormatted}</small></td>
          <td><strong>${c.total_recipients.toLocaleString()}</strong></td>
          <td><span class="text-success font-bold">${c.sent_count.toLocaleString()}</span></td>
          <td><span class="badge-seen"><i class="fa-solid fa-eye"></i> ${c.opened_count.toLocaleString()}</span></td>
          <td><span class="badge-unseen"><i class="fa-solid fa-eye-slash"></i> ${unseen.toLocaleString()}</span></td>
          <td>
            <div class="open-rate-pill">
              <div class="open-rate-bar-bg">
                <div class="open-rate-bar-fill" style="width: ${Math.min(100, openRate)}%;"></div>
              </div>
              <span class="open-rate-pct-text">${openRate}%</span>
            </div>
          </td>
          <td><span class="badge-cell ${statusBadgeClass}">${(c.status || 'COMPLETED').toUpperCase()}</span></td>
          <td>
            <div class="table-actions-cell">
              <button class="btn-action-icon btn-view-detail" data-id="${c.id}" title="View Details & Recipient Open Log">
                <i class="fa-solid fa-chart-simple"></i>
              </button>
              <a href="/api/campaigns/${c.id}/download-original" class="btn-action-icon" title="Download Original Excel" download>
                <i class="fa-solid fa-file-excel"></i>
              </a>
              <a href="/api/campaigns/${c.id}/export-analytics" class="btn-action-icon" title="Export Analytics Report (.xlsx)" download>
                <i class="fa-solid fa-file-export"></i>
              </a>
              <button class="btn-action-icon btn-delete" data-id="${c.id}" title="Delete Record">
                <i class="fa-solid fa-trash"></i>
              </button>
            </div>
          </td>
        </tr>
      `;
    }).join('');

    // Attach Event Listeners to actions
    historyTableBody.querySelectorAll('.btn-view-detail').forEach((btn) => {
      btn.addEventListener('click', () => openCampaignDetailModal(btn.getAttribute('data-id')));
    });

    historyTableBody.querySelectorAll('.btn-delete').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const id = btn.getAttribute('data-id');
        if (confirm('Are you sure you want to permanently delete this campaign dump record?')) {
          try {
            const res = await fetch(`/api/campaigns/${id}`, { method: 'DELETE' });
            if (res.ok) {
              showToast('Campaign record deleted successfully', 'success');
              loadAnalyticsData();
            }
          } catch (err) {
            showToast('Failed to delete campaign', 'error');
          }
        }
      });
    });
  }

  historyFilterBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      historyFilterBtns.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      renderHistoricalCampaignsTable();
    });
  });

  if (campaignSearchInput) {
    campaignSearchInput.addEventListener('input', renderHistoricalCampaignsTable);
  }

  if (btnRefreshAnalytics) {
    btnRefreshAnalytics.addEventListener('click', async () => {
      btnRefreshAnalytics.innerHTML = '<i class="fa-solid fa-rotate fa-spin"></i> Refreshing...';
      await loadAnalyticsData();
      btnRefreshAnalytics.innerHTML = '<i class="fa-solid fa-rotate"></i> Refresh Data';
      showToast('Analytics refreshed!', 'info');
    });
  }

  // ----------------------------------------------------
  // CAMPAIGN DETAIL MODAL LOGIC (3 TABS)
  // ----------------------------------------------------
  async function openCampaignDetailModal(campaignId) {
    // Reset tabs to Tab 1 (Recipients)
    modalTabs.forEach((t) => t.classList.remove('active'));
    modalTabPanes.forEach((p) => (p.style.display = 'none'));
    const firstTab = document.querySelector('.modal-tab-btn[data-tab="tabRecipients"]');
    if (firstTab) firstTab.classList.add('active');
    const firstPane = document.getElementById('tabRecipients');
    if (firstPane) firstPane.style.display = 'block';

    // Reset filter
    state.analytics.currentDetailFilter = 'all';
    modalRecipFilterBtns.forEach((b) => b.classList.toggle('active', b.getAttribute('data-recip-filter') === 'all'));
    if (modalRecipientSearch) modalRecipientSearch.value = '';

    campaignDetailModal.classList.add('active');
    modalCampaignName.textContent = 'Loading Campaign...';
    modalRecipientsTableBody.innerHTML = `<tr><td colspan="7" class="text-center py-4"><i class="fa-solid fa-circle-notch fa-spin"></i> Loading recipient engagement data...</td></tr>`;

    try {
      const res = await fetch(`/api/campaigns/${campaignId}`);
      const data = await res.json();

      if (res.ok && data.success) {
        state.analytics.currentDetailCampaign = data.campaign;
        renderCampaignDetailModal(data.campaign);
      } else {
        throw new Error(data.error || 'Failed to load details');
      }
    } catch (err) {
      showToast(err.message, 'error');
      campaignDetailModal.classList.remove('active');
    }
  }

  function renderCampaignDetailModal(c) {
    modalCampaignName.textContent = c.name;
    const dateFormatted = new Date(c.created_at).toLocaleString();
    modalCampaignMeta.textContent = `Created: ${dateFormatted} • Stored File: ${c.original_filename || 'recipients.xlsx'}`;

    modalMetricTotal.textContent = (c.total_recipients || 0).toLocaleString();
    modalMetricSent.textContent = (c.sent_count || 0).toLocaleString();
    modalMetricOpened.textContent = (c.opened_count || 0).toLocaleString();
    modalMetricUnseen.textContent = (c.unseen_count !== undefined ? c.unseen_count : (c.total_recipients - c.opened_count)).toLocaleString();
    modalMetricOpenRate.textContent = `${c.open_rate_pct || 0}%`;
    modalMetricFailed.textContent = (c.failed_count || 0).toLocaleString();

    // Tab 1: Counts for filter pills
    const all = c.recipients || [];
    const seen = all.filter((r) => r.is_opened === 1);
    const unseen = all.filter((r) => r.is_opened === 0);
    const failed = all.filter((r) => r.send_status === 'failed');

    if (countFilterAll) countFilterAll.textContent = all.length;
    if (countFilterSeen) countFilterSeen.textContent = seen.length;
    if (countFilterUnseen) countFilterUnseen.textContent = unseen.length;
    if (countFilterFailed) countFilterFailed.textContent = failed.length;

    renderModalRecipientsTable();

    // Tab 2: Content Sent
    modalEmailSubject.textContent = c.subject || '(No subject)';
    modalEmailSender.textContent = c.smtp_from || '(Default Sender)';
    modalEmailAttachments.textContent = (c.attachmentNames && c.attachmentNames.length > 0)
      ? c.attachmentNames.join(', ')
      : 'None';
    modalEmailHtmlBody.innerHTML = c.html_body || `<pre style="white-space: pre-wrap;">${c.text_body || ''}</pre>`;

    // Tab 3: Downloads
    btnDownloadOriginalExcel.href = `/api/campaigns/${c.id}/download-original`;
    btnDownloadAnalyticsReport.href = `/api/campaigns/${c.id}/export-analytics`;
  }

  function renderModalRecipientsTable() {
    const c = state.analytics.currentDetailCampaign;
    if (!c || !c.recipients) return;

    const filter = state.analytics.currentDetailFilter || 'all';
    const search = (modalRecipientSearch ? modalRecipientSearch.value : '').toLowerCase().trim();

    let recipients = c.recipients;

    if (filter === 'seen') recipients = recipients.filter((r) => r.is_opened === 1);
    if (filter === 'unseen') recipients = recipients.filter((r) => r.is_opened === 0);
    if (filter === 'failed') recipients = recipients.filter((r) => r.send_status === 'failed');

    if (search) {
      recipients = recipients.filter((r) =>
        (r.email || '').toLowerCase().includes(search) ||
        (r.name || '').toLowerCase().includes(search)
      );
    }

    if (recipients.length === 0) {
      modalRecipientsTableBody.innerHTML = `<tr><td colspan="7" class="text-center py-4 text-muted">No recipients match the selected filter.</td></tr>`;
      return;
    }

    modalRecipientsTableBody.innerHTML = recipients.map((r) => {
      const isOpened = r.is_opened === 1;
      const openedAtFormatted = r.opened_at ? new Date(r.opened_at).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' }) : '--';

      let sendStatusBadge = `<span class="badge-cell valid">SENT</span>`;
      if (r.send_status === 'failed') {
        sendStatusBadge = `<span class="badge-cell invalid" title="${r.error_message || ''}">FAILED</span>`;
      } else if (r.send_status === 'pending') {
        sendStatusBadge = `<span class="badge-cell">PENDING</span>`;
      }

      const seenStatusBadge = isOpened
        ? `<span class="badge-seen"><i class="fa-solid fa-eye"></i> SEEN</span>`
        : `<span class="badge-unseen"><i class="fa-solid fa-eye-slash"></i> UNSEEN</span>`;

      return `
        <tr>
          <td>${r.recipient_index}</td>
          <td><strong>${r.email}</strong></td>
          <td>${r.name || '--'}</td>
          <td>${sendStatusBadge}</td>
          <td>${seenStatusBadge}</td>
          <td><small>${openedAtFormatted}</small></td>
          <td><span class="font-bold">${r.open_count || 0}</span></td>
        </tr>
      `;
    }).join('');
  }

  // Modal Tabs Switching Listener
  modalTabs.forEach((tab) => {
    tab.addEventListener('click', () => {
      modalTabs.forEach((t) => t.classList.remove('active'));
      modalTabPanes.forEach((p) => (p.style.display = 'none'));

      tab.classList.add('active');
      const target = tab.getAttribute('data-tab');
      const targetPane = document.getElementById(target);
      if (targetPane) targetPane.style.display = 'block';
    });
  });

  // Modal Recipient Filter Pills
  modalRecipFilterBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      modalRecipFilterBtns.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      state.analytics.currentDetailFilter = btn.getAttribute('data-recip-filter');
      renderModalRecipientsTable();
    });
  });

  if (modalRecipientSearch) {
    modalRecipientSearch.addEventListener('input', renderModalRecipientsTable);
  }

  if (btnCloseDetailModal) {
    btnCloseDetailModal.addEventListener('click', () => campaignDetailModal.classList.remove('active'));
  }
  if (btnCloseDetailModalFooter) {
    btnCloseDetailModalFooter.addEventListener('click', () => campaignDetailModal.classList.remove('active'));
  }

  // ----------------------------------------------------
  // INTERACTIVE KPI BUTTONS & RECORDS MODAL LOGIC
  // ----------------------------------------------------
  const kpiBtnDumps = document.getElementById('kpiBtnDumps');
  const kpiBtnDispatched = document.getElementById('kpiBtnDispatched');
  const kpiBtnSeen = document.getElementById('kpiBtnSeen');
  const kpiBtnUnseen = document.getElementById('kpiBtnUnseen');

  const kpiDetailModal = document.getElementById('kpiDetailModal');
  const btnCloseKpiModal = document.getElementById('btnCloseKpiModal');
  const btnCloseKpiModalFooter = document.getElementById('btnCloseKpiModalFooter');
  const kpiModalIcon = document.getElementById('kpiModalIcon');
  const kpiModalTitle = document.getElementById('kpiModalTitle');
  const kpiModalSubtitle = document.getElementById('kpiModalSubtitle');
  const kpiModalSearch = document.getElementById('kpiModalSearch');
  const btnCopyKpiEmails = document.getElementById('btnCopyKpiEmails');
  const btnExportKpiCsv = document.getElementById('btnExportKpiCsv');
  const kpiModalTableHead = document.getElementById('kpiModalTableHead');
  const kpiModalTableBody = document.getElementById('kpiModalTableBody');

  let currentKpiType = 'dumps';
  let currentKpiRows = [];

  if (kpiBtnDumps) {
    kpiBtnDumps.addEventListener('click', () => openKpiDetailModal('dumps'));
  }
  if (kpiBtnDispatched) {
    kpiBtnDispatched.addEventListener('click', () => openKpiDetailModal('dispatched'));
  }
  if (kpiBtnSeen) {
    kpiBtnSeen.addEventListener('click', () => openKpiDetailModal('seen'));
  }
  if (kpiBtnUnseen) {
    kpiBtnUnseen.addEventListener('click', () => openKpiDetailModal('unseen'));
  }

  async function openKpiDetailModal(type) {
    currentKpiType = type;
    if (kpiModalSearch) kpiModalSearch.value = '';
    kpiDetailModal.classList.add('active');

    // Configure header based on KPI Type
    if (type === 'dumps') {
      kpiModalIcon.className = 'modal-icon-badge bg-purple';
      kpiModalIcon.innerHTML = '<i class="fa-solid fa-folder-tree"></i>';
      kpiModalTitle.textContent = 'Stored Email Dumps & Spreadsheets Records';
      kpiModalSubtitle.textContent = 'All archived campaigns with their original Excel files and email content';
      btnCopyKpiEmails.style.display = 'none';
      
      currentKpiRows = state.analytics.campaigns || [];
      renderKpiTable();
    } else {
      btnCopyKpiEmails.style.display = 'inline-flex';
      let iconClass = 'bg-blue';
      let iconHtml = '<i class="fa-solid fa-paper-plane"></i>';
      let title = 'All Dispatched Email Records';
      let subtitle = 'Individual recipients dispatched across all campaigns';

      if (type === 'seen') {
        iconClass = 'bg-green';
        iconHtml = '<i class="fa-solid fa-eye"></i>';
        title = 'All Seen (Opened) Email Activity';
        subtitle = 'Recipients who have successfully opened your emails with timestamp audit';
      } else if (type === 'unseen') {
        iconClass = 'bg-amber';
        iconHtml = '<i class="fa-solid fa-envelope-open-text"></i>';
        title = 'All Unseen (Unopened) Email Recipients';
        subtitle = 'Recipients pending open engagement';
      }

      kpiModalIcon.className = `modal-icon-badge ${iconClass}`;
      kpiModalIcon.innerHTML = iconHtml;
      kpiModalTitle.textContent = title;
      kpiModalSubtitle.textContent = 'Loading records from SQLite database...';

      kpiModalTableHead.innerHTML = `<tr><th>#</th><th>Recipient Email</th><th>Name</th><th>Campaign / Dump</th><th>Timestamp</th><th>Status</th><th>Action</th></tr>`;
      kpiModalTableBody.innerHTML = `<tr><td colspan="7" class="text-center py-5"><i class="fa-solid fa-circle-notch fa-spin"></i> Loading records...</td></tr>`;

      try {
        const res = await fetch(`/api/analytics/recipients?type=${type}`);
        const data = await res.json();
        if (res.ok && data.success) {
          currentKpiRows = data.recipients || [];
          kpiModalSubtitle.textContent = `Found ${currentKpiRows.length.toLocaleString()} records stored in SQLite`;
          renderKpiTable();
        } else {
          throw new Error(data.error || 'Failed to fetch records');
        }
      } catch (err) {
        showToast(err.message, 'error');
        kpiModalTableBody.innerHTML = `<tr><td colspan="7" class="text-center text-danger py-4">${err.message}</td></tr>`;
      }
    }
  }

  function renderKpiTable() {
    const search = (kpiModalSearch ? kpiModalSearch.value : '').toLowerCase().trim();

    if (currentKpiType === 'dumps') {
      kpiModalTableHead.innerHTML = `
        <tr>
          <th>#</th>
          <th>Dump / Campaign Name</th>
          <th>Created At</th>
          <th>Stored Excel File</th>
          <th>Recipients</th>
          <th>Dispatched</th>
          <th>Seen (Opened)</th>
          <th>Unseen</th>
          <th style="text-align: right;">Actions</th>
        </tr>
      `;

      let rows = currentKpiRows;
      if (search) {
        rows = rows.filter((c) =>
          (c.name || '').toLowerCase().includes(search) ||
          (c.original_filename || '').toLowerCase().includes(search) ||
          (c.subject || '').toLowerCase().includes(search)
        );
      }

      if (rows.length === 0) {
        kpiModalTableBody.innerHTML = `<tr><td colspan="9" class="text-center py-5 text-muted">No campaign dump records found.</td></tr>`;
        return;
      }

      kpiModalTableBody.innerHTML = rows.map((c, i) => {
        const dateFormatted = new Date(c.created_at).toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' });
        const unseen = c.unseen_count !== undefined ? c.unseen_count : (c.total_recipients - c.opened_count);
        return `
          <tr>
            <td>${i + 1}</td>
            <td>
              <strong>${c.name}</strong>
              <br><small class="text-accent">${c.subject || 'No subject'}</small>
            </td>
            <td><small>${dateFormatted}</small></td>
            <td>
              <span class="badge-cell" style="background: rgba(16, 185, 129, 0.15); color: #34d399;">
                <i class="fa-solid fa-file-excel"></i> ${c.original_filename || 'recipients.xlsx'}
              </span>
            </td>
            <td><strong>${c.total_recipients.toLocaleString()}</strong></td>
            <td><span class="text-success font-bold">${c.sent_count.toLocaleString()}</span></td>
            <td><span class="badge-seen"><i class="fa-solid fa-eye"></i> ${c.opened_count.toLocaleString()}</span></td>
            <td><span class="badge-unseen"><i class="fa-solid fa-eye-slash"></i> ${unseen.toLocaleString()}</span></td>
            <td>
              <div class="table-actions-cell">
                <button class="btn btn-outline btn-sm btn-view-from-kpi" data-id="${c.id}" title="View Email Content & Details">
                  <i class="fa-regular fa-envelope"></i> Content
                </button>
                <a href="/api/campaigns/${c.id}/download-original" class="btn btn-secondary btn-sm" title="Download Excel" download>
                  <i class="fa-solid fa-download"></i> Excel
                </a>
                <a href="/api/campaigns/${c.id}/export-analytics" class="btn btn-primary btn-sm" title="Export Full Report" download>
                  <i class="fa-solid fa-file-export"></i> Report
                </a>
              </div>
            </td>
          </tr>
        `;
      }).join('');

      kpiModalTableBody.querySelectorAll('.btn-view-from-kpi').forEach((btn) => {
        btn.addEventListener('click', () => {
          kpiDetailModal.classList.remove('active');
          openCampaignDetailModal(btn.getAttribute('data-id'));
        });
      });

    } else {
      // Dispatched, Seen, or Unseen Recipient Records
      kpiModalTableHead.innerHTML = `
        <tr>
          <th>#</th>
          <th>Recipient Email</th>
          <th>Name</th>
          <th>Campaign Dump Name</th>
          <th>${currentKpiType === 'seen' ? 'First Opened At' : 'Dispatched At'}</th>
          <th>${currentKpiType === 'seen' ? 'Open Count' : 'Engagement Status'}</th>
          <th style="text-align: right;">Action</th>
        </tr>
      `;

      let rows = currentKpiRows;
      if (search) {
        rows = rows.filter((r) =>
          (r.email || '').toLowerCase().includes(search) ||
          (r.name || '').toLowerCase().includes(search) ||
          (r.campaign_name || '').toLowerCase().includes(search) ||
          (r.campaign_subject || '').toLowerCase().includes(search)
        );
      }

      if (rows.length === 0) {
        kpiModalTableBody.innerHTML = `<tr><td colspan="7" class="text-center py-5 text-muted">No matching records found.</td></tr>`;
        return;
      }

      kpiModalTableBody.innerHTML = rows.map((r, i) => {
        const timeVal = currentKpiType === 'seen' ? r.opened_at : (r.sent_at || r.campaign_created_at);
        const timeFormatted = timeVal ? new Date(timeVal).toLocaleString([], { dateStyle: 'short', timeStyle: 'short' }) : '--';

        let statusCol = `<span class="badge-seen"><i class="fa-solid fa-eye"></i> SEEN</span>`;
        if (currentKpiType === 'seen') {
          statusCol = `<span class="badge-seen font-bold">👁️ ${r.open_count || 1} Opens</span>`;
        } else if (currentKpiType === 'unseen') {
          statusCol = `<span class="badge-unseen"><i class="fa-solid fa-eye-slash"></i> UNSEEN</span>`;
        } else {
          statusCol = r.send_status === 'sent'
            ? `<span class="badge-cell valid">SENT</span>`
            : `<span class="badge-cell invalid">FAILED</span>`;
        }

        return `
          <tr>
            <td>${i + 1}</td>
            <td><strong>${r.email}</strong></td>
            <td>${r.name || '--'}</td>
            <td>
              <span class="text-accent font-bold">${r.campaign_name}</span>
              <br><small class="text-muted">${r.campaign_subject || ''}</small>
            </td>
            <td><small>${timeFormatted}</small></td>
            <td>${statusCol}</td>
            <td>
              <button class="btn btn-outline btn-sm btn-view-from-kpi" data-id="${r.campaign_id}" title="Inspect Campaign Dump">
                <i class="fa-solid fa-folder-open"></i> Dump
              </button>
            </td>
          </tr>
        `;
      }).join('');

      kpiModalTableBody.querySelectorAll('.btn-view-from-kpi').forEach((btn) => {
        btn.addEventListener('click', () => {
          kpiDetailModal.classList.remove('active');
          openCampaignDetailModal(btn.getAttribute('data-id'));
        });
      });
    }
  }

  if (kpiModalSearch) {
    kpiModalSearch.addEventListener('input', renderKpiTable);
  }

  // Copy All Emails to Clipboard
  if (btnCopyKpiEmails) {
    btnCopyKpiEmails.addEventListener('click', () => {
      const search = (kpiModalSearch ? kpiModalSearch.value : '').toLowerCase().trim();
      let rows = currentKpiRows;
      if (search) {
        rows = rows.filter((r) =>
          (r.email || '').toLowerCase().includes(search) ||
          (r.name || '').toLowerCase().includes(search) ||
          (r.campaign_name || '').toLowerCase().includes(search)
        );
      }

      const emails = rows.map((r) => r.email).filter(Boolean);
      if (emails.length === 0) {
        showToast('No emails to copy.', 'error');
        return;
      }

      navigator.clipboard.writeText(emails.join(', '));
      showToast(`Copied ${emails.length.toLocaleString()} email addresses to clipboard! 📋`, 'success');
    });
  }

  // Export CSV / Excel for the currently displayed KPI table
  if (btnExportKpiCsv) {
    btnExportKpiCsv.addEventListener('click', () => {
      const search = (kpiModalSearch ? kpiModalSearch.value : '').toLowerCase().trim();
      let rows = currentKpiRows;

      if (currentKpiType === 'dumps') {
        if (search) {
          rows = rows.filter((c) =>
            (c.name || '').toLowerCase().includes(search) ||
            (c.original_filename || '').toLowerCase().includes(search)
          );
        }
        let csvContent = 'data:text/csv;charset=utf-8,Dump_Name,Subject,Created_At,Original_Excel,Total_Recipients,Sent_Count,Opened_Count,Unseen_Count,Open_Rate_Pct\n';
        rows.forEach((c) => {
          const unseen = c.unseen_count !== undefined ? c.unseen_count : (c.total_recipients - c.opened_count);
          csvContent += `"${(c.name||'').replace(/"/g, '""')}","${(c.subject||'').replace(/"/g, '""')}","${c.created_at}","${c.original_filename}",${c.total_recipients},${c.sent_count},${c.opened_count},${unseen},${c.open_rate_pct}%\n`;
        });
        triggerCsvDownload(csvContent, 'Stored_Email_Dumps_Report.csv');
      } else {
        if (search) {
          rows = rows.filter((r) =>
            (r.email || '').toLowerCase().includes(search) ||
            (r.name || '').toLowerCase().includes(search) ||
            (r.campaign_name || '').toLowerCase().includes(search)
          );
        }
        let csvContent = 'data:text/csv;charset=utf-8,Index,Email,Name,Campaign_Dump_Name,Subject,Send_Status,Sent_At,Seen_Status,Opened_At,Open_Count\n';
        rows.forEach((r, i) => {
          const seenStr = r.is_opened === 1 ? 'SEEN' : 'UNSEEN';
          csvContent += `${i + 1},"${r.email}","${(r.name||'').replace(/"/g, '""')}","${(r.campaign_name||'').replace(/"/g, '""')}","${(r.campaign_subject||'').replace(/"/g, '""')}","${r.send_status}","${r.sent_at||''}","${seenStr}","${r.opened_at||''}",${r.open_count||0}\n`;
        });
        triggerCsvDownload(csvContent, `${currentKpiType.toUpperCase()}_Recipients_Report.csv`);
      }
    });
  }

  function triggerCsvDownload(csvData, filename) {
    const encodedUri = encodeURI(csvData);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', filename);
    document.body.appendChild(link);
    link.click();
    link.remove();
    showToast(`Exported ${filename}!`, 'success');
  }

  if (btnCloseKpiModal) {
    btnCloseKpiModal.addEventListener('click', () => kpiDetailModal.classList.remove('active'));
  }
  if (btnCloseKpiModalFooter) {
    btnCloseKpiModalFooter.addEventListener('click', () => kpiDetailModal.classList.remove('active'));
  }

  // ----------------------------------------------------
  // Documentation / Guide & Test Email Modals
  // ----------------------------------------------------
  if (btnHelpModal) btnHelpModal.addEventListener('click', () => helpModal.classList.add('active'));
  if (btnCloseHelpModal) btnCloseHelpModal.addEventListener('click', () => helpModal.classList.remove('active'));

  if (btnCloseTestModal) btnCloseTestModal.addEventListener('click', () => testEmailModal.classList.remove('active'));
  if (btnCancelTestModal) btnCancelTestModal.addEventListener('click', () => testEmailModal.classList.remove('active'));

  // Universal Modal Dismiss Handlers
  document.querySelectorAll('.modal-backdrop').forEach((backdrop) => {
    backdrop.addEventListener('click', (e) => {
      if (e.target === backdrop) backdrop.classList.remove('active');
    });
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      document.querySelectorAll('.modal-backdrop').forEach((m) => m.classList.remove('active'));
    }
  });

  // ----------------------------------------------------
  // Theme Switching Logic
  // ----------------------------------------------------
  if (themeSelector) {
    themeSelector.addEventListener('change', (e) => {
      const theme = e.target.value;
      document.body.setAttribute('data-theme', theme);
      const isLight = theme.includes('light');
      document.body.classList.toggle('light-theme', isLight);
      document.body.classList.toggle('dark-theme', !isLight);
      localStorage.setItem('automailer_theme', theme);
    });
  }

  if (btnToggleMode) {
    btnToggleMode.addEventListener('click', () => {
      const isLight = document.body.classList.contains('light-theme');
      const newTheme = isLight ? 'cosmic-aurora' : 'clean-light';
      if (themeSelector) {
        themeSelector.value = newTheme;
        themeSelector.dispatchEvent(new Event('change'));
      }
    });
  }

  // Load Saved Theme
  const savedTheme = localStorage.getItem('automailer_theme');
  if (savedTheme && themeSelector) {
    themeSelector.value = savedTheme;
    themeSelector.dispatchEvent(new Event('change'));
  }

  // Initialize SSE and initial analytics count
  initSSE();
  loadAnalyticsData();
});


