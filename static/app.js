/**
 * Bigrock Bulk Mailer Web Application Client
 * Handles real-time status polling, SSE live log streaming, file uploads, and UI interactions.
 */

// Application State
let appState = {
  defaults: null,
  status: 'idle',
  pollInterval: null,
  eventSource: null,
  cachedLogs: [],
  cachedSentRows: [],
  sampleRecipients: [],
  templateHtmlPreview: '',
  subjectsList: [],
};

// DOM Elements
const el = {
  connectionBadge: document.getElementById('connectionBadge'),
  workerStatusBadge: document.getElementById('workerStatusBadge'),
  fromEmail: document.getElementById('fromEmail'),
  emailPassword: document.getElementById('emailPassword'),
  togglePasswordBtn: document.getElementById('togglePasswordBtn'),
  eyeIconOpen: document.getElementById('eyeIconOpen'),
  eyeIconClosed: document.getElementById('eyeIconClosed'),
  delaySeconds: document.getElementById('delaySeconds'),
  showChromeWindow: document.getElementById('showChromeWindow'),
  webmailUrl: document.getElementById('webmailUrl'),

  recipientsFileName: document.getElementById('recipientsFileName'),
  recipientsMeta: document.getElementById('recipientsMeta'),
  recipientsFileInput: document.getElementById('recipientsFileInput'),
  previewRecipientsBtn: document.getElementById('previewRecipientsBtn'),

  templateFileName: document.getElementById('templateFileName'),
  templateMeta: document.getElementById('templateMeta'),
  templateFileInput: document.getElementById('templateFileInput'),
  previewTemplateBtn: document.getElementById('previewTemplateBtn'),

  subjectFileName: document.getElementById('subjectFileName'),
  subjectMeta: document.getElementById('subjectMeta'),
  subjectFileInput: document.getElementById('subjectFileInput'),
  previewSubjectBtn: document.getElementById('previewSubjectBtn'),

  startBtn: document.getElementById('startBtn'),
  pauseBtn: document.getElementById('pauseBtn'),
  stopBtn: document.getElementById('stopBtn'),
  currentStatusText: document.getElementById('currentStatusText'),

  totalCount: document.getElementById('totalCount'),
  sentCount: document.getElementById('sentCount'),
  failedCount: document.getElementById('failedCount'),
  remainingCount: document.getElementById('remainingCount'),
  elapsedTime: document.getElementById('elapsedTime'),
  progressBar: document.getElementById('progressBar'),
  progressPercent: document.getElementById('progressPercent'),

  tableSearchInput: document.getElementById('tableSearchInput'),
  statusFilterSelect: document.getElementById('statusFilterSelect'),
  dateFilterSelect: document.getElementById('dateFilterSelect'),
  datePickerInput: document.getElementById('datePickerInput'),
  exportLogBtn: document.getElementById('exportLogBtn'),
  sentLogTableBody: document.getElementById('sentLogTableBody'),

  consoleTerminal: document.getElementById('consoleTerminal'),
  clearConsoleBtn: document.getElementById('clearConsoleBtn'),

  previewModal: document.getElementById('previewModal'),
  modalTitle: document.getElementById('modalTitle'),
  modalBody: document.getElementById('modalBody'),
  closeModalBtn: document.getElementById('closeModalBtn'),
};

// ---------------------------------------------------------------------------
// INITIALIZATION
// ---------------------------------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {
  initDefaults();
  initPasswordToggle();
  initFileUploads();
  initControlButtons();
  initLogStream();
  initStatusPolling();
  initTableFilters();
  initModal();
});

// Load defaults from backend
async function initDefaults() {
  try {
    const res = await fetch('/api/defaults');
    if (!res.ok) throw new Error('Failed to load defaults');
    const data = await res.json();
    appState.defaults = data;

    // Prefill form
    el.fromEmail.value = data.from_email || '';
    el.webmailUrl.value = data.webmail_url || 'https://us3.webmail.mailhostbox.com';
    el.delaySeconds.value = data.delay_seconds || 30;

    // Recipients info
    if (data.recipients_count > 0) {
      el.recipientsFileName.textContent = data.recipients_filename || 'recipients.csv';
      el.recipientsMeta.textContent = `${data.recipients_count} valid recipient(s) loaded`;
      appState.sampleRecipients = data.sample_recipients || [];
      el.totalCount.textContent = data.recipients_count;
      el.remainingCount.textContent = data.recipients_count;
    } else {
      el.recipientsFileName.textContent = 'No recipients file';
      el.recipientsMeta.textContent = 'Click "Upload New" to select .csv or .xlsx';
    }

    // Template info
    if (data.template_filename) {
      el.templateFileName.textContent = data.template_filename;
      el.templateMeta.textContent = 'Formatting preserved (HTML & text)';
      appState.templateHtmlPreview = data.template_html_preview || '';
    } else {
      el.templateFileName.textContent = 'No template file';
      el.templateMeta.textContent = 'Click "Upload New" to select .docx or .txt';
    }

    // Subjects info
    if (data.subjects_count > 0) {
      el.subjectFileName.textContent = data.subjects_filename || 'subject.docx';
      el.subjectMeta.textContent = `${data.subjects_count} custom subject line(s)`;
      appState.subjectsList = data.subjects_list || [];
    } else {
      el.subjectFileName.textContent = 'Automatic Subjects';
      el.subjectMeta.textContent = 'Will use built-in alternating subjects';
    }

    // Initial sent log fetch
    fetchSentLog();
  } catch (err) {
    console.error('Error loading defaults:', err);
    appendLog(`[Warning] Could not load defaults: ${err.message}`, 'warning');
  }
}

// ---------------------------------------------------------------------------
// PASSWORD TOGGLE
// ---------------------------------------------------------------------------
function initPasswordToggle() {
  el.togglePasswordBtn.addEventListener('click', () => {
    const isPassword = el.emailPassword.type === 'password';
    el.emailPassword.type = isPassword ? 'text' : 'password';
    el.eyeIconOpen.classList.toggle('hidden', isPassword);
    el.eyeIconClosed.classList.toggle('hidden', !isPassword);
  });
}

// ---------------------------------------------------------------------------
// FILE UPLOADS
// ---------------------------------------------------------------------------
function initFileUploads() {
  // Recipients Upload
  el.recipientsFileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const formData = new FormData();
    formData.append('file', file);
    el.recipientsMeta.textContent = 'Uploading & parsing...';

    try {
      const res = await fetch('/api/upload/recipients', { method: 'POST', body: formData });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || 'Upload failed');

      el.recipientsFileName.textContent = data.filename;
      el.recipientsMeta.textContent = `${data.count} valid recipient(s) loaded`;
      appState.sampleRecipients = data.sample || [];
      el.totalCount.textContent = data.count;
      el.remainingCount.textContent = data.count;
      appendLog(`[Success] Loaded ${data.count} recipients from ${data.filename}`, 'success');
    } catch (err) {
      alert(`Error uploading recipients: ${err.message}`);
      el.recipientsMeta.textContent = 'Upload failed. Try again.';
      appendLog(`[Error] Recipients upload failed: ${err.message}`, 'error');
    }
  });

  // Template Upload
  el.templateFileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const formData = new FormData();
    formData.append('file', file);
    el.templateMeta.textContent = 'Uploading & parsing...';

    try {
      const res = await fetch('/api/upload/template', { method: 'POST', body: formData });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || 'Upload failed');

      el.templateFileName.textContent = data.filename;
      el.templateMeta.textContent = 'Template parsed successfully';
      appState.templateHtmlPreview = data.html_preview || '';
      appendLog(`[Success] Loaded template from ${data.filename}`, 'success');
    } catch (err) {
      alert(`Error uploading template: ${err.message}`);
      el.templateMeta.textContent = 'Upload failed. Try again.';
      appendLog(`[Error] Template upload failed: ${err.message}`, 'error');
    }
  });

  // Subjects Upload
  el.subjectFileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const formData = new FormData();
    formData.append('file', file);
    el.subjectMeta.textContent = 'Uploading & parsing...';

    try {
      const res = await fetch('/api/upload/subjects', { method: 'POST', body: formData });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || 'Upload failed');

      el.subjectFileName.textContent = data.filename;
      el.subjectMeta.textContent = `${data.count} custom subject line(s)`;
      appState.subjectsList = data.subjects || [];
      appendLog(`[Success] Loaded ${data.count} subjects from ${data.filename}`, 'success');
    } catch (err) {
      alert(`Error uploading subjects: ${err.message}`);
      el.subjectMeta.textContent = 'Upload failed. Try again.';
      appendLog(`[Error] Subjects upload failed: ${err.message}`, 'error');
    }
  });
}

// ---------------------------------------------------------------------------
// PREVIEW MODALS
// ---------------------------------------------------------------------------
function initModal() {
  el.closeModalBtn.addEventListener('click', () => el.previewModal.classList.add('hidden'));
  el.previewModal.addEventListener('click', (e) => {
    if (e.target === el.previewModal) el.previewModal.classList.add('hidden');
  });

  // Preview Recipients
  el.previewRecipientsBtn.addEventListener('click', () => {
    if (!appState.sampleRecipients.length) {
      alert('No recipients loaded yet to preview.');
      return;
    }
    el.modalTitle.textContent = `Recipients Preview (First ${appState.sampleRecipients.length} Rows)`;
    let html = `
      <table class="preview-table">
        <thead>
          <tr>
            <th>#</th>
            <th>Name</th>
            <th>Email</th>
            <th>Paper Title</th>
            <th>Topic</th>
          </tr>
        </thead>
        <tbody>
    `;
    appState.sampleRecipients.forEach((r, idx) => {
      html += `
        <tr>
          <td>${idx + 1}</td>
          <td><strong>${escapeHtml(r.name || '—')}</strong></td>
          <td>${escapeHtml(r.email)}</td>
          <td>${escapeHtml(r.paper_title || '—')}</td>
          <td>${escapeHtml(r.topic || '—')}</td>
        </tr>
      `;
    });
    html += '</tbody></table>';
    el.modalBody.innerHTML = html;
    el.previewModal.classList.remove('hidden');
  });

  // Preview Template
  el.previewTemplateBtn.addEventListener('click', () => {
    if (!appState.templateHtmlPreview) {
      alert('No template loaded yet to preview.');
      return;
    }
    el.modalTitle.textContent = 'Email Template Preview (Rendered HTML)';
    el.modalBody.innerHTML = `
      <div style="margin-bottom: 0.8rem; font-size: 0.8rem; color: var(--text-muted);">
        Placeholders <code>{name}</code> and <code>{paper_title}</code> will be automatically replaced with each recipient's actual details.
      </div>
      <div class="template-rendered-box">
        ${appState.templateHtmlPreview}
      </div>
    `;
    el.previewModal.classList.remove('hidden');
  });

  // View Subjects
  el.previewSubjectBtn.addEventListener('click', () => {
    el.modalTitle.textContent = 'Subject Lines';
    if (!appState.subjectsList.length) {
      el.modalBody.innerHTML = `
        <p style="color: var(--text-muted);">
          No custom subject file loaded. The mailer will automatically alternate between 5 built-in professional subject patterns customized with each author's name and topic.
        </p>
      `;
    } else {
      let html = '<ul style="padding-left: 1.2rem; display: flex; flex-direction: column; gap: 0.5rem;">';
      appState.subjectsList.forEach(s => {
        html += `<li><strong>${escapeHtml(s)}</strong></li>`;
      });
      html += '</ul>';
      el.modalBody.innerHTML = html;
    }
    el.previewModal.classList.remove('hidden');
  });
}

// ---------------------------------------------------------------------------
// CONTROL BUTTONS (START, PAUSE, STOP)
// ---------------------------------------------------------------------------
function initControlButtons() {
  // Start Campaign
  el.startBtn.addEventListener('click', async () => {
    const from_email = el.fromEmail.value.trim();
    const password = el.emailPassword.value;
    const webmail_url = el.webmailUrl.value.trim();
    const delay_seconds = parseFloat(el.delaySeconds.value) || 30;
    const headless = !el.showChromeWindow.checked;

    if (!from_email || !from_email.includes('@')) {
      alert('Please enter a valid sender email address.');
      el.fromEmail.focus();
      return;
    }
    if (!password) {
      alert('Please enter your email account password to log into Webmail.');
      el.emailPassword.focus();
      return;
    }

    el.startBtn.disabled = true;
    el.startBtn.innerHTML = '<span class="dot"></span> Starting...';

    try {
      const res = await fetch('/api/start', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          from_email,
          password,
          webmail_url,
          delay_seconds,
          headless,
        }),
      });

      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || 'Failed to start campaign');

      appendLog('[System] Campaign initiated successfully!', 'success');
      pollStatus();
    } catch (err) {
      alert(`Error starting campaign: ${err.message}`);
      appendLog(`[Error] Failed to start campaign: ${err.message}`, 'error');
      el.startBtn.disabled = false;
      el.startBtn.innerHTML = `
        <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
          <polygon points="5 3 19 12 5 21 5 3"></polygon>
        </svg>
        Start Campaign
      `;
    }
  });

  // Pause Campaign
  el.pauseBtn.addEventListener('click', async () => {
    try {
      if (appState.status === 'paused') {
        await fetch('/api/resume', { method: 'POST' });
        appendLog('[System] Resuming campaign...', 'info');
      } else {
        await fetch('/api/pause', { method: 'POST' });
        appendLog('[System] Pausing campaign after current step...', 'warning');
      }
      pollStatus();
    } catch (err) {
      console.error('Pause/Resume error:', err);
    }
  });

  // Stop Campaign
  el.stopBtn.addEventListener('click', async () => {
    if (!confirm('Are you sure you want to stop the email sending process?')) return;
    try {
      await fetch('/api/stop', { method: 'POST' });
      appendLog('[System] Stop signal sent. Closing browser...', 'warning');
      pollStatus();
    } catch (err) {
      console.error('Stop error:', err);
    }
  });

  // Export Log CSV
  el.exportLogBtn.addEventListener('click', () => {
    window.location.href = '/api/export-log';
  });

  // Clear Console
  el.clearConsoleBtn.addEventListener('click', () => {
    el.consoleTerminal.innerHTML = '<div class="log-line log-info">[System] Console cleared.</div>';
  });
}

// ---------------------------------------------------------------------------
// STATUS POLLING & METRICS
// ---------------------------------------------------------------------------
function initStatusPolling() {
  pollStatus();
  appState.pollInterval = setInterval(pollStatus, 1500);
}

async function pollStatus() {
  try {
    const res = await fetch('/api/status');
    if (!res.ok) throw new Error('Status poll failed');
    const data = await res.json();
    updateUIWithStatus(data);
    el.connectionBadge.className = 'status-pill pill-online';
    el.connectionBadge.innerHTML = '<span class="dot"></span> Online';
  } catch (err) {
    el.connectionBadge.className = 'status-pill pill-idle';
    el.connectionBadge.innerHTML = '<span class="dot"></span> Offline';
  }
}

function updateUIWithStatus(data) {
  const status = data.status || 'idle';
  appState.status = status;

  // Status Badge
  el.workerStatusBadge.textContent = status.toUpperCase();
  el.workerStatusBadge.className = `status-pill pill-${status}`;

  // Metrics
  const total = data.total || parseInt(el.totalCount.textContent) || 0;
  const sent = data.sent || 0;
  const failed = data.failed || 0;
  const current = data.current || 0;
  const remaining = Math.max(0, total - sent - failed);

  el.totalCount.textContent = total;
  el.sentCount.textContent = sent;
  el.failedCount.textContent = failed;
  el.remainingCount.textContent = remaining;

  // Format Elapsed Time
  const seconds = data.elapsed_seconds || 0;
  const mins = Math.floor(seconds / 60).toString().padStart(2, '0');
  const secs = (seconds % 60).toString().padStart(2, '0');
  el.elapsedTime.textContent = `${mins}:${secs}`;

  // Progress Bar
  const pct = total > 0 ? Math.round(((sent + failed) / total) * 100) : 0;
  el.progressBar.style.width = `${pct}%`;
  el.progressPercent.textContent = `${pct}%`;

  // Buttons state
  if (status === 'running') {
    el.startBtn.disabled = true;
    el.startBtn.innerHTML = '<span class="dot"></span> Running...';
    el.pauseBtn.disabled = false;
    el.pauseBtn.textContent = 'Pause';
    el.stopBtn.disabled = false;
    el.currentStatusText.textContent = data.current_recipient ? `Sending to ${data.current_recipient} (${current}/${total})` : 'Processing...';
  } else if (status === 'paused') {
    el.startBtn.disabled = true;
    el.pauseBtn.disabled = false;
    el.pauseBtn.textContent = 'Resume';
    el.stopBtn.disabled = false;
    el.currentStatusText.textContent = 'Campaign is paused.';
  } else {
    // idle / stopped / completed / error
    el.startBtn.disabled = false;
    el.startBtn.innerHTML = `
      <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
        <polygon points="5 3 19 12 5 21 5 3"></polygon>
      </svg>
      Start Campaign
    `;
    el.pauseBtn.disabled = true;
    el.pauseBtn.textContent = 'Pause';
    el.stopBtn.disabled = true;

    if (status === 'completed') {
      el.currentStatusText.textContent = `Completed! ${sent} sent, ${failed} failed.`;
    } else if (status === 'stopped') {
      el.currentStatusText.textContent = 'Campaign stopped.';
    } else if (status === 'error') {
      el.currentStatusText.textContent = `Error: ${data.error_message || 'Unknown error'}`;
    } else {
      el.currentStatusText.textContent = 'Ready to start.';
    }
  }

  // Refresh sent log periodically while running
  if (status === 'running' || status === 'completed') {
    fetchSentLog();
  }
}

// ---------------------------------------------------------------------------
// SENT LOG TABLE & FILTERING
// ---------------------------------------------------------------------------
function initTableFilters() {
  let debounceTimeout;
  el.tableSearchInput.addEventListener('input', () => {
    clearTimeout(debounceTimeout);
    debounceTimeout = setTimeout(fetchSentLog, 300);
  });

  el.statusFilterSelect.addEventListener('change', fetchSentLog);

  el.dateFilterSelect.addEventListener('change', () => {
    if (el.dateFilterSelect.value) {
      el.datePickerInput.value = '';
    }
    fetchSentLog();
  });

  el.datePickerInput.addEventListener('change', () => {
    if (el.datePickerInput.value) {
      el.dateFilterSelect.value = '';
    }
    fetchSentLog();
  });
}

let loadedDatesKey = '';

async function fetchSentLog() {
  const query = el.tableSearchInput.value.trim();
  const status = el.statusFilterSelect.value;
  const selectedDate = el.datePickerInput.value || el.dateFilterSelect.value;

  try {
    const url = `/api/sent-log?limit=500&query=${encodeURIComponent(query)}&status=${encodeURIComponent(status)}&date=${encodeURIComponent(selectedDate)}`;
    const res = await fetch(url);
    if (!res.ok) return;
    const data = await res.json();
    renderSentTable(data.rows || []);

    // Populate date dropdown if new dates available
    if (data.available_dates && data.available_dates.length) {
      const datesKey = data.available_dates.join(',');
      if (datesKey !== loadedDatesKey) {
        loadedDatesKey = datesKey;
        const currentVal = el.dateFilterSelect.value;
        let optHtml = '<option value="">All Dates</option>';
        data.available_dates.forEach(d => {
          const selected = d === currentVal ? 'selected' : '';
          optHtml += `<option value="${escapeHtml(d)}" ${selected}>${escapeHtml(d)}</option>`;
        });
        el.dateFilterSelect.innerHTML = optHtml;
      }
    }
  } catch (err) {
    console.error('Error fetching sent log:', err);
  }
}

function renderSentTable(rows) {
  if (!rows || rows.length === 0) {
    el.sentLogTableBody.innerHTML = `
      <tr>
        <td colspan="4" class="text-center text-muted">No matching records found.</td>
      </tr>
    `;
    return;
  }

  let html = '';
  rows.forEach(r => {
    const isSent = (r.status || '').toLowerCase() === 'sent';
    const statusClass = isSent ? 'badge-sent' : 'badge-failed';
    const statusLabel = isSent ? 'SENT' : 'FAILED';
    const titleAttr = r.error ? `title="${escapeHtml(r.error)}"` : '';

    html += `
      <tr>
        <td style="white-space: nowrap; font-size: 0.8rem; color: var(--text-muted);">${escapeHtml(r.timestamp || '—')}</td>
        <td>
          <div class="recipient-cell">
            <span class="recipient-name">${escapeHtml(r.name || 'Researcher')}</span>
            <span class="recipient-email">${escapeHtml(r.email || '—')}</span>
          </div>
        </td>
        <td style="max-width: 320px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">
          ${escapeHtml(r.subject || '—')}
        </td>
        <td>
          <span class="badge-status ${statusClass}" ${titleAttr}>
            ${statusLabel}
          </span>
        </td>
      </tr>
    `;
  });

  el.sentLogTableBody.innerHTML = html;
}

// ---------------------------------------------------------------------------
// LIVE SSE LOG STREAM
// ---------------------------------------------------------------------------
function initLogStream() {
  if (appState.eventSource) {
    appState.eventSource.close();
  }

  appState.eventSource = new EventSource('/api/logs/stream');

  appState.eventSource.onmessage = (event) => {
    try {
      const data = JSON.parse(event.data);
      appendLog(data.message, data.level);
    } catch (e) {
      appendLog(event.data, 'info');
    }
  };

  appState.eventSource.onerror = () => {
    // Reconnect silently handled by EventSource
  };
}

function appendLog(message, level = 'info') {
  const line = document.createElement('div');
  line.className = `log-line log-${level}`;
  line.textContent = message;
  el.consoleTerminal.appendChild(line);

  // Auto-scroll
  el.consoleTerminal.scrollTop = el.consoleTerminal.scrollHeight;
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
