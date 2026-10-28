let rows = [];
let currentIndex = 0;
let previewIndex = 0;
let generated = [];
let csvHeaders = [];
let lastTemplateFieldId = 'body';

const $ = (id) => document.getElementById(id);
const STORAGE_KEY = 'outlookMailMergeState';
const DEFAULT_DELAY_MS = 1000;
const TEMPLATE_FIELD_IDS = ['subject', 'cc', 'bcc', 'body'];
const TEMPLATE_PRESETS = {
  courseReminder: {
    subject: 'Reminder for {{Course}}',
    body: 'Hi {{Name}},\n\nThis is a quick reminder about {{Course}}.'
  },
  blank: {
    subject: '',
    body: ''
  }
};
const COMPOSE_MODE_HINTS = {
  outlook: 'Outlook Web: Drafts open in Outlook Web compose tabs (mailto is used if CC/BCC is present).',
  gmail: 'Gmail: Drafts open in Gmail compose tabs with full To, CC, BCC, subject, and body support.'
};

function actionLabels() {
  const isAutoSend = $('autoSend')?.checked ?? false;
  return {
    current: isAutoSend ? 'Open & Send Preview Email' : 'Open Preview Draft',
    selected: isAutoSend ? 'Open & Send Selected Emails' : 'Open Selected Drafts',
    item: isAutoSend ? 'email' : 'draft',
    items: isAutoSend ? 'emails' : 'drafts',
    openedVerb: isAutoSend ? 'Queued' : 'Opened',
    openVerb: isAutoSend ? 'open and send' : 'open',
    openingVerb: isAutoSend ? 'opening and sending' : 'opening'
  };
}

function syncDraftActionLabels() {
  const labels = actionLabels();
  if ($('openNextBtn')) $('openNextBtn').textContent = labels.current;
  if ($('openRangeBtn')) $('openRangeBtn').textContent = labels.selected;
}

function hasOwn(object, key) {
  return Object.prototype.hasOwnProperty.call(object, key);
}

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function parseCSV(text) {
  const records = [];
  let record = [];
  let cur = '';
  let inQuotes = false;
  const source = text.replace(/^\uFEFF/, '');

  for (let i = 0; i < source.length; i++) {
    const ch = source[i];
    const next = source[i + 1];

    if (ch === '"' && inQuotes && next === '"') {
      cur += '"';
      i++;
    } else if (ch === '"') {
      inQuotes = !inQuotes;
    } else if (ch === ',' && !inQuotes) {
      record.push(cur.trim());
      cur = '';
    } else if ((ch === '\n' || ch === '\r') && !inQuotes) {
      if (ch === '\r' && next === '\n') i++;
      record.push(cur.trim());
      if (record.some(value => value !== '')) records.push(record);
      record = [];
      cur = '';
    } else {
      cur += ch;
    }
  }

  if (inQuotes) throw new Error('CSV has an unclosed quoted field.');
  record.push(cur.trim());
  if (record.some(value => value !== '')) records.push(record);

  const lines = records;
  if (lines.length < 2) {
    const hasCommaSeparatedValues = source.split(',').length > 2;
    if (hasCommaSeparatedValues) {
      throw new Error('CSV needs a real line break between the header row and the first data row. Text wrapping in the box does not count; press Enter after the last header, then put John,john@example.com,... on the next line.');
    }
    throw new Error('CSV needs a header row and at least one data row.');
  }

  const headers = lines[0].map(h => h.trim());
  csvHeaders = headers;

  // Smart email column detection (supports Email, E-mail, Mail, 邮箱, Recipient, To, etc.)
  const emailAliases = ['email', 'e-mail', 'mail', '邮箱', 'recipient', 'to', 'student email', 'student_email'];
  let emailKey = headers.find(h => emailAliases.includes(h.toLowerCase()));
  if (!emailKey) {
    emailKey = headers.find(h => /email|邮箱|mail/i.test(h));
  }

  return lines.slice(1).map(line => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = line[i] || '');
    if (emailKey && !obj.Email && !obj.email) {
      obj.Email = obj[emailKey];
    }
    return obj;
  }).filter(r => String(r.Email || r.email || '').trim() !== '');
}

function applyTemplate(template, row) {
  return String(template || '').replace(/{{\s*([^}]+?)\s*}}/g, (_, key) => row[key] ?? row[key.trim()] ?? '');
}

function templateVariables(template) {
  const variables = new Set();
  String(template || '').replace(/{{\s*([^}]+?)\s*}}/g, (_, key) => {
    variables.add(key.trim());
    return '';
  });
  return [...variables];
}

function findMissingTemplateVariables(headers, templates) {
  const headerSet = new Set(headers);
  const missing = new Set();
  templates.flatMap(templateVariables).forEach(variable => {
    if (!headerSet.has(variable)) missing.add(variable);
  });
  return [...missing];
}

function splitRecipients(value) {
  return String(value || '')
    .split(/[;,]/)
    .map(recipient => recipient.trim())
    .filter(Boolean);
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function validateMessages(messages) {
  const issues = [];
  const toCounts = new Map();

  messages.forEach((msg, index) => {
    const rowLabel = `Draft ${index + 1}`;
    const to = String(msg.email || '').trim();
    if (!isValidEmail(to)) issues.push(`${rowLabel}: Invalid email address: ${to || '(blank)'}`);
    if (to) toCounts.set(to.toLowerCase(), (toCounts.get(to.toLowerCase()) || 0) + 1);

    const copyRecipients = [...splitRecipients(msg.cc), ...splitRecipients(msg.bcc)];
    copyRecipients.forEach(recipient => {
      if (!isValidEmail(recipient)) issues.push(`${rowLabel}: Invalid email address: ${recipient}`);
      if (recipient.toLowerCase() === to.toLowerCase()) {
        issues.push(`${rowLabel}: ${recipient} also appears in CC/BCC.`);
      }
    });
  });

  toCounts.forEach((count, email) => {
    if (count > 1) issues.push(`Duplicate To recipient: ${email}`);
  });

  return issues;
}

function getCsvText() {
  return $('csvText') ? $('csvText').value.trim() : '';
}

async function saveState() {
  const state = {
    csvText: $('csvText')?.value ?? '',
    composeMode: $('composeMode')?.value ?? 'outlook',
    autoSend: Boolean($('autoSend')?.checked),
    templatePreset: $('templatePreset')?.value ?? '',
    batchDelay: $('batchDelay')?.value ?? '1000',
    subject: $('subject')?.value ?? '',
    cc: $('cc')?.value ?? '',
    bcc: $('bcc')?.value ?? '',
    body: $('body')?.value ?? '',
    rows,
    csvHeaders,
    generated,
    currentIndex,
    previewIndex,
    hasStarted: generated.length > 0,
    updatedAt: Date.now()
  };
  await chrome.storage.local.set({ [STORAGE_KEY]: state });
}

function updateCcBccVisibility(show) {
  const container = $('ccBccContainer');
  const trigger = $('toggleCcBccBtn');
  if (!container || !trigger) return;
  if (show) {
    container.classList.remove('hidden');
    trigger.textContent = '− Hide CC / BCC';
  } else {
    container.classList.add('hidden');
    trigger.textContent = '+ Add CC / BCC';
  }
}

function updateCsvSummary() {
  const summaryEl = $('csvSummary');
  if (!summaryEl) return;
  if (rows.length > 0) {
    summaryEl.textContent = `✓ ${rows.length} rows loaded`;
    summaryEl.classList.remove('hidden');
  } else {
    summaryEl.classList.add('hidden');
  }
}

async function restoreState() {
  const result = await chrome.storage.local.get(STORAGE_KEY);
  const state = result[STORAGE_KEY];
  if (!state) return;

  if ($('csvText')) $('csvText').value = hasOwn(state, 'csvText') ? state.csvText : $('csvText').value;
  if ($('composeMode')) $('composeMode').value = state.composeMode || 'outlook';
  if ($('autoSend')) $('autoSend').checked = Boolean(state.autoSend);
  if ($('templatePreset')) $('templatePreset').value = state.templatePreset || '';
  if ($('batchDelay')) $('batchDelay').value = state.batchDelay || '1000';
  if ($('subject')) $('subject').value = hasOwn(state, 'subject') ? state.subject : $('subject').value;
  if ($('cc')) $('cc').value = hasOwn(state, 'cc') ? state.cc : '';
  if ($('bcc')) $('bcc').value = hasOwn(state, 'bcc') ? state.bcc : '';
  if ($('body')) $('body').value = hasOwn(state, 'body') ? state.body : $('body').value;

  rows = Array.isArray(state.rows) ? state.rows : [];
  csvHeaders = Array.isArray(state.csvHeaders) ? state.csvHeaders : getHeadersFromRows(rows);
  generated = Array.isArray(state.generated) ? state.generated : [];
  currentIndex = Number.isInteger(state.currentIndex) ? state.currentIndex : 0;
  previewIndex = Number.isInteger(state.previewIndex) ? state.previewIndex : currentIndex;

  if ($('cc')?.value || $('bcc')?.value) {
    updateCcBccVisibility(true);
  }
  updateCsvSummary();
  syncDraftActionLabels();
  renderVariableButtons();
  renderComposeModeHints();

  if (generated.length) {
    previewIndex = clampPreviewIndex(previewIndex);
    if ($('sendControls')) $('sendControls').classList.remove('hidden');
    if ($('previewControls')) $('previewControls').classList.remove('hidden');
    syncRangeInputs();
    const labels = actionLabels();
    renderStatus(`Restored saved progress at ${currentIndex}/${generated.length}. Use preview controls, then click “${labels.current}”.`, 'ok');
    renderPreview();
  }
}

async function loadRows() {
  const file = $('csvFile')?.files?.[0];
  let text = getCsvText();
  if (file) {
    text = await file.text();
    if ($('csvText')) $('csvText').value = text;
  }
  rows = parseCSV(text);
  if (!rows.length) throw new Error('No usable rows found. Make sure your CSV contains an Email column.');
  renderVariableButtons();
  updateCsvSummary();
  await saveState();
  return rows;
}

function makeMessages() {
  const subjectTemplate = $('subject')?.value ?? '';
  const ccTemplate = $('cc')?.value ?? '';
  const bccTemplate = $('bcc')?.value ?? '';
  const bodyTemplate = $('body')?.value ?? '';
  generated = rows.map(row => {
    const email = row.Email || row.email;
    const subject = applyTemplate(subjectTemplate, row);
    const cc = applyTemplate(ccTemplate, row);
    const bcc = applyTemplate(bccTemplate, row);
    const body = applyTemplate(bodyTemplate, row);
    return { email, cc, bcc, subject, body, row };
  });
}

function currentTemplates() {
  return TEMPLATE_FIELD_IDS.map(id => $(id)?.value ?? '');
}

function encodeForOutlook(text) {
  return encodeURIComponent(String(text ?? ''));
}

function outlookComposeUrl(msg) {
  const params = [
    ['to', msg.email],
    ['cc', msg.cc],
    ['bcc', msg.bcc],
    ['subject', msg.subject],
    ['body', msg.body]
  ].filter(([, value]) => String(value ?? '').trim() !== '');

  if ($('autoSend')?.checked) params.push(['mailMergeAutoSend', 'true']);

  return 'https://outlook.office.com/mail/deeplink/compose'
    + '?' + params.map(([key, value]) => key + '=' + encodeForOutlook(value)).join('&');
}

function gmailComposeUrl(msg) {
  const params = [
    ['view', 'cm'],
    ['fs', '1'],
    ['to', msg.email],
    ['cc', msg.cc],
    ['bcc', msg.bcc],
    ['su', msg.subject],
    ['body', msg.body]
  ].filter(([, value]) => String(value ?? '').trim() !== '');

  if ($('autoSend')?.checked) params.push(['mailMergeAutoSend', 'true']);

  return 'https://mail.google.com/mail/?'
    + params.map(([key, value]) => key + '=' + encodeForOutlook(value)).join('&');
}

function mailtoUrl(msg) {
  const params = [
    ['cc', msg.cc],
    ['bcc', msg.bcc],
    ['subject', msg.subject],
    ['body', msg.body]
  ].filter(([, value]) => String(value ?? '').trim() !== '');

  return 'mailto:' + encodeForOutlook(msg.email)
    + '?' + params.map(([key, value]) => key + '=' + encodeForOutlook(value)).join('&');
}

function hasCopyRecipients(msg) {
  return String(msg.cc ?? '').trim() !== '' || String(msg.bcc ?? '').trim() !== '';
}

function composeUrl(msg) {
  if ($('composeMode')?.value === 'gmail') return gmailComposeUrl(msg);

  if ($('autoSend')?.checked && hasCopyRecipients(msg)) {
    throw new Error('Auto-Send with Outlook Mode cannot be used with CC or BCC. Use Gmail Mode or turn off Auto-Send.');
  }
  if ($('autoSend')?.checked) return outlookComposeUrl(msg);
  return hasCopyRecipients(msg) ? mailtoUrl(msg) : outlookComposeUrl(msg);
}

function renderStatus(text, cls = '') {
  if (!$('status')) return;
  $('status').className = cls;
  $('status').textContent = text;
}

function renderComposeModeHints() {
  const mode = $('composeMode')?.value ?? 'outlook';
  if ($('composeModeHint')) $('composeModeHint').textContent = COMPOSE_MODE_HINTS[mode] || '';
  if ($('copyRecipientHint')) {
    $('copyRecipientHint').innerHTML = mode === 'outlook'
      ? 'In Outlook Mode, CC/BCC drafts use your browser&apos;s <code>mailto:</code> handler.'
      : 'In Gmail Mode, CC/BCC are included in the Gmail compose URL.';
  }

  // Sync visual segmented buttons if present in DOM
  if (typeof document !== 'undefined' && document.querySelectorAll) {
    const buttons = document.querySelectorAll('.segmented-btn');
    buttons.forEach(btn => {
      const btnMode = btn.getAttribute('data-mode');
      if (btnMode === mode) btn.classList.add('active');
      else btn.classList.remove('active');
    });
  }
}

function syncRangeInputs() {
  const total = generated.length;
  const next = Math.min(currentIndex + 1, total || 1);
  if ($('rangeStart')) {
    $('rangeStart').max = total || 1;
    $('rangeStart').value = next;
  }
  if ($('rangeEnd')) {
    $('rangeEnd').max = total || 1;
    $('rangeEnd').value = total || 1;
  }
}

function syncPreviewControls() {
  const total = generated.length;
  const hasMessages = total > 0;
  if ($('previewControls')) $('previewControls').classList.toggle('hidden', !hasMessages);
  if ($('previewCounter')) $('previewCounter').textContent = hasMessages ? `${previewIndex + 1} / ${total}` : '';
  if ($('previewPrevBtn')) $('previewPrevBtn').disabled = !hasMessages || previewIndex <= 0;
  if ($('previewNextBtn')) $('previewNextBtn').disabled = !hasMessages || previewIndex >= total - 1;
}

function renderPreflightIssues(issues) {
  if (!issues.length) return '';
  return '\n\nPreflight warnings:\n- ' + issues.slice(0, 8).join('\n- ')
    + (issues.length > 8 ? `\n- ...and ${issues.length - 8} more.` : '');
}

function clampPreviewIndex(index) {
  if (!generated.length) return 0;
  return Math.min(Math.max(index, 0), generated.length - 1);
}

function getHeadersFromRows(rowList) {
  const seen = new Set();
  rowList.forEach(row => {
    Object.keys(row).forEach(key => {
      if (key && !seen.has(key)) seen.add(key);
    });
  });
  return [...seen];
}

function insertAtCursor(field, text) {
  if (!field) return;
  const start = field.selectionStart ?? field.value.length;
  const end = field.selectionEnd ?? field.value.length;
  field.value = field.value.slice(0, start) + text + field.value.slice(end);
  field.focus();
  field.setSelectionRange(start + text.length, start + text.length);
  field.dispatchEvent(new Event('input', { bubbles: true }));
}

function renderVariableButtons() {
  const container = $('variableButtons');
  if (!container) return;
  container.textContent = '';

  csvHeaders.filter(Boolean).forEach(header => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'variable-button';
    button.textContent = `+ {{${header}}}`;
    button.title = `Insert {{${header}}}`;
    button.addEventListener('click', () => {
      insertAtCursor($(lastTemplateFieldId), `{{${header}}}`);
    });
    container.appendChild(button);
  });
}

function renderPreview(title) {
  if (!generated.length || !$('preview')) return;

  previewIndex = clampPreviewIndex(previewIndex);
  const msg = generated[previewIndex];
  if (!msg) return;

  const titleHtml = title ? `<div style="font-weight:700; color:var(--primary); margin-bottom:8px;">${escapeHtml(title)}</div>` : '';
  const ccHtml = msg.cc ? `<div class="preview-field"><span class="preview-field-label">CC:</span><span class="preview-field-val">${escapeHtml(msg.cc)}</span></div>` : '';
  const bccHtml = msg.bcc ? `<div class="preview-field"><span class="preview-field-label">BCC:</span><span class="preview-field-val">${escapeHtml(msg.bcc)}</span></div>` : '';

  $('preview').innerHTML = `
    ${titleHtml}
    <div class="preview-mail-header">
      <div class="preview-field"><span class="preview-field-label">To:</span><span class="preview-field-val preview-tag">${escapeHtml(msg.email)}</span></div>
      ${ccHtml}
      ${bccHtml}
      <div class="preview-field" style="margin-top:6px;"><span class="preview-field-label">Subject:</span><span class="preview-field-val" style="font-weight:600;">${escapeHtml(msg.subject)}</span></div>
    </div>
    <div class="preview-mail-body">${escapeHtml(msg.body)}</div>
  `;
  syncPreviewControls();
}

async function openDraftTab(msg) {
  return new Promise((resolve, reject) => {
    chrome.tabs.create({ url: composeUrl(msg), active: false }, () => {
      if (chrome.runtime?.lastError) reject(new Error(chrome.runtime.lastError.message));
      else resolve();
    });
  });
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function setRangeOpening(isOpening) {
  const labels = actionLabels();
  if ($('openRangeBtn')) {
    $('openRangeBtn').disabled = isOpening;
    $('openRangeBtn').textContent = isOpening ? 'Opening...' : labels.selected;
  }
  if ($('openNextBtn')) $('openNextBtn').disabled = isOpening;
}

function runPreflight() {
  const issues = [
    ...findMissingTemplateVariables(csvHeaders, currentTemplates()),
    ...validateMessages(generated)
  ];

  return issues.map(issue => issue.includes('Invalid') || issue.includes('Duplicate') || issue.includes('appears')
    ? issue
    : `Unknown template variable: {{${issue}}}`);
}

function getSelectedRange() {
  const total = generated.length;
  const start = Number.parseInt($('rangeStart')?.value, 10);
  const end = Number.parseInt($('rangeEnd')?.value, 10);

  if (!Number.isInteger(start) || !Number.isInteger(end)) {
    throw new Error('Please enter a valid From and To range.');
  }
  if (start < 1 || end < 1 || start > total || end > total) {
    throw new Error(`Range must be between 1 and ${total}.`);
  }
  if (start > end) {
    throw new Error('From must be less than or equal to To.');
  }

  return { startIndex: start - 1, endIndex: end - 1 };
}

// Event Listeners Initialization
TEMPLATE_FIELD_IDS.forEach(id => {
  if ($(id)) {
    $(id).addEventListener('focus', () => {
      lastTemplateFieldId = id;
    });
  }
});

['subject', 'cc', 'bcc', 'body'].forEach(id => {
  if ($(id)) {
    $(id).addEventListener('input', () => {
      if (id === 'subject' || id === 'body') {
        if ($('templatePreset')) $('templatePreset').value = '';
      }
      if (rows.length) {
        makeMessages();
        renderPreview();
      }
      saveState();
    });
  }
});

if ($('composeMode')) {
  $('composeMode').addEventListener('change', () => {
    renderComposeModeHints();
    saveState();
  });
}

// Segmented buttons click handling
if (typeof document !== 'undefined' && document.querySelectorAll) {
  const segButtons = document.querySelectorAll('.segmented-btn');
  segButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const mode = btn.getAttribute('data-mode');
      if ($('composeMode')) {
        $('composeMode').value = mode;
        $('composeMode').dispatchEvent(new Event('change', { bubbles: true }));
      }
    });
  });
}

// Collapsible CC/BCC Toggle
if ($('toggleCcBccBtn')) {
  $('toggleCcBccBtn').addEventListener('click', () => {
    const container = $('ccBccContainer');
    const isHidden = container?.classList.contains('hidden');
    updateCcBccVisibility(isHidden);
  });
}

// Raw CSV text toggle
if ($('togglePasteCsvBtn')) {
  $('togglePasteCsvBtn').addEventListener('click', () => {
    const container = $('pasteCsvContainer');
    if (container) {
      container.classList.toggle('hidden');
    }
  });
}

// Drag & drop file upload
const dropZone = $('dropZone');
if (dropZone) {
  ['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropZone.classList.add('dragover');
    });
  });
  ['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropZone.classList.remove('dragover');
    });
  });
  dropZone.addEventListener('drop', async (e) => {
    const dt = e.dataTransfer;
    if (dt && dt.files && dt.files.length) {
      const file = dt.files[0];
      const text = await file.text();
      if ($('csvText')) $('csvText').value = text;
      try {
        await loadRows();
        renderStatus(`CSV loaded successfully from "${file.name}" (${rows.length} rows).`, 'ok');
      } catch (err) {
        renderStatus(err.message, 'error');
      }
    }
  });
}

if ($('csvFile')) {
  $('csvFile').addEventListener('change', async () => {
    try {
      await loadRows();
      renderStatus(`CSV file selected (${rows.length} rows). Click "Generate Drafts".`, 'ok');
    } catch (err) {
      renderStatus(err.message, 'error');
    }
  });
}

if ($('autoSend')) {
  $('autoSend').addEventListener('change', () => {
    syncDraftActionLabels();
    saveState();
  });
}

if ($('batchDelay')) {
  $('batchDelay').addEventListener('change', () => {
    saveState();
  });
}

if ($('templatePreset')) {
  $('templatePreset').addEventListener('change', () => {
    const preset = TEMPLATE_PRESETS[$('templatePreset').value];
    if (!preset) return;
    if ($('subject')) $('subject').value = preset.subject;
    if ($('body')) $('body').value = preset.body;
    if (rows.length) {
      makeMessages();
      renderPreview();
    }
    saveState();
  });
}

if ($('csvText')) {
  $('csvText').addEventListener('input', () => {
    saveState();
  });
}

if ($('clearStateBtn')) {
  $('clearStateBtn').addEventListener('click', async () => {
    if (rows.length && typeof confirm === 'function') {
      if (!confirm('Clear all saved data, CSV rows, and progress?')) return;
    }
    await chrome.storage.local.remove(STORAGE_KEY);
    rows = [];
    currentIndex = 0;
    previewIndex = 0;
    generated = [];
    csvHeaders = [];
    if ($('csvText')) $('csvText').value = '';
    if ($('composeMode')) $('composeMode').value = 'outlook';
    if ($('autoSend')) $('autoSend').checked = false;
    if ($('templatePreset')) $('templatePreset').value = '';
    if ($('batchDelay')) $('batchDelay').value = '1000';
    if ($('subject')) $('subject').value = TEMPLATE_PRESETS.courseReminder.subject;
    if ($('cc')) $('cc').value = '';
    if ($('bcc')) $('bcc').value = '';
    if ($('body')) $('body').value = TEMPLATE_PRESETS.courseReminder.body;
    if ($('sendControls')) $('sendControls').classList.add('hidden');
    if ($('previewControls')) $('previewControls').classList.add('hidden');
    if ($('preview')) $('preview').innerHTML = '';
    updateCcBccVisibility(false);
    updateCsvSummary();
    renderVariableButtons();
    renderComposeModeHints();
    syncDraftActionLabels();
    renderStatus('Saved data cleared. Ready for new merge.', 'ok');
  });
}

if ($('loadBtn')) {
  $('loadBtn').addEventListener('click', async () => {
    try {
      await loadRows();
      makeMessages();
      const issues = runPreflight();
      currentIndex = 0;
      previewIndex = 0;
      await saveState();
      if ($('sendControls')) $('sendControls').classList.remove('hidden');
      if ($('previewControls')) $('previewControls').classList.remove('hidden');
      syncRangeInputs();
      syncDraftActionLabels();
      renderPreview(`Loaded ${generated.length} rows.`);
      const labels = actionLabels();
      renderStatus(
        `Ready: ${generated.length} personalized ${labels.items}. Review preview, then ${labels.openVerb} ${labels.items}.`
          + renderPreflightIssues(issues),
        issues.length ? 'error' : 'ok'
      );
    } catch (e) {
      renderStatus(e.message, 'error');
    }
  });
}

if ($('previewPrevBtn')) {
  $('previewPrevBtn').addEventListener('click', async () => {
    if (!generated.length) return;
    previewIndex = clampPreviewIndex(previewIndex - 1);
    renderPreview();
    await saveState();
  });
}

if ($('previewNextBtn')) {
  $('previewNextBtn').addEventListener('click', async () => {
    if (!generated.length) return;
    previewIndex = clampPreviewIndex(previewIndex + 1);
    renderPreview();
    await saveState();
  });
}

if ($('openNextBtn')) {
  $('openNextBtn').addEventListener('click', async () => {
    const labels = actionLabels();
    if (!generated.length) return renderStatus('Please generate drafts first.', 'error');

    makeMessages();
    const issues = runPreflight();
    if (issues.length) {
      return renderStatus(`Please fix the preflight warnings before ${labels.openingVerb} ${labels.items}.` + renderPreflightIssues(issues), 'error');
    }
    previewIndex = clampPreviewIndex(previewIndex);
    const msg = generated[previewIndex];
    const previousIndex = currentIndex;
    currentIndex = Math.max(currentIndex, previewIndex + 1);
    await saveState();

    try {
      await openDraftTab(msg);
      await saveState();
      syncRangeInputs();
      syncDraftActionLabels();
      renderPreview();
      renderStatus(`${labels.openedVerb} draft ${previewIndex + 1}/${generated.length}: ${msg.email}${$('autoSend')?.checked ? '' : '\n\nReview the opened draft tab before sending.'}`, 'ok');
    } catch (e) {
      currentIndex = previousIndex;
      await saveState();
      syncRangeInputs();
      renderStatus(`Could not ${labels.openVerb} ${labels.item}: ${e.message}`, 'error');
    }
  });
}

if ($('openRangeBtn')) {
  $('openRangeBtn').addEventListener('click', async () => {
    const labels = actionLabels();
    if (!generated.length) return renderStatus('Please generate drafts first.', 'error');

    makeMessages();
    const issues = runPreflight();
    if (issues.length) {
      return renderStatus(`Please fix the preflight warnings before ${labels.openingVerb} ${labels.items}.` + renderPreflightIssues(issues), 'error');
    }
    let range;
    try {
      range = getSelectedRange();
    } catch (e) {
      return renderStatus(e.message, 'error');
    }

    const { startIndex, endIndex } = range;
    const selected = generated.slice(startIndex, endIndex + 1);
    const previousIndex = currentIndex;
    currentIndex = Math.max(currentIndex, endIndex + 1);
    await saveState();

    const delayMs = Number.parseInt($('batchDelay')?.value || `${DEFAULT_DELAY_MS}`, 10) || DEFAULT_DELAY_MS;
    let openedCount = 0;
    setRangeOpening(true);

    const progressContainer = $('progressContainer');
    const progressBar = $('progressBar');
    if (progressContainer) progressContainer.classList.remove('hidden');
    if (progressBar) progressBar.style.width = '0%';

    try {
      for (const [index, msg] of selected.entries()) {
        renderStatus(`${$('autoSend')?.checked ? 'Opening and sending' : 'Opening'} draft ${index + 1}/${selected.length}...`, 'ok');
        await openDraftTab(msg);
        openedCount += 1;

        if (progressBar) {
          const percent = Math.round((openedCount / selected.length) * 100);
          progressBar.style.width = `${percent}%`;
        }

        if (openedCount < selected.length) {
          await sleep(delayMs);
        }
      }
      previewIndex = clampPreviewIndex(currentIndex);
      await saveState();
      syncRangeInputs();
      syncDraftActionLabels();
      renderPreview();
      renderStatus(`${labels.openedVerb} ${selected.length} ${labels.items} (${startIndex + 1}-${endIndex + 1}/${generated.length}).${$('autoSend')?.checked ? '' : '\n\nReview the opened drafts in your browser.'}`, 'ok');
    } catch (e) {
      currentIndex = Math.max(previousIndex, startIndex + openedCount);
      await saveState();
      syncRangeInputs();
      renderStatus(`Stopped after ${labels.openingVerb} ${openedCount}/${selected.length} selected ${labels.items}: ${e.message}`, 'error');
    } finally {
      setRangeOpening(false);
      setTimeout(() => {
        if (progressContainer) progressContainer.classList.add('hidden');
      }, 1500);
    }
  });
}

// Initial restoration
restoreState();
renderComposeModeHints();
syncDraftActionLabels();
