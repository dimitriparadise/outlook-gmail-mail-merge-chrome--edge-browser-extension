let rows = [];
let currentIndex = 0;
let previewIndex = 0;
let generated = [];
let csvHeaders = [];
let lastTemplateFieldId = 'body';

const $ = (id) => document.getElementById(id);
const STORAGE_KEY = 'outlookMailMergeState';
const OPEN_RANGE_DELAY_MS = 500;
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
  outlook: 'Opens Outlook Web drafts. If CC/BCC is present, Outlook Mode uses mailto because Outlook Web may ignore copy recipients.',
  gmail: 'Opens Gmail compose windows directly. Gmail Mode supports To, CC, BCC, subject, and body without mailto.'
};

function hasOwn(object, key) {
  return Object.prototype.hasOwnProperty.call(object, key);
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
  if (lines.length < 2) throw new Error('CSV needs a header row and at least one data row.');
  const headers = lines[0].map(h => h.trim());
  csvHeaders = headers;
  return lines.slice(1).map(line => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = line[i] || '');
    return obj;
  }).filter(r => String(r.Email || r.email || '').trim() !== '');
}

function applyTemplate(template, row) {
  return template.replace(/{{\s*([^}]+?)\s*}}/g, (_, key) => row[key] ?? row[key.trim()] ?? '');
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
  return $('csvText').value.trim();
}

async function saveState() {
  const state = {
    csvText: $('csvText').value,
    composeMode: $('composeMode').value,
    autoSend: $('autoSend').checked,
    templatePreset: $('templatePreset').value,
    subject: $('subject').value,
    cc: $('cc').value,
    bcc: $('bcc').value,
    body: $('body').value,
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

async function restoreState() {
  const result = await chrome.storage.local.get(STORAGE_KEY);
  const state = result[STORAGE_KEY];
  if (!state) return;

  $('csvText').value = hasOwn(state, 'csvText') ? state.csvText : $('csvText').value;
  $('composeMode').value = state.composeMode || 'outlook';
  $('autoSend').checked = Boolean(state.autoSend);
  $('templatePreset').value = state.templatePreset || '';
  $('subject').value = hasOwn(state, 'subject') ? state.subject : $('subject').value;
  $('cc').value = hasOwn(state, 'cc') ? state.cc : '';
  $('bcc').value = hasOwn(state, 'bcc') ? state.bcc : '';
  $('body').value = hasOwn(state, 'body') ? state.body : $('body').value;
  rows = Array.isArray(state.rows) ? state.rows : [];
  csvHeaders = Array.isArray(state.csvHeaders) ? state.csvHeaders : getHeadersFromRows(rows);
  generated = Array.isArray(state.generated) ? state.generated : [];
  currentIndex = Number.isInteger(state.currentIndex) ? state.currentIndex : 0;
  previewIndex = Number.isInteger(state.previewIndex) ? state.previewIndex : currentIndex;
  renderVariableButtons();
  renderComposeModeHints();

  if (generated.length) {
    previewIndex = clampPreviewIndex(previewIndex);
    $('sendControls').classList.remove('hidden');
    $('previewControls').classList.remove('hidden');
    syncRangeInputs();
    renderStatus(`Restored: ${currentIndex}/${generated.length} drafts opened. Click “Open Next Draft” to continue.`, 'ok');
    renderPreview();
  }
}

async function loadRows() {
  const file = $('csvFile').files[0];
  let text = getCsvText();
  if (file) {
    text = await file.text();
    // File input cannot be restored after the popup closes, so copy its text into the textarea.
    $('csvText').value = text;
  }
  rows = parseCSV(text);
  if (!rows.length) throw new Error('No usable rows found. Make sure you have an Email column.');
  renderVariableButtons();
  await saveState();
  return rows;
}

function makeMessages() {
  const subjectTemplate = $('subject').value;
  const ccTemplate = $('cc').value;
  const bccTemplate = $('bcc').value;
  const bodyTemplate = $('body').value;
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
  return TEMPLATE_FIELD_IDS.map(id => $(id).value);
}

function encodeForOutlook(text) {
  // Outlook Web deeplinks need percent-encoded spaces/newlines. URLSearchParams may serialize spaces as '+',
  // which Outlook sometimes shows literally in the compose window.
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

  if ($('autoSend').checked) params.push(['mailMergeAutoSend', 'true']);

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

  if ($('autoSend').checked) params.push(['mailMergeAutoSend', 'true']);

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
  if ($('composeMode').value === 'gmail') return gmailComposeUrl(msg);

  // Outlook Web deeplinks currently ignore cc/bcc in some tenants. mailto handles those fields reliably
  // through the user's configured mail handler, so use it only when copy recipients are present.
  if ($('autoSend').checked && hasCopyRecipients(msg)) {
    throw new Error('Auto-Send with Outlook Mode cannot be used with CC or BCC. Use Gmail Mode or turn off Auto-Send.');
  }
  if ($('autoSend').checked) return outlookComposeUrl(msg);
  return hasCopyRecipients(msg) ? mailtoUrl(msg) : outlookComposeUrl(msg);
}

function renderStatus(text, cls = '') {
  $('status').className = cls;
  $('status').textContent = text;
}

function renderComposeModeHints() {
  const mode = $('composeMode').value;
  $('composeModeHint').textContent = COMPOSE_MODE_HINTS[mode] || '';
  $('copyRecipientHint').innerHTML = mode === 'outlook'
    ? 'In Outlook Mode, CC/BCC drafts use your browser&apos;s <code>mailto:</code> handler.'
    : 'In Gmail Mode, CC/BCC are included in the Gmail compose URL.';
}

function syncRangeInputs() {
  const total = generated.length;
  const next = Math.min(currentIndex + 1, total || 1);
  $('rangeStart').max = total || 1;
  $('rangeEnd').max = total || 1;
  $('rangeStart').value = next;
  $('rangeEnd').value = total || 1;
}

function syncPreviewControls() {
  const total = generated.length;
  const hasMessages = total > 0;
  $('previewControls').classList.toggle('hidden', !hasMessages);
  $('previewCounter').textContent = hasMessages ? `${previewIndex + 1}/${total}` : '';
  $('previewPrevBtn').disabled = !hasMessages || previewIndex <= 0;
  $('previewNextBtn').disabled = !hasMessages || previewIndex >= total - 1;
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
  const start = field.selectionStart ?? field.value.length;
  const end = field.selectionEnd ?? field.value.length;
  field.value = field.value.slice(0, start) + text + field.value.slice(end);
  field.focus();
  field.setSelectionRange(start + text.length, start + text.length);
  field.dispatchEvent(new Event('input', { bubbles: true }));
}

function renderVariableButtons() {
  const container = $('variableButtons');
  container.textContent = '';

  csvHeaders.filter(Boolean).forEach(header => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'variable-button';
    button.textContent = `{{${header}}}`;
    button.title = `Insert {{${header}}}`;
    button.addEventListener('click', () => {
      insertAtCursor($(lastTemplateFieldId), `{{${header}}}`);
    });
    container.appendChild(button);
  });
}

function messagePreview(title, msg) {
  const lines = [
    title,
    `To: ${msg.email}`
  ];
  if (msg.cc) lines.push(`CC: ${msg.cc}`);
  if (msg.bcc) lines.push(`BCC: ${msg.bcc}`);
  lines.push(`Subject: ${msg.subject}`, '', msg.body);
  return lines.join('\n');
}

function renderPreview(title) {
  if (!generated.length) return;

  previewIndex = clampPreviewIndex(previewIndex);
  const msg = generated[previewIndex];
  if (msg) {
    $('preview').textContent = messagePreview(title || `Preview draft ${previewIndex + 1} of ${generated.length}:`, msg);
  }
  syncPreviewControls();
}

async function openDraftTab(msg) {
  return new Promise((resolve, reject) => {
    chrome.tabs.create({ url: composeUrl(msg), active: false }, () => {
      if (chrome.runtime.lastError) reject(new Error(chrome.runtime.lastError.message));
      else resolve();
    });
  });
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function setRangeOpening(isOpening) {
  $('openRangeBtn').disabled = isOpening;
  $('openNextBtn').disabled = isOpening;
  $('openRangeBtn').textContent = isOpening ? 'Opening...' : 'Open Range';
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
  const start = Number.parseInt($('rangeStart').value, 10);
  const end = Number.parseInt($('rangeEnd').value, 10);

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

TEMPLATE_FIELD_IDS.forEach(id => {
  $(id).addEventListener('focus', () => {
    lastTemplateFieldId = id;
  });
});

['subject', 'cc', 'bcc', 'body'].forEach(id => {
  $(id).addEventListener('input', () => {
    if (id === 'subject' || id === 'body') $('templatePreset').value = '';
    if (rows.length) {
      makeMessages();
      renderPreview();
    }
    saveState();
  });
});

$('composeMode').addEventListener('change', () => {
  renderComposeModeHints();
  saveState();
});

$('autoSend').addEventListener('change', () => {
  saveState();
});

$('templatePreset').addEventListener('change', () => {
  const preset = TEMPLATE_PRESETS[$('templatePreset').value];
  if (!preset) return;
  $('subject').value = preset.subject;
  $('body').value = preset.body;
  if (rows.length) {
    makeMessages();
    renderPreview();
  }
  saveState();
});

$('csvText').addEventListener('input', () => {
  saveState();
});

$('clearStateBtn').addEventListener('click', async () => {
  await chrome.storage.local.remove(STORAGE_KEY);
  rows = [];
  currentIndex = 0;
  previewIndex = 0;
  generated = [];
  csvHeaders = [];
  $('csvText').value = '';
  $('composeMode').value = 'outlook';
  $('autoSend').checked = false;
  $('templatePreset').value = '';
  $('subject').value = TEMPLATE_PRESETS.courseReminder.subject;
  $('cc').value = '';
  $('bcc').value = '';
  $('body').value = TEMPLATE_PRESETS.courseReminder.body;
  $('sendControls').classList.add('hidden');
  $('previewControls').classList.add('hidden');
  $('preview').textContent = '';
  renderVariableButtons();
  renderComposeModeHints();
  renderStatus('Saved data cleared. The default template is ready.', 'ok');
});

$('loadBtn').addEventListener('click', async () => {
  try {
    await loadRows();
    makeMessages();
    const issues = runPreflight();
    currentIndex = 0;
    previewIndex = 0;
    await saveState();
    $('sendControls').classList.remove('hidden');
    $('previewControls').classList.remove('hidden');
    syncRangeInputs();
    renderPreview(`Loaded ${generated.length} rows.\n\nFirst preview:`);
    renderStatus(
      `Ready: ${generated.length} personalized drafts. Review the preview, then open drafts.`
        + renderPreflightIssues(issues),
      issues.length ? 'error' : 'ok'
    );
  } catch (e) {
    renderStatus(e.message, 'error');
  }
});

$('previewPrevBtn').addEventListener('click', async () => {
  if (!generated.length) return;
  previewIndex = clampPreviewIndex(previewIndex - 1);
  renderPreview();
  await saveState();
});

$('previewNextBtn').addEventListener('click', async () => {
  if (!generated.length) return;
  previewIndex = clampPreviewIndex(previewIndex + 1);
  renderPreview();
  await saveState();
});

$('openNextBtn').addEventListener('click', async () => {
  if (!generated.length) return renderStatus('Please generate drafts first.', 'error');
  if (currentIndex >= generated.length) return renderStatus('All drafts opened.', 'ok');

  makeMessages();
  const issues = runPreflight();
  if (issues.length) {
    return renderStatus('Please fix the preflight warnings before opening drafts.' + renderPreflightIssues(issues), 'error');
  }
  const msg = generated[currentIndex];
  const openedIndex = currentIndex + 1;
  currentIndex = openedIndex;
  await saveState();

  try {
    await openDraftTab(msg);
    previewIndex = clampPreviewIndex(currentIndex);
    await saveState();
    syncRangeInputs();
    renderPreview();
    renderStatus(`Opened draft ${currentIndex}/${generated.length}: ${msg.email}\n\nKeep clicking “Open Next Draft” to queue the rest, or review the opened drafts before sending.`, 'ok');
  } catch (e) {
    currentIndex = openedIndex - 1;
    await saveState();
    syncRangeInputs();
    renderStatus(`Could not open draft: ${e.message}`, 'error');
  }
});

$('openRangeBtn').addEventListener('click', async () => {
  if (!generated.length) return renderStatus('Please generate drafts first.', 'error');

  makeMessages();
  const issues = runPreflight();
  if (issues.length) {
    return renderStatus('Please fix the preflight warnings before opening drafts.' + renderPreflightIssues(issues), 'error');
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

  let openedCount = 0;
  setRangeOpening(true);
  try {
    for (const [index, msg] of selected.entries()) {
      renderStatus(`Opening draft ${index + 1}/${selected.length} from selected range...`, 'ok');
      await openDraftTab(msg);
      openedCount += 1;
      if (openedCount < selected.length) {
        await sleep(OPEN_RANGE_DELAY_MS);
      }
    }
    previewIndex = clampPreviewIndex(currentIndex);
    await saveState();
    syncRangeInputs();
    renderPreview();
    renderStatus(`Opened ${selected.length} drafts (${startIndex + 1}-${endIndex + 1}/${generated.length}).\n\nReview the opened drafts before sending.`, 'ok');
  } catch (e) {
    currentIndex = Math.max(previousIndex, startIndex + openedCount);
    await saveState();
    syncRangeInputs();
    renderStatus(`Stopped after opening ${openedCount}/${selected.length} selected drafts: ${e.message}`, 'error');
  } finally {
    setRangeOpening(false);
  }
});

restoreState();
renderComposeModeHints();
