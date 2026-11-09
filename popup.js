let rows = [];
// Indices of drafts that have been opened, so batches never skip or repeat a recipient.
let openedRows = new Set();
// True while drafts are being opened; blocks starting a second run at the same time.
let isOpening = false;
let previewIndex = 0;
let generated = [];
let csvHeaders = [];
let lastTemplateFieldId = 'body';
// { start, end, at } for the most recently opened range, shown under the batch button.
let lastBatch = null;
// Keys of review notes the user marked as fine.
let dismissedNotes = [];
// Set when the raw CSV text is edited by hand, so the next Generate starts progress over.
let listEdited = false;
// Column picker state while an Excel file is being imported.
let mapping = null;
// Step-by-step mode: which step is shown, and whether the mode is on (otherwise everything is on one page).
let wizardStep = 1;
let wizardOn = true;

const $ = (id) => document.getElementById(id);
const STORAGE_KEY = 'outlookMailMergeState';
const DEFAULT_DELAY_MS = 1000;
const DEFAULT_BATCH_SIZE = 50;
const MAX_BATCH_SIZE = 500;
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
const PLAN_TYPE = 'mail-merge-plan';
const PLAN_VERSION = 1;
const BATCH_DELAY_OPTIONS = ['500', '1000', '2000'];
const EMAIL_HEADER_ALIASES = ['email', 'e-mail', 'mail', '邮箱', 'recipient', 'to', 'student email', 'student_email', 'correo', 'correo electrónico'];
const NAME_COLUMN_PATTERN = /name|名|nombre/i;

function isAutoSendOn() {
  return $('autoSend')?.checked ?? false;
}

// Message keys ending in "Send" hold the Auto-Send wording of the same message.
function tm(key, vars) {
  return t(isAutoSendOn() ? `${key}Send` : key, vars);
}

function syncDraftActionLabels() {
  if ($('openNextBtn')) $('openNextBtn').textContent = tm('openPreview');
  if ($('openRangeBtn')) $('openRangeBtn').textContent = tm('openRange');
  if ($('sendTestBtn')) $('sendTestBtn').textContent = tm('sendTest');
  renderBatchControls();
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

function detectEmailHeader(headers) {
  return headers.find(h => EMAIL_HEADER_ALIASES.includes(String(h).trim().toLowerCase()))
    || headers.find(h => /email|邮箱|mail|correo/i.test(h))
    || null;
}

function parseCSV(text) {
  const records = [];
  let record = [];
  let cur = '';
  let inQuotes = false;
  const source = text.replace(/^﻿/, '');

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

  if (inQuotes) throw new Error(t('csvUnclosedQuote'));
  record.push(cur.trim());
  if (record.some(value => value !== '')) records.push(record);

  const lines = records;
  if (lines.length < 2) {
    const hasCommaSeparatedValues = source.split(',').length > 2;
    if (hasCommaSeparatedValues) throw new Error(t('csvNeedsLineBreak'));
    throw new Error(t('csvNeedsRows'));
  }

  const headers = lines[0].map(h => h.trim());
  csvHeaders = headers;
  const emailKey = detectEmailHeader(headers);

  return lines.slice(1).map(line => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = line[i] || '');
    if (emailKey && !obj.Email && !obj.email) {
      obj.Email = obj[emailKey];
    }
    return obj;
  }).filter(r => String(r.Email || r.email || '').trim() !== '');
}

function csvCell(value) {
  const text = String(value ?? '');
  return /[",\r\n]/.test(text) || text !== text.trim() ? `"${text.replace(/"/g, '""')}"` : text;
}

function rowsToCSV(headers, rowList) {
  return [headers, ...rowList.map(row => headers.map(h => row[h] ?? ''))]
    .map(cells => cells.map(csvCell).join(','))
    .join('\n');
}

// {{Name}} inserts a column value; {{Name|there}} uses "there" when that value is empty.
const VARIABLE_PATTERN = /{{\s*([^}|]+?)\s*(?:\|([^}]*))?}}/g;

function applyTemplate(template, row) {
  return String(template || '').replace(VARIABLE_PATTERN, (_, key, fallback) => {
    const value = String(row[key] ?? '');
    return value.trim() !== '' || fallback === undefined ? value : fallback.trim();
  });
}

function templateVariables(template) {
  return [...new Set([...String(template || '').matchAll(VARIABLE_PATTERN)].map(match => match[1].trim()))];
}

// Variables used at least once without a {{Name|default}}, so an empty value would show as a gap.
function variablesWithoutDefault(templates) {
  return new Set(templates.flatMap(template => [...String(template || '').matchAll(VARIABLE_PATTERN)]
    .filter(match => match[2] === undefined)
    .map(match => match[1].trim())));
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

  messages.forEach((msg, index) => {
    const row = t('draftLabel', { n: index + 1 });
    const to = String(msg.email || '').trim();
    if (!isValidEmail(to)) issues.push(t('issueInvalidEmail', { row, email: to || t('blank') }));

    const copyRecipients = [...splitRecipients(msg.cc), ...splitRecipients(msg.bcc)];
    copyRecipients.forEach(recipient => {
      if (!isValidEmail(recipient)) issues.push(t('issueInvalidEmail', { row, email: recipient }));
      if (recipient.toLowerCase() === to.toLowerCase()) {
        issues.push(t('issueCopyOverlap', { row, email: recipient }));
      }
    });
  });

  // The same address on several rows is allowed (e.g. one parent of two students); see reviewRecipients.
  return issues;
}

// --- Recipient review notes: soft checks that never block sending ---

function findFirstNameHeader(headers) {
  return headers.find(h => /^(first[\s_-]*name|given[\s_-]*name|preferred[\s_-]*name|nombre|名|名字)$/i.test(String(h).trim()))
    || headers.find(h => /^(name|姓名|full[\s_-]*name)$/i.test(String(h).trim()))
    || null;
}

function emailDomain(email) {
  const at = String(email).lastIndexOf('@');
  return at < 0 ? '' : String(email).slice(at + 1).trim().toLowerCase();
}

function asciiLetters(text) {
  return String(text).toLowerCase().normalize('NFD').replace(/[^a-z]/g, '');
}

function reviewRecipients(rowList, headers, templates) {
  const used = new Set(templates.flatMap(templateVariables));
  const needsValue = variablesWithoutDefault(templates);
  const emailHeader = detectEmailHeader(headers) || 'Email';
  const nameHeader = findFirstNameHeader(headers);
  const isGivenName = nameHeader && !/^(name|姓名|full[\s_-]*name)$/i.test(nameHeader.trim());
  const emails = rowList.map(row => String(row.Email || row.email || '').trim());
  const emailCounts = new Map();
  emails.forEach(email => email && emailCounts.set(email.toLowerCase(), (emailCounts.get(email.toLowerCase()) || 0) + 1));

  const domainCounts = new Map();
  emails.map(emailDomain).forEach(domain => domain && domainCounts.set(domain, (domainCounts.get(domain) || 0) + 1));
  const [mainDomain, mainCount] = [...domainCounts.entries()].sort((a, b) => b[1] - a[1])[0] || ['', 0];
  const checkDomains = rowList.length >= 5 && mainCount / rowList.length >= 0.8;

  // A name that is not part of the address local part, e.g. "Jhoune-Lynn" for lynn.duplan@...
  const mismatched = rowList.map((row, index) => {
    if (!nameHeader) return false;
    const first = asciiLetters(String(row[nameHeader] ?? '').trim().split(/\s+/)[0]);
    const local = asciiLetters(emails[index].split('@')[0]);
    return Boolean(first && local && !local.includes(first));
  });
  // If many addresses do not contain names at all (e.g. jsmith3@), the check is just noise.
  const checkNames = mismatched.filter(Boolean).length <= Math.max(1, rowList.length * 0.3);

  const notes = [];
  rowList.forEach((row, index) => {
    const email = emails[index];
    const add = (type, field, message) => notes.push({
      key: `${type}:${field}:${email.toLowerCase()}`,
      index,
      field,
      email,
      value: row[field] ?? '',
      message
    });

    needsValue.forEach(variable => {
      if (headers.includes(variable) && String(row[variable] ?? '').trim() === '') add('empty', variable, t('noteEmpty', { name: variable }));
    });
    const copies = emailCounts.get(email.toLowerCase()) || 0;
    if (copies > 1) add('duplicate', emailHeader, t('noteDuplicate', { count: copies }));
    if (nameHeader && checkNames && mismatched[index]) {
      add('mismatch', nameHeader, t('noteNameMismatch', { name: nameHeader }));
    } else if (isGivenName && used.has(nameHeader) && /\s/.test(String(row[nameHeader] ?? '').trim())) {
      add('multiword', nameHeader, t('noteMultiWord', { name: nameHeader }));
    }
    const domain = emailDomain(email);
    if (checkDomains && domain && domain !== mainDomain) add('domain', emailHeader, t('noteDomain', { domain, main: mainDomain }));
  });

  return notes.filter(note => !dismissedNotes.includes(note.key));
}

function renderReviewPanel() {
  const panel = $('reviewPanel');
  if (!panel) return;
  const notes = generated.length ? reviewRecipients(rows, csvHeaders, currentTemplates()) : [];
  panel.classList.toggle('hidden', !notes.length);
  if (!notes.length) {
    panel.innerHTML = '';
    return;
  }

  const shown = notes.slice(0, 30);
  panel.innerHTML = `
    <div class="review-title">${escapeHtml(t('reviewTitle', { count: notes.length }))}</div>
    <p class="hint">${escapeHtml(t('reviewHint'))}</p>
    ${shown.map(note => `
      <div class="review-item">
        <div class="review-meta"><span class="review-row">${escapeHtml(t('rowLabel', { n: note.index + 1 }))}</span><span class="review-email">${escapeHtml(note.email)}</span></div>
        <div class="review-message">${escapeHtml(note.message)}</div>
        <div class="review-edit">
          <input type="text" value="${escapeHtml(note.value)}" data-row="${note.index}" data-field="${escapeHtml(note.field)}" aria-label="${escapeHtml(note.field)}" />
          <button type="button" class="review-dismiss" data-dismiss="${escapeHtml(note.key)}">${escapeHtml(t('reviewDismiss'))}</button>
        </div>
      </div>`).join('')}
    ${notes.length > shown.length ? `<p class="hint">${escapeHtml(t('reviewMore', { count: notes.length - shown.length }))}</p>` : ''}
  `;
}

async function updateRecipientField(index, field, value) {
  if (!rows[index] || !csvHeaders.includes(field)) return false;
  const nextValue = String(value ?? '').trim();
  // Blanking an address would drop the row and shift every later draft's position.
  if (!nextValue && field === (detectEmailHeader(csvHeaders) || 'Email')) {
    renderReviewPanel();
    renderStatus(t('emailRequired'), 'error');
    return false;
  }
  rows[index][field] = nextValue;
  if ($('csvText')) $('csvText').value = rowsToCSV(csvHeaders, rows);
  // Re-parse so the derived Email field follows an edited email column.
  rows = parseCSV(getCsvText());
  makeMessages();
  renderPreview();
  renderReviewPanel();
  renderBatchControls();
  await saveState();
  return true;
}

function getCsvText() {
  return $('csvText') ? $('csvText').value.trim() : '';
}

async function saveState() {
  const state = {
    uiLanguage,
    csvText: $('csvText')?.value ?? '',
    composeMode: $('composeMode')?.value ?? 'outlook',
    autoSend: Boolean($('autoSend')?.checked),
    templatePreset: $('templatePreset')?.value ?? '',
    batchDelay: $('batchDelay')?.value ?? '1000',
    batchSize: getBatchSize(),
    testEmail: $('testEmail')?.value ?? '',
    subject: $('subject')?.value ?? '',
    cc: $('cc')?.value ?? '',
    bcc: $('bcc')?.value ?? '',
    body: $('body')?.value ?? '',
    rows,
    csvHeaders,
    generated,
    openedRows: [...openedRows],
    previewIndex,
    lastBatch,
    dismissedNotes,
    listEdited,
    wizardStep,
    wizardOn,
    hasStarted: generated.length > 0,
    updatedAt: Date.now()
  };
  await chrome.storage.local.set({ [STORAGE_KEY]: state });
}

function updateCcBccVisibility(show) {
  const container = $('ccBccContainer');
  if (!container) return;
  container.classList.toggle('hidden', !show);
  renderCcBccToggle();
}

function renderCcBccToggle() {
  const container = $('ccBccContainer');
  const trigger = $('toggleCcBccBtn');
  if (!container || !trigger) return;
  trigger.textContent = container.classList.contains('hidden') ? t('addCcBcc') : t('hideCcBcc');
}

function updateCsvSummary() {
  const summaryEl = $('csvSummary');
  if (!summaryEl) return;
  // While the Excel column picker is open, the loaded count would describe the previous list.
  if (rows.length > 0 && !mapping) {
    summaryEl.textContent = t('rowsLoaded', { count: rows.length });
    summaryEl.classList.remove('hidden');
  } else {
    summaryEl.classList.add('hidden');
  }
}

function applyLanguageSelection(lang) {
  setUiLanguage(lang);
  if ($('uiLanguage')) $('uiLanguage').value = uiLanguage;
  applyStaticTranslations();
  renderDynamicText();
}

// Re-renders every piece of text that JavaScript writes, e.g. after the language changes.
function renderDynamicText() {
  renderWizard();
  syncDraftActionLabels();
  renderComposeModeHints();
  renderCcBccToggle();
  updateCsvSummary();
  renderVariableButtons();
  renderMappingPanel();
  renderReviewPanel();
  if (generated.length) renderPreview();
}

async function restoreState() {
  const result = await chrome.storage.local.get(STORAGE_KEY);
  const state = result[STORAGE_KEY];
  if (!state) return;

  if (state.uiLanguage) applyLanguageSelection(state.uiLanguage);
  if ($('csvText')) $('csvText').value = hasOwn(state, 'csvText') ? state.csvText : $('csvText').value;
  if ($('composeMode')) $('composeMode').value = state.composeMode || 'outlook';
  if ($('autoSend')) $('autoSend').checked = Boolean(state.autoSend);
  if ($('templatePreset')) $('templatePreset').value = state.templatePreset || '';
  if ($('batchDelay')) $('batchDelay').value = state.batchDelay || '1000';
  if ($('batchSize')) $('batchSize').value = String(state.batchSize || DEFAULT_BATCH_SIZE);
  if ($('testEmail')) $('testEmail').value = state.testEmail || '';
  if ($('subject')) $('subject').value = hasOwn(state, 'subject') ? state.subject : $('subject').value;
  if ($('cc')) $('cc').value = hasOwn(state, 'cc') ? state.cc : '';
  if ($('bcc')) $('bcc').value = hasOwn(state, 'bcc') ? state.bcc : '';
  if ($('body')) $('body').value = hasOwn(state, 'body') ? state.body : $('body').value;

  rows = Array.isArray(state.rows) ? state.rows : [];
  csvHeaders = Array.isArray(state.csvHeaders) ? state.csvHeaders : getHeadersFromRows(rows);
  generated = Array.isArray(state.generated) ? state.generated : [];
  openedRows = new Set(Array.isArray(state.openedRows)
    ? state.openedRows
    // 0.1.x saved only the highest opened position.
    : Array.from({ length: Number.isInteger(state.currentIndex) ? state.currentIndex : 0 }, (_, i) => i));
  previewIndex = Number.isInteger(state.previewIndex) ? state.previewIndex : 0;
  lastBatch = state.lastBatch || null;
  dismissedNotes = Array.isArray(state.dismissedNotes) ? state.dismissedNotes : [];
  listEdited = Boolean(state.listEdited);
  wizardOn = state.wizardOn !== false;
  wizardStep = Number.isInteger(state.wizardStep) ? state.wizardStep : 1;

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
    renderStatus(t('restored', { done: openedCount(), total: generated.length }), 'ok');
    renderPreview();
    renderReviewPanel();
  }
  renderWizard();
}

function resetProgress() {
  generated = [];
  openedRows = new Set();
  previewIndex = 0;
  lastBatch = null;
  if ($('sendControls')) $('sendControls').classList.add('hidden');
  if ($('previewControls')) $('previewControls').classList.add('hidden');
  if ($('preview')) $('preview').innerHTML = '';
  renderReviewPanel();
}

// Parses the CSV text box. A fresh list (new file, plan, or Excel import) starts progress over.
async function loadRows({ fresh = false } = {}) {
  const parsed = parseCSV(getCsvText());
  if (!parsed.length) throw new Error(t('noRows'));
  rows = parsed;
  if (fresh) {
    dismissedNotes = [];
    listEdited = false;
    resetProgress();
  }
  renderVariableButtons();
  updateCsvSummary();
  renderWizard();
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

  if (isAutoSendOn()) params.push(['mailMergeAutoSend', 'true']);

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

  if (isAutoSendOn()) params.push(['mailMergeAutoSend', 'true']);

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

  if (isAutoSendOn() && hasCopyRecipients(msg)) {
    throw new Error(t('autoSendCcBlocked'));
  }
  if (isAutoSendOn()) return outlookComposeUrl(msg);
  return hasCopyRecipients(msg) ? mailtoUrl(msg) : outlookComposeUrl(msg);
}

function renderStatus(text, cls = '') {
  if (!$('status')) return;
  $('status').className = cls;
  $('status').textContent = text;
  // Errors get a pointer to the GitHub issues page.
  if ($('statusHelp')) $('statusHelp').classList.toggle('hidden', cls !== 'error');
}

function renderComposeModeHints() {
  const mode = $('composeMode')?.value ?? 'outlook';
  if ($('composeModeHint')) $('composeModeHint').textContent = mode === 'gmail' ? t('composeHintGmail') : t('composeHintOutlook');
  if ($('copyRecipientHint')) $('copyRecipientHint').textContent = mode === 'gmail' ? t('copyHintGmail') : t('copyHintOutlook');

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
  const next = (nextBatchIndices(total, 1)[0] ?? total - 1) + 1 || 1;
  if ($('rangeStart')) {
    $('rangeStart').max = total || 1;
    $('rangeStart').value = next;
  }
  if ($('rangeEnd')) {
    $('rangeEnd').max = total || 1;
    $('rangeEnd').value = total || 1;
  }
  renderBatchControls();
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
  return `\n\n${t('preflightHeader')}\n- ` + issues.slice(0, 8).join('\n- ')
    + (issues.length > 8 ? `\n- ${t('preflightMore', { count: issues.length - 8 })}` : '');
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

// --- Column suggestions while typing {{ ---

// Returns { start, query } when the text just before the cursor is an unfinished "{{...".
function variableQueryAt(text, cursor) {
  const match = /{{\s*([^{}|]*)$/.exec(String(text).slice(0, cursor));
  return match ? { start: match.index, query: match[1] } : null;
}

function variableSuggestions(headers, query) {
  const needle = String(query).trim().toLowerCase();
  return headers.filter(Boolean).filter(header => header.toLowerCase().includes(needle)).slice(0, 8);
}

// Replaces the unfinished "{{que" with "{{Header}}", absorbing the rest of a variable already closed after the cursor.
function applySuggestion(text, start, cursor, header) {
  const after = text.slice(cursor);
  const tail = /^[^{}\n]*}}/.exec(after);
  const insert = `{{${header}}}`;
  return {
    text: text.slice(0, start) + insert + (tail ? after.slice(tail[0].length) : after),
    cursor: start + insert.length
  };
}

let suggest = null;

function closeSuggestions() {
  suggest = null;
  if ($('variableSuggest')) $('variableSuggest').classList.add('hidden');
}

function renderSuggestions() {
  const box = $('variableSuggest');
  if (!box || !suggest) return;
  box.innerHTML = suggest.items.map((header, i) => `<div class="variable-suggest-item${i === suggest.active ? ' active' : ''}" role="option" aria-selected="${i === suggest.active}" data-header="${escapeHtml(header)}">{{${escapeHtml(header)}}}</div>`).join('');
  box.classList.remove('hidden');
}

function updateSuggestions(field) {
  if (!field || !$('variableSuggest')) return;
  const query = variableQueryAt(field.value, field.selectionStart ?? field.value.length);
  const items = query ? variableSuggestions(csvHeaders, query.query) : [];
  if (!items.length) return closeSuggestions();
  const active = suggest?.fieldId === field.id ? Math.min(suggest.active, items.length - 1) : 0;
  suggest = { fieldId: field.id, start: query.start, items, active };
  // Show the list right under the field being typed in.
  if ($('variableSuggest').previousElementSibling !== field) field.insertAdjacentElement?.('afterend', $('variableSuggest'));
  renderSuggestions();
}

function chooseSuggestion(header) {
  const field = suggest && $(suggest.fieldId);
  if (!field) return;
  const result = applySuggestion(field.value, suggest.start, field.selectionStart ?? field.value.length, header);
  closeSuggestions();
  field.value = result.text;
  field.focus();
  field.setSelectionRange(result.cursor, result.cursor);
  field.dispatchEvent(new Event('input', { bubbles: true }));
}

// --- Live warnings for variables that are not in the list ---

function normalizeName(name) {
  return String(name).toLowerCase().replace(/[\s_-]+/g, '');
}

function editDistance(a, b) {
  let previous = Array.from({ length: b.length + 1 }, (_, j) => j);
  for (let i = 1; i <= a.length; i++) {
    const current = [i];
    for (let j = 1; j <= b.length; j++) {
      current[j] = Math.min(previous[j] + 1, current[j - 1] + 1, previous[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    previous = current;
  }
  return previous[b.length];
}

function closestHeader(name, headers) {
  const target = normalizeName(name);
  let best = null;
  let bestScore = Infinity;
  headers.filter(Boolean).forEach(header => {
    const score = editDistance(target, normalizeName(header));
    if (score < bestScore) {
      best = header;
      bestScore = score;
    }
  });
  return bestScore <= Math.max(1, Math.floor(target.length / 4)) ? best : null;
}

function renderTemplateWarnings() {
  const box = $('templateWarnings');
  if (!box) return;
  const missing = csvHeaders.length ? findMissingTemplateVariables(csvHeaders, currentTemplates()) : [];
  box.classList.toggle('hidden', !missing.length);
  box.textContent = missing.map(name => {
    const suggestion = closestHeader(name, csvHeaders);
    return suggestion ? t('templateUnknownSuggest', { name, suggestion }) : t('templateUnknown', { name });
  }).join('\n');
}

// --- Sample data for a first try ---

const SAMPLE_CSV = [
  'FirstName,LastName,Email,Course',
  'Ada,Lovelace,ada.lovelace@example.edu,ISOM 210',
  'Grace,Hopper,grace.hopper@example.edu,ISOM 340',
  'Bob,Smith,robert.smith@example.edu,ISOM 210',
  ',Turing,alan.turing@example.edu,ISOM 340',
  'Katherine,Johnson,katherine.johnson@example.edu,ISOM 210',
  'Katherine,Johnson,katherine.johnson@example.edu,ISOM 340',
  'Margaret,Hamilton,margaret.hamilton@gmail.com,ISOM 210'
].join('\n');

async function loadSampleData() {
  if (isOpening) return;
  if (rows.length && typeof confirm === 'function' && !confirm(t('sampleConfirm'))) return;
  await applyPlan({
    name: '',
    csv: SAMPLE_CSV,
    subject: t('sampleSubject'),
    cc: '',
    bcc: '',
    body: t('sampleBody'),
    composeMode: $('composeMode')?.value === 'gmail' ? 'gmail' : 'outlook',
    batchDelay: '2000',
    batchSize: 3
  });
  if (await generateDrafts([t('sampleLoaded', { count: rows.length })])) setWizardStep(3);
}

// --- Step-by-step mode ---

// Steps unlock in order: a list is needed for the email step, and drafts for review and send.
function maxReachableStep() {
  if (generated.length) return 4;
  if (rows.length && !mapping) return 2;
  return 1;
}

function renderWizard() {
  const app = $('app');
  if (!app) return;
  const reachable = maxReachableStep();
  wizardStep = Math.min(Math.max(1, wizardStep), reachable);
  app.classList.toggle('wizard', wizardOn);
  if (app.dataset) app.dataset.currentStep = String(wizardStep);

  for (let n = 1; n <= 4; n++) {
    const tab = $(`stepTab${n}`);
    if (!tab) continue;
    tab.classList.toggle('active', n === wizardStep);
    tab.classList.toggle('done', n < wizardStep);
    tab.disabled = n > reachable;
    if (n === wizardStep) tab.setAttribute?.('aria-current', 'step');
    else tab.removeAttribute?.('aria-current');
  }
  if ($('wizardBack')) $('wizardBack').classList.toggle('hidden', wizardStep === 1);
  if ($('wizardNext')) {
    $('wizardNext').classList.toggle('hidden', wizardStep === 4);
    $('wizardNext').textContent = t(`wizardNext${Math.min(wizardStep, 3)}`);
    $('wizardNext').disabled = isOpening || (wizardStep === 1 && reachable < 2);
  }
  if ($('wizardHint')) $('wizardHint').textContent = wizardStep === 1 && reachable < 2 ? t('wizardNeedList') : '';
  if ($('wizardToggle')) $('wizardToggle').textContent = wizardOn ? t('showAllSteps') : t('showStepByStep');
}

function setWizardStep(step) {
  wizardStep = step;
  closeSuggestions();
  renderWizard();
  saveState();
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
    button.title = t('insertVariable', { name: header });
    button.addEventListener('click', () => {
      insertAtCursor($(lastTemplateFieldId), `{{${header}}}`);
    });
    container.appendChild(button);
  });
  renderTemplateWarnings();
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
      <div class="preview-field"><span class="preview-field-label">${escapeHtml(t('previewTo'))}</span><span class="preview-field-val preview-tag">${escapeHtml(msg.email)}</span></div>
      ${ccHtml}
      ${bccHtml}
      <div class="preview-field" style="margin-top:6px;"><span class="preview-field-label">${escapeHtml(t('previewSubject'))}</span><span class="preview-field-val" style="font-weight:600;">${escapeHtml(msg.subject)}</span></div>
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

function setRangeOpening(opening) {
  isOpening = opening;
  renderWizard();
  for (const id of ['openRangeBtn', 'openBatchBtn', 'openNextBtn', 'sendTestBtn']) {
    if ($(id)) $(id).disabled = opening;
  }
  if (opening && $('openRangeBtn')) $('openRangeBtn').textContent = t('openingButton');
  if (!opening) syncDraftActionLabels();
}

function runPreflight() {
  return [
    ...findMissingTemplateVariables(csvHeaders, currentTemplates()).map(name => t('unknownVariable', { name })),
    ...validateMessages(generated)
  ];
}

// Rebuilds the messages from the current template and reports blocking problems, if any.
function readyToOpen() {
  if (!generated.length) {
    renderStatus(t('generateFirst'), 'error');
    return false;
  }
  makeMessages();
  const issues = runPreflight();
  if (issues.length) {
    renderStatus(tm('fixPreflight') + renderPreflightIssues(issues), 'error');
    return false;
  }
  return true;
}

function getSelectedRange() {
  const total = generated.length;
  const start = Number.parseInt($('rangeStart')?.value, 10);
  const end = Number.parseInt($('rangeEnd')?.value, 10);

  if (!Number.isInteger(start) || !Number.isInteger(end)) {
    throw new Error(t('rangeInvalid'));
  }
  if (start < 1 || end < 1 || start > total || end > total) {
    throw new Error(t('rangeBounds', { total }));
  }
  if (start > end) {
    throw new Error(t('rangeOrder'));
  }

  return { startIndex: start - 1, endIndex: end - 1 };
}

// --- Batches ---

function getBatchSize() {
  const size = Number.parseInt($('batchSize')?.value, 10);
  return Number.isInteger(size) && size >= 1 ? Math.min(size, MAX_BATCH_SIZE) : DEFAULT_BATCH_SIZE;
}

function openedCount() {
  return [...openedRows].filter(index => index < generated.length).length;
}

// The next drafts to open, in list order, skipping any that were already opened.
function nextBatchIndices(total, size) {
  const indices = [];
  for (let i = 0; i < total && indices.length < size; i++) {
    if (!openedRows.has(i)) indices.push(i);
  }
  return indices;
}

function formatTime(timestamp) {
  return new Date(timestamp).toLocaleTimeString(uiLocale(), { hour: 'numeric', minute: '2-digit' });
}

function renderBatchControls() {
  const total = generated.length;
  const size = getBatchSize();
  const done = openedCount();
  const next = nextBatchIndices(total, size);

  if ($('batchProgress')) $('batchProgress').textContent = total ? t('batchProgress', { done, total }) : '';
  if ($('openBatchBtn')) {
    $('openBatchBtn').disabled = isOpening || !next.length;
    $('openBatchBtn').textContent = next.length
      ? tm('openBatch', {
        n: Math.floor(done / size) + 1,
        count: Math.max(1, Math.ceil(total / size)),
        start: next[0] + 1,
        end: next[next.length - 1] + 1
      })
      : t('allBatchesDone');
  }
  if ($('lastBatchNote')) {
    $('lastBatchNote').textContent = lastBatch ? t('lastBatchNote', { start: lastBatch.start, end: lastBatch.end, time: formatTime(lastBatch.at) }) : '';
  }
  if ($('resetProgressBtn')) $('resetProgressBtn').classList.toggle('hidden', done === 0);
}

async function openDraftRange(indices) {
  if (isOpening || !indices.length) return;
  setRangeOpening(true);
  const start = indices[0] + 1;
  const end = indices[indices.length - 1] + 1;
  const delayMs = Number.parseInt($('batchDelay')?.value || `${DEFAULT_DELAY_MS}`, 10) || DEFAULT_DELAY_MS;
  let openedNow = 0;

  const progressContainer = $('progressContainer');
  const progressBar = $('progressBar');
  if (progressContainer) progressContainer.classList.remove('hidden');
  if (progressBar) progressBar.style.width = '0%';

  try {
    for (const [n, index] of indices.entries()) {
      renderStatus(tm('openingProgress', { n: n + 1, total: indices.length }), 'ok');
      await openDraftTab(generated[index]);
      openedRows.add(index);
      openedNow += 1;
      // Save after every tab, so closing the popup mid-batch never marks unopened drafts as done.
      await saveState();

      if (progressBar) {
        const percent = Math.round((openedNow / indices.length) * 100);
        progressBar.style.width = `${percent}%`;
      }

      if (openedNow < indices.length) {
        await sleep(delayMs);
      }
    }
    lastBatch = { start, end, at: Date.now() };
    previewIndex = clampPreviewIndex(nextBatchIndices(generated.length, 1)[0] ?? previewIndex);
    await saveState();
    renderPreview();
    renderStatus(tm('rangeDone', { count: indices.length, start, end, total: generated.length }), 'ok');
  } catch (e) {
    await saveState();
    renderStatus(tm('rangeStopped', { opened: openedNow, count: indices.length, error: e.message }), 'error');
  } finally {
    setRangeOpening(false);
    syncRangeInputs();
    setTimeout(() => {
      if (progressContainer) progressContainer.classList.add('hidden');
    }, 1500);
  }
}

// Builds the drafts from the list and template. Returns true when there are no blocking problems.
async function generateDrafts(extraLines = []) {
  try {
    const restart = listEdited;
    await loadRows({ fresh: restart });
    makeMessages();
    const issues = runPreflight();
    // Regenerating after a template edit keeps progress, so a batch is never opened twice.
    previewIndex = clampPreviewIndex(previewIndex);
    await saveState();
    if ($('sendControls')) $('sendControls').classList.remove('hidden');
    if ($('previewControls')) $('previewControls').classList.remove('hidden');
    syncRangeInputs();
    syncDraftActionLabels();
    renderPreview(t('previewLoaded', { count: generated.length }));
    renderReviewPanel();
    const notes = reviewRecipients(rows, csvHeaders, currentTemplates());
    const lines = [...extraLines, tm('ready', { count: generated.length })];
    if (openedCount() > 0) lines.push(t('progressKept', { done: openedCount(), total: generated.length }));
    if (notes.length) lines.push(t('readyReviewNote', { count: notes.length }));
    renderStatus(lines.join('\n') + renderPreflightIssues(issues), issues.length ? 'error' : 'ok');
    renderWizard();
    return !issues.length;
  } catch (e) {
    renderStatus(e.message, 'error');
    return false;
  }
}

// --- Test email to yourself ---

function buildTestMessage(msg, testEmail) {
  return { ...msg, email: testEmail, cc: '', bcc: '', subject: `[TEST] ${msg.subject}` };
}

// --- Excel import with a column picker ---

function uniqueHeaders(headerRow, width) {
  const used = new Set();
  return Array.from({ length: width }, (_, i) => {
    const base = String(headerRow[i] ?? '').trim() || t('columnFallback', { n: i + 1 });
    let name = base;
    // Check every name already given out, so "Name, Name, Name (2)" cannot produce two "Name (2)".
    for (let n = 2; used.has(name); n++) name = `${base} (${n})`;
    used.add(name);
    return name;
  });
}

function sheetTable(sheet) {
  const nonEmpty = sheet.rows.filter(row => row.some(cell => String(cell).trim() !== ''));
  const width = Math.max(0, ...nonEmpty.map(row => row.length));
  return { headers: uniqueHeaders(nonEmpty[0] || [], width), dataRows: nonEmpty.slice(1) };
}

function guessEmailColumn(headers, dataRows) {
  const byName = detectEmailHeader(headers);
  if (byName) return headers.indexOf(byName);
  // Otherwise pick the column where most values look like email addresses.
  let best = -1;
  let bestCount = 0;
  headers.forEach((_, col) => {
    const count = dataRows.filter(row => isValidEmail(String(row[col] ?? '').trim())).length;
    if (count > bestCount) {
      best = col;
      bestCount = count;
    }
  });
  return best;
}

function mappingToCSV(headers, dataRows, emailColumn, columns) {
  const order = [...columns]
    .filter(col => col !== emailColumn && headers[col].toLowerCase() !== 'email')
    .sort((a, b) => a - b);
  const outRows = dataRows
    .map(row => [...order.map(col => row[col] ?? ''), row[emailColumn] ?? ''])
    .filter(cells => String(cells[cells.length - 1]).trim() !== '');
  return [[...order.map(col => headers[col]), 'Email'], ...outRows]
    .map(cells => cells.map(csvCell).join(','))
    .join('\n');
}

function mappingTable() {
  return sheetTable(mapping.sheets[mapping.sheetIndex]);
}

function resetMappingColumns() {
  const { headers, dataRows } = mappingTable();
  mapping.emailColumn = guessEmailColumn(headers, dataRows);
  // Start with name columns only; anything else is opt-in.
  mapping.columns = new Set(headers.map((_, i) => i).filter(i => i !== mapping.emailColumn && NAME_COLUMN_PATTERN.test(headers[i])));
}

function openMapping(fileName, sheets) {
  const sheetIndex = sheets.reduce((best, sheet, i) => (sheet.rows.length > sheets[best].rows.length ? i : best), 0);
  mapping = { fileName, sheets, sheetIndex, emailColumn: -1, columns: new Set() };
  resetMappingColumns();
  renderMappingPanel();
}

function renderMappingPanel() {
  const panel = $('mappingPanel');
  if (!panel) return;
  updateCsvSummary();
  renderWizard();
  if (!mapping) {
    panel.classList.add('hidden');
    panel.innerHTML = '';
    return;
  }

  const { headers, dataRows } = mappingTable();
  const sample = col => String(dataRows.find(row => String(row[col] ?? '').trim() !== '')?.[col] ?? '').replace(/\s+/g, ' ').slice(0, 40);
  const withEmail = mapping.emailColumn >= 0 ? dataRows.filter(row => String(row[mapping.emailColumn] ?? '').trim() !== '').length : 0;
  const sheetPicker = mapping.sheets.length > 1
    ? `<label for="mappingSheet">${escapeHtml(t('mappingSheet'))}</label>
       <select id="mappingSheet" data-mapping="sheet">${mapping.sheets.map((sheet, i) => `<option value="${i}"${i === mapping.sheetIndex ? ' selected' : ''}>${escapeHtml(sheet.name)}</option>`).join('')}</select>`
    : '';

  panel.classList.remove('hidden');
  panel.innerHTML = `
    <div class="mapping-title">${escapeHtml(t('mappingTitle', { name: mapping.fileName }))}</div>
    ${sheetPicker}
    <label for="mappingEmail">${escapeHtml(t('mappingEmail'))}</label>
    <select id="mappingEmail" data-mapping="email">
      ${mapping.emailColumn < 0 ? '<option value="-1" selected>—</option>' : ''}
      ${headers.map((header, i) => `<option value="${i}"${i === mapping.emailColumn ? ' selected' : ''}>${escapeHtml(header)}</option>`).join('')}
    </select>
    <label id="mappingColumnsLabel">${escapeHtml(t('mappingColumns'))}</label>
    <p class="hint">${escapeHtml(t('mappingColumnsHint'))}</p>
    <div class="mapping-columns" role="group" aria-labelledby="mappingColumnsLabel">
      ${headers.map((header, i) => (i === mapping.emailColumn ? '' : `
        <label class="mapping-column">
          <input type="checkbox" data-mapping="column" value="${i}"${mapping.columns.has(i) ? ' checked' : ''} />
          <span class="mapping-column-name">${escapeHtml(header)}</span>
          <span class="mapping-sample">${escapeHtml(sample(i))}</span>
        </label>`)).join('')}
    </div>
    <p class="hint">${escapeHtml(t('mappingCount', { count: withEmail, total: dataRows.length }))}</p>
    <div class="mapping-actions">
      <button type="button" class="mapping-apply" data-mapping="apply">${escapeHtml(t('mappingApply'))}</button>
      <button type="button" class="mapping-cancel" data-mapping="cancel">${escapeHtml(t('mappingCancel'))}</button>
    </div>
  `;
}

async function applyMapping() {
  if (!mapping) return;
  if (mapping.emailColumn < 0) return renderStatus(t('mappingNoEmail'), 'error');
  const { headers, dataRows } = mappingTable();
  const fileName = mapping.fileName;
  try {
    if ($('csvText')) $('csvText').value = mappingToCSV(headers, dataRows, mapping.emailColumn, mapping.columns);
    await loadRows({ fresh: true });
    mapping = null;
    renderMappingPanel();
    renderStatus(t('xlsxLoaded', { name: fileName, count: rows.length }), 'ok');
  } catch (err) {
    renderStatus(err.message, 'error');
  }
}

async function handleRecipientFile(file) {
  if (!file) return;
  try {
    if (/\.json$/i.test(file.name)) {
      await importPlanFile(file);
    } else if (/\.xlsx$/i.test(file.name)) {
      let sheets;
      try {
        sheets = (await readXlsx(await file.arrayBuffer())).filter(sheet => sheet.rows.some(row => row.some(cell => String(cell).trim() !== '')));
      } catch {
        throw new Error(t('xlsxInvalid'));
      }
      if (!sheets.length) throw new Error(t('xlsxEmpty'));
      openMapping(file.name, sheets);
      renderStatus(t('xlsxPickColumns', { name: file.name }), 'ok');
    } else if (/\.xls$/i.test(file.name)) {
      throw new Error(t('xlsOldFormat'));
    } else {
      if ($('csvText')) $('csvText').value = await file.text();
      mapping = null;
      renderMappingPanel();
      await loadRows({ fresh: true });
      renderStatus(t('csvLoaded', { name: file.name, count: rows.length }), 'ok');
    }
  } catch (err) {
    renderStatus(err.message, 'error');
  } finally {
    // Allow choosing the same file again.
    if ($('csvFile')) $('csvFile').value = '';
  }
}

// --- Email plans ---
// Plan files bundle the recipient CSV, templates, and compose settings.
// Auto-Send is never stored in a plan and is always turned off on import.
function buildPlan() {
  const csv = getCsvText();
  if (!csv) throw new Error(t('planExportNoCsv'));
  return {
    type: PLAN_TYPE,
    version: PLAN_VERSION,
    exportedAt: new Date().toISOString(),
    csv,
    subject: $('subject')?.value ?? '',
    cc: $('cc')?.value ?? '',
    bcc: $('bcc')?.value ?? '',
    body: $('body')?.value ?? '',
    settings: {
      composeMode: $('composeMode')?.value ?? 'outlook',
      batchDelay: Number.parseInt($('batchDelay')?.value || `${DEFAULT_DELAY_MS}`, 10) || DEFAULT_DELAY_MS,
      batchSize: getBatchSize()
    }
  };
}

function parsePlan(text) {
  let plan;
  try {
    plan = JSON.parse(String(text ?? '').replace(/^﻿/, ''));
  } catch {
    throw new Error(t('planInvalidJson'));
  }
  if (!plan || typeof plan !== 'object' || plan.type !== PLAN_TYPE) {
    throw new Error(t('planWrongType'));
  }
  if (plan.version !== PLAN_VERSION) {
    throw new Error(t('planVersion', { version: plan.version }));
  }

  // Text fields may also be written as an array of lines, which is easier to edit by hand.
  const textField = key => {
    const value = plan[key] ?? '';
    if (Array.isArray(value) && value.every(line => typeof line === 'string')) return value.join('\n');
    if (typeof value !== 'string') throw new Error(t('planFieldText', { key }));
    return value;
  };

  const csv = textField('csv');
  if (!csv.trim()) throw new Error(t('planNoCsv'));
  const settings = plan.settings && typeof plan.settings === 'object' ? plan.settings : {};
  const batchDelay = String(settings.batchDelay ?? DEFAULT_DELAY_MS);
  const batchSize = Number(settings.batchSize);

  return {
    name: typeof plan.name === 'string' ? plan.name.trim() : '',
    csv,
    subject: textField('subject'),
    cc: textField('cc'),
    bcc: textField('bcc'),
    body: textField('body'),
    composeMode: settings.composeMode === 'gmail' ? 'gmail' : 'outlook',
    batchDelay: BATCH_DELAY_OPTIONS.includes(batchDelay) ? batchDelay : String(DEFAULT_DELAY_MS),
    batchSize: Number.isInteger(batchSize) && batchSize >= 1 && batchSize <= MAX_BATCH_SIZE ? batchSize : DEFAULT_BATCH_SIZE
  };
}

async function applyPlan(plan) {
  const previousHeaders = csvHeaders;
  const planRows = parseCSV(plan.csv);
  if (!planRows.length) {
    csvHeaders = previousHeaders;
    throw new Error(t('planNoRows'));
  }

  rows = planRows;
  dismissedNotes = [];
  listEdited = false;
  mapping = null;
  resetProgress();

  // Clear a previously chosen CSV file so it cannot be mistaken for the plan's list.
  if ($('csvFile')) $('csvFile').value = '';
  if ($('csvText')) $('csvText').value = plan.csv;
  if ($('subject')) $('subject').value = plan.subject;
  if ($('cc')) $('cc').value = plan.cc;
  if ($('bcc')) $('bcc').value = plan.bcc;
  if ($('body')) $('body').value = plan.body;
  if ($('templatePreset')) $('templatePreset').value = '';
  if ($('composeMode')) $('composeMode').value = plan.composeMode;
  if ($('batchDelay')) $('batchDelay').value = plan.batchDelay;
  if ($('batchSize')) $('batchSize').value = String(plan.batchSize);
  if ($('autoSend')) $('autoSend').checked = false;

  updateCcBccVisibility(Boolean(plan.cc || plan.bcc));
  updateCsvSummary();
  renderVariableButtons();
  renderMappingPanel();
  renderComposeModeHints();
  syncDraftActionLabels();
  renderWizard();
  await saveState();
  return rows.length;
}

async function importPlanFile(file) {
  try {
    const plan = parsePlan(await file.text());
    const count = await applyPlan(plan);
    renderStatus(plan.name ? t('planImportedNamed', { name: plan.name, count }) : t('planImportedFile', { name: file.name, count }), 'ok');
  } catch (err) {
    renderStatus(err.message, 'error');
  }
}

function planFileName(date = new Date()) {
  const pad = value => String(value).padStart(2, '0');
  return `mail-merge-plan-${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}.json`;
}

function downloadPlan(plan) {
  const blob = new Blob([JSON.stringify(plan, null, 2) + '\n'], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  const fileName = planFileName();
  link.href = url;
  link.download = fileName;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  return fileName;
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
      updateSuggestions($(id));
      renderTemplateWarnings();
      saveState();
    });
    $(id).addEventListener('keydown', (event) => {
      if (!suggest || suggest.fieldId !== id) return;
      const count = suggest.items.length;
      if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
        event.preventDefault();
        suggest.active = (suggest.active + (event.key === 'ArrowDown' ? 1 : -1) + count) % count;
        renderSuggestions();
      } else if (event.key === 'Enter' || event.key === 'Tab') {
        event.preventDefault();
        chooseSuggestion(suggest.items[suggest.active]);
      } else if (event.key === 'Escape') {
        event.preventDefault();
        closeSuggestions();
      }
    });
    $(id).addEventListener('blur', closeSuggestions);
  }
});

if ($('variableSuggest')) {
  // mousedown (not click) so the field keeps focus while a suggestion is picked.
  $('variableSuggest').addEventListener('mousedown', (event) => {
    const header = event.target?.closest?.('[data-header]')?.dataset.header ?? event.target?.dataset?.header;
    if (!header) return;
    event.preventDefault();
    chooseSuggestion(header);
  });
}

if ($('tryDemoBtn')) {
  $('tryDemoBtn').addEventListener('click', () => loadSampleData());
}

if ($('uiLanguage')) {
  $('uiLanguage').addEventListener('change', async () => {
    applyLanguageSelection($('uiLanguage').value);
    await saveState();
  });
}

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
    updateCcBccVisibility($('ccBccContainer')?.classList.contains('hidden'));
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
    const file = e.dataTransfer?.files?.[0];
    if (file) await handleRecipientFile(file);
  });
}

if ($('csvFile')) {
  $('csvFile').addEventListener('change', async () => {
    await handleRecipientFile($('csvFile').files?.[0]);
  });
}

if ($('mappingPanel')) {
  $('mappingPanel').addEventListener('change', (event) => {
    const kind = event.target?.dataset?.mapping;
    if (!mapping || !kind) return;
    if (kind === 'sheet') {
      mapping.sheetIndex = Number(event.target.value);
      resetMappingColumns();
    } else if (kind === 'email') {
      mapping.emailColumn = Number(event.target.value);
      mapping.columns.delete(mapping.emailColumn);
    } else if (kind === 'column') {
      const col = Number(event.target.value);
      if (event.target.checked) mapping.columns.add(col);
      else mapping.columns.delete(col);
      // Leave the list as it is so the scroll position stays put.
      return;
    }
    renderMappingPanel();
  });
  $('mappingPanel').addEventListener('click', async (event) => {
    const kind = event.target?.dataset?.mapping;
    if (kind === 'apply') await applyMapping();
    if (kind === 'cancel') {
      mapping = null;
      renderMappingPanel();
      renderStatus(t('mappingCancelled'), 'ok');
    }
  });
}

if ($('reviewPanel')) {
  $('reviewPanel').addEventListener('change', async (event) => {
    const { row, field } = event.target?.dataset || {};
    if (row === undefined || !field) return;
    if (await updateRecipientField(Number(row), field, event.target.value)) {
      renderStatus(t('recipientUpdated', { n: Number(row) + 1 }), 'ok');
    }
  });
  $('reviewPanel').addEventListener('click', async (event) => {
    const key = event.target?.dataset?.dismiss;
    if (!key) return;
    dismissedNotes = [...new Set([...dismissedNotes, key])];
    renderReviewPanel();
    await saveState();
  });
}

if ($('importPlanBtn')) {
  $('importPlanBtn').addEventListener('click', () => {
    $('planFile')?.click();
  });
}

if ($('planFile')) {
  $('planFile').addEventListener('change', async () => {
    const file = $('planFile').files?.[0];
    if (!file) return;
    await importPlanFile(file);
    $('planFile').value = '';
  });
}

if ($('exportPlanBtn')) {
  $('exportPlanBtn').addEventListener('click', () => {
    try {
      const fileName = downloadPlan(buildPlan());
      renderStatus(t('planExported', { name: fileName }), 'ok');
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

if ($('batchSize')) {
  $('batchSize').addEventListener('change', () => {
    renderBatchControls();
    saveState();
  });
}

if ($('testEmail')) {
  $('testEmail').addEventListener('input', () => {
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
    renderTemplateWarnings();
    saveState();
  });
}

if ($('csvText')) {
  $('csvText').addEventListener('input', () => {
    listEdited = true;
    saveState();
  });
}

if ($('clearStateBtn')) {
  $('clearStateBtn').addEventListener('click', async () => {
    if (rows.length && typeof confirm === 'function') {
      if (!confirm(t('confirmClear'))) return;
    }
    await chrome.storage.local.remove(STORAGE_KEY);
    rows = [];
    csvHeaders = [];
    dismissedNotes = [];
    listEdited = false;
    mapping = null;
    resetProgress();
    if ($('csvText')) $('csvText').value = '';
    if ($('composeMode')) $('composeMode').value = 'outlook';
    if ($('autoSend')) $('autoSend').checked = false;
    if ($('templatePreset')) $('templatePreset').value = '';
    if ($('batchDelay')) $('batchDelay').value = '1000';
    if ($('batchSize')) $('batchSize').value = String(DEFAULT_BATCH_SIZE);
    if ($('testEmail')) $('testEmail').value = '';
    if ($('subject')) $('subject').value = TEMPLATE_PRESETS.courseReminder.subject;
    if ($('cc')) $('cc').value = '';
    if ($('bcc')) $('bcc').value = '';
    if ($('body')) $('body').value = TEMPLATE_PRESETS.courseReminder.body;
    updateCcBccVisibility(false);
    updateCsvSummary();
    renderVariableButtons();
    renderMappingPanel();
    renderComposeModeHints();
    syncDraftActionLabels();
    wizardStep = 1;
    renderWizard();
    renderStatus(t('cleared'), 'ok');
    // Keep the chosen language after clearing.
    await saveState();
  });
}

if ($('loadBtn')) {
  $('loadBtn').addEventListener('click', () => generateDrafts());
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

if ($('sendTestBtn')) {
  $('sendTestBtn').addEventListener('click', async () => {
    if (isOpening) return;
    if (!generated.length) return renderStatus(t('generateFirst'), 'error');
    const testEmail = ($('testEmail')?.value ?? '').trim();
    if (!isValidEmail(testEmail)) return renderStatus(t('testEmailInvalid'), 'error');
    makeMessages();
    previewIndex = clampPreviewIndex(previewIndex);
    try {
      await openDraftTab(buildTestMessage(generated[previewIndex], testEmail));
      await saveState();
      renderStatus(tm('testOpened', { n: previewIndex + 1, email: testEmail }), 'ok');
    } catch (e) {
      renderStatus(t('testFailed', { error: e.message }), 'error');
    }
  });
}

if ($('openBatchBtn')) {
  $('openBatchBtn').addEventListener('click', async () => {
    if (isOpening || !readyToOpen()) return;
    const indices = nextBatchIndices(generated.length, getBatchSize());
    if (!indices.length) return renderStatus(t('allBatchesDone'), 'ok');
    await openDraftRange(indices);
  });
}

if ($('resetProgressBtn')) {
  $('resetProgressBtn').addEventListener('click', async () => {
    if (isOpening) return;
    openedRows = new Set();
    lastBatch = null;
    syncRangeInputs();
    await saveState();
    renderStatus(t('progressReset'), 'ok');
  });
}

if ($('openNextBtn')) {
  $('openNextBtn').addEventListener('click', async () => {
    if (isOpening || !readyToOpen()) return;
    previewIndex = clampPreviewIndex(previewIndex);
    const msg = generated[previewIndex];

    try {
      await openDraftTab(msg);
      openedRows.add(previewIndex);
      await saveState();
      syncRangeInputs();
      syncDraftActionLabels();
      renderPreview();
      renderStatus(tm('openedOne', { n: previewIndex + 1, total: generated.length, email: msg.email }), 'ok');
    } catch (e) {
      syncRangeInputs();
      renderStatus(tm('openFailed', { error: e.message }), 'error');
    }
  });
}

if ($('openRangeBtn')) {
  $('openRangeBtn').addEventListener('click', async () => {
    if (isOpening || !readyToOpen()) return;
    let range;
    try {
      range = getSelectedRange();
    } catch (e) {
      return renderStatus(e.message, 'error');
    }
    await openDraftRange(Array.from({ length: range.endIndex - range.startIndex + 1 }, (_, k) => range.startIndex + k));
  });
}

if ($('wizardNext')) {
  $('wizardNext').addEventListener('click', async () => {
    if (isOpening) return;
    // Moving on from the email step builds the drafts; stay put if something blocks sending.
    if (wizardStep === 2) {
      if (await generateDrafts()) setWizardStep(3);
      return;
    }
    setWizardStep(wizardStep + 1);
  });
}

if ($('wizardBack')) {
  $('wizardBack').addEventListener('click', () => setWizardStep(wizardStep - 1));
}

for (let n = 1; n <= 4; n++) {
  if ($(`stepTab${n}`)) $(`stepTab${n}`).addEventListener('click', () => setWizardStep(n));
}

if ($('wizardToggle')) {
  $('wizardToggle').addEventListener('click', () => {
    wizardOn = !wizardOn;
    renderWizard();
    saveState();
  });
}

// Initial restoration
applyLanguageSelection(detectUiLanguage());
restoreState();
