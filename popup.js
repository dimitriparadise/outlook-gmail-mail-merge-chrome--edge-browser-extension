let rows = [];
let currentIndex = 0;
let generated = [];
let csvHeaders = [];
let lastTemplateFieldId = 'body';

const $ = (id) => document.getElementById(id);
const STORAGE_KEY = 'outlookMailMergeState';
const TEMPLATE_FIELD_IDS = ['subject', 'cc', 'bcc', 'body'];

function parseCSV(text) {
  const lines = text.replace(/^\uFEFF/, '').split(/\r?\n/).filter(l => l.trim() !== '');
  if (lines.length < 2) throw new Error('CSV needs a header row and at least one data row.');

  const parseLine = (line) => {
    const result = [];
    let cur = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      const next = line[i + 1];
      if (ch === '"' && inQuotes && next === '"') { cur += '"'; i++; }
      else if (ch === '"') inQuotes = !inQuotes;
      else if (ch === ',' && !inQuotes) { result.push(cur.trim()); cur = ''; }
      else cur += ch;
    }
    result.push(cur.trim());
    return result;
  };

  const headers = parseLine(lines[0]).map(h => h.trim());
  csvHeaders = headers;
  return lines.slice(1).map(line => {
    const values = parseLine(line);
    const obj = {};
    headers.forEach((h, i) => obj[h] = values[i] || '');
    return obj;
  }).filter(r => r.Email || r.email);
}

function applyTemplate(template, row) {
  return template.replace(/{{\s*([^}]+?)\s*}}/g, (_, key) => row[key] ?? row[key.trim()] ?? '');
}

function getCsvText() {
  return $('csvText').value.trim();
}

async function saveState() {
  const state = {
    csvText: $('csvText').value,
    subject: $('subject').value,
    cc: $('cc').value,
    bcc: $('bcc').value,
    body: $('body').value,
    rows,
    csvHeaders,
    generated,
    currentIndex,
    hasStarted: generated.length > 0,
    updatedAt: Date.now()
  };
  await chrome.storage.local.set({ [STORAGE_KEY]: state });
}

async function restoreState() {
  const result = await chrome.storage.local.get(STORAGE_KEY);
  const state = result[STORAGE_KEY];
  if (!state) return;

  $('csvText').value = state.csvText || $('csvText').value;
  $('subject').value = state.subject || $('subject').value;
  $('cc').value = state.cc || '';
  $('bcc').value = state.bcc || '';
  $('body').value = state.body || $('body').value;
  rows = Array.isArray(state.rows) ? state.rows : [];
  csvHeaders = Array.isArray(state.csvHeaders) ? state.csvHeaders : getHeadersFromRows(rows);
  generated = Array.isArray(state.generated) ? state.generated : [];
  currentIndex = Number.isInteger(state.currentIndex) ? state.currentIndex : 0;
  renderVariableButtons();

  if (generated.length) {
    $('sendControls').classList.remove('hidden');
    renderStatus(`Restored: ${currentIndex}/${generated.length} drafts opened. Click “Open Next Draft” to continue.`, 'ok');
    const next = generated[Math.min(currentIndex, generated.length - 1)];
    if (next) {
      $('preview').textContent = messagePreview('Next draft preview:', next);
    }
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

  return 'https://outlook.office.com/mail/deeplink/compose'
    + '?' + params.map(([key, value]) => key + '=' + encodeForOutlook(value)).join('&');
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

function renderStatus(text, cls = '') {
  $('status').className = cls;
  $('status').textContent = text;
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

TEMPLATE_FIELD_IDS.forEach(id => {
  $(id).addEventListener('focus', () => {
    lastTemplateFieldId = id;
  });
});

['csvText', 'subject', 'cc', 'bcc', 'body'].forEach(id => {
  $(id).addEventListener('input', () => {
    // If templates/CSV change, keep typed content but require Start Merge again to regenerate drafts.
    saveState();
  });
});

$('previewBtn').addEventListener('click', async () => {
  try {
    await loadRows();
    makeMessages();
    await saveState();
    const first = generated[0];
    $('preview').textContent = `Loaded ${generated.length} rows.\n\n` + messagePreview('First preview:', first);
    renderStatus('Preview ready.', 'ok');
  } catch (e) {
    renderStatus(e.message, 'error');
  }
});

$('startBtn').addEventListener('click', async () => {
  try {
    await loadRows();
    makeMessages();
    currentIndex = 0;
    await saveState();
    $('sendControls').classList.remove('hidden');
    renderStatus(`Ready: ${generated.length} personalized drafts. Click “Open Next Draft.”`, 'ok');
  } catch (e) {
    renderStatus(e.message, 'error');
  }
});

$('openNextBtn').addEventListener('click', async () => {
  if (!generated.length) return renderStatus('Please start merge first.', 'error');
  if (currentIndex >= generated.length) return renderStatus('All drafts opened.', 'ok');

  const msg = generated[currentIndex];
  const openedIndex = currentIndex + 1;
  currentIndex = openedIndex;
  await saveState();

  chrome.tabs.create({ url: outlookComposeUrl(msg), active: false }, async () => {
    if (chrome.runtime.lastError) {
      currentIndex = openedIndex - 1;
      await saveState();
      renderStatus(`Could not open draft: ${chrome.runtime.lastError.message}`, 'error');
      return;
    }

    renderStatus(`Opened draft ${currentIndex}/${generated.length} in a background tab: ${msg.email}\n\nKeep clicking “Open Next Draft” to queue the rest, or open the new Outlook tabs to review and send.`, 'ok');
  });
});

$('copyBodyBtn').addEventListener('click', async () => {
  if (!generated.length) return renderStatus('Please start merge first.', 'error');
  const idx = Math.min(currentIndex, generated.length - 1);
  await navigator.clipboard.writeText(generated[idx].body);
  renderStatus(`Copied body for row ${idx + 1}.`, 'ok');
});

restoreState();
