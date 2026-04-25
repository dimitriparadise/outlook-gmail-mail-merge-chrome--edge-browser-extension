let rows = [];
let currentIndex = 0;
let generated = [];

const $ = (id) => document.getElementById(id);
const STORAGE_KEY = 'outlookMailMergeState';

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
    body: $('body').value,
    rows,
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
  $('body').value = state.body || $('body').value;
  rows = Array.isArray(state.rows) ? state.rows : [];
  generated = Array.isArray(state.generated) ? state.generated : [];
  currentIndex = Number.isInteger(state.currentIndex) ? state.currentIndex : 0;

  if (generated.length) {
    $('sendControls').classList.remove('hidden');
    renderStatus(`Restored: ${currentIndex}/${generated.length} drafts opened. Click “Open Next Draft” to continue.`, 'ok');
    const next = generated[Math.min(currentIndex, generated.length - 1)];
    if (next) {
      $('preview').textContent = `Next draft preview:\nTo: ${next.email}\nSubject: ${next.subject}\n\n${next.body}`;
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
  await saveState();
  return rows;
}

function makeMessages() {
  const subjectTemplate = $('subject').value;
  const bodyTemplate = $('body').value;
  generated = rows.map(row => {
    const email = row.Email || row.email;
    const subject = applyTemplate(subjectTemplate, row);
    const body = applyTemplate(bodyTemplate, row);
    return { email, subject, body, row };
  });
}

function encodeForOutlook(text) {
  // Outlook Web deeplinks need percent-encoded spaces/newlines. URLSearchParams may serialize spaces as '+',
  // which Outlook sometimes shows literally in the compose window.
  return encodeURIComponent(String(text ?? ''));
}

function outlookComposeUrl(msg) {
  return 'https://outlook.office.com/mail/deeplink/compose'
    + '?to=' + encodeForOutlook(msg.email)
    + '&subject=' + encodeForOutlook(msg.subject)
    + '&body=' + encodeForOutlook(msg.body);
}

function mailtoUrl(msg) {
  return 'mailto:' + encodeForOutlook(msg.email)
    + '?subject=' + encodeForOutlook(msg.subject)
    + '&body=' + encodeForOutlook(msg.body);
}

function renderStatus(text, cls = '') {
  $('status').className = cls;
  $('status').textContent = text;
}

['csvText', 'subject', 'body'].forEach(id => {
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
    $('preview').textContent = `Loaded ${generated.length} rows.\n\nFirst preview:\nTo: ${first.email}\nSubject: ${first.subject}\n\n${first.body}`;
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
