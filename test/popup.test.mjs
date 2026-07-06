import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

function createElement(id, initial = {}) {
  const listeners = new Map();
  return {
    id,
    value: initial.value ?? '',
    checked: initial.checked ?? false,
    textContent: '',
    innerHTML: '',
    className: '',
    files: [],
    max: '',
    disabled: false,
    classList: {
      add() {},
      remove() {},
      toggle() {}
    },
    addEventListener(type, handler) {
      if (!listeners.has(type)) listeners.set(type, []);
      listeners.get(type).push(handler);
    },
    async dispatchEvent(event) {
      for (const handler of listeners.get(event.type) || []) {
        await handler(event);
      }
    },
    async click() {
      await this.dispatchEvent({ type: 'click' });
    },
    appendChild() {},
    focus() {},
    setSelectionRange() {},
    get selectionStart() {
      return this.value.length;
    },
    get selectionEnd() {
      return this.value.length;
    }
  };
}

function loadPopup(state) {
  const elements = new Map();
  const defaults = {
    csvText: { value: '' },
    composeMode: { value: 'outlook' },
    autoSend: { checked: false },
    subject: { value: 'Reminder for {{Course}}' },
    cc: { value: '' },
    bcc: { value: '' },
    body: { value: 'Hi {{Name}},\n\nThis is a quick reminder about {{Course}}.' }
  };
  const ids = [
    'csvText',
    'composeMode',
    'autoSend',
    'subject',
    'cc',
    'bcc',
    'body',
    'variableButtons',
    'sendControls',
    'previewControls',
    'rangeStart',
    'rangeEnd',
    'status',
    'preview',
    'previewCounter',
    'previewPrevBtn',
    'previewNextBtn',
    'loadBtn',
    'openNextBtn',
    'openRangeBtn',
    'copyRecipientHint',
    'composeModeHint',
    'csvFile'
  ];
  for (const id of ids) elements.set(id, createElement(id, defaults[id]));

  const stored = state === undefined ? {} : { outlookMailMergeState: state };
  const saved = [];
  const openedTabs = [];
  const context = {
    console,
    Event: class Event {
      constructor(type, options = {}) {
        this.type = type;
        this.bubbles = Boolean(options.bubbles);
      }
    },
    document: {
      getElementById(id) {
        if (!elements.has(id)) elements.set(id, createElement(id));
        return elements.get(id);
      },
      createElement(tag) {
        return createElement(tag);
      }
    },
    chrome: {
      storage: {
        local: {
          async get() {
            return stored;
          },
          async set(next) {
            saved.push(next);
            Object.assign(stored, next);
          }
        }
      },
      tabs: {
        create(options, callback) {
          openedTabs.push(options);
          callback();
        }
      },
      runtime: {}
    }
  };
  vm.createContext(context);
  const script = fs.readFileSync(new URL('../popup.js', import.meta.url), 'utf8');
  vm.runInContext(`${script}
globalThis.__popupTestApi = {
  parseCSV,
  applyTemplate,
  validateMessages: typeof validateMessages === 'function' ? validateMessages : undefined,
  findMissingTemplateVariables: typeof findMissingTemplateVariables === 'function' ? findMissingTemplateVariables : undefined,
  syncDraftActionLabels,
  restoreState,
  makeMessages,
  composeUrl
};`, context);
  return { context, elements, saved, openedTabs, api: context.__popupTestApi };
}

test('restoreState preserves intentionally blank subject and body', async () => {
  const { elements, api } = loadPopup({
    csvText: '',
    composeMode: 'outlook',
    autoSend: false,
    subject: '',
    cc: '',
    bcc: '',
    body: '',
    rows: [],
    csvHeaders: [],
    generated: [],
    currentIndex: 0,
    previewIndex: 0
  });

  await api.restoreState();

  assert.equal(elements.get('subject').value, '');
  assert.equal(elements.get('body').value, '');
});

test('parseCSV supports quoted fields containing newlines', () => {
  const { api } = loadPopup();

  const rows = api.parseCSV('Name,Email,Notes\nJane,jane@example.com,"Line one\nLine two"');

  assert.equal(rows.length, 1);
  assert.equal(rows[0].Notes, 'Line one\nLine two');
});

test('parseCSV explains that wrapped text still needs a real line break', () => {
  const { api } = loadPopup();

  assert.throws(
    () => api.parseCSV('Name, Email, Course, Section, DueDate, John, john@example.com, ISOM 210, A, Friday'),
    /real line break.*press Enter/i
  );
});

test('findMissingTemplateVariables reports variables that are not CSV headers', () => {
  const { api } = loadPopup();

  const missing = api.findMissingTemplateVariables(
    ['Name', 'Email', 'Course'],
    ['Hi {{Name}}', 'Reminder for {{Coursee}}']
  );

  assert.deepEqual(Array.from(missing), ['Coursee']);
});

test('validateMessages reports invalid, duplicate, and overlapping recipients', () => {
  const { api } = loadPopup();
  const messages = [
    {
      email: 'student@example.com',
      cc: 'student@example.com, bad-email',
      bcc: '',
      subject: 'Hello',
      body: 'Body'
    },
    {
      email: 'student@example.com',
      cc: '',
      bcc: '',
      subject: 'Hello',
      body: 'Body'
    }
  ];

  const issues = api.validateMessages(messages);

  assert(issues.some(issue => issue.includes('Invalid email address: bad-email')));
  assert(issues.some(issue => issue.includes('Duplicate To recipient: student@example.com')));
  assert(issues.some(issue => issue.includes('also appears in CC/BCC')));
});

test('composeUrl blocks Outlook auto-send when CC or BCC is present', () => {
  const { elements, api } = loadPopup();
  elements.get('composeMode').value = 'outlook';
  elements.get('autoSend').checked = true;

  assert.throws(
    () => api.composeUrl({
      email: 'student@example.com',
      cc: 'ta@example.com',
      bcc: '',
      subject: 'Hello',
      body: 'Body'
    }),
    /Auto-Send with Outlook Mode cannot be used with CC or BCC/
  );
});

test('syncDraftActionLabels changes button text when Auto-Send is enabled', () => {
  const { elements, api } = loadPopup();

  elements.get('autoSend').checked = true;
  api.syncDraftActionLabels();

  assert.equal(elements.get('openNextBtn').textContent, 'Open & Send Preview Email');
  assert.equal(elements.get('openRangeBtn').textContent, 'Open & Send Selected Emails');
});

test('syncDraftActionLabels uses draft wording when Auto-Send is disabled', () => {
  const { elements, api } = loadPopup();

  elements.get('autoSend').checked = false;
  api.syncDraftActionLabels();

  assert.equal(elements.get('openNextBtn').textContent, 'Open Preview Draft');
  assert.equal(elements.get('openRangeBtn').textContent, 'Open Selected Drafts');
});

test('single open button opens the currently previewed draft', async () => {
  const generated = [
    { email: 'first@example.com', cc: '', bcc: '', subject: 'First', body: 'Body', row: {} },
    { email: 'second@example.com', cc: '', bcc: '', subject: 'Second', body: 'Body', row: {} }
  ];
  const { elements, openedTabs, api } = loadPopup({
    csvText: 'Name,Email\nFirst,first@example.com\nSecond,second@example.com',
    composeMode: 'gmail',
    autoSend: false,
    templatePreset: '',
    subject: '{{Name}}',
    cc: '',
    bcc: '',
    body: 'Body',
    rows: [
      { Name: 'First', Email: 'first@example.com' },
      { Name: 'Second', Email: 'second@example.com' }
    ],
    csvHeaders: ['Name', 'Email'],
    generated,
    currentIndex: 0,
    previewIndex: 1
  });

  await api.restoreState();
  await elements.get('openNextBtn').click();

  assert.equal(openedTabs.length, 1);
  assert.match(openedTabs[0].url, /to=second%40example\.com/);
  assert.doesNotMatch(openedTabs[0].url, /first%40example\.com/);
});
