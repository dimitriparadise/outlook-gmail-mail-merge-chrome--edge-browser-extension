import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

function createElement(id, initial = {}) {
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
    addEventListener() {},
    appendChild() {},
    focus() {},
    setSelectionRange() {},
    dispatchEvent() {},
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
        create(_options, callback) {
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
  restoreState,
  makeMessages,
  composeUrl
};`, context);
  return { context, elements, saved, api: context.__popupTestApi };
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
