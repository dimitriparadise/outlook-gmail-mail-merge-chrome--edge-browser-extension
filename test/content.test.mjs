import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

function loadContent(buttons) {
  const context = {
    console,
    window: {
      location: {
        href: 'https://mail.google.com/mail/?mailMergeAutoSend=true',
        hostname: 'mail.google.com'
      },
      close() {}
    },
    setInterval() {
      return 1;
    },
    clearInterval() {},
    setTimeout() {},
    document: {
      querySelector(selector) {
        if (selector === '[data-mail-merge-target="true"]') {
          return buttons.find(button => button.targeted) || null;
        }
        return null;
      },
      querySelectorAll(selector) {
        if (selector === '[data-mail-merge-target="true"] div[role="button"]') {
          return buttons.filter(button => button.targeted);
        }
        if (selector === 'div[role="button"]') return buttons;
        return [];
      }
    }
  };
  vm.createContext(context);
  const script = fs.readFileSync(new URL('../content.js', import.meta.url), 'utf8');
  vm.runInContext(`${script}
globalThis.__contentTestApi = { findAndClickSend };`, context);
  return context.__contentTestApi;
}

test('findAndClickSend clicks only the send button inside the mail merge target compose', () => {
  const clicks = [];
  const unrelatedSend = {
    targeted: false,
    textContent: '',
    getAttribute(name) {
      return name === 'aria-label' ? 'Send' : null;
    },
    click() {
      clicks.push('unrelated');
    }
  };
  const targetSend = {
    targeted: true,
    textContent: '',
    getAttribute(name) {
      return name === 'aria-label' ? 'Send' : null;
    },
    click() {
      clicks.push('target');
    }
  };

  const api = loadContent([unrelatedSend, targetSend]);

  assert.equal(api.findAndClickSend(), true);
  assert.deepEqual(clicks, ['target']);
});

test('findAndClickSend supports Chinese "发送" aria-label', () => {
  const clicks = [];
  const targetSendChinese = {
    targeted: true,
    textContent: '',
    getAttribute(name) {
      return name === 'aria-label' ? '发送 (Ctrl-Enter)' : null;
    },
    click() {
      clicks.push('chinese-send');
    }
  };

  const api = loadContent([targetSendChinese]);
  assert.equal(api.findAndClickSend(), true);
  assert.deepEqual(clicks, ['chinese-send']);
});

// Minimal DOM tree supporting only the selectors content.js uses on Outlook.
function el(tag, { attrs = {}, value, text = '', onClick } = {}, children = []) {
  const node = {
    tag,
    attrs: { ...attrs },
    value,
    parentElement: null,
    children,
    get textContent() {
      return text + children.map(child => child.textContent).join('');
    },
    getAttribute(name) {
      return name in node.attrs ? node.attrs[name] : null;
    },
    setAttribute(name, val) {
      node.attrs[name] = String(val);
    },
    click() {
      onClick?.();
    },
    querySelectorAll(selector) {
      return descendants(node).filter(child => matches(child, selector));
    }
  };
  for (const child of children) child.parentElement = node;
  return node;
}

function descendants(node) {
  return node.children.flatMap(child => [child, ...descendants(child)]);
}

function isTargeted(node) {
  return node.getAttribute('data-mail-merge-target') === 'true';
}

function matches(node, selector) {
  if (selector === 'input' || selector === 'button') return node.tag === selector;
  if (selector === '[role="dialog"], form') return node.getAttribute('role') === 'dialog' || node.tag === 'form';
  if (selector === '[data-mail-merge-target="true"]') return isTargeted(node);
  if (selector === '[data-mail-merge-target="true"] button') {
    if (node.tag !== 'button') return false;
    for (let parent = node.parentElement; parent; parent = parent.parentElement) {
      if (isTargeted(parent)) return true;
    }
    return false;
  }
  throw new Error(`Unsupported selector in test DOM: ${selector}`);
}

function loadOutlookContent(body, subject) {
  const search = `?subject=${encodeURIComponent(subject)}&mailMergeAutoSend=true`;
  const context = {
    console: { log() {}, warn() {}, error() {} },
    URLSearchParams,
    window: {
      location: {
        href: `https://outlook.office.com/mail/deeplink/compose${search}`,
        hostname: 'outlook.office.com',
        search
      },
      close() {}
    },
    setInterval() {
      return 1;
    },
    clearInterval() {},
    setTimeout() {},
    document: {
      querySelector(selector) {
        return [body, ...descendants(body)].find(node => matches(node, selector)) || null;
      },
      querySelectorAll(selector) {
        return [body, ...descendants(body)].filter(node => matches(node, selector));
      }
    }
  };
  vm.createContext(context);
  const script = fs.readFileSync(new URL('../content.js', import.meta.url), 'utf8');
  vm.runInContext(`${script}
globalThis.__contentTestApi = { findAndClickSend };`, context);
  return context.__contentTestApi;
}

test('findAndClickSend finds the Outlook full-page compose by its subject input value', () => {
  const clicks = [];
  const body = el('body', { attrs: { role: 'application' } }, [
    el('div', {}, [
      el('button', { attrs: { 'aria-label': 'Send' }, onClick: () => clicks.push('unrelated') }, [])
    ]),
    el('div', {}, [
      el('div', {}, [
        el('button', { attrs: { 'aria-label': 'Send' }, text: 'Send', onClick: () => clicks.push('target') }, []),
        el('button', { attrs: { 'aria-label': 'More send options' }, onClick: () => clicks.push('more') }, [])
      ]),
      el('div', {}, [
        el('input', { attrs: { 'aria-label': 'Subject' }, value: 'Hello Ada' }, [])
      ])
    ])
  ]);

  const api = loadOutlookContent(body, 'Hello Ada');
  assert.equal(api.findAndClickSend(), true);
  assert.deepEqual(clicks, ['target']);
});

test('findAndClickSend does not click in Outlook when no subject input matches', () => {
  const clicks = [];
  const body = el('body', {}, [
    el('div', {}, [
      el('button', { attrs: { 'aria-label': 'Send' }, onClick: () => clicks.push('send') }, []),
      el('input', { attrs: { 'aria-label': 'Subject' }, value: 'Another draft' }, [])
    ])
  ]);

  const api = loadOutlookContent(body, 'Hello Ada');
  assert.equal(api.findAndClickSend(), false);
  assert.deepEqual(clicks, []);
});

test('findAndClickSend still uses the Outlook compose dialog that contains the subject', () => {
  const clicks = [];
  const body = el('body', {}, [
    el('div', { attrs: { role: 'dialog' }, text: 'Other draft' }, [
      el('button', { attrs: { title: 'Send' }, onClick: () => clicks.push('other-dialog') }, [])
    ]),
    el('div', { attrs: { role: 'dialog' }, text: 'Hello Ada' }, [
      el('button', { attrs: { title: '发送' }, onClick: () => clicks.push('target-dialog') }, [])
    ])
  ]);

  const api = loadOutlookContent(body, 'Hello Ada');
  assert.equal(api.findAndClickSend(), true);
  assert.deepEqual(clicks, ['target-dialog']);
});
