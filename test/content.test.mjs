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
