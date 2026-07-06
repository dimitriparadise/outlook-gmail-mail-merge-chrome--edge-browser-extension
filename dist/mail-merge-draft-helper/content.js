// content.js

function isAutoSendEnabled() {
  return window.location.href.includes('mailMergeAutoSend=true');
}

function decodedParam(name) {
  const params = new URLSearchParams(window.location.search);
  return params.get(name) || '';
}

function markGmailTargetCompose() {
  const subject = decodedParam('su');
  const composeContainers = [...document.querySelectorAll('div[role="dialog"], table[role="presentation"]')];

  if (!subject && composeContainers.length !== 1) return false;

  for (const container of composeContainers) {
    const subjectInput = container.querySelector('input[name="subjectbox"]');
    if (!subject || (subjectInput && subjectInput.value === subject)) {
      container.setAttribute('data-mail-merge-target', 'true');
      return true;
    }
  }

  return false;
}

function markOutlookTargetCompose() {
  const subject = decodedParam('subject');
  const composeContainers = [...document.querySelectorAll('[role="dialog"], form')];

  if (!subject && composeContainers.length !== 1) return false;

  for (const container of composeContainers) {
    const text = container.textContent || '';
    if (!subject || text.includes(subject)) {
      container.setAttribute('data-mail-merge-target', 'true');
      return true;
    }
  }

  return false;
}

function ensureTargetCompose(isGmail, isOutlook) {
  if (document.querySelector('[data-mail-merge-target="true"]')) return true;
  if (isGmail) return markGmailTargetCompose();
  if (isOutlook) return markOutlookTargetCompose();
  return false;
}

function findAndClickSend() {
  const isGmail = window.location.hostname.includes('mail.google.com');
  const isOutlook = window.location.hostname.includes('outlook.office.com') || window.location.hostname.includes('outlook.live.com');

  let sendButton = null;

  if (!ensureTargetCompose(isGmail, isOutlook)) {
    console.warn('[MailMerge] Could not identify the target compose window. Auto-send will wait.');
    return false;
  }

  if (isGmail) {
    // Gmail send button usually has role="button" and aria-label starting with "Send"
    const buttons = document.querySelectorAll('[data-mail-merge-target="true"] div[role="button"]');
    for (const btn of buttons) {
      const label = btn.getAttribute('aria-label');
      // Sometimes it's "Send" or "Send ‪(Ctrl-Enter)‬"
      if (label && label.toLowerCase().startsWith('send')) {
        sendButton = btn;
        break;
      }
    }
  } else if (isOutlook) {
    // Outlook send button usually has aria-label="Send" or title="Send"
    const buttons = document.querySelectorAll('[data-mail-merge-target="true"] button');
    for (const btn of buttons) {
      const label = btn.getAttribute('aria-label') || btn.getAttribute('title') || btn.textContent;
      if (label && label.trim().toLowerCase() === 'send') {
        sendButton = btn;
        break;
      }
    }
  }

  if (sendButton) {
    console.log('[MailMerge] Found send button, clicking it now...');
    sendButton.click();
    return true;
  }

  return false;
}

if (isAutoSendEnabled()) {
  console.log('[MailMerge] Auto-send is enabled for this draft. Waiting for Send button...');
  
  const maxAttempts = 30; // Try for 15 seconds (30 * 500ms)
  let attempts = 0;

  const intervalId = setInterval(() => {
    attempts++;
    
    if (findAndClickSend()) {
      clearInterval(intervalId);
      console.log('[MailMerge] Clicked send. Will attempt to close tab in 3 seconds.');
      
      // Attempt to close the tab after a delay to ensure it sends
      setTimeout(() => {
        try {
          window.close();
        } catch (e) {
          console.error('[MailMerge] Could not close window:', e);
        }
      }, 3000);
    } else if (attempts >= maxAttempts) {
      clearInterval(intervalId);
      console.error('[MailMerge] Could not find Send button after 15 seconds. Aborting auto-send.');
    }
  }, 500);
}
