// content.js

function isAutoSendEnabled() {
  return window.location.href.includes('mailMergeAutoSend=true');
}

function findAndClickSend() {
  const isGmail = window.location.hostname.includes('mail.google.com');
  const isOutlook = window.location.hostname.includes('outlook.office.com') || window.location.hostname.includes('outlook.live.com');

  let sendButton = null;

  if (isGmail) {
    // Gmail send button usually has role="button" and aria-label starting with "Send"
    const buttons = document.querySelectorAll('div[role="button"]');
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
    const buttons = document.querySelectorAll('button');
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
