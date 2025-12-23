/**
 * Gmail Lookalike Detector - Content Script
 * v3.6.0 - Fixed false positives: homoglyphs, public domain matching
 */

(function() {
  'use strict';
  
  // Storage keys
  const STATS_KEY = 'emailLookalikeStats';
  const BLOCKED_KEY = 'emailLookalikeBlocked';
  const SEEN_KEY = 'emailLookalikeSeenSenders';
  
  // Known contacts list (loaded per-account)
  let knownContacts = [];
  
  // Blocked senders list
  let blockedSenders = [];
  
  // Seen senders list
  let seenSenders = [];
  
  // Stats
  let emailsScanned = 0;
  
  // Current Gmail account being viewed
  let currentAccount = null;
  
  // Is this account set up?
  let accountSetup = false;
  
  // Debounce timer
  let checkEmailTimeout = null;
  let lastCheckedUrl = null;
  
  // Flag to prevent recheck when adding from warning modal
  let skipNextRecheck = false;
  
  // Account detection retry count
  let accountDetectionRetries = 0;
  const MAX_ACCOUNT_RETRIES = 10;
  
  /**
   * Detect the current Gmail account from the page
   */
  function detectCurrentAccount() {
    // Method 1: data-email attribute
    const profileButton = document.querySelector('[data-email]');
    if (profileButton) {
      const email = profileButton.getAttribute('data-email');
      if (email && email.includes('@')) {
        return email.toLowerCase().trim();
      }
    }
    
    // Method 2: aria-label/data-tooltip with email
    const tooltipElements = document.querySelectorAll('[data-tooltip*="@"], [aria-label*="@"]');
    for (const el of tooltipElements) {
      const tooltip = el.getAttribute('data-tooltip') || el.getAttribute('aria-label');
      const emailMatch = tooltip.match(/([a-zA-Z0-9._+-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/);
      if (emailMatch) {
        return emailMatch[1].toLowerCase().trim();
      }
    }
    
    // Method 3: Google Bar area
    const gbArea = document.querySelector('.gb_d, .gb_ua, .gb_A');
    if (gbArea) {
      const ariaLabel = gbArea.getAttribute('aria-label');
      if (ariaLabel) {
        const emailMatch = ariaLabel.match(/([a-zA-Z0-9._+-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/);
        if (emailMatch) {
          return emailMatch[1].toLowerCase().trim();
        }
      }
    }
    
    // Method 4: Page scripts
    const scripts = document.querySelectorAll('script');
    for (const script of scripts) {
      const content = script.textContent || '';
      const emailMatch = content.match(/"([a-zA-Z0-9._+-]+@gmail\.com)"/);
      if (emailMatch) {
        return emailMatch[1].toLowerCase().trim();
      }
    }
    
    // Method 5: Any element with email
    const allElements = document.querySelectorAll('[aria-label], [title], [data-hovercard-id]');
    for (const el of allElements) {
      const text = el.getAttribute('aria-label') || el.getAttribute('title') || el.getAttribute('data-hovercard-id') || '';
      if (text.includes('@gmail.com') || text.includes('@googlemail.com')) {
        const emailMatch = text.match(/([a-zA-Z0-9._+-]+@(?:gmail|googlemail)\.com)/i);
        if (emailMatch) {
          return emailMatch[1].toLowerCase().trim();
        }
      }
    }
    
    return null;
  }
  
  /**
   * Initialize the detector
   */
  async function init() {
    console.log('[Lookalike Detector] v3.6.0 Initializing...');
    
    // Detect current account
    currentAccount = detectCurrentAccount();
    
    if (!currentAccount) {
      if (accountDetectionRetries < MAX_ACCOUNT_RETRIES) {
        accountDetectionRetries++;
        console.log('[Lookalike Detector] Account not detected, retry', accountDetectionRetries);
        setTimeout(init, 1000);
        return;
      } else {
        console.log('[Lookalike Detector] Could not detect account');
        showAccountDetectionError();
        return;
      }
    }
    
    console.log('[Lookalike Detector] Current account:', currentAccount);
    
    // Check if this account has contacts synced
    const setupStatus = await checkAccountSetup();
    
    if (!setupStatus) {
      console.log('[Lookalike Detector] Account not set up, showing setup prompt');
      showSetupPrompt();
      return;
    }
    
    // Load contacts for this account
    await loadKnownContacts();
    await loadBlockedSenders();
    await loadSeenSenders();
    await loadStats();
    
    // Start monitoring
    observeGmail();
    
    // Listen for URL changes
    window.addEventListener('hashchange', () => {
      setTimeout(checkCurrentEmail, 500);
    });
    
    // Listen for storage changes
    chrome.storage.onChanged.addListener((changes, areaName) => {
      if (areaName === 'local') {
        const contactsKey = `contacts_${currentAccount}`;
        if (changes[contactsKey]) {
          loadKnownContacts().then(() => {
            if (skipNextRecheck) {
              skipNextRecheck = false;
              return;
            }
            lastCheckedUrl = null;
            setTimeout(checkCurrentEmail, 1000);
          });
        }
      }
    });
    
    console.log('[Lookalike Detector] Ready for account:', currentAccount);
  }
  
  /**
   * Check if current account has been set up
   */
  async function checkAccountSetup() {
    return new Promise((resolve) => {
      chrome.runtime.sendMessage(
        { action: 'isAccountSetup', email: currentAccount },
        (response) => {
          if (chrome.runtime.lastError) {
            resolve(false);
            return;
          }
          accountSetup = response && response.isSetup;
          resolve(accountSetup);
        }
      );
    });
  }
  
  /**
   * Show setup prompt
   */
  function showSetupPrompt() {
    const existing = document.querySelector('.lookalike-setup-prompt');
    if (existing) existing.remove();
    
    const prompt = document.createElement('div');
    prompt.className = 'lookalike-setup-prompt';
    prompt.innerHTML = `
      <div style="
        position: fixed;
        top: 80px;
        right: 20px;
        z-index: 999999;
        background: white;
        border-radius: 12px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.15);
        padding: 20px;
        max-width: 350px;
        font-family: 'Google Sans', Roboto, Arial, sans-serif;
      ">
        <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 12px;">
          <span style="font-size: 32px;">🛡️</span>
          <h3 style="margin: 0; font-size: 16px; color: #1a1a2e;">Protect This Account</h3>
        </div>
        <p style="margin: 0 0 12px 0; font-size: 14px; color: #666;">
          Email Fraud Alert needs to sync your contacts for <strong>${currentAccount}</strong> to detect lookalike emails.
        </p>
        <p style="margin: 0 0 16px 0; font-size: 12px; color: #999;">
          Your contacts stay private and are never sent to our servers.
        </p>
        <div style="display: flex; gap: 10px;">
          <button id="efa-setup-btn" style="
            flex: 1;
            padding: 10px 16px;
            background: linear-gradient(135deg, #d4a84b 0%, #c49a3a 100%);
            color: #1a1a2e;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
          ">🔗 Connect Contacts</button>
          <button id="efa-later-btn" style="
            padding: 10px 16px;
            background: #f0f0f0;
            color: #666;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            cursor: pointer;
          ">Later</button>
        </div>
      </div>
    `;
    
    document.body.appendChild(prompt);
    
    document.getElementById('efa-setup-btn').addEventListener('click', async () => {
      const btn = document.getElementById('efa-setup-btn');
      btn.textContent = '⏳ Connecting...';
      btn.disabled = true;
      
      try {
        await syncContactsForCurrentAccount();
        prompt.remove();
        accountSetup = true;
        await loadKnownContacts();
        observeGmail();
        showSetupSuccess();
      } catch (error) {
        btn.textContent = '🔗 Connect Contacts';
        btn.disabled = false;
        showSetupError(error.message);
      }
    });
    
    document.getElementById('efa-later-btn').addEventListener('click', () => {
      prompt.remove();
    });
  }
  
  /**
   * Sync contacts for current account
   */
  async function syncContactsForCurrentAccount() {
    return new Promise((resolve, reject) => {
      chrome.runtime.sendMessage(
        { action: 'syncContacts', email: currentAccount },
        (response) => {
          if (chrome.runtime.lastError) {
            reject(new Error(chrome.runtime.lastError.message));
            return;
          }
          if (response && response.success) {
            resolve(response.contacts);
          } else {
            reject(new Error(response ? response.error : 'Failed to sync contacts'));
          }
        }
      );
    });
  }
  
  /**
   * Show setup success
   */
  function showSetupSuccess() {
    const msg = document.createElement('div');
    msg.style.cssText = `
      position: fixed; top: 80px; right: 20px; z-index: 999999;
      background: #F0FDF4; color: #166534; border: 1px solid #86EFAC;
      padding: 16px 24px; border-radius: 8px; font-family: 'Google Sans', Roboto, Arial, sans-serif;
      font-size: 14px; box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    `;
    msg.textContent = `✓ Email Fraud Alert is now protecting ${currentAccount}`;
    document.body.appendChild(msg);
    setTimeout(() => msg.remove(), 5000);
  }
  
  /**
   * Show setup error
   */
  function showSetupError(message) {
    const msg = document.createElement('div');
    msg.style.cssText = `
      position: fixed; top: 80px; right: 20px; z-index: 999999;
      background: #FEF2F2; color: #991B1B; border: 1px solid #FECACA;
      padding: 16px 24px; border-radius: 8px; font-family: 'Google Sans', Roboto, Arial, sans-serif;
      font-size: 14px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); max-width: 350px;
    `;
    msg.innerHTML = `<strong>Setup failed:</strong> ${message}`;
    document.body.appendChild(msg);
    setTimeout(() => msg.remove(), 10000);
  }
  
  /**
   * Show account detection error
   */
  function showAccountDetectionError() {
    const msg = document.createElement('div');
    msg.style.cssText = `
      position: fixed; top: 80px; right: 20px; z-index: 999999;
      background: #FEF3C7; color: #92400E; border: 1px solid #FCD34D;
      padding: 16px 24px; border-radius: 8px; font-family: 'Google Sans', Roboto, Arial, sans-serif;
      font-size: 14px; box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    `;
    msg.textContent = '⚠ Email Fraud Alert could not detect your Gmail account. Please refresh the page.';
    document.body.appendChild(msg);
    setTimeout(() => msg.remove(), 10000);
  }
  
  /**
   * Load known contacts for current account
   */
  async function loadKnownContacts() {
    return new Promise((resolve) => {
      chrome.runtime.sendMessage(
        { action: 'getContacts', email: currentAccount },
        (response) => {
          if (chrome.runtime.lastError) {
            knownContacts = [];
            resolve();
            return;
          }
          knownContacts = (response && response.contacts) || [];
          
          // Add current account to known contacts so lookalikes of YOUR email are detected
          if (currentAccount && !knownContacts.includes(currentAccount.toLowerCase())) {
            knownContacts.push(currentAccount.toLowerCase());
          }
          
          console.log('[Lookalike Detector] Loaded', knownContacts.length, 'contacts for', currentAccount);
          resolve();
        }
      );
    });
  }
  
  /**
   * Add contact for current account
   */
  async function addKnownContact(email) {
    const normalized = email.toLowerCase().trim();
    if (!knownContacts.includes(normalized)) {
      knownContacts.push(normalized);
      return new Promise((resolve) => {
        chrome.runtime.sendMessage(
          { action: 'addContact', email: currentAccount, contact: normalized },
          () => resolve()
        );
      });
    }
    return Promise.resolve();
  }
  
  /**
   * Load blocked senders
   */
  async function loadBlockedSenders() {
    return new Promise((resolve) => {
      chrome.storage.local.get([BLOCKED_KEY], (result) => {
        blockedSenders = result[BLOCKED_KEY] || [];
        resolve();
      });
    });
  }
  
  /**
   * Load stats
   */
  async function loadStats() {
    return new Promise((resolve) => {
      chrome.storage.local.get([STATS_KEY], (result) => {
        const stats = result[STATS_KEY] || { emailsScanned: 0 };
        emailsScanned = stats.emailsScanned || 0;
        resolve();
      });
    });
  }
  
  /**
   * Block a sender
   */
  async function blockSender(email) {
    const normalized = email.toLowerCase().trim();
    if (!blockedSenders.includes(normalized)) {
      blockedSenders.push(normalized);
      return new Promise((resolve) => {
        chrome.storage.local.set({ [BLOCKED_KEY]: blockedSenders }, resolve);
      });
    }
    return Promise.resolve();
  }
  
  /**
   * Load seen senders
   */
  async function loadSeenSenders() {
    return new Promise((resolve) => {
      chrome.storage.local.get([SEEN_KEY], (result) => {
        seenSenders = result[SEEN_KEY] || [];
        resolve();
      });
    });
  }
  
  /**
   * Mark sender as seen
   */
  async function markSenderAsSeen(email) {
    const normalized = email.toLowerCase().trim();
    if (!seenSenders.includes(normalized)) {
      seenSenders.push(normalized);
      return new Promise((resolve) => {
        chrome.storage.local.set({ [SEEN_KEY]: seenSenders }, resolve);
      });
    }
    return Promise.resolve();
  }
  
  /**
   * Monitor Gmail for opened emails
   */
  function observeGmail() {
    setTimeout(() => {
      const emailView = document.querySelector('.AO') || document.querySelector('.nH') || document.body;
      
      const observer = new MutationObserver(() => {
        if (checkEmailTimeout) clearTimeout(checkEmailTimeout);
        checkEmailTimeout = setTimeout(checkCurrentEmail, 500);
      });
      
      observer.observe(emailView, { childList: true, subtree: true });
    }, 2000);
    
    setTimeout(checkCurrentEmail, 1000);
  }
  
  /**
   * Check if viewing an email
   */
  function isViewingEmail() {
    const hash = window.location.hash;
    const emailViewPattern = /#[^/]+\/[a-zA-Z0-9]{10,}/;
    const searchEmailPattern = /#search\/[^/]+\/[a-zA-Z0-9]{10,}/;
    return emailViewPattern.test(hash) || searchEmailPattern.test(hash);
  }
  
  /**
   * Check the currently opened email - collects ALL warnings
   */
  function checkCurrentEmail() {
    if (!accountSetup || !currentAccount) return;
    
    const currentUrl = window.location.hash;
    
    if (!isViewingEmail()) {
      removeExistingWarnings();
      lastCheckedUrl = null;
      return;
    }
    
    // If same email URL, don't recheck
    if (currentUrl === lastCheckedUrl) {
      return;
    }
    
    // New email - clear any old warnings
    removeExistingWarnings();
    lastCheckedUrl = currentUrl;
    
    // Find sender email - must be in the CURRENT email view
    let senderEmail = null;
    let senderName = null;
    
    const emailContainer = document.querySelector('.adn.ads') || 
                          document.querySelector('.h7') ||
                          document.querySelector('.nH.hx');
    
    if (emailContainer) {
      const emailAttr = emailContainer.querySelector('span[email]');
      if (emailAttr && emailAttr.getAttribute('email')) {
        senderEmail = emailAttr.getAttribute('email');
        senderName = emailAttr.getAttribute('name') || null;
      }
      
      if (!senderEmail) {
        const hovercard = emailContainer.querySelector('[data-hovercard-id]');
        if (hovercard) {
          const hovercardId = hovercard.getAttribute('data-hovercard-id');
          if (hovercardId && hovercardId.includes('@')) {
            senderEmail = hovercardId;
          }
        }
      }
      
      if (!senderEmail) {
        const headerElements = emailContainer.querySelectorAll('.gD, .go, .g2');
        for (const el of headerElements) {
          const text = el.textContent || '';
          const emailMatch = text.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/);
          if (emailMatch) {
            senderEmail = emailMatch[1];
            break;
          }
        }
      }
    }
    
    // Fallback to document-wide search
    if (!senderEmail) {
      const emailAttr = document.querySelector('span[email]');
      if (emailAttr && emailAttr.getAttribute('email')) {
        senderEmail = emailAttr.getAttribute('email');
        senderName = emailAttr.getAttribute('name') || null;
      }
    }
    
    if (!senderEmail) return;
    
    senderEmail = senderEmail.toLowerCase().trim();
    
    // Get sender name if not found
    if (!senderName) {
      const gdElement = document.querySelector('.gD');
      if (gdElement) {
        senderName = gdElement.getAttribute('name') || gdElement.textContent;
      }
    }
    
    console.log('[Lookalike Detector] Checking:', senderEmail, 'Name:', senderName, '(', knownContacts.length, 'contacts)');
    
    incrementScannedCount();
    
    // ========================================
    // COLLECT ALL WARNINGS - NO EARLY RETURNS
    // ========================================
    const warnings = [];
    
    // CHECK 1: Self-impersonation (someone using dot-variant of YOUR email)
    const selfImpersonation = checkSelfImpersonation(senderEmail);
    if (selfImpersonation) {
      warnings.push({
        type: 'self-impersonation',
        severity: 'critical',
        title: 'Someone Is Impersonating YOU',
        description: 'This email appears to be from a dot-variant of your own email address. Gmail treats dots as identical, but this could be a scammer pretending to be you.',
        fromEmail: senderEmail,
        toEmail: currentAccount,
        labelFrom: 'Sender',
        labelTo: 'Your email',
        reason: 'Gmail ignores dots, so this address could be used to impersonate you'
      });
    }
    
    // CHECK 2: Reply-To mismatch
    const replyToEmail = getReplyToAddress();
    if (replyToEmail && replyToEmail.toLowerCase().trim() !== senderEmail) {
      warnings.push({
        type: 'replyto-mismatch',
        severity: 'critical',
        title: 'Reply-To Hijacking',
        description: 'Replies will go to a different address than the sender. This is a common fraud tactic.',
        fromEmail: senderEmail,
        toEmail: replyToEmail.toLowerCase().trim(),
        labelFrom: 'From',
        labelTo: 'Replies go to'
      });
    }
    
    // CHECK 3: Lookalike detection (check BEFORE display name)
    // Skip if sender is already in contacts (they're a known coworker, not an impersonator)
    let isLookalike = false;
    if (!knownContacts.includes(senderEmail)) {
      // First check for Gmail dot-variant impersonation of contacts
      const dotVariant = checkContactDotVariant(senderEmail);
      if (dotVariant) {
        isLookalike = true;
        warnings.push({
          type: 'lookalike',
          severity: 'critical',
          title: 'Lookalike Email Address',
          description: 'This is a dot-variant of a contact\'s email. Gmail treats dots as identical, so this could be impersonation.',
          fromEmail: senderEmail,
          toEmail: dotVariant.contactEmail,
          labelFrom: 'Suspicious',
          labelTo: 'Your contact',
          reason: 'Gmail dot-trick: same address with different dot placement'
        });
      }
      
      // Then check for other lookalike patterns
      if (!isLookalike) {
        const matchResult = window.LookalikeDetector.findLookalike(senderEmail, knownContacts);
        if (matchResult) {
          // Filter out false positives
          const publicDomains = ['gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'aol.com', 
                                 'icloud.com', 'mail.com', 'protonmail.com', 'zoho.com', 'yandex.com'];
          const incomingDomain = senderEmail.split('@')[1] || '';
          const matchedDomain = matchResult.knownEmail.split('@')[1] || '';
          const incomingUsername = senderEmail.split('@')[0] || '';
          const matchedUsername = matchResult.knownEmail.split('@')[0] || '';
          
          // Calculate how different the usernames actually are
          const usernameDiff = levenshteinDistance(incomingUsername, matchedUsername);
          
          // False positive check:
          // Only skip if BOTH on same public domain AND usernames are very different (>4 chars)
          // If usernames are similar (≤4 chars), it's likely impersonation - flag it!
          const bothPublicSameDomain = publicDomains.includes(incomingDomain) && incomingDomain === matchedDomain;
          const usernamesSimilar = usernameDiff <= 4;
          
          // Flag if: usernames are similar, OR different domains (real lookalike attack)
          // Skip if: same public domain AND usernames are very different (just different people)
          const isFalsePositive = bothPublicSameDomain && !usernamesSimilar;
          
          if (!isFalsePositive) {
            isLookalike = true;
            warnings.push({
              type: 'lookalike',
              severity: 'critical',
              title: 'Lookalike Email Address',
              description: 'This email is nearly identical to someone in your contacts, but slightly different.',
              fromEmail: matchResult.incomingEmail,
              toEmail: matchResult.knownEmail,
              labelFrom: 'Suspicious',
              labelTo: 'Your contact',
              reason: matchResult.reasons[0] || 'Similar email detected'
            });
          }
        }
      }
    }
    
    // CHECK 4: Display name impersonation (only if NOT already a lookalike)
    // Also skip if display name IS the sender's email (not spoofing, just using email as name)
    // Also skip if sender is already in contacts (they're trusted)
    const displayNameIsEmail = senderName && senderName.toLowerCase().includes('@') && 
                               senderName.toLowerCase().includes(senderEmail.split('@')[0]);
    if (!isLookalike && senderName && !displayNameIsEmail && !knownContacts.includes(senderEmail)) {
      const impersonationMatch = checkDisplayNameImpersonation(senderName, senderEmail);
      if (impersonationMatch) {
        warnings.push({
          type: 'display-name',
          severity: 'critical',
          title: 'Display Name Spoofing',
          description: 'The name looks trustworthy, but the actual email address is completely different.',
          fromEmail: senderName,
          toEmail: senderEmail,
          labelFrom: 'Display Name',
          labelTo: 'Actual Email',
          reason: impersonationMatch.isUsernameImpersonation 
            ? `Name mimics "${impersonationMatch.keyword}" but email doesn't match`
            : `Name contains "${impersonationMatch.keyword}" but email is NOT from ${impersonationMatch.trustedDomain}`
        });
      }
    }
    
    // CHECK 5: Deceptive TLD
    const deceptiveTLD = checkDeceptiveTLD(senderEmail);
    if (deceptiveTLD) {
      warnings.push({
        type: 'deceptive-tld',
        severity: 'critical',
        title: 'Deceptive Domain',
        description: 'This domain is designed to look legitimate but is registered elsewhere.',
        fromEmail: senderEmail,
        toEmail: deceptiveTLD.fakingAs,
        labelFrom: 'Actual',
        labelTo: 'Looks like',
        reason: deceptiveTLD.warning
      });
    }
    
    // CHECK 6: Homoglyph/Unicode attack detection
    const homoglyphCheck = checkHomoglyphs(senderEmail);
    if (homoglyphCheck) {
      warnings.push({
        type: 'homoglyph',
        severity: 'critical',
        title: 'Invisible Character Trick',
        description: 'This email contains deceptive characters that look identical to normal letters.',
        fromEmail: senderEmail,
        toEmail: homoglyphCheck.detail,
        labelFrom: 'Email',
        labelTo: 'Hidden chars',
        reason: 'Cyrillic or special characters used to impersonate a legitimate address'
      });
    }
    
    // CHECK 7: Wire fraud keywords (always check)
    const wireKeywords = checkForWireInstructions();
    if (wireKeywords.length > 0) {
      warnings.push({
        type: 'wire-fraud',
        severity: 'critical',
        title: 'Dangerous Keywords Detected',
        description: 'This email contains terms commonly used in wire fraud.',
        keywords: wireKeywords,
        isWireFraud: true
      });
    }
    
    // CHECK 8: Self-name in subject or display name (personalized spam)
    const selfNameCheck = checkSelfNameUsed(senderEmail, senderName);
    if (selfNameCheck) {
      warnings.push({
        type: 'self-name-spam',
        severity: 'critical',
        title: 'Personalized Spam Detected',
        description: 'Your name appears in this email from an unknown sender — a common spam tactic.',
        fromEmail: senderName || senderEmail,
        toEmail: `Contains: "${selfNameCheck.found}"`,
        labelFrom: 'Sender',
        labelTo: 'Your name used',
        reason: selfNameCheck.reason
      });
    }
    
    // CHECK 9: Via domain deceptive TLD
    const viaCheck = checkViaDomain();
    if (viaCheck) {
      warnings.push({
        type: 'via-deceptive',
        severity: 'critical',
        title: 'Suspicious Sending Server',
        description: 'This email was sent through a server with a deceptive domain.',
        fromEmail: senderEmail,
        toEmail: viaCheck.viaDomain,
        labelFrom: 'Sender',
        labelTo: 'Sent via',
        reason: viaCheck.warning
      });
    }
    
    // CHECK 10: Gibberish/random sender detection (only truly random like QIQXFjNGdELwgWF)
    const gibberishCheck = checkGibberishSender(senderEmail, senderName);
    if (gibberishCheck && !knownContacts.includes(senderEmail)) {
      warnings.push({
        type: 'gibberish-sender',
        severity: 'critical',
        title: 'Suspicious Sender Address',
        description: 'This email address appears to be randomly generated.',
        fromEmail: senderEmail,
        toEmail: gibberishCheck.suspicious,
        labelFrom: 'Sender',
        labelTo: 'Suspicious part',
        reason: gibberishCheck.reason
      });
    }
    
    // CHECK 11: Spam/scam subject patterns (require 2+ matches to avoid false positives)
    const spamCheck = checkSpamPatterns();
    if (spamCheck && spamCheck.patterns.length >= 2 && !knownContacts.includes(senderEmail)) {
      warnings.push({
        type: 'spam-patterns',
        severity: 'critical',
        title: 'Scam Pattern Detected',
        description: 'This email contains multiple phrases commonly used in scams.',
        keywords: spamCheck.patterns,
        isWireFraud: true // reuse the keyword display style
      });
    }
    
    // ========================================
    // DISPLAY ALL WARNINGS IN ONE MODAL
    // ========================================
    if (warnings.length > 0) {
      console.log('[Lookalike Detector] Found', warnings.length, 'warning(s):', warnings.map(w => w.type).join(', '));
      showCombinedWarning(warnings, senderEmail);
    } else if (knownContacts.includes(senderEmail)) {
      console.log('[Lookalike Detector] ✓ Email is in known contacts:', senderEmail);
    }
  }
  
  /**
   * Get Reply-To address
   */
  function getReplyToAddress() {
    const headerRows = document.querySelectorAll('.ajA, .ajz, .gH, .gI');
    for (const row of headerRows) {
      const text = row.textContent || '';
      if (text.toLowerCase().includes('reply-to')) {
        const emailMatch = text.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/);
        if (emailMatch) return emailMatch[1];
      }
    }
    return null;
  }
  
  /**
   * Check display name impersonation
   */
  function checkDisplayNameImpersonation(displayName, senderEmail) {
    if (!displayName) return null;
    
    const normalizedName = displayName.toLowerCase().replace(/[^a-z0-9]/g, '');
    const senderDomain = senderEmail.split('@')[1] || '';
    
    const publicDomains = new Set([
      'gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'aol.com',
      'icloud.com', 'mail.com', 'protonmail.com', 'zoho.com', 'yandex.com'
    ]);
    
    const domainKeywords = new Map();
    
    for (const contact of knownContacts) {
      const domain = contact.split('@')[1];
      if (domain && !publicDomains.has(domain.toLowerCase())) {
        const domainName = domain.split('.')[0].toLowerCase();
        if (domainName.length >= 5) {
          domainKeywords.set(domainName, domain.toLowerCase());
        }
      }
    }
    
    if (domainKeywords.has(senderDomain.split('.')[0].toLowerCase())) {
      return null;
    }
    
    for (const [keyword, trustedDomain] of domainKeywords) {
      if (normalizedName.includes(keyword)) {
        return { keyword, trustedDomain };
      }
    }
    
    // NEW CHECK: Display name looks like a known contact's email username
    // BUT skip if sender is from a trusted domain
    const senderDomainLower = senderDomain.toLowerCase();
    const trustedDomains = new Set();
    for (const contact of knownContacts) {
      const d = contact.split('@')[1];
      if (d && !publicDomains.has(d.toLowerCase())) {
        trustedDomains.add(d.toLowerCase());
      }
    }
    
    // Skip username check if sender is from a trusted domain
    if (!trustedDomains.has(senderDomainLower)) {
      for (const contact of knownContacts) {
        const contactUsername = contact.split('@')[0].toLowerCase().replace(/[^a-z0-9]/g, '');
        if (contactUsername.length >= 5 && normalizedName.includes(contactUsername)) {
          if (senderEmail.toLowerCase() !== contact.toLowerCase()) {
            return { 
              keyword: contact.split('@')[0], 
              trustedDomain: contact,
              isUsernameImpersonation: true
            };
          }
        }
      }
    }
    
    return null;
  }
  
  /**
   * Check deceptive TLD
   */
  function checkDeceptiveTLD(email) {
    const domain = email.split('@')[1];
    if (!domain) return null;
    
    const patterns = [
      { pattern: /\.com\.co$/i, readable: '.com.co (Colombia)', fake: '.com' },
      { pattern: /\.com\.br$/i, readable: '.com.br (Brazil)', fake: '.com' },
      { pattern: /\.com\.mx$/i, readable: '.com.mx (Mexico)', fake: '.com' },
      { pattern: /\.com\.au$/i, readable: '.com.au (Australia)', fake: '.com' },
      { pattern: /\.com\.cn$/i, readable: '.com.cn (China)', fake: '.com' },
      { pattern: /\.com\.ru$/i, readable: '.com.ru (Russia)', fake: '.com' },
      { pattern: /\.com\.ng$/i, readable: '.com.ng (Nigeria)', fake: '.com' },
    ];
    
    for (const p of patterns) {
      if (p.pattern.test(domain)) {
        const fakeDomain = domain.replace(p.pattern, p.fake);
        return {
          domain: domain,
          readable: p.readable,
          fakingAs: fakeDomain,
          warning: `This domain "${domain}" looks like "${fakeDomain}" but is registered in ${p.readable.split('(')[1].replace(')', '')}`
        };
      }
    }
    
    return null;
  }
  
  /**
   * Check for homoglyph/unicode attacks (Cyrillic characters that look like Latin)
   * NOTE: We removed '0' and '1' because normal numbers in emails trigger false positives
   */
  function checkHomoglyphs(email) {
    const homoglyphs = {
      'а': 'a', 'е': 'e', 'о': 'o', 'р': 'p', 'с': 'c', 'х': 'x',
      'і': 'i', 'ј': 'j', 'ѕ': 's', 'ԁ': 'd', 'ɡ': 'g', 'ո': 'n',
      'ν': 'v', 'ѡ': 'w', 'у': 'y', 'һ': 'h', 'ⅼ': 'l', 'ｍ': 'm',
      '！': '!', '＠': '@'
      // Removed '0': 'o' and '1': 'l' - normal digits cause false positives
    };
    
    const found = [];
    for (const [homoglyph, latin] of Object.entries(homoglyphs)) {
      if (email.includes(homoglyph)) {
        found.push(`"${homoglyph}" looks like "${latin}"`);
      }
    }
    
    return found.length > 0 ? { found: found, detail: found.join(', ') } : null;
  }

  /**
   * Check for wire instructions
   */
  function checkForWireInstructions() {
    const wireKeywords = [
      // Wire/Money Movement
      'wire transfer', 'wire instructions', 'wiring instructions',
      'wire information', 'wire details', 'updated wire',
      'new wire', 'wire account', 'wire funds',
      'ach transfer', 'direct deposit',
      'zelle', 'venmo', 'cryptocurrency', 'bitcoin',
      'send funds', 'transfer funds', 'remit funds',
      
      // Bank/Account Info
      'bank account', 'account number', 'routing number',
      'aba number', 'swift code', 'iban',
      'bank statement', 'voided check', 'beneficiary',
      
      // Account Changes (red flag)
      'updated bank', 'new bank', 'changed bank',
      'updated payment', 'new payment info',
      'changed account', 'new account details',
      'payment update', 'revised instructions',
      'please update your records',
      
      // Real Estate / Escrow
      'closing funds', 'earnest money', 'escrow funds',
      'wire to', 'remittance', 'wire payment',
      
      // Legal / Attorney
      'settlement funds', 'settlement payment',
      'retainer', 'trust account', 'iolta',
      'client funds', 'case settlement',
      'court filing fee', 'legal fee',
      
      // Secrecy Red Flags
      'keep this confidential', 'keep this quiet',
      'dont mention this', 'between us',
      'dont tell anyone', 'private matter',
      'off the record', 'handle personally',
      
      // Sensitive Data Requests
      'social security', 'ssn', 'tax id',
      'w-9', 'w9', 'ein number',
      'login credentials', 'password reset',
      
      // Authority Impersonation
      'ceo request', 'cfo request', 'owner request',
      'boss asked', 'executive request', 'president asked'
    ];
    
    const bodySelectors = ['.a3s.aiL', '.ii.gt', '.a3s', 'div[data-message-id]'];
    let emailBody = null;
    for (const selector of bodySelectors) {
      emailBody = document.querySelector(selector);
      if (emailBody && emailBody.innerText.trim()) break;
    }
    
    const subjectEl = document.querySelector('.hP, h2.hP');
    
    let textToCheck = '';
    if (emailBody) textToCheck += emailBody.innerText.toLowerCase() + ' ';
    if (subjectEl) textToCheck += subjectEl.innerText.toLowerCase();
    
    if (!textToCheck.trim()) return [];
    
    const found = [];
    for (const keyword of wireKeywords) {
      if (textToCheck.includes(keyword.toLowerCase())) {
        found.push(keyword);
      }
    }
    
    return found;
  }
  
  /**
   * Calculate Levenshtein distance between two strings
   */
  function levenshteinDistance(a, b) {
    if (a.length === 0) return b.length;
    if (b.length === 0) return a.length;
    
    const matrix = [];
    
    for (let i = 0; i <= b.length; i++) {
      matrix[i] = [i];
    }
    
    for (let j = 0; j <= a.length; j++) {
      matrix[0][j] = j;
    }
    
    for (let i = 1; i <= b.length; i++) {
      for (let j = 1; j <= a.length; j++) {
        if (b.charAt(i - 1) === a.charAt(j - 1)) {
          matrix[i][j] = matrix[i - 1][j - 1];
        } else {
          matrix[i][j] = Math.min(
            matrix[i - 1][j - 1] + 1,
            matrix[i][j - 1] + 1,
            matrix[i - 1][j] + 1
          );
        }
      }
    }
    
    return matrix[b.length][a.length];
  }

  /**
   * Check if sender is using a dot-variant of the user's own email (impersonating them)
   */
  function checkSelfImpersonation(senderEmail) {
    if (!currentAccount) return null;
    
    const senderParts = senderEmail.toLowerCase().split('@');
    const userParts = currentAccount.toLowerCase().split('@');
    
    // Only check Gmail addresses (Gmail ignores dots)
    if (senderParts[1] !== 'gmail.com' || userParts[1] !== 'gmail.com') return null;
    
    // If exact match, not impersonation
    if (senderEmail.toLowerCase() === currentAccount.toLowerCase()) return null;
    
    // Remove all dots from usernames and compare
    const senderNoDots = senderParts[0].replace(/\./g, '');
    const userNoDots = userParts[0].replace(/\./g, '');
    
    // If same without dots but different with dots = dot-variant impersonation
    if (senderNoDots === userNoDots && senderParts[0] !== userParts[0]) {
      return {
        senderEmail: senderEmail,
        userEmail: currentAccount
      };
    }
    
    return null;
  }

  /**
   * Check if sender is using a dot-variant of a known contact's Gmail (impersonating them)
   */
  function checkContactDotVariant(senderEmail) {
    const senderParts = senderEmail.toLowerCase().split('@');
    
    // Only check Gmail addresses
    if (senderParts[1] !== 'gmail.com') return null;
    
    const senderNoDots = senderParts[0].replace(/\./g, '');
    
    for (const contact of knownContacts) {
      const contactParts = contact.toLowerCase().split('@');
      
      // Only compare Gmail to Gmail
      if (contactParts[1] !== 'gmail.com') continue;
      
      // Skip exact matches
      if (senderEmail.toLowerCase() === contact.toLowerCase()) continue;
      
      const contactNoDots = contactParts[0].replace(/\./g, '');
      
      // Same without dots but different with dots = dot-variant
      if (senderNoDots === contactNoDots && senderParts[0] !== contactParts[0]) {
        return {
          senderEmail: senderEmail,
          contactEmail: contact
        };
      }
    }
    
    return null;
  }

  /**
   * Check if user's own name is being used in subject or display name (personalized spam)
   */
  function checkSelfNameUsed(senderEmail, senderName) {
    if (!currentAccount) return null;
    if (knownContacts.includes(senderEmail)) return null; // Skip known contacts
    
    // Extract username from current account
    const myUsername = currentAccount.split('@')[0].toLowerCase();
    const myNameParts = myUsername.replace(/[^a-z]/g, ' ').split(' ').filter(p => p.length >= 4);
    
    // Check subject line
    const subjectEl = document.querySelector('.hP, h2.hP');
    const subject = subjectEl ? subjectEl.innerText.toLowerCase() : '';
    
    // Check display name
    const displayName = (senderName || '').toLowerCase();
    
    // Look for username or name parts in subject
    if (subject.includes(myUsername) || subject.includes(myUsername.replace('.', ''))) {
      return { found: myUsername, reason: 'Your username appears in the subject line' };
    }
    
    for (const part of myNameParts) {
      if (subject.includes(part) && !displayName.includes(part)) {
        return { found: part, reason: `Your name "${part}" appears in the subject line` };
      }
    }
    
    return null;
  }
  
  /**
   * Check the "via" domain for deceptive TLDs
   */
  function checkViaDomain() {
    // Look for "via" text in email header
    const headerArea = document.querySelector('.gH, .gI, .ha');
    if (!headerArea) return null;
    
    const headerText = headerArea.innerText || '';
    const viaMatch = headerText.match(/via\s+([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i);
    
    if (!viaMatch) return null;
    
    const viaDomain = viaMatch[1].toLowerCase();
    
    // Check for deceptive TLDs in via domain
    const deceptivePatterns = [
      { pattern: /\.com\.br$/i, country: 'Brazil' },
      { pattern: /\.com\.co$/i, country: 'Colombia' },
      { pattern: /\.com\.mx$/i, country: 'Mexico' },
      { pattern: /\.com\.ng$/i, country: 'Nigeria' },
      { pattern: /\.com\.ru$/i, country: 'Russia' },
      { pattern: /\.com\.cn$/i, country: 'China' },
    ];
    
    for (const dp of deceptivePatterns) {
      if (dp.pattern.test(viaDomain)) {
        return {
          viaDomain: viaDomain,
          warning: `Email sent through server in ${dp.country} (${viaDomain})`
        };
      }
    }
    
    return null;
  }
  
  /**
   * Check for gibberish/random sender addresses (STRICT - only truly random)
   */
  function checkGibberishSender(email, displayName) {
    const originalUsername = email.split('@')[0];
    const username = originalUsername.toLowerCase();
    const domain = email.split('@')[1] || '';
    
    // Check for mixed case gibberish in original (like QIQXFjNGdELwgWF)
    // Must have multiple uppercase letters mixed with lowercase
    const upperCount = (originalUsername.match(/[A-Z]/g) || []).length;
    const lowerCount = (originalUsername.match(/[a-z]/g) || []).length;
    const hasMixedCaseGibberish = upperCount >= 3 && lowerCount >= 3 && 
                                   /[A-Z][a-z][A-Z]|[a-z][A-Z][a-z]/.test(originalUsername);
    
    // Check if BOTH username AND domain have consonant clusters (like qiqxfjngdelwgwf.com)
    const usernameConsonants = /[bcdfghjklmnpqrstvwxz]{6,}/i.test(username);
    const domainParts = domain.split('.');
    const domainConsonants = domainParts.some(p => /[bcdfghjklmnpqrstvwxz]{6,}/i.test(p));
    const bothHaveConsonantClusters = usernameConsonants && domainConsonants;
    
    // Check for very long random-looking username (15+ chars, no common patterns)
    const veryLongRandom = username.length >= 15 && 
                           !/^[a-z]+\d*$/.test(username) && // not "johnsmith123"
                           !/[aeiou]{2,}/i.test(username); // no vowel pairs (real words have these)
    
    if (hasMixedCaseGibberish) {
      return { suspicious: originalUsername, reason: 'Username has random mixed-case pattern' };
    }
    
    if (bothHaveConsonantClusters) {
      return { suspicious: email, reason: 'Both username and domain appear randomly generated' };
    }
    
    if (veryLongRandom) {
      return { suspicious: username, reason: 'Username appears randomly generated' };
    }
    
    return null;
  }
  
  /**
   * Check for common spam/scam subject patterns
   */
  function checkSpamPatterns() {
    const subjectEl = document.querySelector('.hP, h2.hP');
    const subject = subjectEl ? subjectEl.innerText.toLowerCase() : '';
    
    if (!subject) return null;
    
    const spamPatterns = [
      { pattern: /congrat(s|ulations)/i, name: 'Congratulations' },
      { pattern: /you('ve| have) (been |)(selected|chosen|won)/i, name: 'You have been selected' },
      { pattern: /\d+%\s*(match|bonus|off|discount)/i, name: 'Percentage bonus' },
      { pattern: /claim (your|now|today)/i, name: 'Claim now' },
      { pattern: /act (now|fast|immediately)/i, name: 'Act now urgency' },
      { pattern: /limited time/i, name: 'Limited time' },
      { pattern: /winner|winning/i, name: 'Winner notification' },
      { pattern: /prize|reward|gift card/i, name: 'Prize/reward' },
      { pattern: /verify (your |)(account|identity)/i, name: 'Verify account' },
      { pattern: /suspended|locked out/i, name: 'Account suspended' },
      { pattern: /urgent.{0,10}(action|response|attention)/i, name: 'Urgent action' },
      { pattern: /bitcoin|crypto|investment opportunity/i, name: 'Crypto scam' },
      { pattern: /inheritance|beneficiary|next of kin/i, name: 'Inheritance scam' },
    ];
    
    const found = [];
    for (const sp of spamPatterns) {
      if (sp.pattern.test(subject)) {
        found.push(sp.name);
      }
    }
    
    if (found.length > 0) {
      return { patterns: found };
    }
    
    return null;
  }

  /**
   * Remove existing warnings
   */
  function removeExistingWarnings() {
    document.querySelectorAll('.lookalike-warning-banner').forEach(el => el.remove());
    document.querySelectorAll('.lookalike-warning-backdrop').forEach(el => el.remove());
    document.querySelectorAll('.lookalike-wire-banner').forEach(el => el.remove());
  }
  
  /**
   * Increment scanned count
   */
  function incrementScannedCount() {
    emailsScanned++;
    chrome.storage.local.get([STATS_KEY], (result) => {
      const stats = result[STATS_KEY] || { emailsScanned: 0 };
      stats.emailsScanned = emailsScanned;
      chrome.storage.local.set({ [STATS_KEY]: stats });
    });
  }
  
  /**
   * Show combined warning modal with all detected issues
   */
  function showCombinedWarning(warnings, senderEmail) {
    const backdrop = document.createElement('div');
    backdrop.className = 'lookalike-warning-backdrop';
    
    const banner = document.createElement('div');
    banner.className = 'lookalike-warning-banner';
    
    // Build warning items HTML
    const warningItemsHtml = warnings.map((w, index) => {
      let contentHtml = '';
      
      if (w.isWireFraud) {
        // Wire fraud/spam patterns have keyword tags
        const keywordTags = w.keywords.slice(0, 5).map(k => 
          `<span class="lookalike-keyword-tag">${k}</span>`
        ).join('');
        contentHtml = `
          <div class="lookalike-keywords">
            <div class="lookalike-keywords-list">${keywordTags}</div>
          </div>
          <div class="lookalike-info-box">
            <p class="lookalike-info-text">
              <strong>Be careful:</strong> Verify this email is legitimate before clicking links, downloading attachments, or taking any action.
            </p>
          </div>
        `;
      } else if (w.fromEmail && w.toEmail) {
        // Email comparison warnings
        const toClass = (w.type === 'replyto-mismatch' || w.type === 'display-name') ? 'lookalike-dangerous' : 'lookalike-known';
        contentHtml = `
          <div class="lookalike-comparison">
            <div class="lookalike-email-row">
              <span class="lookalike-email-label">${w.labelFrom}</span>
              <div class="lookalike-email lookalike-suspicious">${w.fromEmail}</div>
            </div>
            <div class="lookalike-email-row">
              <span class="lookalike-email-label">${w.labelTo}</span>
              <div class="lookalike-email ${toClass}">${w.toEmail}</div>
            </div>
            ${w.reason ? `<p class="lookalike-reason">${w.reason}</p>` : ''}
          </div>
        `;
      }
      
      return `
        <div class="lookalike-warning-item ${w.severity}" data-type="${w.type}">
          <div class="lookalike-warning-item-header">
            <span class="lookalike-warning-item-icon">${w.type === 'wire-fraud' ? '💰' : '⚠'}</span>
            <div class="lookalike-warning-item-title">
              <strong>${w.title}</strong>
              <p>${w.description}</p>
            </div>
          </div>
          ${contentHtml}
        </div>
      `;
    }).join('');
    
    const warningCount = warnings.length;
    
    banner.innerHTML = `
      <div class="lookalike-warning-container warning-critical">
        <div class="lookalike-warning-inner">
          <div class="lookalike-header">
            <div class="lookalike-icon">🚨</div>
            <div class="lookalike-header-content">
              <h3 class="lookalike-title">${warningCount} Security Warning${warningCount > 1 ? 's' : ''} Detected</h3>
              <a href="https://emailfraudalert.com/learn.html" target="_blank" class="lookalike-learn-link" style="display:block;font-size:13px;color:#0066cc;text-decoration:none;margin-top:4px;">See how this scam works →</a>
            </div>
            <button class="lookalike-close" aria-label="Dismiss warning">×</button>
          </div>
          
          <div class="lookalike-warnings-list">
            ${warningItemsHtml}
          </div>
          
          <div class="lookalike-actions">
            <button class="lookalike-btn lookalike-btn-safe" data-action="trust">
              ✓ Add to Safe Senders
            </button>
            <button class="lookalike-btn lookalike-btn-block" data-action="block">
              🚫 Block This Sender
            </button>
          </div>
        </div>
      </div>
    `;
    
    document.body.appendChild(backdrop);
    document.body.appendChild(banner);
    
    const removeWarning = () => {
      backdrop.remove();
      banner.remove();
    };
    
    banner.querySelector('.lookalike-close').addEventListener('click', removeWarning);
    
    banner.querySelector('[data-action="trust"]').addEventListener('click', async () => {
      skipNextRecheck = true;
      await addKnownContact(senderEmail);
      removeWarning();
    });
    
    banner.querySelector('[data-action="block"]').addEventListener('click', async () => {
      await blockSender(senderEmail);
      removeWarning();
    });
  }
  
  /**
   * Listen for messages from popup
   */
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'getCurrentAccount') {
      sendResponse({ account: currentAccount, isSetup: accountSetup });
    }
    return true;
  });
  
  // Start
  init();
})();
