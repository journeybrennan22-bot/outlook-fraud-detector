/**
 * Gmail Lookalike Detector - Content Script
 * v3.9.0 - Organization impersonation, keyword categories, false positive fixes
 */

(function() {
  'use strict';
  
  // Storage keys
  const STATS_KEY = 'emailLookalikeStats';
  const BLOCKED_KEY = 'emailLookalikeBlocked';
  const SEEN_KEY = 'emailLookalikeSeenSenders';
  
  // Trusted domains (skip many checks for these)
  const TRUSTED_DOMAINS = ['baynac.com', 'purelogicescrow.com', 'journeyinsurance.com'];
  
  // ============================================
  // ORGANIZATION IMPERSONATION TARGETS
  // Maps commonly impersonated entities to their legitimate domains
  // Only checks DISPLAY NAME (not subject/body to avoid false positives)
  // ============================================
  const IMPERSONATION_TARGETS = {
    // US Government - Federal
    "social security": ["ssa.gov"],
    "social security administration": ["ssa.gov"],
    "internal revenue service": ["irs.gov"],
    "irs": ["irs.gov"],
    "treasury department": ["treasury.gov"],
    "us treasury": ["treasury.gov"],
    "department of treasury": ["treasury.gov"],
    "medicare": ["medicare.gov", "cms.gov"],
    "medicaid": ["medicaid.gov", "cms.gov"],
    "federal bureau of investigation": ["fbi.gov"],
    "fbi": ["fbi.gov"],
    "veterans affairs": ["va.gov"],
    "department of veterans affairs": ["va.gov"],
    "va benefits": ["va.gov"],
    "federal trade commission": ["ftc.gov"],
    "ftc": ["ftc.gov"],
    "department of homeland security": ["dhs.gov"],
    "homeland security": ["dhs.gov"],
    "uscis": ["uscis.gov"],
    "us citizenship": ["uscis.gov"],
    "department of justice": ["justice.gov", "usdoj.gov"],
    "department of labor": ["dol.gov"],
    "small business administration": ["sba.gov"],
    "sba": ["sba.gov"],
    "federal housing administration": ["hud.gov"],
    "hud": ["hud.gov"],
    "student aid": ["studentaid.gov", "ed.gov"],
    "fafsa": ["studentaid.gov", "ed.gov"],
    "department of education": ["ed.gov"],

    // Shipping & Postal
    "usps": ["usps.com"],
    "postal service": ["usps.com"],
    "us postal service": ["usps.com"],
    "united states postal": ["usps.com"],
    "ups": ["ups.com"],
    "united parcel service": ["ups.com"],
    "fedex": ["fedex.com"],
    "federal express": ["fedex.com"],
    "dhl": ["dhl.com"],

    // Major Banks - Full names only (not standalone like "chase")
    "wells fargo": ["wellsfargo.com", "wf.com", "notify.wellsfargo.com"],
    "bank of america": ["bankofamerica.com", "bofa.com"],
    "chase bank": ["chase.com", "jpmorganchase.com"],
    "jpmorgan chase": ["chase.com", "jpmorganchase.com"],
    "jpmorgan": ["chase.com", "jpmorganchase.com"],
    "citibank": ["citi.com", "citibank.com"],
    "citigroup": ["citi.com", "citibank.com"],
    "us bank": ["usbank.com"],
    "u.s. bank": ["usbank.com"],
    "pnc bank": ["pnc.com"],
    "capital one": ["capitalone.com"],
    "td bank": ["td.com", "tdbank.com"],
    "truist": ["truist.com"],
    "regions bank": ["regions.com"],
    "fifth third bank": ["53.com"],
    "huntington bank": ["huntington.com"],
    "ally bank": ["ally.com"],
    "discover bank": ["discover.com"],
    "american express": ["americanexpress.com", "amex.com", "aexp.com"],
    "navy federal": ["navyfederal.org"],
    "navy federal credit union": ["navyfederal.org"],
    "usaa": ["usaa.com"],

    // Tech Companies - PHRASES ONLY (not standalone words like "microsoft" or "apple")
    "microsoft support": ["microsoft.com"],
    "microsoft account": ["microsoft.com", "live.com"],
    "microsoft security": ["microsoft.com"],
    "apple support": ["apple.com"],
    "apple id": ["apple.com"],
    "apple security": ["apple.com"],
    "google support": ["google.com"],
    "google account": ["google.com"],
    "google security": ["google.com"],
    "amazon support": ["amazon.com"],
    "amazon account": ["amazon.com"],
    "amazon security": ["amazon.com"],
    "netflix support": ["netflix.com"],
    "netflix account": ["netflix.com"],

    // Document Signing / Business Tools - Keep full detection
    "docusign": ["docusign.com", "docusign.net"],
    "adobe sign": ["adobe.com", "adobesign.com"],
    "intuit": ["intuit.com"],
    "quickbooks": ["intuit.com", "quickbooks.com"],
    "turbotax": ["intuit.com", "turbotax.com"],

    // Payment Platforms
    "paypal": ["paypal.com"],
    "venmo": ["venmo.com"],
    "zelle": ["zellepay.com"],
    "cash app": ["cash.app", "square.com"],
    "cashapp": ["cash.app", "square.com"],

    // Credit Bureaus
    "equifax": ["equifax.com"],
    "experian": ["experian.com"],
    "transunion": ["transunion.com"],

    // Title & Escrow Companies
    "fidelity national title": ["fnf.com", "fntg.com"],
    "first american title": ["firstam.com"],
    "first american": ["firstam.com"],
    "chicago title": ["chicagotitle.com", "fnf.com"],
    "stewart title": ["stewart.com"],
    "old republic title": ["oldrepublictitle.com", "oldrepublic.com"]
  };

  // Suspicious words commonly used in fake domains
  // REMOVED: 'team' (too common in legitimate business names)
  const SUSPICIOUS_DOMAIN_WORDS = [
    'secure', 'security', 'verify', 'verification', 'login', 'signin', 'signon',
    'alert', 'alerts', 'support', 'helpdesk', 'service', 'services',
    'account', 'accounts', 'update', 'confirm', 'confirmation',
    'billing', 'payment', 'invoice', 'refund', 'claim',
    'unlock', 'suspended', 'locked', 'validate',
    'official', 'authentic', 'legit', 'real', 'genuine',
    'dept', 'department', 'center', 'centre',
    'online', 'web', 'portal', 'access', 'customer'
  ];

  // Suspicious display name patterns (suggest impersonation)
  const SUSPICIOUS_DISPLAY_PATTERNS = [
    'security', 'fraud', 'alert', 'support', 'helpdesk', 'help desk',
    'customer service', 'account team', 'billing', 'verification',
    'department', 'official', 'admin', 'administrator',
    'no-reply', 'noreply', 'do not reply', 'automated',
    'urgent', 'important', 'action required', 'immediate'
  ];

  // Deceptive TLDs that look like .com but aren't
  // REMOVED: '.co' (legitimate Colombia TLD, heavily marketed as .com alternative)
  const DECEPTIVE_TLDS = [
    '.com.co', '.com.br', '.com.mx', '.com.ar', '.com.au', '.com.ng',
    '.com.pk', '.com.ph', '.com.ua', '.com.ve', '.com.vn', '.com.tr',
    '.net.co', '.net.br', '.org.co', '.co.uk.com', '.us.com',
    '.cm', '.cc', '.ru', '.cn', '.tk', '.ml', '.ga', '.cf'
  ];

  // ============================================
  // KEYWORD CATEGORIES WITH EXPLANATIONS
  // ============================================
  const KEYWORD_CATEGORIES = {
    'Wire & Payment Methods': {
      keywords: [
        'wire transfer', 'wire instructions', 'wiring instructions',
        'wire information', 'wire details', 'updated wire',
        'new wire', 'wire account', 'wire funds',
        'ach transfer', 'direct deposit',
        'zelle', 'venmo', 'cryptocurrency', 'bitcoin',
        'send funds', 'transfer funds', 'remit funds',
        'wire to', 'remittance', 'wire payment'
      ],
      explanation: 'Emails requesting money transfers are prime targets for fraud. Always verify payment requests by calling a known number before sending funds.'
    },
    'Banking Details': {
      keywords: [
        'bank account', 'account number', 'routing number',
        'aba number', 'swift code', 'iban',
        'bank statement', 'voided check', 'beneficiary'
      ],
      explanation: 'Requests for banking information via email are risky. Scammers use this data to redirect payments or steal funds.'
    },
    'Account Changes': {
      keywords: [
        'updated bank', 'new bank', 'changed bank',
        'updated payment', 'new payment info',
        'changed account', 'new account details',
        'payment update', 'revised instructions',
        'please update your records'
      ],
      explanation: 'Last-minute changes to payment details are the #1 sign of wire fraud. Always verify changes by phone before proceeding.'
    },
    'Real Estate & Legal': {
      keywords: [
        'closing funds', 'earnest money', 'escrow funds',
        'settlement funds', 'settlement payment',
        'retainer', 'trust account', 'iolta',
        'client funds', 'case settlement',
        'court filing fee', 'legal fee'
      ],
      explanation: 'Real estate and legal transactions are heavily targeted by scammers. Verify all payment instructions directly with your escrow officer or attorney.'
    },
    'Secrecy Tactics': {
      keywords: [
        'keep this confidential', 'keep this quiet',
        'dont mention this', 'between us',
        'dont tell anyone', 'private matter',
        'off the record', 'handle personally'
      ],
      explanation: "Requests for secrecy are a major red flag. Legitimate transactions don't require you to bypass normal verification procedures."
    },
    'Sensitive Data Requests': {
      keywords: [
        'social security', 'ssn', 'tax id',
        'W-9', 'W9', 'ein number',
        'login credentials', 'password reset',
        'verify your account', 'verify immediately',
        'confirm your identity', 'verify your identity'
      ],
      explanation: 'Requests for sensitive personal information via email may be phishing attempts. Verify the request through a known phone number.'
    },
    'Authority Impersonation': {
      keywords: [
        'ceo request', 'cfo request', 'owner request',
        'boss asked', 'executive request', 'president asked'
      ],
      explanation: 'Scammers impersonate executives to pressure urgent payments. Verify any unusual requests directly with the person through a known channel.'
    },
    'Urgency Tactics': {
      keywords: [
        'act now', 'urgent action required',
        'account suspended', 'account will be closed',
        'unusual activity', 'suspicious activity', 'unauthorized access',
        'action required within', 'expires today', 'last chance'
      ],
      explanation: 'False urgency is a common fraud tactic designed to prevent you from verifying details. Legitimate requests allow time to confirm.'
    }
  };

  // Build flat keyword list for detection
  const WIRE_FRAUD_KEYWORDS = Object.values(KEYWORD_CATEGORIES).flatMap(cat => cat.keywords);

  // Helper function to get explanation for a keyword
  function getKeywordExplanation(keyword) {
    const lowerKeyword = keyword.toLowerCase();
    for (const [category, data] of Object.entries(KEYWORD_CATEGORIES)) {
      if (data.keywords.some(k => k.toLowerCase() === lowerKeyword)) {
        return {
          category: category,
          explanation: data.explanation
        };
      }
    }
    return {
      category: 'Suspicious Content',
      explanation: 'This email contains terms that may indicate fraud. Verify any requests through a known phone number.'
    };
  }

  // Homoglyph characters (Cyrillic)
  const HOMOGLYPHS = {
    'а': 'a', 'е': 'e', 'о': 'o', 'р': 'p', 'с': 'c', 'х': 'x',
    'і': 'i', 'ј': 'j', 'ѕ': 's', 'ԁ': 'd', 'ɡ': 'g', 'ո': 'n',
    'ν': 'v', 'ѡ': 'w', 'у': 'y', 'һ': 'h', 'ⅼ': 'l', 'ｍ': 'm'
  };

  // ============================================
  // STATE
  // ============================================
  let knownContacts = [];
  let blockedSenders = [];
  let seenSenders = [];
  let emailsScanned = 0;
  let currentAccount = null;
  let accountSetup = false;
  let checkEmailTimeout = null;
  let lastCheckedUrl = null;
  let skipNextRecheck = false;
  let accountDetectionRetries = 0;
  const MAX_ACCOUNT_RETRIES = 10;

  // ============================================
  // HELPER FUNCTIONS
  // ============================================
  
  function isTrustedDomain(domain) {
    return TRUSTED_DOMAINS.includes(domain.toLowerCase());
  }

  function escapeRegex(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function formatEntityName(name) {
    return name.split(' ').map(word => 
      word.charAt(0).toUpperCase() + word.slice(1)
    ).join(' ');
  }

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

  // ============================================
  // ACCOUNT DETECTION
  // ============================================
  
  function detectCurrentAccount() {
    const profileButton = document.querySelector('[data-email]');
    if (profileButton) {
      const email = profileButton.getAttribute('data-email');
      if (email && email.includes('@')) {
        return email.toLowerCase().trim();
      }
    }
    
    const tooltipElements = document.querySelectorAll('[data-tooltip*="@"], [aria-label*="@"]');
    for (const el of tooltipElements) {
      const tooltip = el.getAttribute('data-tooltip') || el.getAttribute('aria-label');
      const emailMatch = tooltip.match(/([a-zA-Z0-9._+-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/);
      if (emailMatch) {
        return emailMatch[1].toLowerCase().trim();
      }
    }
    
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
    
    return null;
  }

  // ============================================
  // DETECTION FUNCTIONS
  // ============================================

  /**
   * Detect organization impersonation - ONLY checks display name
   */
  function detectOrganizationImpersonation(displayName, senderDomain) {
    if (!displayName || !senderDomain) return null;
    if (isTrustedDomain(senderDomain)) return null;
    
    const searchText = displayName.toLowerCase();
    
    for (const [entityName, legitimateDomains] of Object.entries(IMPERSONATION_TARGETS)) {
      const entityPattern = new RegExp(`\\b${escapeRegex(entityName)}\\b`, 'i');
      
      if (entityPattern.test(searchText)) {
        const isLegitimate = legitimateDomains.some(legit => {
          return senderDomain === legit || senderDomain.endsWith(`.${legit}`);
        });
        
        if (!isLegitimate) {
          return {
            entityClaimed: formatEntityName(entityName),
            senderDomain: senderDomain,
            legitimateDomains: legitimateDomains,
            message: `Sender claims to be "${formatEntityName(entityName)}" but email comes from ${senderDomain}. Legitimate emails come from: ${legitimateDomains.join(', ')}`
          };
        }
      }
    }
    
    return null;
  }

  /**
   * Detect deceptive TLDs
   */
  function detectDeceptiveTLD(domain) {
    const domainLower = domain.toLowerCase();
    for (const tld of DECEPTIVE_TLDS) {
      if (domainLower.endsWith(tld)) {
        return tld;
      }
    }
    return null;
  }

  /**
   * Detect suspicious domain patterns
   * Only flags hyphenated domains if they contain suspicious words
   */
  function detectSuspiciousDomain(domain) {
    const domainLower = domain.toLowerCase();
    const domainName = domainLower.split('.')[0];
    
    // Only check hyphenated domains if they contain suspicious words
    if (domainName.includes('-')) {
      const parts = domainName.split('-');
      for (const part of parts) {
        for (const word of SUSPICIOUS_DOMAIN_WORDS) {
          if (part === word) {
            return {
              pattern: word,
              reason: `Domain contains "-${word}" which is commonly used in phishing attacks`
            };
          }
        }
      }
      // Don't flag ALL hyphenated domains - only those with suspicious words
      return null;
    }
    
    // Check for suspicious words as suffixes in non-hyphenated domains
    for (const word of SUSPICIOUS_DOMAIN_WORDS) {
      if (domainName.endsWith(word) && domainName !== word && domainName.length > word.length + 3) {
        return {
          pattern: word,
          reason: `Domain ends with "${word}" which is commonly used in phishing attacks`
        };
      }
    }
    
    return null;
  }

  /**
   * Detect suspicious display names
   */
  function detectSuspiciousDisplayName(displayName, senderDomain) {
    if (!displayName) return null;
    
    const nameLower = displayName.toLowerCase();
    const domainLower = senderDomain.toLowerCase();
    
    const genericDomains = [
      'gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'aol.com',
      'icloud.com', 'mail.com', 'protonmail.com', 'zoho.com', 'yandex.com',
      'live.com', 'msn.com', 'me.com', 'inbox.com'
    ];
    
    const isGenericDomain = genericDomains.includes(domainLower);
    
    for (const pattern of SUSPICIOUS_DISPLAY_PATTERNS) {
      if (nameLower.includes(pattern) && isGenericDomain) {
        return {
          pattern: pattern,
          reason: `Display name contains "${pattern}" but email is from ${senderDomain}. Legitimate companies don't use free email services.`
        };
      }
    }
    
    // Only flag high-risk words from generic domains
    const companyPatterns = ['security', 'billing', 'account', 'verification', 'fraud alert', 'helpdesk'];
    for (const pattern of companyPatterns) {
      if (nameLower.includes(pattern) && isGenericDomain) {
        return {
          pattern: pattern,
          reason: `"${displayName}" sounds official but is from a free email provider (${senderDomain})`
        };
      }
    }
    
    return null;
  }

  /**
   * Detect display name impersonation of trusted domains
   */
  function detectDisplayNameImpersonation(displayName, senderDomain) {
    if (!displayName) return null;
    
    const nameLower = displayName.toLowerCase();
    
    for (const domain of TRUSTED_DOMAINS) {
      if (nameLower.includes(domain) && senderDomain !== domain) {
        return {
          reason: `The display name shows a different email address than the actual sender.`,
          impersonatedDomain: domain
        };
      }
    }
    
    const emailPattern = /[\w.-]+@[\w.-]+\.\w+/;
    const match = displayName.match(emailPattern);
    if (match) {
      const nameEmail = match[0].toLowerCase();
      if (!nameEmail.includes(senderDomain)) {
        return {
          reason: `The display name shows a different email address than the actual sender.`,
          impersonatedDomain: nameEmail
        };
      }
    }
    
    return null;
  }

  /**
   * Detect homoglyphs
   */
  function detectHomoglyphs(email) {
    let found = [];
    for (const [homoglyph, latin] of Object.entries(HOMOGLYPHS)) {
      if (email.includes(homoglyph)) {
        found.push(`"${homoglyph}" looks like "${latin}"`);
      }
    }
    return found.length > 0 ? found.join(', ') : null;
  }

  /**
   * Detect lookalike domains (similar to trusted domains)
   */
  function detectLookalikeDomain(domain) {
    for (const trusted of TRUSTED_DOMAINS) {
      const distance = levenshteinDistance(domain, trusted);
      if (distance > 0 && distance <= 2) {
        return { trustedDomain: trusted, distance: distance };
      }
    }
    return null;
  }

  /**
   * Detect wire fraud keywords
   */
  function detectWireFraudKeywords(content) {
    const found = [];
    for (const keyword of WIRE_FRAUD_KEYWORDS) {
      if (content.toLowerCase().includes(keyword.toLowerCase())) {
        found.push(keyword);
      }
    }
    return found;
  }

  /**
   * Detect contact lookalike - skip if sender is on trusted domain
   */
  function detectContactLookalike(senderEmail) {
    const parts = senderEmail.toLowerCase().split('@');
    if (parts.length !== 2) return null;
    
    const senderLocal = parts[0];
    const senderDomain = parts[1];
    
    // Skip lookalike detection for trusted domains
    if (isTrustedDomain(senderDomain)) return null;
    
    const publicDomains = ['gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'aol.com', 
                           'icloud.com', 'mail.com', 'protonmail.com', 'zoho.com', 'yandex.com'];
    
    for (const contact of knownContacts) {
      if (contact === senderEmail) continue;
      
      const contactParts = contact.toLowerCase().split('@');
      if (contactParts.length !== 2) continue;
      
      const contactLocal = contactParts[0];
      const contactDomain = contactParts[1];
      
      const usernameDiff = levenshteinDistance(senderLocal, contactLocal);
      
      // Same domain, similar username (1-4 chars different)
      if (senderDomain === contactDomain) {
        if (usernameDiff > 0 && usernameDiff <= 4) {
          return {
            incomingEmail: senderEmail,
            matchedContact: contact,
            reason: `Username is ${usernameDiff} character${usernameDiff > 1 ? 's' : ''} different`
          };
        }
      }
      
      // Different domains - check domain similarity
      const bothPublicSameDomain = publicDomains.includes(senderDomain) && 
                                    senderDomain === contactDomain;
      
      if (!bothPublicSameDomain || usernameDiff <= 4) {
        const domainDistance = levenshteinDistance(senderDomain, contactDomain);
        if (domainDistance > 0 && domainDistance <= 2) {
          return {
            incomingEmail: senderEmail,
            matchedContact: contact,
            reason: `Domain is ${domainDistance} character${domainDistance > 1 ? 's' : ''} different`
          };
        }
      }
    }
    
    return null;
  }

  // ============================================
  // EMAIL EXTRACTION
  // ============================================
  
  function extractSenderInfo() {
    // Try multiple selectors for sender info
    const senderSelectors = [
      '.gD[email]',
      '[email]',
      '.go[email]',
      'span[email]'
    ];
    
    let senderElement = null;
    for (const selector of senderSelectors) {
      senderElement = document.querySelector(selector);
      if (senderElement) break;
    }
    
    if (!senderElement) return null;
    
    const email = senderElement.getAttribute('email')?.toLowerCase().trim();
    const displayName = senderElement.getAttribute('name') || senderElement.innerText || '';
    
    if (!email) return null;
    
    return { email, displayName };
  }

  function extractEmailContent() {
    const bodySelectors = [
      '.a3s.aiL',
      '.ii.gt',
      '.a3s',
      'div[data-message-id]',
      '.nH .aHU .a3s'
    ];
    
    let emailBody = null;
    for (const selector of bodySelectors) {
      emailBody = document.querySelector(selector);
      if (emailBody && emailBody.innerText.trim()) break;
    }
    
    const subjectEl = document.querySelector('.hP, h2.hP');
    
    let content = '';
    if (emailBody) content += emailBody.innerText + ' ';
    if (subjectEl) content += subjectEl.innerText;
    
    return content;
  }

  function extractReplyTo() {
    // Gmail doesn't easily expose Reply-To, but we can check expanded headers
    const headerRows = document.querySelectorAll('.ajy, .gH .gI');
    for (const row of headerRows) {
      if (row.innerText.toLowerCase().includes('reply-to')) {
        const emailMatch = row.innerText.match(/([a-zA-Z0-9._+-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/);
        if (emailMatch) return emailMatch[1].toLowerCase();
      }
    }
    return null;
  }

  // ============================================
  // MAIN ANALYSIS
  // ============================================
  
  function checkCurrentEmail() {
    console.log('DEBUG: checkCurrentEmail called');
    const sender = extractSenderInfo();
    console.log('DEBUG: sender =', sender);
    if (!sender) {
      console.log('DEBUG: No sender found, exiting');
      return;
    }
    
    const senderEmail = sender.email;
    const displayName = sender.displayName;
    const senderDomain = senderEmail.split('@')[1] || '';
    const content = extractEmailContent();
    console.log('DEBUG: content length =', content.length);
    console.log('DEBUG: content preview =', content.substring(0, 200));
    const replyTo = extractReplyTo();
    
    const isKnownContact = knownContacts.includes(senderEmail);
    
    const warnings = [];
    
    // 1. Reply-To Mismatch (only flag if different domain)
    if (replyTo && replyTo !== senderEmail) {
      const replyToDomain = replyTo.split('@')[1] || '';
      if (replyToDomain !== senderDomain) {
        warnings.push({
          type: 'replyto-mismatch',
          title: 'Reply-To Mismatch',
          description: 'Replies will go to a different address than the sender.',
          senderEmail: senderEmail,
          matchedEmail: replyTo
        });
      }
    }
    
    // 2. Organization Impersonation (display name only)
    if (!isTrustedDomain(senderDomain)) {
      const orgImpersonation = detectOrganizationImpersonation(displayName, senderDomain);
      if (orgImpersonation) {
        warnings.push({
          type: 'org-impersonation',
          title: 'Organization Impersonation',
          description: orgImpersonation.message,
          senderEmail: senderEmail,
          entityClaimed: orgImpersonation.entityClaimed,
          legitimateDomains: orgImpersonation.legitimateDomains
        });
      }
    }
    
    // 3. Deceptive TLD
    const deceptiveTld = detectDeceptiveTLD(senderDomain);
    if (deceptiveTld) {
      warnings.push({
        type: 'deceptive-tld',
        title: 'Deceptive Domain',
        description: `This domain uses "${deceptiveTld}" which is designed to look like a legitimate .com address.`,
        senderEmail: senderEmail,
        matchedEmail: deceptiveTld
      });
    }
    
    // 4. Suspicious Domain
    const suspiciousDomain = detectSuspiciousDomain(senderDomain);
    if (suspiciousDomain) {
      warnings.push({
        type: 'suspicious-domain',
        title: 'Suspicious Domain',
        description: suspiciousDomain.reason,
        senderEmail: senderEmail,
        matchedEmail: suspiciousDomain.pattern
      });
    }
    
    // 5. Display Name Suspicion
    if (!isKnownContact) {
      const displaySuspicion = detectSuspiciousDisplayName(displayName, senderDomain);
      if (displaySuspicion) {
        warnings.push({
          type: 'display-name-suspicion',
          title: 'Suspicious Display Name',
          description: displaySuspicion.reason,
          senderEmail: senderEmail,
          matchedEmail: displaySuspicion.pattern
        });
      }
    }
    
    // 6. Display Name Impersonation (trusted domains)
    if (!isKnownContact) {
      const impersonation = detectDisplayNameImpersonation(displayName, senderDomain);
      if (impersonation) {
        warnings.push({
          type: 'impersonation',
          title: 'Display Name Impersonation',
          description: impersonation.reason,
          senderEmail: senderEmail,
          matchedEmail: impersonation.impersonatedDomain
        });
      }
    }
    
    // 7. Homoglyphs
    const homoglyph = detectHomoglyphs(senderEmail);
    if (homoglyph) {
      warnings.push({
        type: 'homoglyph',
        title: 'Invisible Character Trick',
        description: 'This email contains deceptive characters that look identical to normal letters.',
        senderEmail: senderEmail,
        detail: homoglyph
      });
    }
    
    // 8. Lookalike Domain
    const lookalike = detectLookalikeDomain(senderDomain);
    if (lookalike) {
      warnings.push({
        type: 'lookalike-domain',
        title: 'Lookalike Domain',
        description: `This domain is similar to ${lookalike.trustedDomain}`,
        senderEmail: senderEmail,
        matchedEmail: lookalike.trustedDomain
      });
    }
    
    // 9. Wire Fraud Keywords
    const wireKeywords = detectWireFraudKeywords(content);
    if (wireKeywords.length > 0) {
      const keywordInfo = getKeywordExplanation(wireKeywords[0]);
      warnings.push({
        type: 'wire-fraud',
        title: 'Dangerous Keywords Detected',
        description: 'This email contains terms commonly used in wire fraud.',
        keywords: wireKeywords,
        keywordCategory: keywordInfo.category,
        keywordExplanation: keywordInfo.explanation
      });
    }
    
    // 10. Contact Lookalike
    if (!isKnownContact) {
      const contactLookalike = detectContactLookalike(senderEmail);
      if (contactLookalike) {
        warnings.push({
          type: 'contact-lookalike',
          title: 'Lookalike Email Address',
          description: 'This email is nearly identical to someone in your contacts, but slightly different.',
          senderEmail: contactLookalike.incomingEmail,
          matchedEmail: contactLookalike.matchedContact,
          reason: contactLookalike.reason
        });
      }
    }
    
    // Display warnings if any
    if (warnings.length > 0) {
      showWarning(warnings, senderEmail);
    }
    
    emailsScanned++;
  }

  // ============================================
  // UI - WARNING DISPLAY
  // ============================================
  
  function removeExistingWarnings() {
    document.querySelectorAll('.lookalike-warning-banner').forEach(el => el.remove());
    document.querySelectorAll('.lookalike-warning-backdrop').forEach(el => el.remove());
  }

  function showWarning(warnings, senderEmail) {
    removeExistingWarnings();
    
    const backdrop = document.createElement('div');
    backdrop.className = 'lookalike-warning-backdrop';
    
    const banner = document.createElement('div');
    banner.className = 'lookalike-warning-banner';
    
    const warningItemsHtml = warnings.map(w => {
      let detailHtml = '';
      
      if (w.type === 'wire-fraud' && w.keywords) {
        const keywordTags = w.keywords.slice(0, 5).map(k => 
          `<span class="lookalike-keyword-tag">${k}</span>`
        ).join('');
        detailHtml = `
          <div class="lookalike-warning-keywords">
            <div class="lookalike-keywords-label">Triggered by:</div>
            <div class="lookalike-keywords-list">${keywordTags}</div>
          </div>
          <div class="lookalike-warning-advice">
            <strong>Why this matters:</strong> ${w.keywordExplanation}
          </div>
        `;
      } else if (w.type === 'org-impersonation') {
        detailHtml = `
          <div class="lookalike-warning-emails">
            <div class="lookalike-email-row">
              <span class="lookalike-label">Claims to be:</span>
              <span class="lookalike-value safe">${w.entityClaimed}</span>
            </div>
            <div class="lookalike-email-row">
              <span class="lookalike-label">Actually from:</span>
              <span class="lookalike-value suspicious">${w.senderEmail}</span>
            </div>
            <div class="lookalike-email-row">
              <span class="lookalike-label">Legitimate domains:</span>
              <span class="lookalike-value safe">${w.legitimateDomains.join(', ')}</span>
            </div>
          </div>
        `;
      } else if (w.senderEmail && w.matchedEmail) {
        const matchLabel = w.type === 'replyto-mismatch' ? 'Replies go to' : 
                           w.type === 'impersonation' ? 'Display name shows' : 'Similar to';
        detailHtml = `
          <div class="lookalike-warning-emails">
            <div class="lookalike-email-row">
              <span class="lookalike-label">Sender:</span>
              <span class="lookalike-value suspicious">${w.senderEmail}</span>
            </div>
            <div class="lookalike-email-row">
              <span class="lookalike-label">${matchLabel}:</span>
              <span class="lookalike-value safe">${w.matchedEmail}</span>
            </div>
            ${w.reason ? `<div class="lookalike-reason">${w.reason}</div>` : ''}
          </div>
        `;
      } else if (w.detail) {
        detailHtml = `<div class="lookalike-detail">${w.detail}</div>`;
      }
      
      return `
        <div class="lookalike-warning-item">
          <div class="lookalike-warning-title">${w.title}</div>
          <div class="lookalike-warning-description">${w.description}</div>
          ${detailHtml}
        </div>
      `;
    }).join('');
    
    banner.innerHTML = `
      <div class="lookalike-warning-container">
        <div class="lookalike-warning-inner">
          <div class="lookalike-header">
            <div class="lookalike-icon">⚠️</div>
            <div class="lookalike-header-content">
              <h3 class="lookalike-title">${warnings.length} Warning${warnings.length > 1 ? 's' : ''} Detected</h3>
              <a href="https://emailfraudalert.com/learn.html?v=2" target="_blank" class="lookalike-learn-link">See how this scam works →</a>
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

  // ============================================
  // STORAGE FUNCTIONS
  // ============================================
  
  async function loadData() {
    return new Promise(resolve => {
      chrome.storage.local.get([STATS_KEY, BLOCKED_KEY, SEEN_KEY, `contacts_${currentAccount}`], (result) => {
        emailsScanned = result[STATS_KEY]?.emailsScanned || 0;
        blockedSenders = result[BLOCKED_KEY] || [];
        seenSenders = result[SEEN_KEY] || [];
        knownContacts = result[`contacts_${currentAccount}`] || [];
        accountSetup = knownContacts.length > 0;
        resolve();
      });
    });
  }

  async function addKnownContact(email) {
    if (!knownContacts.includes(email)) {
      knownContacts.push(email);
      await chrome.storage.local.set({ [`contacts_${currentAccount}`]: knownContacts });
    }
  }

  async function blockSender(email) {
    if (!blockedSenders.includes(email)) {
      blockedSenders.push(email);
      await chrome.storage.local.set({ [BLOCKED_KEY]: blockedSenders });
    }
  }

  // ============================================
  // URL MONITORING
  // ============================================
  
  function isViewingEmail() {
    const url = window.location.href;
    // Gmail email view URLs contain #inbox/ or #sent/ or similar followed by an ID
    return /#[a-z]+\/[A-Za-z0-9]+/.test(url);
  }

  function onUrlChange() {
    const currentUrl = window.location.href;
    if (currentUrl === lastCheckedUrl) return;
    lastCheckedUrl = currentUrl;
    
    removeExistingWarnings();
    
    if (isViewingEmail()) {
      clearTimeout(checkEmailTimeout);
      checkEmailTimeout = setTimeout(() => {
        if (!skipNextRecheck) {
          checkCurrentEmail();
        }
        skipNextRecheck = false;
      }, 500);
    }
  }

  // ============================================
  // INITIALIZATION
  // ============================================
  
  async function init() {
    // Detect account
    currentAccount = detectCurrentAccount();
    
    if (!currentAccount) {
      if (accountDetectionRetries < MAX_ACCOUNT_RETRIES) {
        accountDetectionRetries++;
        setTimeout(init, 1000);
        return;
      }
      console.log('Email Fraud Alert: Could not detect Gmail account');
      return;
    }
    
    // Load data
    await loadData();
    
    // Set up URL monitoring
    window.addEventListener('hashchange', onUrlChange);
    
    // Also use MutationObserver for SPA navigation
    const observer = new MutationObserver(() => {
      setTimeout(onUrlChange, 100);
    });
    observer.observe(document.body, { childList: true, subtree: true });
    
    // Initial check
    onUrlChange();
    
    console.log('Email Fraud Alert v3.9.0 loaded for', currentAccount);
  }

  // Listen for messages from popup
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'getCurrentAccount') {
      sendResponse({ account: currentAccount, isSetup: accountSetup });
    }
    if (request.action === 'resync') {
      loadData().then(() => {
        sendResponse({ success: true, contacts: knownContacts.length });
      });
      return true;
    }
    return true;
  });

  // Start
  init();
})();
