// Email Fraud Detector - Outlook Web Add-in
// Version 3.2.0 - Organization impersonation detection

// ============================================
// CONFIGURATION
// ============================================
const CONFIG = {
    clientId: '622f0452-d622-45d1-aab3-3a2026389dd3',
    redirectUri: 'https://journeybrennan22-bot.github.io/outlook-fraud-detector/src/taskpane.html',
    scopes: ['User.Read', 'Contacts.Read'],
    trustedDomains: ['baynac.com', 'purelogicescrow.com', 'journeyinsurance.com']
};

// Suspicious words commonly used in fake domains
const SUSPICIOUS_DOMAIN_WORDS = [
    'secure', 'security', 'verify', 'verification', 'login', 'signin', 'signon',
    'alert', 'alerts', 'support', 'helpdesk', 'service', 'services',
    'account', 'accounts', 'update', 'confirm', 'confirmation',
    'billing', 'payment', 'invoice', 'refund', 'claim',
    'unlock', 'suspended', 'locked', 'verify', 'validate',
    'official', 'authentic', 'legit', 'real', 'genuine',
    'team', 'dept', 'department', 'center', 'centre',
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
const DECEPTIVE_TLDS = [
    '.com.co', '.com.br', '.com.mx', '.com.ar', '.com.au', '.com.ng',
    '.com.pk', '.com.ph', '.com.ua', '.com.ve', '.com.vn', '.com.tr',
    '.net.co', '.net.br', '.org.co', '.co.uk.com', '.us.com',
    '.co', '.cm', '.cc', '.ru', '.cn', '.tk', '.ml', '.ga', '.cf'
];

// ============================================
// ORGANIZATION IMPERSONATION TARGETS
// Maps commonly impersonated entities to their legitimate domains
// ============================================
const IMPERSONATION_TARGETS = {
    // US Government - Federal
    "social security": ["ssa.gov"],
    "social security administration": ["ssa.gov"],
    "internal revenue service": ["irs.gov"],
    "irs": ["irs.gov"],
    "treasury department": ["treasury.gov"],
    "us treasury": ["treasury.gov"],
    "medicare": ["medicare.gov", "cms.gov"],
    "medicaid": ["medicaid.gov", "cms.gov"],
    "veterans affairs": ["va.gov"],
    "federal trade commission": ["ftc.gov"],
    "ftc": ["ftc.gov"],
    "homeland security": ["dhs.gov"],
    "uscis": ["uscis.gov"],
    "department of justice": ["justice.gov"],
    "fbi": ["fbi.gov"],
    "usps": ["usps.com"],
    "us postal service": ["usps.com"],
    
    // Major Banks
    "chase bank": ["chase.com"],
    "jpmorgan": ["jpmorgan.com", "chase.com"],
    "bank of america": ["bankofamerica.com"],
    "wells fargo": ["wellsfargo.com"],
    "citibank": ["citi.com", "citibank.com"],
    "citi": ["citi.com", "citibank.com"],
    "us bank": ["usbank.com"],
    "capital one": ["capitalone.com"],
    "pnc bank": ["pnc.com"],
    "td bank": ["td.com", "tdbank.com"],
    "truist": ["truist.com"],
    
    // Payment & Financial Services
    "paypal": ["paypal.com"],
    "venmo": ["venmo.com"],
    "zelle": ["zellepay.com"],
    "square": ["squareup.com", "square.com"],
    "stripe": ["stripe.com"],
    
    // Title & Escrow (relevant to your industry)
    "first american title": ["firstam.com"],
    "fidelity national title": ["fnf.com"],
    "chicago title": ["ctic.com"],
    "old republic title": ["oldrepublictitle.com"],
    "stewart title": ["stewart.com"],
    
    // Tech Companies (commonly spoofed)
    "microsoft": ["microsoft.com"],
    "apple": ["apple.com"],
    "google": ["google.com"],
    "amazon": ["amazon.com"],
    "netflix": ["netflix.com"],
    "docusign": ["docusign.com", "docusign.net"]
};

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

// Homoglyph characters (Cyrillic only)
const HOMOGLYPHS = {
    'а': 'a', 'е': 'e', 'о': 'o', 'р': 'p', 'с': 'c', 'х': 'x',
    'і': 'i', 'ј': 'j', 'ѕ': 's', 'ԁ': 'd', 'ɡ': 'g', 'ո': 'n',
    'ν': 'v', 'ѡ': 'w', 'у': 'y', 'һ': 'h', 'ⅼ': 'l', 'ｍ': 'm',
    '！': '!', '＠': '@'
};

// ============================================
// STATE
// ============================================
let msalInstance = null;
let knownContacts = new Set();
let currentUserEmail = null;
let currentItemId = null;
let isAutoScanEnabled = true;

// ============================================
// INITIALIZATION
// ============================================
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        initializeMsal();
        setupEventHandlers();
        analyzeCurrentEmail();
        setupAutoScan();
    }
});

function initializeMsal() {
    const msalConfig = {
        auth: {
            clientId: CONFIG.clientId,
            redirectUri: CONFIG.redirectUri,
            authority: 'https://login.microsoftonline.com/common'
        },
        cache: {
            cacheLocation: 'sessionStorage',
            storeAuthStateInCookie: false
        }
    };
    msalInstance = new msal.PublicClientApplication(msalConfig);
}

function setupEventHandlers() {
    document.getElementById('retry-btn').addEventListener('click', analyzeCurrentEmail);
}

function setupAutoScan() {
    if (Office.context.mailbox.addHandlerAsync) {
        Office.context.mailbox.addHandlerAsync(
            Office.EventType.ItemChanged,
            onItemChanged,
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log('Auto-scan enabled');
                }
            }
        );
    }
}

function onItemChanged() {
    if (isAutoScanEnabled) {
        analyzeCurrentEmail();
    }
}

// ============================================
// AUTHENTICATION & DATA FETCHING
// ============================================
async function getAccessToken() {
    if (!msalInstance) return null;
    
    const accounts = msalInstance.getAllAccounts();
    
    try {
        if (accounts.length > 0) {
            const response = await msalInstance.acquireTokenSilent({
                scopes: CONFIG.scopes,
                account: accounts[0]
            });
            return response.accessToken;
        } else {
            const response = await msalInstance.acquireTokenPopup({
                scopes: CONFIG.scopes
            });
            return response.accessToken;
        }
    } catch (error) {
        console.log('Auth error:', error);
        return null;
    }
}

/**
 * Fetch contacts from Microsoft Graph (contacts only - fast)
 */
async function fetchContacts(token) {
    const contacts = [];
    
    try {
        let url = 'https://graph.microsoft.com/v1.0/me/contacts?$top=500&$select=emailAddresses';
        
        while (url) {
            const response = await fetch(url, {
                headers: { 'Authorization': `Bearer ${token}` }
            });
            
            if (!response.ok) break;
            
            const data = await response.json();
            
            if (data.value) {
                data.value.forEach(contact => {
                    if (contact.emailAddresses) {
                        contact.emailAddresses.forEach(email => {
                            if (email.address) {
                                contacts.push(email.address.toLowerCase());
                            }
                        });
                    }
                });
            }
            
            url = data['@odata.nextLink'] || null;
        }
        
        console.log('Fetched', contacts.length, 'contacts');
    } catch (error) {
        console.log('Contacts fetch error:', error);
    }
    
    return contacts;
}

/**
 * Fetch contacts and add current user email
 */
async function fetchAllKnownContacts() {
    const token = await getAccessToken();
    if (!token) return;
    
    console.log('Fetching contacts...');
    
    const contacts = await fetchContacts(token);
    
    // Add all to the knownContacts set
    contacts.forEach(e => knownContacts.add(e));
    
    // Add current user's email so lookalikes of their own address are detected
    if (currentUserEmail) {
        knownContacts.add(currentUserEmail.toLowerCase());
    }
    
    console.log('Total known contacts:', knownContacts.size);
}

// ============================================
// MAIN ANALYSIS
// ============================================
async function analyzeCurrentEmail() {
    showLoading();
    
    try {
        // Get current user email
        currentUserEmail = Office.context.mailbox.userProfile.emailAddress;
        
        // Fetch contacts if not already loaded
        if (knownContacts.size === 0) {
            await fetchAllKnownContacts();
        }
        
        // Get email data
        const item = Office.context.mailbox.item;
        const from = item.from;
        const subject = item.subject;
        
        // Get body and headers
        item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
            // Try to get Reply-To from internet headers
            if (item.getAllInternetHeadersAsync) {
                item.getAllInternetHeadersAsync((headerResult) => {
                    let replyTo = null;
                    if (headerResult.status === Office.AsyncResultStatus.Succeeded) {
                        const headers = headerResult.value;
                        const replyToMatch = headers.match(/^Reply-To:\s*(.+)$/mi);
                        if (replyToMatch) {
                            const emailMatch = replyToMatch[1].match(/<([^>]+)>/) || replyToMatch[1].match(/([^\s,]+@[^\s,]+)/);
                            if (emailMatch) {
                                replyTo = emailMatch[1].trim();
                            }
                        }
                    }
                    
                    const emailData = {
                        from: from,
                        subject: subject || '',
                        body: bodyResult.status === Office.AsyncResultStatus.Succeeded ? bodyResult.value : '',
                        replyTo: replyTo
                    };
                    
                    performAnalysis(emailData);
                });
            } else {
                const emailData = {
                    from: from,
                    subject: subject || '',
                    body: bodyResult.status === Office.AsyncResultStatus.Succeeded ? bodyResult.value : '',
                    replyTo: null
                };
                
                performAnalysis(emailData);
            }
        });
    } catch (error) {
        showError(error.message);
    }
}

function performAnalysis(emailData) {
    const warnings = [];
    
    const senderEmail = emailData.from.emailAddress.toLowerCase();
    const senderDomain = senderEmail.split('@')[1] || '';
    const displayName = emailData.from.displayName || '';
    const subject = emailData.subject || '';
    const body = emailData.body || '';
    const fullContent = (subject + ' ' + body).toLowerCase();
    
    // Skip if sender is in known contacts
    const isKnownContact = knownContacts.has(senderEmail);
    
    // 1. Reply-To Mismatch (only flag if different domain)
    if (emailData.replyTo && emailData.replyTo.toLowerCase() !== senderEmail) {
        const replyToDomain = emailData.replyTo.toLowerCase().split('@')[1] || '';
        if (replyToDomain !== senderDomain) {
            warnings.push({
                type: 'replyto-mismatch',
                severity: 'critical',
                title: 'Reply-To Mismatch',
                description: 'Replies will go to a different address than the sender.',
                senderEmail: senderEmail,
                matchedEmail: emailData.replyTo.toLowerCase()
            });
        }
    }
    
    // 2. Deceptive TLD Detection
    const deceptiveTld = detectDeceptiveTLD(senderDomain);
    if (deceptiveTld) {
        warnings.push({
            type: 'deceptive-tld',
            severity: 'critical',
            title: 'Deceptive Domain',
            description: `This domain uses "${deceptiveTld}" which is designed to look like a legitimate .com address.`,
            senderEmail: senderEmail,
            matchedEmail: deceptiveTld
        });
    }
    
    // 3. Suspicious Domain Pattern Detection (NEW - pattern based)
    const suspiciousDomain = detectSuspiciousDomain(senderDomain);
    if (suspiciousDomain) {
        warnings.push({
            type: 'suspicious-domain',
            severity: 'critical',
            title: 'Suspicious Domain',
            description: suspiciousDomain.reason,
            senderEmail: senderEmail,
            matchedEmail: suspiciousDomain.pattern
        });
    }
    
    // 4. Display Name Suspicion (NEW - pattern based)
    if (!isKnownContact) {
        const displaySuspicion = detectSuspiciousDisplayName(displayName, senderDomain);
        if (displaySuspicion) {
            warnings.push({
                type: 'display-name-suspicion',
                severity: 'critical',
                title: 'Suspicious Display Name',
                description: displaySuspicion.reason,
                senderEmail: senderEmail,
                matchedEmail: displaySuspicion.pattern
            });
        }
    }
    
    // 5. Display Name Impersonation (trusted domains)
    if (!isKnownContact) {
        const impersonation = detectDisplayNameImpersonation(displayName, senderDomain);
        if (impersonation) {
            warnings.push({
                type: 'impersonation',
                severity: 'critical',
                title: 'Display Name Impersonation',
                description: impersonation.reason,
                senderEmail: senderEmail,
                matchedEmail: impersonation.impersonatedDomain
            });
        }
    }
    
    // 6. Homoglyph/Unicode Detection
    const homoglyph = detectHomoglyphs(senderEmail);
    if (homoglyph) {
        warnings.push({
            type: 'homoglyph',
            severity: 'critical',
            title: 'Invisible Character Trick',
            description: 'This email contains deceptive characters that look identical to normal letters.',
            senderEmail: senderEmail,
            matchedEmail: homoglyph,
            detail: homoglyph
        });
    }
    
    // 7. Lookalike Domain Detection (your trusted domains)
    const lookalike = detectLookalikeDomain(senderDomain);
    if (lookalike) {
        warnings.push({
            type: 'lookalike-domain',
            severity: 'critical',
            title: 'Lookalike Domain',
            description: `This domain is similar to ${lookalike.trustedDomain}`,
            senderEmail: senderEmail,
            matchedEmail: lookalike.trustedDomain
        });
    }
    
    // 8. Fraud Keywords - now with contextual explanations
    const wireKeywords = detectWireFraudKeywords(fullContent);
    if (wireKeywords.length > 0) {
        const keywordInfo = getKeywordExplanation(wireKeywords[0]);
        warnings.push({
            type: 'wire-fraud',
            severity: 'critical',
            title: 'Dangerous Keywords Detected',
            description: 'This email contains terms commonly used in wire fraud.',
            keywords: wireKeywords,
            isWireFraud: true,
            keywordCategory: keywordInfo.category,
            keywordExplanation: keywordInfo.explanation
        });
    }
    
    // 9. Contact Lookalike Detection (skip if known contact)
    if (!isKnownContact) {
        const contactLookalike = detectContactLookalike(senderEmail);
        if (contactLookalike) {
            warnings.push({
                type: 'contact-lookalike',
                severity: 'critical',
                title: 'Lookalike Email Address',
                description: 'This email is nearly identical to someone in your contacts, but slightly different.',
                senderEmail: contactLookalike.incomingEmail,
                matchedEmail: contactLookalike.matchedContact,
                reason: contactLookalike.reason
            });
        }
    }
    
    // 10. Organization Impersonation Detection (SSA, IRS, banks, etc.)
    // Skip for trusted domains - no need to check our own company
    const orgImpersonation = isTrustedDomain(senderDomain) ? null : detectOrganizationImpersonation(displayName, subject, senderDomain);
    if (orgImpersonation) {
        warnings.push({
            type: 'org-impersonation',
            severity: 'critical',
            title: 'Organization Impersonation',
            description: `This email claims to be from ${orgImpersonation.entityClaimed} but is not coming from a legitimate domain.`,
            senderEmail: senderEmail,
            matchedEmail: orgImpersonation.legitimateDomains.join(', '),
            claimedEntity: orgImpersonation.entityClaimed,
            legitimateDomains: orgImpersonation.legitimateDomains
        });
    }
    
    displayResults(warnings, senderEmail);
}

// ============================================
// DETECTION FUNCTIONS
// ============================================

/**
 * Detect organization impersonation (e.g., fake SSA, IRS, banks)
 * Checks if display name or subject claims to be a known entity
 * but email comes from a non-legitimate domain
 */
function detectOrganizationImpersonation(displayName, subject, senderDomain) {
    const searchText = ((displayName || '') + ' ' + (subject || '')).toLowerCase();
    const domainLower = senderDomain.toLowerCase();
    
    for (const [entity, legitimateDomains] of Object.entries(IMPERSONATION_TARGETS)) {
        if (searchText.includes(entity)) {
            // Check if sender domain matches any legitimate domain
            const isLegitimate = legitimateDomains.some(legit => 
                domainLower === legit || domainLower.endsWith('.' + legit)
            );
            
            if (!isLegitimate) {
                // Format entity name for display (capitalize)
                const formattedEntity = entity.split(' ')
                    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
                    .join(' ');
                
                return {
                    entityClaimed: formattedEntity,
                    senderDomain: senderDomain,
                    legitimateDomains: legitimateDomains
                };
            }
        }
    }
    
    return null;
}

/**
 * Detect deceptive TLDs like .com.co, .com.br, etc.
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
 * Detect suspicious domain patterns (hyphenated domains with security words)
 */
function detectSuspiciousDomain(domain) {
    const domainLower = domain.toLowerCase();
    const domainName = domainLower.split('.')[0]; // Get name before TLD
    
    // Check for hyphenated domains with suspicious words
    if (domainName.includes('-')) {
        const parts = domainName.split('-');
        for (const part of parts) {
            for (const word of SUSPICIOUS_DOMAIN_WORDS) {
                if (part === word || part.includes(word)) {
                    return {
                        pattern: word,
                        reason: `Domain contains "-${word}" which is commonly used in phishing attacks`
                    };
                }
            }
        }
        
        // Any hyphenated domain is slightly suspicious
        return {
            pattern: 'hyphenated domain',
            reason: `Hyphenated domains like "${domainName}" are commonly used in phishing. Verify this sender.`
        };
    }
    
    // Check for suspicious words anywhere in non-hyphenated domain
    for (const word of SUSPICIOUS_DOMAIN_WORDS) {
        // Only flag if the word is a suffix (like "chasesecure.com" or "paypalverify.com")
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
 * Detect suspicious display names that suggest impersonation
 */
function detectSuspiciousDisplayName(displayName, senderDomain) {
    if (!displayName) return null;
    
    const nameLower = displayName.toLowerCase();
    const domainLower = senderDomain.toLowerCase();
    
    // List of generic/free email domains
    const genericDomains = [
        'gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'aol.com',
        'icloud.com', 'mail.com', 'protonmail.com', 'zoho.com', 'yandex.com',
        'live.com', 'msn.com', 'me.com', 'inbox.com'
    ];
    
    const isGenericDomain = genericDomains.includes(domainLower);
    
    // Check for suspicious patterns in display name
    for (const pattern of SUSPICIOUS_DISPLAY_PATTERNS) {
        if (nameLower.includes(pattern)) {
            // If display name has official-sounding words but comes from generic email
            if (isGenericDomain) {
                return {
                    pattern: pattern,
                    reason: `Display name contains "${pattern}" but email is from ${senderDomain}. Legitimate companies don't use free email services.`
                };
            }
        }
    }
    
    // Check if display name looks like a company but domain doesn't match
    // Only flag truly suspicious words - not common business words like "service", "team", "support"
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

function detectDisplayNameImpersonation(displayName, senderDomain) {
    if (!displayName) return null;
    
    const nameLower = displayName.toLowerCase();
    
    // Check if display name contains a trusted domain
    for (const domain of CONFIG.trustedDomains) {
        if (nameLower.includes(domain) && senderDomain !== domain) {
            return {
                reason: `The display name shows a different email address than the actual sender.`,
                impersonatedDomain: domain
            };
        }
    }
    
    // Check for email-like patterns in display name
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

function detectHomoglyphs(email) {
    let found = [];
    for (const [homoglyph, latin] of Object.entries(HOMOGLYPHS)) {
        if (email.includes(homoglyph)) {
            found.push(`"${homoglyph}" looks like "${latin}"`);
        }
    }
    return found.length > 0 ? found.join(', ') : null;
}

function detectLookalikeDomain(domain) {
    for (const trusted of CONFIG.trustedDomains) {
        const distance = levenshteinDistance(domain, trusted);
        if (distance > 0 && distance <= 2) {
            return { trustedDomain: trusted, distance: distance };
        }
    }
    return null;
}

function detectWireFraudKeywords(content) {
    const found = [];
    for (const keyword of WIRE_FRAUD_KEYWORDS) {
        if (content.includes(keyword.toLowerCase())) {
            found.push(keyword);
        }
    }
    return found;
}

function detectContactLookalike(incomingEmail) {
    const incomingParts = parseEmailParts(incomingEmail);
    if (!incomingParts) return null;
    
    const publicDomains = ['gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'aol.com', 
                           'icloud.com', 'mail.com', 'protonmail.com', 'zoho.com', 'yandex.com'];
    
    for (const contactEmail of knownContacts) {
        if (contactEmail === incomingEmail) continue;
        
        const contactParts = parseEmailParts(contactEmail);
        if (!contactParts) continue;
        
        // Calculate username difference
        const usernameDiff = levenshteinDistance(incomingParts.local, contactParts.local);
        
        // Same domain, similar username (1-4 chars different)
        if (incomingParts.domain === contactParts.domain) {
            if (usernameDiff > 0 && usernameDiff <= 4) {
                return {
                    incomingEmail: incomingEmail,
                    matchedContact: contactEmail,
                    reason: `Username is ${usernameDiff} character${usernameDiff > 1 ? 's' : ''} different`
                };
            }
        }
        
        // Different domains - check if domain is similar (1-2 chars different)
        // But skip if both on same public domain with very different usernames
        const bothPublicSameDomain = publicDomains.includes(incomingParts.domain) && 
                                      incomingParts.domain === contactParts.domain;
        
        if (!bothPublicSameDomain || usernameDiff <= 4) {
            const domainDistance = levenshteinDistance(incomingParts.domain, contactParts.domain);
            if (domainDistance > 0 && domainDistance <= 2) {
                return {
                    incomingEmail: incomingEmail,
                    matchedContact: contactEmail,
                    reason: `Domain is ${domainDistance} character${domainDistance > 1 ? 's' : ''} different`
                };
            }
        }
    }
    
    return null;
}

function parseEmailParts(email) {
    const parts = email.toLowerCase().split('@');
    if (parts.length !== 2) return null;
    return { local: parts[0], domain: parts[1], full: email.toLowerCase() };
}

function isTrustedDomain(domain) {
    return CONFIG.trustedDomains.includes(domain.toLowerCase());
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
// UI FUNCTIONS
// ============================================
function showLoading() {
    document.getElementById('loading').classList.remove('hidden');
    document.getElementById('results').classList.add('hidden');
    document.getElementById('error').classList.add('hidden');
    document.body.className = '';
}

function showError(message) {
    document.getElementById('loading').classList.add('hidden');
    document.getElementById('results').classList.add('hidden');
    document.getElementById('error').classList.remove('hidden');
    document.getElementById('error-message').textContent = message;
    document.body.className = '';
}

function displayResults(warnings, senderEmail) {
    document.getElementById('loading').classList.add('hidden');
    document.getElementById('error').classList.add('hidden');
    document.getElementById('results').classList.remove('hidden');
    
    // Count warnings by severity
    const criticalCount = warnings.filter(w => w.severity === 'critical').length;
    const mediumCount = warnings.filter(w => w.severity === 'medium').length;
    
    // Set body background and status badge
    document.body.classList.remove('status-critical', 'status-medium', 'status-info', 'status-safe');
    
    const statusBadge = document.getElementById('status-badge');
    const statusIcon = statusBadge.querySelector('.status-icon');
    const statusText = statusBadge.querySelector('.status-text');
    
    if (criticalCount > 0 || mediumCount > 0) {
        const totalWarnings = criticalCount + mediumCount;
        document.body.classList.add('status-critical');
        statusBadge.className = 'status-badge danger';
        statusIcon.textContent = '🚨';
        statusText.textContent = `${totalWarnings} Warning${totalWarnings > 1 ? 's' : ''} Detected`;
    } else {
        document.body.classList.add('status-safe');
        statusBadge.className = 'status-badge safe';
        statusIcon.textContent = '✅';
        statusText.textContent = 'No Issues Detected';
    }
    
    // Display learn link
    const learnLink = document.getElementById('learn-link');
    if (learnLink) {
        if (warnings.length > 0) {
            learnLink.classList.remove('hidden');
        } else {
            learnLink.classList.add('hidden');
        }
    }
    
    // Display warnings
    const warningsSection = document.getElementById('warnings-section');
    const warningsList = document.getElementById('warnings-list');
    const warningsFooter = document.getElementById('warnings-footer');
    const safeMessage = document.getElementById('safe-message');
    
    if (warnings.length > 0) {
        warningsSection.classList.remove('hidden');
        if (warningsFooter) warningsFooter.classList.remove('hidden');
        if (safeMessage) safeMessage.classList.add('hidden');
        
        warningsList.innerHTML = warnings.map(w => {
            let emailHtml = '';
            if (w.isWireFraud && w.keywords) {
                const keywordTags = w.keywords.slice(0, 5).map(k => 
                    `<span class="keyword-tag">${k}</span>`
                ).join('');
                emailHtml = `
                    <div class="warning-keywords-section">
                        <div class="warning-keywords-label">Triggered by:</div>
                        <div class="warning-keywords">${keywordTags}</div>
                    </div>
                    <div class="warning-advice">
                        <strong>Why this matters:</strong> ${w.keywordExplanation}
                    </div>
                `;
            } else if (w.type === 'org-impersonation') {
                emailHtml = `
                    <div class="warning-emails">
                        <div class="warning-email-row">
                            <span class="warning-email-label">Claims to be:</span>
                            <span class="warning-email-value known">${w.claimedEntity}</span>
                        </div>
                        <div class="warning-email-row">
                            <span class="warning-email-label">Actually from:</span>
                            <span class="warning-email-value suspicious">${w.senderEmail}</span>
                        </div>
                        <div class="warning-email-row">
                            <span class="warning-email-label">Legitimate domains:</span>
                            <span class="warning-email-value known">${w.matchedEmail}</span>
                        </div>
                    </div>
                `;
            } else if (w.senderEmail && w.matchedEmail) {
                const matchLabel = w.type === 'replyto-mismatch' ? 'Replies go to' : 
                                   w.type === 'deceptive-tld' ? 'Deceptive TLD' : 
                                   w.type === 'suspicious-domain' ? 'Pattern' :
                                   w.type === 'display-name-suspicion' ? 'Pattern' :
                                   w.type === 'impersonation' ? 'Display name shows' : 'Similar to';
                emailHtml = `
                    <div class="warning-emails">
                        <div class="warning-email-row">
                            <span class="warning-email-label">Sender:</span>
                            <span class="warning-email-value suspicious">${w.senderEmail}</span>
                        </div>
                        <div class="warning-email-row">
                            <span class="warning-email-label">${matchLabel}:</span>
                            <span class="warning-email-value known">${w.matchedEmail}</span>
                        </div>
                        ${w.reason ? `<div class="warning-reason">${w.reason}</div>` : ''}
                    </div>
                `;
            } else if (w.detail) {
                emailHtml = `<div class="warning-detail">${w.detail}</div>`;
            }
            
            return `
                <div class="warning-item ${w.severity}">
                    <div class="warning-title">${w.title}</div>
                    <div class="warning-description">${w.description}</div>
                    ${emailHtml}
                </div>
            `;
        }).join('');
        
        // Setup safe sender button
        const safeSenderBtn = document.getElementById('safe-sender-btn');
        if (safeSenderBtn) {
            safeSenderBtn.onclick = () => {
                knownContacts.add(senderEmail);
                displayResults([], senderEmail);
            };
        }
    } else {
        warningsSection.classList.add('hidden');
        if (warningsFooter) warningsFooter.classList.add('hidden');
        if (safeMessage) safeMessage.classList.remove('hidden');
    }
}
