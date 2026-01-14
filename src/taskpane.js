// Email Fraud Detector - Outlook Web Add-in
// Version 3.2.8 - Fixed hidden class display issue

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
// REMOVED: '.co' (legitimate Colombia TLD)
const DECEPTIVE_TLDS = [
    '.com.co', '.com.br', '.com.mx', '.com.ar', '.com.au', '.com.ng',
    '.com.pk', '.com.ph', '.com.ua', '.com.ve', '.com.vn', '.com.tr',
    '.net.co', '.net.br', '.org.co', '.co.uk.com', '.us.com',
    '.cm', '.cc', '.ru', '.cn', '.tk', '.ml', '.ga', '.cf'
];

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

    // Tech Companies - PHRASES ONLY (not standalone words)
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

    // Document Signing / Business Tools
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
            'act now', 'urgent action required', 'action required',
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
let authInProgress = false;
let contactsFetched = false;

// ============================================
// INITIALIZATION
// ============================================
Office.onReady(async (info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log('Email Fraud Detector v3.2.8 starting...');
        await initializeMsal();
        setupEventHandlers();
        analyzeCurrentEmail();
        setupAutoScan();
    }
});

async function initializeMsal() {
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
    
    // CRITICAL: Clear any stuck interaction state from previous sessions
    try {
        await msalInstance.handleRedirectPromise();
        console.log('MSAL initialized, cleared any pending auth');
    } catch (e) {
        console.log('MSAL init note:', e.message);
    }
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
    if (authInProgress) {
        console.log('Auth already in progress, skipping');
        return null;
    }
    
    const accounts = msalInstance.getAllAccounts();
    
    try {
        if (accounts.length > 0) {
            // Try silent auth first
            const response = await msalInstance.acquireTokenSilent({
                scopes: CONFIG.scopes,
                account: accounts[0]
            });
            return response.accessToken;
        } else {
            // Need interactive auth - guard against multiple attempts
            authInProgress = true;
            try {
                const response = await msalInstance.acquireTokenPopup({
                    scopes: CONFIG.scopes
                });
                authInProgress = false;
                return response.accessToken;
            } catch (popupError) {
                authInProgress = false;
                throw popupError;
            }
        }
    } catch (error) {
        console.log('Auth error:', error);
        authInProgress = false;
        
        // If interaction_in_progress, try to clear it
        if (error.errorCode === 'interaction_in_progress') {
            console.log('Clearing stuck auth state...');
            try {
                sessionStorage.clear();
                await msalInstance.handleRedirectPromise();
            } catch (e) {
                // Ignore cleanup errors
            }
        }
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
    if (contactsFetched) return; // Don't retry if already attempted
    
    const token = await getAccessToken();
    if (!token) {
        console.log('No token available, continuing without contacts');
        contactsFetched = true; // Mark as attempted
        return;
    }
    
    console.log('Fetching contacts...');
    
    const contacts = await fetchContacts(token);
    
    // Add all to the knownContacts set
    contacts.forEach(e => knownContacts.add(e));
    
    // Add current user's email so lookalikes of their own address are detected
    if (currentUserEmail) {
        knownContacts.add(currentUserEmail.toLowerCase());
    }
    
    console.log('Total known contacts:', knownContacts.size);
    contactsFetched = true;
}

// ============================================
// HELPER FUNCTIONS
// ============================================
function isTrustedDomain(domain) {
    return CONFIG.trustedDomains.includes(domain.toLowerCase());
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
 * Detect suspicious display names - only high-risk words from generic domains
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
    
    for (const domain of CONFIG.trustedDomains) {
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
    for (const trusted of CONFIG.trustedDomains) {
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
// MAIN ANALYSIS
// ============================================
async function analyzeCurrentEmail() {
    showLoading();
    
    try {
        // Get current user email
        currentUserEmail = Office.context.mailbox.userProfile.emailAddress;
        
        // Try to fetch contacts (won't block if auth fails)
        if (knownContacts.size === 0 && !contactsFetched) {
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
                        subject: subject,
                        body: bodyResult.value || '',
                        replyTo: replyTo
                    };
                    
                    processEmail(emailData);
                });
            } else {
                const emailData = {
                    from: from,
                    subject: subject,
                    body: bodyResult.value || '',
                    replyTo: null
                };
                
                processEmail(emailData);
            }
        });
        
    } catch (error) {
        console.log('Analysis error:', error);
        showError('Unable to analyze email. Please try again.');
    }
}

function processEmail(emailData) {
    const senderEmail = emailData.from.emailAddress.toLowerCase();
    const displayName = emailData.from.displayName || '';
    const senderDomain = senderEmail.split('@')[1] || '';
    const content = (emailData.subject || '') + ' ' + (emailData.body || '');
    const replyTo = emailData.replyTo;
    
    const isKnownContact = knownContacts.has(senderEmail);
    
    const warnings = [];
    
    // 1. Reply-To Mismatch (only flag if different domain)
    if (replyTo && replyTo.toLowerCase() !== senderEmail) {
        const replyToDomain = replyTo.split('@')[1] || '';
        if (replyToDomain.toLowerCase() !== senderDomain) {
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
    
    // 10. Contact Lookalike (only if contacts were loaded)
    if (!isKnownContact && knownContacts.size > 0) {
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
    
    // Display results
    if (warnings.length > 0) {
        showWarnings(warnings, senderEmail, displayName);
    } else {
        showSafe(senderEmail, displayName);
    }
}

// ============================================
// UI FUNCTIONS
// ============================================
function showLoading() {
    document.getElementById('loading').classList.remove('hidden');
    document.getElementById('results').classList.add('hidden');
    document.getElementById('error').classList.add('hidden');
    // Reset body class
    document.body.className = '';
}

function showError(message) {
    document.getElementById('loading').classList.add('hidden');
    document.getElementById('results').classList.add('hidden');
    document.getElementById('error').classList.remove('hidden');
    document.getElementById('error-message').textContent = message;
    document.body.className = '';
}

function showSafe(email, displayName) {
    document.getElementById('loading').classList.add('hidden');
    document.getElementById('error').classList.add('hidden');
    document.getElementById('results').classList.remove('hidden');
    
    // Set body background to green
    document.body.className = 'status-safe';
    
    // Update status badge
    const statusBadge = document.getElementById('status-badge');
    statusBadge.className = 'status-badge safe';
    statusBadge.querySelector('.status-icon').textContent = '✓';
    statusBadge.querySelector('.status-text').textContent = 'No Warnings Detected';
    
    // Hide warnings section, show safe message
    document.getElementById('warnings-section').classList.add('hidden');
    document.getElementById('safe-message').classList.remove('hidden');
}

function showWarnings(warnings, senderEmail, displayName) {
    document.getElementById('loading').classList.add('hidden');
    document.getElementById('error').classList.add('hidden');
    document.getElementById('results').classList.remove('hidden');
    
    // Determine severity - critical if any warning, could be refined later
    const hasCritical = warnings.some(w => 
        w.type === 'contact-lookalike' || 
        w.type === 'homoglyph' || 
        w.type === 'org-impersonation' ||
        w.type === 'lookalike-domain'
    );
    
    // Set body background
    document.body.className = hasCritical ? 'status-critical' : 'status-critical';
    
    // Update status badge
    const statusBadge = document.getElementById('status-badge');
    statusBadge.className = 'status-badge danger';
    statusBadge.querySelector('.status-icon').textContent = '⚠️';
    statusBadge.querySelector('.status-text').textContent = `${warnings.length} Warning${warnings.length > 1 ? 's' : ''} Detected`;
    
    // Build warning items HTML
    const warningItemsHtml = warnings.map(w => {
        let detailHtml = '';
        const severity = 'critical'; // Could vary based on warning type
        
        if (w.type === 'wire-fraud' && w.keywords) {
            const keywordTags = w.keywords.slice(0, 5).map(k => 
                `<span class="keyword-tag">${k}</span>`
            ).join('');
            detailHtml = `
                <div class="warning-keywords-section">
                    <div class="warning-keywords-label">Triggered by:</div>
                    <div class="warning-keywords">${keywordTags}</div>
                </div>
                <div class="warning-advice">
                    <strong>Why this matters:</strong> ${w.keywordExplanation}
                </div>
            `;
        } else if (w.type === 'org-impersonation') {
            detailHtml = `
                <div class="warning-emails">
                    <div class="warning-email-row">
                        <span class="warning-email-label">Claims to be:</span>
                        <span class="warning-email-value known">${w.entityClaimed}</span>
                    </div>
                    <div class="warning-email-row">
                        <span class="warning-email-label">Actually from:</span>
                        <span class="warning-email-value suspicious">${w.senderEmail}</span>
                    </div>
                    <div class="warning-email-row">
                        <span class="warning-email-label">Legitimate domains:</span>
                        <span class="warning-email-value known">${w.legitimateDomains.join(', ')}</span>
                    </div>
                </div>
            `;
        } else if (w.senderEmail && w.matchedEmail) {
            const matchLabel = w.type === 'replyto-mismatch' ? 'Replies go to' : 
                               w.type === 'impersonation' ? 'Display name shows' : 'Similar to';
            detailHtml = `
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
            detailHtml = `<div class="warning-reason">${w.detail}</div>`;
        }
        
        return `
            <div class="warning-item ${severity}">
                <div class="warning-title">${w.title}</div>
                <div class="warning-description">${w.description}</div>
                ${detailHtml}
            </div>
        `;
    }).join('');
    
    // Populate warnings list
    document.getElementById('warnings-list').innerHTML = warningItemsHtml;
    
    // Show warnings section and footer, hide safe message
    document.getElementById('warnings-section').classList.remove('hidden');
    document.getElementById('warnings-footer').classList.remove('hidden');
    document.getElementById('safe-message').classList.add('hidden');
}
