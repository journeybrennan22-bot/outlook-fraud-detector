// Email Fraud Detector - Outlook Web Add-in
// Version 3.8.0 - Updated: On-behalf-of detection, deduplicated brand/org warnings

// ============================================
// CONFIGURATION
// ============================================
const CONFIG = {
    clientId: '622f0452-d622-45d1-aab3-3a2026389dd3',
    redirectUri: 'https://journeybrennan22-bot.github.io/outlook-fraud-detector/src/taskpane.html',
    scopes: ['User.Read', 'Contacts.Read'],
    trustedDomains: []
};

// ============================================
// COUNTRY CODE TLD LOOKUP
// Maps country-code TLDs to country names
// ============================================
const COUNTRY_CODE_TLDS = {
    // Compound TLDs (check these first - more specific)
    '.com.ar': 'Argentina', '.com.au': 'Australia', '.com.br': 'Brazil',
    '.com.cn': 'China', '.com.co': 'Colombia', '.com.mx': 'Mexico',
    '.com.ng': 'Nigeria', '.com.pk': 'Pakistan', '.com.ph': 'Philippines',
    '.com.tr': 'Turkey', '.com.ua': 'Ukraine', '.com.ve': 'Venezuela',
    '.com.vn': 'Vietnam', '.co.uk': 'United Kingdom', '.co.za': 'South Africa',
    '.co.in': 'India', '.co.jp': 'Japan', '.co.kr': 'South Korea',
    '.co.nz': 'New Zealand', '.net.br': 'Brazil', '.net.co': 'Colombia',
    '.org.br': 'Brazil', '.org.co': 'Colombia', '.org.uk': 'United Kingdom',
    '.co.uk.com': 'United Kingdom', '.us.com': 'United States',
    
    // Single ccTLDs
    '.ar': 'Argentina', '.au': 'Australia', '.at': 'Austria',
    '.be': 'Belgium', '.br': 'Brazil', '.ca': 'Canada',
    '.ch': 'Switzerland', '.cl': 'Chile', '.cn': 'China',
    '.co': 'Colombia', '.cz': 'Czech Republic', '.de': 'Germany',
    '.dk': 'Denmark', '.es': 'Spain', '.fi': 'Finland',
    '.fr': 'France', '.gr': 'Greece', '.hk': 'Hong Kong',
    '.hu': 'Hungary', '.id': 'Indonesia', '.ie': 'Ireland',
    '.il': 'Israel', '.in': 'India', '.it': 'Italy',
    '.jp': 'Japan', '.kr': 'South Korea', '.mx': 'Mexico',
    '.my': 'Malaysia', '.nl': 'Netherlands', '.no': 'Norway',
    '.nz': 'New Zealand', '.pe': 'Peru', '.ph': 'Philippines',
    '.pk': 'Pakistan', '.pl': 'Poland', '.pt': 'Portugal',
    '.ro': 'Romania', '.ru': 'Russia', '.sa': 'Saudi Arabia',
    '.se': 'Sweden', '.sg': 'Singapore', '.th': 'Thailand',
    '.tr': 'Turkey', '.tw': 'Taiwan', '.ua': 'Ukraine',
    '.uk': 'United Kingdom', '.us': 'United States', '.ve': 'Venezuela', '.vn': 'Vietnam',
    '.za': 'South Africa', '.ng': 'Nigeria', '.ke': 'Kenya',
    '.eg': 'Egypt', '.ae': 'United Arab Emirates',
    
    // Suspicious/commonly abused TLDs
    '.tk': 'Tokelau', '.ml': 'Mali', '.ga': 'Gabon',
    '.cf': 'Central African Republic', '.gq': 'Equatorial Guinea',
    '.cm': 'Cameroon', '.cc': 'Cocos Islands', '.ws': 'Samoa',
    '.pw': 'Palau', '.top': 'Generic (often abused)', '.xyz': 'Generic (often abused)',
    '.buzz': 'Generic (often abused)', '.icu': 'Generic (often abused)',
    '.biz': 'Generic (often abused)', '.info': 'Generic (often abused)',
    '.shop': 'Generic (often abused)', '.club': 'Generic (often abused)'
};

// TLDs to flag as international senders (subset that warrants warning)
const INTERNATIONAL_TLDS = [
    '.com.co', '.com.br', '.com.mx', '.com.ar', '.com.au', '.com.ng',
    '.com.pk', '.com.ph', '.com.ua', '.com.ve', '.com.vn', '.com.tr',
    '.net.co', '.net.br', '.org.co', '.us',
    '.cm', '.cc', '.ru', '.cn', '.tk', '.ml', '.ga', '.cf', '.gq', '.pw'
];

// Fake country-lookalike TLDs (commercial services mimicking real TLDs)
const FAKE_COUNTRY_TLDS = ['.us.com', '.co.uk.com', '.eu.com', '.de.com', '.br.com'];

// Suspicious words commonly used in fake domains
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

// ============================================
// NEW v3.5.0: PHISHING URGENCY KEYWORDS
// Different from wire fraud - these are account/deletion threats
// ============================================
const PHISHING_URGENCY_KEYWORDS = [
    // Account threats
    'account locked', 'account suspended', 'account disabled',
    'account will be', 'account has been',
    'access suspended', 'access revoked', 'access denied',
    // Deletion threats
    'will be deleted', 'scheduled for deletion', 'permanently removed',
    'permanently deleted', 'files will be lost', 'data will be erased',
    'photos will be deleted', 'videos will be deleted',
    // Storage/limit scams
    'storage limit', 'storage full', 'critical limit',
    'quota exceeded', 'mailbox full', 'inbox full',
    // Urgency phrases
    'final notice', 'final warning', 'last chance',
    'immediate action', 'act immediately',
    'expires today', 'expires soon',
    'within 24 hours', 'within 48 hours',
    // Verification scams
    'verify your account', 'confirm your identity',
    'verify your email', 'verify your information',
    'update your payment', 'payment failed', 'payment declined',
    'billing problem', 'billing issue',
    // Subscription scams
    'subscription expired', 'renew your subscription',
    'membership expired', 'unable to renew',
    // Block threats
    'we\'ve blocked', 'has been blocked', 'temporarily blocked'
];

// ============================================
// BRAND IMPERSONATION DETECTION (CONTENT-BASED)
// ============================================
const BRAND_CONTENT_DETECTION = {
    'docusign': {
        keywords: ['docusign'],
        legitimateDomains: ['docusign.com', 'docusign.net']
    },
    'microsoft': {
        keywords: ['microsoft 365', 'microsoft-365', 'office 365', 'office-365', 'sharepoint', 'onedrive', 'microsoft account', 'microsoft teams'],
        legitimateDomains: ['microsoft.com', 'office.com', 'sharepoint.com', 'onedrive.com', 'live.com', 'outlook.com', 'office365.com', 'teams.mail.microsoft']
    },
    'google': {
        keywords: ['google drive', 'google docs', 'google account', 'google workspace'],
        legitimateDomains: ['google.com', 'gmail.com', 'googlemail.com']
    },
    'amazon': {
        keywords: ['amazon prime', 'amazon account', 'amazon order', 'amazon.com order'],
        legitimateDomains: ['amazon.com', 'amazon.co.uk', 'amazon.ca', 'amazonses.com']
    },
    'paypal': {
        keywords: ['paypal'],
        legitimateDomains: ['paypal.com']
    },
    'netflix': {
        keywords: ['netflix'],
        legitimateDomains: ['netflix.com']
    },
    'adobe sign': {
        keywords: ['adobe sign', 'adobesign'],
        legitimateDomains: ['adobe.com', 'adobesign.com', 'echosign.com']
    },
    'dropbox': {
        keywords: ['dropbox', 'dropbox sign', 'hellosign'],
        legitimateDomains: ['dropbox.com', 'hellosign.com', 'dropboxmail.com']
    },
    'apple': {
        keywords: ['apple id', 'icloud account', 'apple account'],
        legitimateDomains: ['apple.com', 'icloud.com']
    },
    'facebook': {
        keywords: ['facebook account', 'meta account', 'facebook security'],
        legitimateDomains: ['facebook.com', 'meta.com', 'facebookmail.com']
    },
    'linkedin': {
        keywords: ['linkedin account', 'linkedin invitation', 'linkedin message'],
        legitimateDomains: ['linkedin.com']
    },
    'yahoo': {
        keywords: ['yahoo account', 'yahoo mail', 'yahoo security'],
        legitimateDomains: ['yahoo.com', 'yahoomail.com']
    },
    'mcafee': {
        keywords: ['mcafee'],
        legitimateDomains: ['mcafee.com']
    },
    'coinbase': {
        keywords: ['coinbase'],
        legitimateDomains: ['coinbase.com']
    },
    'dhl': {
        keywords: ['dhl express', 'dhl shipment', 'dhl delivery', 'dhl package'],
        legitimateDomains: ['dhl.com', 'dhl.de']
    },
    'fedex': {
        keywords: ['fedex', 'federal express'],
        legitimateDomains: ['fedex.com']
    },
    'ups': {
        keywords: ['ups package', 'ups delivery', 'ups shipment', 'united parcel'],
        legitimateDomains: ['ups.com']
    },
    'usps': {
        keywords: ['usps', 'postal service', 'usps delivery', 'usps package'],
        legitimateDomains: ['usps.com']
    },
    'zelle': {
        keywords: ['zelle'],
        legitimateDomains: ['zellepay.com', 'zelle.com']
    },
    'venmo': {
        keywords: ['venmo'],
        legitimateDomains: ['venmo.com']
    },
    'cashapp': {
        keywords: ['cash app', 'cashapp'],
        legitimateDomains: ['cash.app', 'square.com', 'squareup.com']
    },
    'quickbooks': {
        keywords: ['quickbooks', 'intuit'],
        legitimateDomains: ['intuit.com', 'quickbooks.com']
    },
    'zoom': {
        keywords: ['zoom meeting', 'zoom invitation', 'zoom account'],
        legitimateDomains: ['zoom.us', 'zoom.com']
    },
    // v3.7.0: Major Retailers
    'walmart': {
        keywords: ['walmart', 'wal-mart'],
        legitimateDomains: ['walmart.com']
    },
    'target': {
        keywords: ['target order', 'target account', 'target registry', 'target circle'],
        legitimateDomains: ['target.com']
    },
    'costco': {
        keywords: ['costco', 'costco wholesale'],
        legitimateDomains: ['costco.com']
    },
    'best buy': {
        keywords: ['best buy', 'bestbuy', 'geek squad'],
        legitimateDomains: ['bestbuy.com']
    },
    'home depot': {
        keywords: ['home depot'],
        legitimateDomains: ['homedepot.com']
    },
    'lowes': {
        keywords: ['lowe\'s', 'lowes'],
        legitimateDomains: ['lowes.com']
    },
    'ebay': {
        keywords: ['ebay'],
        legitimateDomains: ['ebay.com']
    },
    // v3.8.0: Government Agencies
    'dmv': {
        keywords: ['department of motor vehicles', 'dmv service desk', 'dmv appointment', 'dmv registration'],
        legitimateDomains: ['.gov']
    },
    'irs': {
        keywords: ['internal revenue service', 'irs refund', 'irs audit', 'tax return', 'irs notice'],
        legitimateDomains: ['irs.gov']
    },
    'social security': {
        keywords: ['social security administration', 'social security number', 'ssa benefit', 'social security statement'],
        legitimateDomains: ['ssa.gov']
    }
};

// ============================================
// ORGANIZATION IMPERSONATION TARGETS
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

    // Major Banks
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

    // Tech Companies - PHRASES ONLY
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
    "old republic title": ["oldrepublictitle.com", "oldrepublic.com"],

    // v3.7.0: Major Retailers
    "walmart": ["walmart.com"],
    "walmart customer support": ["walmart.com"],
    "target": ["target.com"],
    "costco": ["costco.com"],
    "costco wholesale": ["costco.com"],
    "best buy": ["bestbuy.com"],
    "geek squad": ["bestbuy.com", "geeksquad.com"],
    "home depot": ["homedepot.com"],
    "lowes": ["lowes.com"],
    "ebay": ["ebay.com"],

    // v3.8.0: State Government Agencies
    "dmv": [".gov"],
    "dmv service desk": [".gov"],
    "department of motor vehicles": [".gov"],
    "motor vehicles": [".gov"],
    "state tax board": [".gov"],
    "franchise tax board": [".gov"],
    "edd": [".gov"],
    "employment development": [".gov"],
    "unemployment insurance": [".gov"],
    "child support services": [".gov"],
    "department of revenue": [".gov"],
    "state attorney general": [".gov"],
    "attorney general": [".gov"]

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
    console.log('Email Fraud Detector v3.8.0 script loaded, host:', info.host);
    if (info.host === Office.HostType.Outlook) {
        console.log('Email Fraud Detector v3.8.0 initializing for Outlook...');
        await initializeMsal();
        setupEventHandlers();
        analyzeCurrentEmail();
        setupAutoScan();
        console.log('Email Fraud Detector v3.8.0 ready');
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
            const response = await msalInstance.acquireTokenSilent({
                scopes: CONFIG.scopes,
                account: accounts[0]
            });
            return response.accessToken;
        } else {
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

async function fetchAllKnownContacts() {
    if (contactsFetched) return;
    
    const token = await getAccessToken();
    if (!token) {
        console.log('No token available, continuing without contacts');
        contactsFetched = true;
        return;
    }
    
    console.log('Fetching contacts...');
    
    const contacts = await fetchContacts(token);
    
    contacts.forEach(e => knownContacts.add(e));
    
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
 * NEW v3.5.0: Detect recipient self-spoofing
 * Catches when display name matches the recipient's own email/name
 */
function detectRecipientSpoofing(displayName, senderEmail) {
    if (!displayName || !currentUserEmail) return null;
    
    const displayLower = displayName.toLowerCase().trim();
    const recipientLower = currentUserEmail.toLowerCase();
    const recipientUsername = recipientLower.split('@')[0];
    
    // Clean display name and recipient (remove dots, underscores, spaces)
    const displayCleaned = displayLower.replace(/[\.\-_\s]/g, '');
    const recipientCleaned = recipientUsername.replace(/[\.\-_\s]/g, '');
    
    // Check if display name contains recipient's username (or vice versa)
    if (displayCleaned.length >= 4 && recipientCleaned.length >= 4) {
        if (displayCleaned.includes(recipientCleaned) || recipientCleaned.includes(displayCleaned)) {
            // Make sure it's not actually FROM the recipient
            const senderLower = senderEmail.toLowerCase();
            if (!senderLower.includes(recipientUsername)) {
                return {
                    displayName: displayName,
                    recipientEmail: currentUserEmail
                };
            }
        }
    }
    
    return null;
}

/**
 * NEW v3.5.0: Detect phishing urgency keywords
 * Different from wire fraud - focuses on account threats
 */
function detectPhishingUrgency(bodyText, subject) {
    if (!bodyText && !subject) return null;
    
    const textToCheck = ((subject || '') + ' ' + (bodyText || '')).toLowerCase();
    const foundKeywords = [];
    
    for (const keyword of PHISHING_URGENCY_KEYWORDS) {
        if (textToCheck.includes(keyword.toLowerCase())) {
            foundKeywords.push(keyword);
        }
    }
    
    // Need at least 2 urgency keywords to trigger
    if (foundKeywords.length >= 2) {
        return {
            keywords: foundKeywords.slice(0, 4)
        };
    }
    
    return null;
}

/**
 * NEW v3.5.0: Detect gibberish domain
 * Uses scoring system to avoid false positives
 */
function detectGibberishDomain(email) {
    if (!email) return null;
    
    const parts = email.split('@');
    if (parts.length !== 2) return null;
    
    const domain = parts[1].toLowerCase();
    const domainParts = domain.split('.');
    if (domainParts.length < 2) return null;
    
    const mainPart = domainParts[0];
    
    let suspicionScore = 0;
    const reasons = [];
    
    // Check 1: High digit ratio in main domain part
    if (mainPart.length > 5) {
        const digitCount = (mainPart.match(/\d/g) || []).length;
        const digitRatio = digitCount / mainPart.length;
        if (digitRatio > 0.3) {
            suspicionScore += 2;
            reasons.push('high number ratio');
        }
    }
    
    // Check 2: Multiple random-looking subdomains
    if (domainParts.length >= 3) {
        let gibberishSubdomains = 0;
        const subdomains = domainParts.slice(0, -1);
        
        for (const sub of subdomains) {
            const hasDigits = /\d/.test(sub);
            const hasLetters = /[a-z]/i.test(sub);
            const isShortAndRandom = sub.length > 4 && hasDigits && hasLetters;
            const containsNoWords = !/(mail|web|app|api|www|cdn|img|static|secure|login|account|cloud|storage)/i.test(sub);
            
            if (isShortAndRandom && containsNoWords) {
                gibberishSubdomains++;
            }
        }
        
        if (gibberishSubdomains >= 2) {
            suspicionScore += 3;
            reasons.push('multiple random subdomains');
        }
    }
    
    // Check 3: Suspicious TLD combined with other signals
    const suspiciousTLDs = ['.us', '.tk', '.ml', '.ga', '.cf', '.gq', '.pw', '.cc', '.ws', '.top', '.xyz', '.buzz', '.biz', '.info', '.shop', '.club', '.icu'];
    const tld = '.' + domainParts[domainParts.length - 1];
    if (suspiciousTLDs.includes(tld) && suspicionScore > 0) {
        suspicionScore += 1;
        reasons.push('suspicious TLD (' + tld + ')');
    }
    
    // Check 4: No vowels in main domain part
    if (mainPart.length > 6) {
        const vowelCount = (mainPart.match(/[aeiou]/gi) || []).length;
        if (vowelCount === 0) {
            suspicionScore += 2;
            reasons.push('no vowels in domain');
        }
    }
    
    // Need score of 3+ to trigger
    if (suspicionScore >= 3) {
        return {
            domain: domain,
            reasons: reasons
        };
    }
    
    return null;
}

/**
 * Detect brand impersonation based on email CONTENT
 */
function detectBrandImpersonation(subject, body, senderDomain) {
    console.log('BRAND CHECK CALLED - Domain:', senderDomain);
    
    const contentLower = ((subject || '') + ' ' + (body || '')).toLowerCase();
    
    for (const [brandName, config] of Object.entries(BRAND_CONTENT_DETECTION)) {
        const mentionsBrand = config.keywords.some(keyword => 
            contentLower.includes(keyword.toLowerCase())
        );
        
        if (mentionsBrand) {
            if (!senderDomain) {
                return {
                    brandName: formatEntityName(brandName),
                    senderDomain: '(invalid or hidden sender)',
                    legitimateDomains: config.legitimateDomains
                };
            }
            
            const domainLower = senderDomain.toLowerCase();
            const isLegitimate = config.legitimateDomains.some(legit => {
                // v3.8.0: Handle suffix patterns like ".gov" for government agencies
                if (legit.startsWith('.')) {
                    return domainLower.endsWith(legit);
                }
                return domainLower === legit || domainLower.endsWith(`.${legit}`);
            });
            
            if (!isLegitimate) {
                return {
                    brandName: formatEntityName(brandName),
                    senderDomain: senderDomain,
                    legitimateDomains: config.legitimateDomains
                };
            }
        }
    }
    
    return null;
}

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
                // v3.8.0: Handle suffix patterns like ".gov" for government agencies
                if (legit.startsWith('.')) {
                    return senderDomain.endsWith(legit);
                }
                return senderDomain === legit || senderDomain.endsWith(`.${legit}`);
            });
            
            if (!isLegitimate) {
                // v3.8.0: Better messaging for government agencies
                const hasGovSuffix = legitimateDomains.some(d => d === '.gov');
                const displayDomains = hasGovSuffix ? 'official .gov domains' : legitimateDomains.join(', ');
                return {
                    entityClaimed: formatEntityName(entityName),
                    senderDomain: senderDomain,
                    legitimateDomains: legitimateDomains,
                    message: `Sender claims to be "${formatEntityName(entityName)}" but email comes from ${senderDomain}. Legitimate emails come from: ${displayDomains}`
                };
            }
        }
    }
    
    return null;
}

/**
 * v3.6.0: Detect international sender (formerly "deceptive TLD")
 * Returns country information for foreign domains
 */
function detectInternationalSender(domain) {
    const domainLower = domain.toLowerCase();
    
    // Check compound TLDs first (more specific)
    for (const [tld, country] of Object.entries(COUNTRY_CODE_TLDS)) {
        if (tld.includes('.') && tld.split('.').length > 2) {
            // This is a compound TLD like .com.br
            if (domainLower.endsWith(tld)) {
                // Only flag if it's in our warning list
                if (INTERNATIONAL_TLDS.some(t => domainLower.endsWith(t))) {
                    return { tld, country };
                }
            }
        }
    }
    
    // Then check single ccTLDs
    for (const tld of INTERNATIONAL_TLDS) {
        if (domainLower.endsWith(tld)) {
            const country = COUNTRY_CODE_TLDS[tld] || 'Unknown';
            return { tld, country };
        }
    }
    
    return null;
}

/**
 * Detect suspicious domain patterns
 */
function detectSuspiciousDomain(domain) {
    const domainLower = domain.toLowerCase();
    
    // Check for fake country-lookalike TLDs first
    for (const fakeTld of FAKE_COUNTRY_TLDS) {
        if (domainLower.endsWith(fakeTld)) {
            return {
                pattern: fakeTld,
                reason: `This email was sent from a domain ending in ${fakeTld}. This domain extension is designed to look like a legitimate country domain but is not. Proceed with caution.`
            };
        }
    }
    
    // v3.7.0: Check for suspicious generic TLDs (.biz, .info, .shop, etc.)
    const suspiciousGenericTLDs = ['.biz', '.info', '.shop', '.club', '.top', '.xyz', '.buzz', '.icu'];
    const tld = '.' + domainLower.split('.').pop();
    if (suspiciousGenericTLDs.includes(tld)) {
        return {
            pattern: tld,
            reason: `This email was sent from a domain ending in ${tld}. Domains ending in ${tld} have been identified by Spamhaus and Symantec as frequently used in spam and phishing campaigns. Proceed with caution.`
        };
    }
    
    // Get the registrable domain (the part someone can register, not subdomains)
    // e.g., "mailcenter.usaa.com" -> "usaa", "secure-login.co.uk" -> "secure-login"
    const registrableName = getRegistrableDomainName(domainLower);
    
    if (registrableName.includes('-')) {
        const parts = registrableName.split('-');
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
        return null;
    }
    
    for (const word of SUSPICIOUS_DOMAIN_WORDS) {
        if (registrableName.endsWith(word) && registrableName !== word && registrableName.length > word.length + 3) {
            return {
                pattern: word,
                reason: `Domain ends with "${word}" which is commonly used in phishing attacks`
            };
        }
    }
    
    return null;
}

/**
 * Extract the registrable domain name (without TLD) from a full domain
 * e.g., "mailcenter.usaa.com" -> "usaa", "secure-login.co.uk" -> "secure-login"
 */
function getRegistrableDomainName(domain) {
    const compoundTlds = ['.co.uk', '.co.za', '.co.in', '.co.jp', '.co.kr', '.co.nz', 
                         '.com.ar', '.com.au', '.com.br', '.com.cn', '.com.co', '.com.mx',
                         '.com.ng', '.com.pk', '.com.ph', '.com.tr', '.com.ua', '.com.ve', '.com.vn',
                         '.net.br', '.net.co', '.org.br', '.org.co', '.org.uk'];
    
    // Check for compound TLDs first
    for (const tld of compoundTlds) {
        if (domain.endsWith(tld)) {
            const withoutTld = domain.slice(0, -tld.length);
            const parts = withoutTld.split('.');
            return parts[parts.length - 1];
        }
    }
    
    // Standard TLD - get the second-to-last part
    const parts = domain.split('.');
    if (parts.length >= 2) {
        return parts[parts.length - 2];
    }
    return parts[0];
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
 * Detect lookalike domains
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
 * Detect contact lookalike
 */
function detectContactLookalike(senderEmail) {
    const parts = senderEmail.toLowerCase().split('@');
    if (parts.length !== 2) return null;
    
    const senderLocal = parts[0];
    const senderDomain = parts[1];
    
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
        
        if (senderDomain === contactDomain) {
            if (usernameDiff > 0 && usernameDiff <= 4) {
                return {
                    incomingEmail: senderEmail,
                    matchedContact: contact,
                    reason: `Username is ${usernameDiff} character${usernameDiff > 1 ? 's' : ''} different`
                };
            }
        }
        
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
        currentUserEmail = Office.context.mailbox.userProfile.emailAddress;
        
        if (knownContacts.size === 0 && !contactsFetched) {
            await fetchAllKnownContacts();
        }
        
        const item = Office.context.mailbox.item;
        const from = item.from;
        const subject = item.subject;
        
        item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
            if (item.getAllInternetHeadersAsync) {
                item.getAllInternetHeadersAsync((headerResult) => {
                    let replyTo = null;
                    let senderHeader = null;
                    if (headerResult.status === Office.AsyncResultStatus.Succeeded) {
                        const headers = headerResult.value;
                        const replyToMatch = headers.match(/^Reply-To:\s*(.+)$/mi);
                        if (replyToMatch) {
                            const emailMatch = replyToMatch[1].match(/<([^>]+)>/) || replyToMatch[1].match(/([^\s,]+@[^\s,]+)/);
                            if (emailMatch) {
                                replyTo = emailMatch[1].trim();
                            }
                        }
                        // v3.8.0: Parse Sender header for "on behalf of" detection
                        const senderMatch = headers.match(/^Sender:\s*(.+)$/mi);
                        if (senderMatch) {
                            const senderEmailMatch = senderMatch[1].match(/<([^>]+)>/) || senderMatch[1].match(/([^\s,]+@[^\s,]+)/);
                            if (senderEmailMatch) {
                                senderHeader = senderEmailMatch[1].trim();
                            }
                        }
                    }
                    
                    const emailData = {
                        from: from,
                        subject: subject,
                        body: bodyResult.value || '',
                        replyTo: replyTo,
                        senderHeader: senderHeader
                    };
                    
                    processEmail(emailData);
                });
            } else {
                const emailData = {
                    from: from,
                    subject: subject,
                    body: bodyResult.value || '',
                    replyTo: null,
                    senderHeader: null
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
    const senderHeader = emailData.senderHeader;
    
    const isKnownContact = knownContacts.has(senderEmail);
    
    const warnings = [];
    
    // ============================================
    // NEW v3.5.0 CHECKS (run first - highest priority)
    // ============================================
    
    // Check for recipient self-spoofing
    const recipientSpoof = detectRecipientSpoofing(displayName, senderEmail);
    if (recipientSpoof) {
        warnings.push({
            type: 'recipient-spoof',
            severity: 'critical',
            title: 'Sender Impersonating You',
            description: 'The sender is using YOUR name as their display name. This is a common phishing tactic.',
            senderEmail: senderEmail,
            matchedEmail: recipientSpoof.displayName
        });
    }
    
    // Check for phishing urgency keywords
    const phishingUrgency = detectPhishingUrgency(emailData.body, emailData.subject);
    if (phishingUrgency) {
        warnings.push({
            type: 'phishing-urgency',
            severity: 'critical',
            title: 'Phishing Language Detected',
            description: 'This email uses fear tactics commonly found in phishing scams.',
            keywords: phishingUrgency.keywords,
            keywordCategory: 'Phishing Tactics',
            keywordExplanation: 'Scammers use threats of account deletion, suspension, or data loss to pressure you into clicking malicious links. Legitimate companies rarely threaten immediate action via email.'
        });
    }
    
    // Check for gibberish domain
    const gibberishDomain = detectGibberishDomain(senderEmail);
    if (gibberishDomain) {
        warnings.push({
            type: 'gibberish-domain',
            severity: 'critical',
            title: 'Suspicious Random Domain',
            description: `This email comes from a domain that appears to be randomly generated (${gibberishDomain.reasons.join(', ')}). Legitimate companies use recognizable domain names.`,
            senderEmail: senderEmail,
            matchedEmail: gibberishDomain.domain
        });
    }

    // ============================================
    // EXISTING CHECKS
    // ============================================
    
    // 1. Reply-To Mismatch
    if (replyTo && replyTo.toLowerCase() !== senderEmail) {
        const replyToDomain = replyTo.split('@')[1] || '';
        if (replyToDomain.toLowerCase() !== senderDomain) {
            warnings.push({
                type: 'replyto-mismatch',
                severity: 'medium',
                title: 'Reply-To Mismatch',
                description: 'Replies will go to a different address than the sender.',
                senderEmail: senderEmail,
                matchedEmail: replyTo
            });
        }
    }
    
    // v3.8.0: On-Behalf-Of / Sender Mismatch
    if (senderHeader) {
        const senderHeaderLower = senderHeader.toLowerCase();
        const senderHeaderDomain = senderHeaderLower.split('@')[1] || '';
        if (senderHeaderDomain && senderHeaderDomain !== senderDomain) {
            warnings.push({
                type: 'on-behalf-of',
                severity: 'medium',
                title: 'Sent On Behalf Of Another Domain',
                description: 'This email was sent by one domain on behalf of a completely different domain. This is a common tactic used to disguise the true origin of an email.',
                senderEmail: senderHeader,
                matchedEmail: senderEmail
            });
        }
    }
    
    // 2. Brand Impersonation
    const brandImpersonation = detectBrandImpersonation(emailData.subject, emailData.body, senderDomain);
    if (brandImpersonation) {
        warnings.push({
            type: 'brand-impersonation',
            severity: 'critical',
            title: 'Brand Impersonation Suspected',
            description: `This email references ${brandImpersonation.brandName} but was NOT sent from a verified ${brandImpersonation.brandName} domain.`,
            senderEmail: senderEmail,
            senderDomain: senderDomain,
            brandClaimed: brandImpersonation.brandName,
            legitimateDomains: brandImpersonation.legitimateDomains
        });
    }
    
    // 3. Organization Impersonation
    // v3.8.0: Skip if brand impersonation already caught this brand (avoids duplicate warnings)
    if (!isTrustedDomain(senderDomain)) {
        const orgImpersonation = detectOrganizationImpersonation(displayName, senderDomain);
        if (orgImpersonation) {
            // Only add if brand impersonation didn't already flag the same entity
            const brandAlreadyCaught = brandImpersonation && 
                brandImpersonation.brandName.toLowerCase() === orgImpersonation.entityClaimed.toLowerCase();
            if (!brandAlreadyCaught) {
                warnings.push({
                    type: 'org-impersonation',
                    severity: 'critical',
                    title: 'Organization Impersonation',
                    description: orgImpersonation.message,
                    senderEmail: senderEmail,
                    entityClaimed: orgImpersonation.entityClaimed,
                    legitimateDomains: orgImpersonation.legitimateDomains
                });
            }
        }
    }
    
    // 4. International Sender (formerly Deceptive TLD)
    const internationalSender = detectInternationalSender(senderDomain);
    if (internationalSender) {
        warnings.push({
            type: 'international-sender',
            severity: 'medium',
            title: 'International Sender',
            description: '',
            senderEmail: senderEmail,
            senderDomain: senderDomain,
            country: internationalSender.country,
            tld: internationalSender.tld
        });
    }
    
    // 5. Suspicious Domain
    const suspiciousDomain = detectSuspiciousDomain(senderDomain);
    if (suspiciousDomain) {
        warnings.push({
            type: 'suspicious-domain',
            severity: 'medium',
            title: 'Suspicious Domain',
            description: suspiciousDomain.reason
        });
    }
    
    // 6. Display Name Suspicion
    if (!isKnownContact) {
        const displaySuspicion = detectSuspiciousDisplayName(displayName, senderDomain);
        if (displaySuspicion) {
            warnings.push({
                type: 'display-name-suspicion',
                severity: 'medium',
                title: 'Suspicious Display Name',
                description: displaySuspicion.reason,
                senderEmail: senderEmail,
                matchedEmail: displaySuspicion.pattern
            });
        }
    }
    
    // 7. Display Name Impersonation
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
    
    // 8. Homoglyphs
    const homoglyph = detectHomoglyphs(senderEmail);
    if (homoglyph) {
        warnings.push({
            type: 'homoglyph',
            severity: 'critical',
            title: 'Invisible Character Trick',
            description: 'This email contains deceptive characters that look identical to normal letters.',
            senderEmail: senderEmail,
            detail: homoglyph
        });
    }
    
    // 9. Lookalike Domain
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
    
    // 10. Wire Fraud Keywords
    const wireKeywords = detectWireFraudKeywords(content);
    if (wireKeywords.length > 0) {
        const keywordInfo = getKeywordExplanation(wireKeywords[0]);
        warnings.push({
            type: 'wire-fraud',
            severity: 'critical',
            title: 'Dangerous Keywords Detected',
            description: 'This email contains terms commonly used in wire fraud.',
            keywords: wireKeywords,
            keywordCategory: keywordInfo.category,
            keywordExplanation: keywordInfo.explanation
        });
    }
    
    // 11. Contact Lookalike
    if (!isKnownContact && knownContacts.size > 0) {
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
    
    displayResults(warnings);
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

/**
 * Helper to wrap domain in nowrap span
 */
function wrapDomain(domain) {
    return `<span style="white-space: nowrap;">${domain}</span>`;
}

/**
 * Format domains list with proper styling
 */
function formatDomainsList(domains) {
    return domains.map(d => wrapDomain(d)).join(', ');
}

function displayResults(warnings) {
    document.getElementById('loading').classList.add('hidden');
    document.getElementById('error').classList.add('hidden');
    document.getElementById('results').classList.remove('hidden');
    
    const criticalCount = warnings.filter(w => w.severity === 'critical').length;
    const mediumCount = warnings.filter(w => w.severity === 'medium').length;
    
    document.body.classList.remove('status-critical', 'status-medium', 'status-info', 'status-safe');
    
    const statusBadge = document.getElementById('status-badge');
    const statusIcon = statusBadge.querySelector('.status-icon');
    const statusText = statusBadge.querySelector('.status-text');
    
    if (criticalCount > 0 || mediumCount > 0) {
        const totalWarnings = criticalCount + mediumCount;
        document.body.classList.add('status-critical');
        statusBadge.className = 'status-badge danger';
        statusIcon.textContent = '🚨';
        statusText.textContent = `${totalWarnings} Issue${totalWarnings > 1 ? 's' : ''} Found`;
    } else {
        document.body.classList.add('status-safe');
        statusBadge.className = 'status-badge safe';
        statusIcon.textContent = '✅';
        statusText.textContent = 'No Issues Detected';
    }
    
    const warningsSection = document.getElementById('warnings-section');
    const warningsList = document.getElementById('warnings-list');
    const warningsFooter = document.getElementById('warnings-footer');
    const safeMessage = document.getElementById('safe-message');
    
    if (warnings.length > 0) {
        // Sort warnings by priority (highest threat first)
        const WARNING_PRIORITY = {
            'replyto-mismatch': 1,
            'on-behalf-of': 2,
            'impersonation': 3,
            'recipient-spoof': 4,
            'contact-lookalike': 5,
            'brand-impersonation': 6,
            'org-impersonation': 7,
            'suspicious-domain': 8,
            'via-routing': 9,
            'gibberish-domain': 10,
            'lookalike-domain': 11,
            'homoglyph': 12,
            'display-name-suspicion': 13,
            'international-sender': 14,
            'wire-fraud': 15,
            'phishing-urgency': 16
        };
        warnings.sort((a, b) => (WARNING_PRIORITY[a.type] || 99) - (WARNING_PRIORITY[b.type] || 99));
        
        warningsSection.classList.remove('hidden');
        warningsFooter.classList.remove('hidden');
        safeMessage.classList.add('hidden');
        
        warningsList.innerHTML = warnings.map(w => {
            let emailHtml = '';
            
            if ((w.type === 'wire-fraud' || w.type === 'phishing-urgency') && w.keywords) {
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
                            <span class="warning-email-value known">${w.entityClaimed}</span>
                        </div>
                        <div class="warning-email-row">
                            <span class="warning-email-label">Actually from:</span>
                            <span class="warning-email-value suspicious" style="white-space: nowrap;">${w.senderEmail}</span>
                        </div>
                        <div class="warning-email-row">
                            <span class="warning-email-label">Legitimate domains:</span>
                            <span class="warning-email-value known">${formatDomainsList(w.legitimateDomains)}</span>
                        </div>
                    </div>
                `;
            } else if (w.type === 'brand-impersonation') {
                // v3.6.0: Three-row format with nowrap domains
                emailHtml = `
                    <div class="warning-emails">
                        <div class="warning-email-row">
                            <span class="warning-email-label">This email claims to be from:</span>
                            <span class="warning-email-value known">${w.brandClaimed}</span>
                        </div>
                        <div class="warning-email-row">
                            <span class="warning-email-label">But is actually from:</span>
                            <span class="warning-email-value suspicious">${wrapDomain(w.senderDomain)}</span>
                        </div>
                        <div class="warning-email-row">
                            <span class="warning-email-label">Legitimate domains:</span>
                            <span class="warning-email-value known">${formatDomainsList(w.legitimateDomains)}</span>
                        </div>
                    </div>
                `;
            } else if (w.type === 'international-sender') {
                // v3.6.0: Simplified international sender format - country on second line
                emailHtml = `
                    <div class="warning-international-info">
                        <p>This sender's email address includes a country code: ${w.tld}<br>(${w.country})</p>
                        <p style="margin-top: 8px;">Be careful, this could be a phishing attempt.</p>
                        <p style="margin-top: 8px;">Most legitimate business emails use .com domains.</p>
                    </div>
                `;
            } else if (w.type === 'impersonation') {
                emailHtml = `
                    <div class="warning-emails">
                        <div class="warning-email-row">
                            <span class="warning-email-label">This email claims to be from:</span>
                            <span class="warning-email-value known" style="white-space: nowrap;">${w.matchedEmail}</span>
                        </div>
                        <div class="warning-email-row">
                            <span class="warning-email-label">But is actually from:</span>
                            <span class="warning-email-value suspicious" style="white-space: nowrap;">${w.senderEmail}</span>
                        </div>
                    </div>
                `;
            } else if (w.type === 'recipient-spoof') {
                emailHtml = `
                    <div class="warning-emails">
                        <div class="warning-email-row">
                            <span class="warning-email-label">Display name shows:</span>
                            <span class="warning-email-value suspicious">${w.matchedEmail}</span>
                        </div>
                        <div class="warning-email-row">
                            <span class="warning-email-label">But actually from:</span>
                            <span class="warning-email-value suspicious" style="white-space: nowrap;">${w.senderEmail}</span>
                        </div>
                    </div>
                `;
            } else if (w.senderEmail && w.matchedEmail) {
                const matchLabel = w.type === 'replyto-mismatch' ? 'Replies go to' : w.type === 'on-behalf-of' ? 'On behalf of' : w.type === 'gibberish-domain' ? 'Domain' : 'Similar to';
                emailHtml = `
                    <div class="warning-emails">
                        <div class="warning-email-row">
                            <span class="warning-email-label">Sender:</span>
                            <span class="warning-email-value suspicious" style="white-space: nowrap;">${w.senderEmail}</span>
                        </div>
                        <div class="warning-email-row">
                            <span class="warning-email-label">${matchLabel}:</span>
                            <span class="warning-email-value ${w.type === 'gibberish-domain' ? 'suspicious' : 'known'}" style="white-space: nowrap;">${w.matchedEmail}</span>
                        </div>
                        ${w.reason ? `<div class="warning-reason">${w.reason}</div>` : ''}
                    </div>
                `;
            } else if (w.detail) {
                emailHtml = `<div class="warning-reason">${w.detail}</div>`;
            }
            
            return `
                <div class="warning-item ${w.severity}">
                    <div class="warning-title">${w.title}</div>
                    <div class="warning-description">${w.description}</div>
                    ${emailHtml}
                </div>
            `;
        }).join('');
    } else {
        warningsSection.classList.add('hidden');
        warningsFooter.classList.add('hidden');
        safeMessage.classList.remove('hidden');
    }
}
