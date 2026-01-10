// Email Fraud Detector - Outlook Web Add-in
// Version 3.2.0 - Organization impersonation detection (FIXED)

// ============================================
// CONFIGURATION
// ============================================
const CONFIG = {
    clientId: '622f0452-d622-45d1-aab3-3a2026389dd3',
    redirectUri: 'https://journeybrennan22-bot.github.io/outlook-fraud-detector/src/taskpane.html',
    scopes: ['User.Read', 'Contacts.Read'],
    trustedDomains: ['baynac.com', 'purelogicescrow.com', 'journeyinsurance.com']
};

// ============================================
// ORGANIZATION IMPERSONATION TARGETS
// Maps commonly impersonated entities to their legitimate domains
// ============================================
const IMPERSONATION_TARGETS = {
    // US Government - Federal
    "social security": ["ssa.gov"],
    "social security administration": ["ssa.gov"],
    "ssa": ["ssa.gov"],
    "internal revenue service": ["irs.gov"],
    "irs": ["irs.gov"],
    "treasury department": ["treasury.gov"],
    "us treasury": ["treasury.gov"],
    "department of treasury": ["treasury.gov"],
    "medicare": ["medicare.gov", "cms.gov"],
    "medicaid": ["medicaid.gov", "cms.gov"],
    "cms": ["cms.gov"],
    "federal bureau of investigation": ["fbi.gov"],
    "fbi": ["fbi.gov"],
    "veterans affairs": ["va.gov"],
    "department of veterans affairs": ["va.gov"],
    "va benefits": ["va.gov"],
    "federal trade commission": ["ftc.gov"],
    "ftc": ["ftc.gov"],
    "department of homeland security": ["dhs.gov"],
    "homeland security": ["dhs.gov"],
    "dhs": ["dhs.gov"],
    "immigration": ["uscis.gov"],
    "uscis": ["uscis.gov"],
    "us citizenship": ["uscis.gov"],
    "department of justice": ["justice.gov", "usdoj.gov"],
    "doj": ["justice.gov"],
    "department of labor": ["dol.gov"],
    "unemployment benefits": ["dol.gov"],
    "small business administration": ["sba.gov"],
    "sba": ["sba.gov"],
    "sba loan": ["sba.gov"],
    "federal housing administration": ["hud.gov"],
    "fha": ["hud.gov"],
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
    "wells fargo": ["wellsfargo.com"],
    "bank of america": ["bankofamerica.com", "bofa.com"],
    "bofa": ["bankofamerica.com", "bofa.com"],
    "chase": ["chase.com", "jpmorganchase.com"],
    "jpmorgan": ["chase.com", "jpmorganchase.com"],
    "jp morgan": ["chase.com", "jpmorganchase.com"],
    "citibank": ["citi.com", "citibank.com"],
    "citi": ["citi.com", "citibank.com"],
    "citigroup": ["citi.com", "citibank.com"],
    "us bank": ["usbank.com"],
    "u.s. bank": ["usbank.com"],
    "pnc": ["pnc.com"],
    "pnc bank": ["pnc.com"],
    "capital one": ["capitalone.com"],
    "td bank": ["td.com", "tdbank.com"],
    "truist": ["truist.com"],
    "regions": ["regions.com"],
    "regions bank": ["regions.com"],
    "fifth third": ["53.com"],
    "fifth third bank": ["53.com"],
    "huntington bank": ["huntington.com"],
    "huntington": ["huntington.com"],
    "ally bank": ["ally.com"],
    "ally": ["ally.com"],
    "discover": ["discover.com"],
    "discover bank": ["discover.com"],
    "american express": ["americanexpress.com", "amex.com"],
    "amex": ["americanexpress.com", "amex.com"],
    "navy federal": ["navyfederal.org"],
    "navy federal credit union": ["navyfederal.org"],
    "usaa": ["usaa.com"],

    // Tech Companies
    "microsoft": ["microsoft.com", "office.com", "live.com", "outlook.com"],
    "office 365": ["microsoft.com", "office.com"],
    "microsoft 365": ["microsoft.com", "office.com"],
    "outlook": ["microsoft.com", "outlook.com"],
    "onedrive": ["microsoft.com"],
    "google": ["google.com", "gmail.com"],
    "gmail": ["google.com", "gmail.com"],
    "google drive": ["google.com"],
    "apple": ["apple.com", "icloud.com"],
    "icloud": ["apple.com", "icloud.com"],
    "apple id": ["apple.com"],
    "itunes": ["apple.com"],
    "app store": ["apple.com"],
    "amazon": ["amazon.com", "amazon.co.uk", "amazonaws.com"],
    "amazon prime": ["amazon.com"],
    "aws": ["amazon.com", "amazonaws.com"],
    "meta": ["meta.com", "facebook.com", "fb.com"],
    "facebook": ["facebook.com", "fb.com", "meta.com"],
    "instagram": ["instagram.com", "meta.com"],
    "whatsapp": ["whatsapp.com", "meta.com"],
    "linkedin": ["linkedin.com"],
    "twitter": ["twitter.com", "x.com"],
    "x.com": ["twitter.com", "x.com"],
    "netflix": ["netflix.com"],
    "spotify": ["spotify.com"],
    "zoom": ["zoom.us"],
    "zoom meeting": ["zoom.us"],
    "dropbox": ["dropbox.com"],

    // Document Signing / Business Tools
    "docusign": ["docusign.com", "docusign.net"],
    "adobe": ["adobe.com"],
    "adobe sign": ["adobe.com", "adobesign.com"],
    "acrobat": ["adobe.com"],
    "intuit": ["intuit.com"],
    "quickbooks": ["intuit.com", "quickbooks.com"],
    "turbotax": ["intuit.com", "turbotax.com"],
    "salesforce": ["salesforce.com"],

    // Payment Platforms
    "paypal": ["paypal.com"],
    "venmo": ["venmo.com"],
    "zelle": ["zellepay.com"],
    "cash app": ["cash.app", "square.com"],
    "cashapp": ["cash.app", "square.com"],
    "square": ["square.com", "squareup.com"],
    "stripe": ["stripe.com"],
    "wise": ["wise.com"],
    "transferwise": ["wise.com"],

    // Telecoms
    "verizon": ["verizon.com", "verizonwireless.com"],
    "at&t": ["att.com"],
    "att": ["att.com"],
    "t-mobile": ["t-mobile.com"],
    "tmobile": ["t-mobile.com"],
    "comcast": ["comcast.com", "xfinity.com"],
    "xfinity": ["comcast.com", "xfinity.com"],
    "spectrum": ["spectrum.com", "charter.com"],

    // Credit Bureaus
    "equifax": ["equifax.com"],
    "experian": ["experian.com"],
    "transunion": ["transunion.com"],
    "credit karma": ["creditkarma.com"],

    // Title & Escrow Companies
    "fidelity national title": ["fnf.com", "fntg.com"],
    "fidelity title": ["fnf.com", "fntg.com"],
    "first american title": ["firstam.com"],
    "first american": ["firstam.com"],
    "chicago title": ["chicagotitle.com", "fnf.com"],
    "stewart title": ["stewart.com"],
    "old republic title": ["oldrepublictitle.com", "oldrepublic.com"],

    // Mortgage / Lending
    "fannie mae": ["fanniemae.com"],
    "freddie mac": ["freddiemac.com"],
    "rocket mortgage": ["rocketmortgage.com", "quickenloans.com"],
    "quicken loans": ["rocketmortgage.com", "quickenloans.com"],
    "united wholesale mortgage": ["uwm.com"],
    "uwm": ["uwm.com"],
    "loandepot": ["loandepot.com"],
    "loan depot": ["loandepot.com"],
    "better mortgage": ["better.com"],

    // Real Estate Platforms
    "zillow": ["zillow.com"],
    "redfin": ["redfin.com"],
    "realtor.com": ["realtor.com"],
    "trulia": ["trulia.com"],

    // Crypto / Investment
    "coinbase": ["coinbase.com"],
    "robinhood": ["robinhood.com"],
    "fidelity investments": ["fidelity.com"],
    "fidelity": ["fidelity.com"],
    "charles schwab": ["schwab.com"],
    "schwab": ["schwab.com"],
    "vanguard": ["vanguard.com"],
    "e*trade": ["etrade.com"],
    "etrade": ["etrade.com"],
    "td ameritrade": ["tdameritrade.com"],

    // E-Commerce / Retail
    "walmart": ["walmart.com"],
    "target": ["target.com"],
    "ebay": ["ebay.com"],
    "best buy": ["bestbuy.com"],
    "costco": ["costco.com"],

    // International - UK
    "hmrc": ["gov.uk"],
    "nhs": ["nhs.uk"],
    "dvla": ["gov.uk"],
    "royal mail": ["royalmail.com"],
    "barclays": ["barclays.co.uk", "barclays.com"],
    "hsbc": ["hsbc.co.uk", "hsbc.com"],
    "lloyds": ["lloydsbank.com"],
    "natwest": ["natwest.com"],

    // International - Canada
    "canada revenue agency": ["canada.ca"],
    "cra": ["canada.ca"],
    "service canada": ["canada.ca"],
    "canada post": ["canadapost.ca", "canadapost-postescanada.ca"],
    "rbc": ["rbc.com", "royalbank.com"],
    "td canada": ["td.com"],
    "scotiabank": ["scotiabank.com"],
    "bmo": ["bmo.com"],
    "cibc": ["cibc.com"],

    // International - Australia
    "ato": ["ato.gov.au"],
    "australian taxation office": ["ato.gov.au"],
    "centrelink": ["servicesaustralia.gov.au"],
    "australia post": ["auspost.com.au"],
    "commbank": ["commbank.com.au"],
    "commonwealth bank": ["commbank.com.au"],
    "westpac": ["westpac.com.au"],
    "anz": ["anz.com.au", "anz.com"],
    "nab": ["nab.com.au"]
};

// ============================================
// WIRE FRAUD KEYWORDS
// ============================================
const WIRE_FRAUD_KEYWORDS = [
    'wire transfer', 'wire instructions', 'wiring instructions',
    'bank account', 'account number', 'routing number',
    'ach transfer', 'direct deposit', 'beneficiary',
    'updated bank', 'new bank', 'changed bank',
    'updated payment', 'new payment info',
    'closing funds', 'earnest money', 'escrow funds',
    'settlement funds', 'wire to', 'remittance',
    'keep this confidential', 'keep this quiet',
    'dont mention this', 'between us',
    'social security', 'ssn', 'tax id',
    'W-9', 'W9', 'ein number',
    'zelle', 'venmo', 'cryptocurrency', 'bitcoin'
];

// ============================================
// DECEPTIVE TLD PATTERNS
// ============================================
const DECEPTIVE_TLDS = [
    { pattern: /\.com\.co$/i, readable: '.com.co (Colombia)', fake: '.com' },
    { pattern: /\.com\.br$/i, readable: '.com.br (Brazil)', fake: '.com' },
    { pattern: /\.com\.mx$/i, readable: '.com.mx (Mexico)', fake: '.com' },
    { pattern: /\.com\.ng$/i, readable: '.com.ng (Nigeria)', fake: '.com' },
    { pattern: /\.com\.ru$/i, readable: '.com.ru (Russia)', fake: '.com' },
    { pattern: /\.com\.cn$/i, readable: '.com.cn (China)', fake: '.com' },
    { pattern: /\.org\.uk$/i, readable: '.org.uk (UK)', fake: '.org' }
];

// ============================================
// GLOBAL STATE
// ============================================
let knownContacts = new Set();
let isAuthenticated = false;
let accessToken = null;
let autoScanEnabled = true;

// ============================================
// INITIALIZATION
// ============================================
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById('scan-btn').onclick = scanCurrentEmail;
        document.getElementById('connect-btn').onclick = connectToMicrosoft;
        
        // Check for stored token
        const storedToken = localStorage.getItem('msalToken');
        if (storedToken) {
            accessToken = storedToken;
            isAuthenticated = true;
            updateAuthUI(true);
            fetchContacts();
        }
        
        // Auto-scan on email change
        Office.context.mailbox.addHandlerAsync(
            Office.EventType.ItemChanged,
            () => {
                if (autoScanEnabled) {
                    setTimeout(scanCurrentEmail, 500);
                }
            }
        );
        
        // Initial scan
        if (autoScanEnabled) {
            console.log('Auto-scan enabled');
            setTimeout(scanCurrentEmail, 1000);
        }
    }
});

// ============================================
// AUTHENTICATION
// ============================================
async function connectToMicrosoft() {
    const btn = document.getElementById('connect-btn');
    btn.textContent = 'Connecting...';
    btn.disabled = true;
    
    try {
        // Use popup auth
        const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?` +
            `client_id=${CONFIG.clientId}` +
            `&response_type=token` +
            `&redirect_uri=${encodeURIComponent(CONFIG.redirectUri)}` +
            `&scope=${encodeURIComponent(CONFIG.scopes.join(' '))}` +
            `&response_mode=fragment` +
            `&prompt=consent`;
        
        const popup = window.open(authUrl, 'auth', 'width=500,height=600');
        
        // Listen for the redirect
        const checkPopup = setInterval(() => {
            try {
                if (popup.closed) {
                    clearInterval(checkPopup);
                    btn.textContent = '🔗 Connect Microsoft';
                    btn.disabled = false;
                    return;
                }
                
                const hash = popup.location.hash;
                if (hash && hash.includes('access_token')) {
                    clearInterval(checkPopup);
                    popup.close();
                    
                    // Parse token from hash
                    const params = new URLSearchParams(hash.substring(1));
                    accessToken = params.get('access_token');
                    
                    if (accessToken) {
                        localStorage.setItem('msalToken', accessToken);
                        isAuthenticated = true;
                        updateAuthUI(true);
                        fetchContacts();
                    }
                }
            } catch (e) {
                // Cross-origin error - popup still on Microsoft domain
            }
        }, 500);
        
    } catch (error) {
        console.error('Auth error:', error);
        btn.textContent = '🔗 Connect Microsoft';
        btn.disabled = false;
        showError('Authentication failed: ' + error.message);
    }
}

function updateAuthUI(connected) {
    const btn = document.getElementById('connect-btn');
    const status = document.getElementById('auth-status');
    
    if (connected) {
        btn.textContent = '✓ Connected';
        btn.disabled = true;
        btn.style.background = '#4CAF50';
        if (status) status.textContent = 'Connected to Microsoft Graph';
    } else {
        btn.textContent = '🔗 Connect Microsoft';
        btn.disabled = false;
        btn.style.background = '';
        if (status) status.textContent = '';
    }
}

// ============================================
// CONTACTS
// ============================================
async function fetchContacts() {
    if (!accessToken) return;
    
    console.log('Fetching contacts...');
    
    try {
        const response = await fetch('https://graph.microsoft.com/v1.0/me/contacts?$top=500&$select=emailAddresses', {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            const contacts = data.value || [];
            
            contacts.forEach(contact => {
                if (contact.emailAddresses) {
                    contact.emailAddresses.forEach(email => {
                        if (email.address) {
                            knownContacts.add(email.address.toLowerCase());
                        }
                    });
                }
            });
            
            console.log('Fetched', knownContacts.size, 'contacts');
            
            // Also add trusted domains
            CONFIG.trustedDomains.forEach(domain => {
                knownContacts.add(`@${domain}`);
            });
            
            console.log('Total known contacts:', knownContacts.size);
        } else if (response.status === 401) {
            // Token expired
            localStorage.removeItem('msalToken');
            isAuthenticated = false;
            accessToken = null;
            updateAuthUI(false);
        }
    } catch (error) {
        console.error('Error fetching contacts:', error);
    }
}

// ============================================
// HELPER FUNCTIONS
// ============================================

/**
 * Escape special regex characters in a string
 */
function escapeRegex(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Format entity name for display (capitalize first letters)
 */
function formatEntityName(name) {
    return name.split(' ').map(word => 
        word.charAt(0).toUpperCase() + word.slice(1)
    ).join(' ');
}

/**
 * Detect organization impersonation
 * Checks if sender claims to be a known organization but sends from wrong domain
 */
function detectOrganizationImpersonation(senderEmail, displayName, subject) {
    const senderDomain = senderEmail.split('@')[1]?.toLowerCase();
    if (!senderDomain) return null;
    
    // Combine display name and subject for searching
    const searchText = `${displayName} ${subject}`.toLowerCase();
    
    // Check each impersonation target
    for (const [entityName, legitimateDomains] of Object.entries(IMPERSONATION_TARGETS)) {
        // Use word boundary matching to prevent false positives (e.g., "chase" in "purchase")
        const entityPattern = new RegExp(`\\b${escapeRegex(entityName)}\\b`, 'i');
        
        if (entityPattern.test(searchText)) {
            // Check if sender domain matches any legitimate domain
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
 * Check for deceptive TLD
 */
function checkDeceptiveTLD(email) {
    const domain = email.split('@')[1];
    if (!domain) return null;
    
    for (const tld of DECEPTIVE_TLDS) {
        if (tld.pattern.test(domain)) {
            const fakeDomain = domain.replace(tld.pattern, tld.fake);
            return {
                domain: domain,
                readable: tld.readable,
                fakingAs: fakeDomain,
                warning: `Domain "${domain}" looks like "${fakeDomain}" but is registered elsewhere`
            };
        }
    }
    
    return null;
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
 * Check for lookalike domain
 */
function checkLookalikeDomain(senderEmail) {
    const senderDomain = senderEmail.split('@')[1]?.toLowerCase();
    if (!senderDomain) return null;
    
    for (const contact of knownContacts) {
        if (contact.startsWith('@')) continue; // Skip domain entries
        
        const contactDomain = contact.split('@')[1]?.toLowerCase();
        if (!contactDomain || contactDomain === senderDomain) continue;
        
        const distance = levenshteinDistance(senderDomain, contactDomain);
        if (distance > 0 && distance <= 2) {
            return {
                senderDomain: senderDomain,
                similarTo: contactDomain,
                distance: distance
            };
        }
    }
    
    return null;
}

// ============================================
// EMAIL SCANNING
// ============================================
function scanCurrentEmail() {
    const item = Office.context.mailbox.item;
    
    if (!item) {
        showError('No email selected');
        return;
    }
    
    showLoading();
    
    try {
        const from = item.from;
        const subject = item.subject;
        
        // Get body
        item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
            // Get reply-to if available
            if (item.internetHeaders) {
                item.internetHeaders.getAsync(['Reply-To'], (headerResult) => {
                    let replyTo = null;
                    if (headerResult.status === Office.AsyncResultStatus.Succeeded) {
                        const headers = headerResult.value;
                        if (headers && headers['Reply-To']) {
                            replyTo = headers['Reply-To'];
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
    
    // 2. Organization Impersonation Detection
    const orgImpersonation = detectOrganizationImpersonation(senderEmail, displayName, subject);
    if (orgImpersonation) {
        warnings.push({
            type: 'org-impersonation',
            severity: 'critical',
            title: 'Organization Impersonation',
            description: orgImpersonation.message,
            senderEmail: senderEmail,
            matchedEmail: orgImpersonation.legitimateDomains.join(', '),
            entityClaimed: orgImpersonation.entityClaimed
        });
    }
    
    // 3. Deceptive TLD
    const deceptiveTLD = checkDeceptiveTLD(senderEmail);
    if (deceptiveTLD) {
        warnings.push({
            type: 'deceptive-tld',
            severity: 'critical',
            title: 'Deceptive Domain',
            description: deceptiveTLD.warning,
            senderEmail: senderEmail,
            matchedEmail: deceptiveTLD.fakingAs
        });
    }
    
    // 4. Lookalike Domain (only if not known contact)
    if (!isKnownContact) {
        const lookalike = checkLookalikeDomain(senderEmail);
        if (lookalike) {
            warnings.push({
                type: 'lookalike',
                severity: 'critical',
                title: 'Lookalike Domain',
                description: `Domain "${lookalike.senderDomain}" is similar to "${lookalike.similarTo}"`,
                senderEmail: senderEmail,
                matchedEmail: lookalike.similarTo
            });
        }
    }
    
    // 5. Wire Fraud Keywords
    const foundKeywords = [];
    for (const keyword of WIRE_FRAUD_KEYWORDS) {
        if (fullContent.includes(keyword.toLowerCase())) {
            foundKeywords.push(keyword);
        }
    }
    
    if (foundKeywords.length > 0) {
        warnings.push({
            type: 'wire-fraud',
            severity: 'critical',
            title: 'Dangerous Keywords Detected',
            description: 'This email contains terms commonly used in wire fraud.',
            keywords: foundKeywords
        });
    }
    
    // Display results
    displayResults(warnings, emailData);
}

// ============================================
// UI FUNCTIONS
// ============================================
function showLoading() {
    const results = document.getElementById('results');
    results.innerHTML = '<div class="loading">Analyzing email...</div>';
}

function showError(message) {
    const results = document.getElementById('results');
    results.innerHTML = `<div class="error">Error: ${message}</div>`;
}

function displayResults(warnings, emailData) {
    const results = document.getElementById('results');
    
    if (warnings.length === 0) {
        results.innerHTML = `
            <div class="safe-banner">
                <div class="safe-icon">✓</div>
                <div class="safe-text">No threats detected</div>
            </div>
            <div class="sender-info">
                <strong>From:</strong> ${emailData.from.displayName}<br>
                <strong>Email:</strong> ${emailData.from.emailAddress}
            </div>
        `;
        return;
    }
    
    // Has warnings
    const warningCount = warnings.length;
    
    let html = `
        <div class="warning-header">
            <div class="warning-icon">🚨</div>
            <div class="warning-count">${warningCount} Warning${warningCount > 1 ? 's' : ''} Detected</div>
        </div>
    `;
    
    for (const warning of warnings) {
        html += `<div class="warning-item ${warning.severity}">`;
        html += `<div class="warning-title">${warning.title}</div>`;
        html += `<div class="warning-description">${warning.description}</div>`;
        
        if (warning.type === 'wire-fraud' && warning.keywords) {
            html += `<div class="warning-keywords">`;
            html += `<span class="keyword-label">TRIGGERED BY:</span>`;
            for (const kw of warning.keywords.slice(0, 5)) {
                html += `<span class="keyword-tag">${kw}</span>`;
            }
            html += `</div>`;
        } else if (warning.type === 'org-impersonation') {
            html += `<div class="warning-details">`;
            html += `<div class="detail-row"><span class="detail-label">Claims to be:</span> <span class="detail-value entity">${warning.entityClaimed}</span></div>`;
            html += `<div class="detail-row"><span class="detail-label">Actually from:</span> <span class="detail-value suspicious">${warning.senderEmail}</span></div>`;
            html += `<div class="detail-row"><span class="detail-label">Legitimate domains:</span> <span class="detail-value safe">${warning.matchedEmail}</span></div>`;
            html += `</div>`;
        } else if (warning.senderEmail && warning.matchedEmail) {
            html += `<div class="warning-details">`;
            html += `<div class="detail-row"><span class="detail-label">From:</span> <span class="detail-value suspicious">${warning.senderEmail}</span></div>`;
            html += `<div class="detail-row"><span class="detail-label">Expected:</span> <span class="detail-value safe">${warning.matchedEmail}</span></div>`;
            html += `</div>`;
        }
        
        html += `</div>`;
    }
    
    html += `
        <div class="learn-more">
            <a href="https://emailfraudalert.com/learn.html" target="_blank">Learn how scams work →</a>
        </div>
    `;
    
    results.innerHTML = html;
}
