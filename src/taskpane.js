// Email Fraud Detector - Outlook Web Add-in
// Version 3.2.3 - Works with existing HTML structure

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
        console.log('[EFA] Office ready, initializing...');
        
        // Set up retry button
        const retryBtn = document.getElementById('retry-btn');
        if (retryBtn) {
            retryBtn.onclick = scanCurrentEmail;
        }
        
        // Check for stored token
        const storedToken = localStorage.getItem('msalToken');
        if (storedToken) {
            accessToken = storedToken;
            isAuthenticated = true;
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
            console.log('[EFA] Auto-scan enabled');
            setTimeout(scanCurrentEmail, 1000);
        }
    }
});

// ============================================
// CONTACTS
// ============================================
async function fetchContacts() {
    if (!accessToken) return;
    
    console.log('[EFA] Fetching contacts...');
    
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
            
            console.log('[EFA] Fetched', knownContacts.size, 'contacts');
            
            CONFIG.trustedDomains.forEach(domain => {
                knownContacts.add(`@${domain}`);
            });
        } else if (response.status === 401) {
            localStorage.removeItem('msalToken');
            isAuthenticated = false;
            accessToken = null;
        }
    } catch (error) {
        console.error('[EFA] Error fetching contacts:', error);
    }
}

// ============================================
// HELPER FUNCTIONS
// ============================================

function escapeRegex(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function formatEntityName(name) {
    return name.split(' ').map(word => 
        word.charAt(0).toUpperCase() + word.slice(1)
    ).join(' ');
}

function detectOrganizationImpersonation(senderEmail, displayName, subject) {
    const senderDomain = senderEmail.split('@')[1];
    if (!senderDomain) return null;
    
    const senderDomainLower = senderDomain.toLowerCase();
    const searchText = `${displayName || ''} ${subject || ''}`.toLowerCase();
    
    for (const [entityName, legitimateDomains] of Object.entries(IMPERSONATION_TARGETS)) {
        const entityPattern = new RegExp(`\\b${escapeRegex(entityName)}\\b`, 'i');
        
        if (entityPattern.test(searchText)) {
            const isLegitimate = legitimateDomains.some(legit => {
                return senderDomainLower === legit || senderDomainLower.endsWith(`.${legit}`);
            });
            
            if (!isLegitimate) {
                return {
                    entityClaimed: formatEntityName(entityName),
                    senderDomain: senderDomainLower,
                    legitimateDomains: legitimateDomains,
                    message: `Sender claims to be "${formatEntityName(entityName)}" but email comes from ${senderDomainLower}. Legitimate emails come from: ${legitimateDomains.join(', ')}`
                };
            }
        }
    }
    
    return null;
}

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

function checkLookalikeDomain(senderEmail) {
    const senderDomain = senderEmail.split('@')[1];
    if (!senderDomain) return null;
    
    const senderDomainLower = senderDomain.toLowerCase();
    
    for (const contact of knownContacts) {
        if (contact.startsWith('@')) continue;
        
        const contactDomain = contact.split('@')[1];
        if (!contactDomain) continue;
        
        const contactDomainLower = contactDomain.toLowerCase();
        if (contactDomainLower === senderDomainLower) continue;
        
        const distance = levenshteinDistance(senderDomainLower, contactDomainLower);
        if (distance > 0 && distance <= 2) {
            return {
                senderDomain: senderDomainLower,
                similarTo: contactDomainLower,
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
    console.log('[EFA] Scanning email...');
    
    const item = Office.context.mailbox.item;
    
    if (!item) {
        console.log('[EFA] No email selected');
        showError('No email selected');
        return;
    }
    
    showLoading();
    
    try {
        const from = item.from;
        const subject = item.subject || '';
        
        console.log('[EFA] From:', from);
        console.log('[EFA] Subject:', subject);
        
        item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
            console.log('[EFA] Body result status:', bodyResult.status);
            
            const body = bodyResult.status === Office.AsyncResultStatus.Succeeded ? bodyResult.value : '';
            
            if (item.internetHeaders && typeof item.internetHeaders.getAsync === 'function') {
                item.internetHeaders.getAsync(['Reply-To'], (headerResult) => {
                    let replyTo = null;
                    if (headerResult.status === Office.AsyncResultStatus.Succeeded && headerResult.value) {
                        replyTo = headerResult.value['Reply-To'] || null;
                    }
                    
                    const emailData = {
                        from: from,
                        subject: subject,
                        body: body,
                        replyTo: replyTo
                    };
                    
                    performAnalysis(emailData);
                });
            } else {
                const emailData = {
                    from: from,
                    subject: subject,
                    body: body,
                    replyTo: null
                };
                
                performAnalysis(emailData);
            }
        });
    } catch (error) {
        console.error('[EFA] Scan error:', error);
        showError(error.message);
    }
}

function performAnalysis(emailData) {
    console.log('[EFA] Performing analysis...');
    
    try {
        const warnings = [];
        
        const senderEmail = (emailData.from.emailAddress || '').toLowerCase();
        const senderDomain = senderEmail.split('@')[1] || '';
        const displayName = emailData.from.displayName || '';
        const subject = emailData.subject || '';
        const body = emailData.body || '';
        const fullContent = (subject + ' ' + body).toLowerCase();
        
        console.log('[EFA] Sender:', senderEmail);
        console.log('[EFA] Display name:', displayName);
        
        const isKnownContact = knownContacts.has(senderEmail);
        
        // 1. Reply-To Mismatch
        if (emailData.replyTo) {
            const replyToLower = emailData.replyTo.toLowerCase();
            if (replyToLower !== senderEmail) {
                const replyToDomain = replyToLower.split('@')[1] || '';
                if (replyToDomain !== senderDomain) {
                    warnings.push({
                        type: 'replyto-mismatch',
                        severity: 'critical',
                        title: 'Reply-To Mismatch',
                        description: 'Replies will go to a different address than the sender.',
                        senderEmail: senderEmail,
                        matchedEmail: replyToLower
                    });
                }
            }
        }
        
        // 2. Organization Impersonation Detection
        const orgImpersonation = detectOrganizationImpersonation(senderEmail, displayName, subject);
        if (orgImpersonation) {
            console.log('[EFA] Org impersonation detected:', orgImpersonation);
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
        
        // 4. Lookalike Domain
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
        
        console.log('[EFA] Warnings found:', warnings.length);
        displayResults(warnings, emailData);
        
    } catch (error) {
        console.error('[EFA] Analysis error:', error);
        showError('Analysis failed: ' + error.message);
    }
}

// ============================================
// UI FUNCTIONS - Works with existing HTML structure
// ============================================

function showLoading() {
    const loading = document.getElementById('loading');
    const results = document.getElementById('results');
    const error = document.getElementById('error');
    
    if (loading) loading.classList.remove('hidden');
    if (results) results.classList.add('hidden');
    if (error) error.classList.add('hidden');
}

function showError(message) {
    const loading = document.getElementById('loading');
    const results = document.getElementById('results');
    const error = document.getElementById('error');
    const errorMessage = document.getElementById('error-message');
    
    if (loading) loading.classList.add('hidden');
    if (results) results.classList.add('hidden');
    if (error) error.classList.remove('hidden');
    if (errorMessage) errorMessage.textContent = message;
}

function clearElement(el) {
    while (el.firstChild) {
        el.removeChild(el.firstChild);
    }
}

function displayResults(warnings, emailData) {
    console.log('[EFA] Displaying results...');
    
    const loading = document.getElementById('loading');
    const results = document.getElementById('results');
    const error = document.getElementById('error');
    const statusBadge = document.getElementById('status-badge');
    const warningsSection = document.getElementById('warnings-section');
    const warningsList = document.getElementById('warnings-list');
    const warningsFooter = document.getElementById('warnings-footer');
    const safeMessage = document.getElementById('safe-message');
    
    // Hide loading and error, show results
    if (loading) loading.classList.add('hidden');
    if (error) error.classList.add('hidden');
    if (results) results.classList.remove('hidden');
    
    if (warnings.length === 0) {
        // SAFE STATE
        console.log('[EFA] Showing safe state');
        
        if (statusBadge) {
            const icon = statusBadge.querySelector('.status-icon');
            const text = statusBadge.querySelector('.status-text');
            statusBadge.className = 'status-badge safe';
            if (icon) icon.textContent = '✓';
            if (text) text.textContent = 'No threats detected';
        }
        
        if (warningsSection) warningsSection.classList.add('hidden');
        if (safeMessage) safeMessage.classList.remove('hidden');
        
    } else {
        // WARNING STATE
        console.log('[EFA] Showing warning state');
        
        if (statusBadge) {
            const icon = statusBadge.querySelector('.status-icon');
            const text = statusBadge.querySelector('.status-text');
            statusBadge.className = 'status-badge danger';
            if (icon) icon.textContent = '🚨';
            if (text) text.textContent = warnings.length + ' Warning' + (warnings.length > 1 ? 's' : '') + ' Detected';
        }
        
        if (safeMessage) safeMessage.classList.add('hidden');
        if (warningsSection) warningsSection.classList.remove('hidden');
        if (warningsFooter) warningsFooter.classList.remove('hidden');
        
        // Clear and populate warnings list
        if (warningsList) {
            clearElement(warningsList);
            
            for (const warning of warnings) {
                const warningItem = document.createElement('div');
                warningItem.className = 'warning-item ' + warning.severity;
                
                const warningTitle = document.createElement('div');
                warningTitle.className = 'warning-title';
                warningTitle.textContent = warning.title;
                warningItem.appendChild(warningTitle);
                
                const warningDesc = document.createElement('div');
                warningDesc.className = 'warning-description';
                warningDesc.textContent = warning.description;
                warningItem.appendChild(warningDesc);
                
                if (warning.type === 'wire-fraud' && warning.keywords) {
                    const keywordsDiv = document.createElement('div');
                    keywordsDiv.className = 'warning-keywords';
                    
                    const keywordLabel = document.createElement('span');
                    keywordLabel.className = 'keyword-label';
                    keywordLabel.textContent = 'TRIGGERED BY: ';
                    keywordsDiv.appendChild(keywordLabel);
                    
                    for (const kw of warning.keywords.slice(0, 5)) {
                        const keywordTag = document.createElement('span');
                        keywordTag.className = 'keyword-tag';
                        keywordTag.textContent = kw;
                        keywordsDiv.appendChild(keywordTag);
                    }
                    
                    warningItem.appendChild(keywordsDiv);
                } else if (warning.type === 'org-impersonation') {
                    const detailsDiv = document.createElement('div');
                    detailsDiv.className = 'warning-details';
                    
                    // Claims to be
                    const row1 = document.createElement('div');
                    row1.className = 'detail-row';
                    const label1 = document.createElement('span');
                    label1.className = 'detail-label';
                    label1.textContent = 'Claims to be: ';
                    row1.appendChild(label1);
                    const value1 = document.createElement('span');
                    value1.className = 'detail-value entity';
                    value1.textContent = warning.entityClaimed;
                    row1.appendChild(value1);
                    detailsDiv.appendChild(row1);
                    
                    // Actually from
                    const row2 = document.createElement('div');
                    row2.className = 'detail-row';
                    const label2 = document.createElement('span');
                    label2.className = 'detail-label';
                    label2.textContent = 'Actually from: ';
                    row2.appendChild(label2);
                    const value2 = document.createElement('span');
                    value2.className = 'detail-value suspicious';
                    value2.textContent = warning.senderEmail;
                    row2.appendChild(value2);
                    detailsDiv.appendChild(row2);
                    
                    // Legitimate domains
                    const row3 = document.createElement('div');
                    row3.className = 'detail-row';
                    const label3 = document.createElement('span');
                    label3.className = 'detail-label';
                    label3.textContent = 'Legitimate domains: ';
                    row3.appendChild(label3);
                    const value3 = document.createElement('span');
                    value3.className = 'detail-value safe';
                    value3.textContent = warning.matchedEmail;
                    row3.appendChild(value3);
                    detailsDiv.appendChild(row3);
                    
                    warningItem.appendChild(detailsDiv);
                } else if (warning.senderEmail && warning.matchedEmail) {
                    const detailsDiv = document.createElement('div');
                    detailsDiv.className = 'warning-details';
                    
                    const row1 = document.createElement('div');
                    row1.className = 'detail-row';
                    const label1 = document.createElement('span');
                    label1.className = 'detail-label';
                    label1.textContent = 'From: ';
                    row1.appendChild(label1);
                    const value1 = document.createElement('span');
                    value1.className = 'detail-value suspicious';
                    value1.textContent = warning.senderEmail;
                    row1.appendChild(value1);
                    detailsDiv.appendChild(row1);
                    
                    const row2 = document.createElement('div');
                    row2.className = 'detail-row';
                    const label2 = document.createElement('span');
                    label2.className = 'detail-label';
                    label2.textContent = 'Expected: ';
                    row2.appendChild(label2);
                    const value2 = document.createElement('span');
                    value2.className = 'detail-value safe';
                    value2.textContent = warning.matchedEmail;
                    row2.appendChild(value2);
                    detailsDiv.appendChild(row2);
                    
                    warningItem.appendChild(detailsDiv);
                }
                
                warningsList.appendChild(warningItem);
            }
        }
    }
    
    console.log('[EFA] Results displayed successfully');
}
