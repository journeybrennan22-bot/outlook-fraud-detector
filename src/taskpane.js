// Email Fraud Detector - Outlook Web Add-in
// Version 2.9.0 - Full scam detection (all domains)

// ============================================
// CONFIGURATION
// ============================================
const CONFIG = {
    clientId: '622f0452-d622-45d1-aab3-3a2026389dd3',
    redirectUri: 'https://journeybrennan22-bot.github.io/outlook-fraud-detector/src/taskpane.html',
    scopes: ['User.Read', 'Contacts.Read'],
    trustedDomains: ['baynac.com', 'purelogicescrow.com', 'journeyinsurance.com']
};

// Major brands that scammers commonly impersonate
const PROTECTED_BRANDS = [
    // Banks
    'chase', 'wellsfargo', 'bankofamerica', 'citi', 'citibank', 'usbank', 
    'capitalone', 'pnc', 'tdbank', 'regions', 'suntrust', 'truist',
    'schwab', 'fidelity', 'vanguard', 'merrill', 'morganstanley',
    
    // Tech Companies
    'microsoft', 'apple', 'google', 'amazon', 'meta', 'facebook',
    'netflix', 'spotify', 'adobe', 'dropbox', 'zoom', 'slack',
    
    // Payment Services
    'paypal', 'venmo', 'zelle', 'cashapp', 'stripe', 'square',
    
    // Email Providers
    'gmail', 'outlook', 'yahoo', 'hotmail', 'icloud', 'protonmail',
    
    // Government
    'irs', 'ssa', 'medicare', 'socialsecurity', 'dmv', 'usps',
    
    // Shipping
    'fedex', 'ups', 'dhl', 'usps',
    
    // E-commerce
    'ebay', 'walmart', 'target', 'costco', 'bestbuy', 'homedepot',
    
    // Social Media
    'twitter', 'instagram', 'linkedin', 'tiktok', 'snapchat',
    
    // Real Estate / Title
    'firstamerican', 'fidelitynational', 'oldrepublic', 'stewartitle',
    'chicagotitle', 'northamerican'
];

// Deceptive TLDs that look like .com but aren't
const DECEPTIVE_TLDS = [
    '.com.co', '.com.br', '.com.mx', '.com.ar', '.com.au', '.com.ng',
    '.com.pk', '.com.ph', '.com.ua', '.com.ve', '.com.vn', '.com.tr',
    '.net.co', '.net.br', '.org.co', '.co.uk.com', '.us.com',
    '.co', '.cm', '.cc', '.ru', '.cn', '.tk', '.ml', '.ga', '.cf'
];

// Fraud keywords
const WIRE_FRAUD_KEYWORDS = [
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
    'boss asked', 'executive request', 'president asked',
    
    // Urgency Phrases
    'verify your account', 'verify immediately', 'act now',
    'urgent action required', 'account suspended', 'account will be closed',
    'unusual activity', 'suspicious activity', 'unauthorized access',
    'confirm your identity', 'verify your identity',
    'action required within', 'expires today', 'last chance'
];

// Homoglyph characters (Cyrillic only - removed 0/1 to avoid false positives)
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
    
    // 1. Reply-To Mismatch
    if (emailData.replyTo && emailData.replyTo.toLowerCase() !== senderEmail) {
        warnings.push({
            type: 'replyto-mismatch',
            severity: 'critical',
            title: 'Reply-To Mismatch',
            description: 'Replies will go to a different address than the sender.',
            senderEmail: senderEmail,
            matchedEmail: emailData.replyTo.toLowerCase()
        });
    }
    
    // 2. Deceptive TLD Detection (checks all domains)
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
    
    // 3. Brand Impersonation Detection (checks major brands)
    const brandImpersonation = detectBrandImpersonation(senderDomain, displayName);
    if (brandImpersonation) {
        warnings.push({
            type: 'brand-impersonation',
            severity: 'critical',
            title: 'Possible Brand Impersonation',
            description: brandImpersonation.reason,
            senderEmail: senderEmail,
            matchedEmail: brandImpersonation.brand
        });
    }
    
    // 4. Display Name Impersonation (skip if known contact)
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
    
    // 5. Homoglyph/Unicode Detection
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
    
    // 6. Lookalike Domain Detection (your trusted domains)
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
    
    // 7. Fraud Keywords
    const wireKeywords = detectWireFraudKeywords(fullContent);
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
    
    // 8. Contact Lookalike Detection (skip if known contact)
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
    
    displayResults(warnings, senderEmail);
}

// ============================================
// DETECTION FUNCTIONS
// ============================================

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
 * Detect brand impersonation in domain or display name
 */
function detectBrandImpersonation(domain, displayName) {
    const domainLower = domain.toLowerCase();
    const displayLower = (displayName || '').toLowerCase();
    
    // Get the domain name without TLD
    const domainParts = domainLower.split('.');
    const domainName = domainParts[0];
    
    for (const brand of PROTECTED_BRANDS) {
        // Check for hyphenated brand domains (e.g., paypal-secure.com, chase-alert.com)
        if (domainName.includes(brand) && domainName !== brand) {
            // Make sure it's not the legitimate domain
            const legitimateDomain = brand + '.com';
            if (domainLower !== legitimateDomain && !domainLower.endsWith('.' + legitimateDomain)) {
                return {
                    brand: brand,
                    reason: `Domain contains "${brand}" but is not the official ${brand}.com`
                };
            }
        }
        
        // Check for brand in display name but different domain
        if (displayLower.includes(brand)) {
            const legitimateDomain = brand + '.com';
            // If display name mentions brand but domain is different
            if (!domainLower.includes(brand) && domainLower !== legitimateDomain) {
                return {
                    brand: brand,
                    reason: `Display name mentions "${brand}" but email is from "${domain}"`
                };
            }
        }
        
        // Check for lookalike brand domains (1-2 char difference)
        const distance = levenshteinDistance(domainName, brand);
        if (distance > 0 && distance <= 2 && domainName.length >= 4) {
            return {
                brand: brand,
                reason: `Domain "${domainName}" looks similar to "${brand}" (${distance} character${distance > 1 ? 's' : ''} different)`
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
                reason: `Display name contains "${domain}" but email is from "${senderDomain}"`,
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
                reason: `Display name shows "${nameEmail}" but actual sender domain is "${senderDomain}"`,
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
    
    if (criticalCount > 0) {
        const totalWarnings = criticalCount + mediumCount;
        document.body.classList.add('status-critical');
        statusBadge.className = 'status-badge danger';
        statusIcon.textContent = '🚨';
        statusText.textContent = `${totalWarnings} Warning${totalWarnings > 1 ? 's' : ''} Detected`;
    } else if (mediumCount > 0) {
        document.body.classList.add('status-medium');
        statusBadge.className = 'status-badge warning';
        statusIcon.textContent = '⚠️';
        statusText.textContent = `${mediumCount} Warning${mediumCount > 1 ? 's' : ''} Detected`;
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
                    <div class="warning-keywords">${keywordTags}</div>
                    <div class="warning-advice">
                        <strong>Be careful:</strong> Verify this email is legitimate before clicking links, downloading attachments, or taking any action.
                    </div>
                `;
            } else if (w.senderEmail && w.matchedEmail) {
                const matchLabel = w.type === 'replyto-mismatch' ? 'Replies go to' : 
                                   w.type === 'brand-impersonation' ? 'Impersonating' :
                                   w.type === 'deceptive-tld' ? 'Deceptive TLD' : 'Similar to';
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
