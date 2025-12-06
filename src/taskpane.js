// Email Fraud Detector - Outlook Web Add-in
// Version 2.4.0

// ============================================
// CONFIGURATION
// ============================================
const CONFIG = {
    clientId: '622f0452-d622-45d1-aab3-3a2026389dd3',
    redirectUri: 'https://journeybrennan22-bot.github.io/outlook-fraud-detector/src/taskpane.html',
    scopes: ['User.Read', 'Contacts.Read'],
    trustedDomains: ['baynac.com', 'purelogicescrow.com', 'journeyinsurance.com']
};

// Fraud keywords
const WIRE_FRAUD_KEYWORDS = [
    // Wire/Money Movement
    'wire', 'wiring',
    'ach transfer', 'direct deposit',
    'zelle', 'venmo', 'cryptocurrency', 'bitcoin',
    'send funds', 'transfer funds', 'remit funds',
    
    // Bank/Account Info
    'bank account', 'account number', 'routing number',
    'aba number', 'swift code', 'iban',
    'bank statement', 'voided check',
    
    // Account Changes (red flag)
    'updated bank', 'new bank', 'changed bank',
    'updated payment', 'new payment info',
    'changed account', 'new account details',
    'payment update', 'revised instructions',
    'please update your records',
    
    // Real Estate / Escrow
    'closing funds', 'earnest money', 'escrow funds',
    
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

// Homoglyph characters
const HOMOGLYPHS = {
    'а': 'a', 'е': 'e', 'о': 'o', 'р': 'p', 'с': 'c', 'х': 'x',
    'і': 'i', 'ј': 'j', 'ѕ': 's', 'ԁ': 'd', 'ɡ': 'g', 'ո': 'n',
    'ν': 'v', 'ѡ': 'w', 'у': 'y', 'һ': 'h', 'ⅼ': 'l', 'ｍ': 'm',
    '0': 'o', '1': 'l', '！': '!', '＠': '@'
};

// ============================================
// STATE
// ============================================
let msalInstance = null;
let knownSenders = new Set();
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
    document.getElementById('rescan-btn').addEventListener('click', analyzeCurrentEmail);
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

function toggleScanDetails() {
    const content = document.getElementById('scan-results');
    const icon = document.getElementById('collapse-icon');
    content.classList.toggle('collapsed');
    icon.classList.toggle('collapsed');
}

// ============================================
// AUTHENTICATION & CONTACTS
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

async function fetchContacts() {
    const token = await getAccessToken();
    if (!token) return;
    
    try {
        const response = await fetch('https://graph.microsoft.com/v1.0/me/contacts?$top=500&$select=emailAddresses', {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        
        if (response.ok) {
            const data = await response.json();
            data.value.forEach(contact => {
                if (contact.emailAddresses) {
                    contact.emailAddresses.forEach(email => {
                        if (email.address) {
                            knownSenders.add(email.address.toLowerCase());
                        }
                    });
                }
            });
        }
    } catch (error) {
        console.log('Contacts fetch error:', error);
    }
}

// ============================================
// MAIN ANALYSIS
// ============================================
async function analyzeCurrentEmail() {
    showLoading();
    
    try {
        // Fetch contacts if not already loaded
        if (knownSenders.size === 0) {
            await fetchContacts();
        }
        
        // Get email data - in read mode, from and subject are direct properties
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
                        // Parse Reply-To header
                        const replyToMatch = headers.match(/^Reply-To:\s*(.+)$/mi);
                        if (replyToMatch) {
                            // Extract email from header (handle "Name <email>" format)
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
                // Fallback if getAllInternetHeadersAsync not available
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
    const scanResults = [];
    
    const senderEmail = emailData.from.emailAddress.toLowerCase();
    const senderDomain = senderEmail.split('@')[1] || '';
    const displayName = emailData.from.displayName || '';
    const subject = emailData.subject || '';
    const body = emailData.body || '';
    const fullContent = (subject + ' ' + body).toLowerCase();
    
    // 1. Reply-To Mismatch (MEDIUM)
    if (emailData.replyTo && emailData.replyTo.toLowerCase() !== senderEmail) {
        warnings.push({
            type: 'replyto-mismatch',
            severity: 'medium',
            title: 'Reply-To Mismatch',
            description: 'Replies will go to a different address than the sender.',
            senderEmail: senderEmail,
            matchedEmail: emailData.replyTo.toLowerCase()
        });
        scanResults.push({ check: 'Reply-To Match', status: 'fail' });
    } else {
        scanResults.push({ check: 'Reply-To Match', status: 'pass' });
    }
    
    // 2. Display Name Impersonation (CRITICAL)
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
        scanResults.push({ check: 'Display Name Check', status: 'fail' });
    } else {
        scanResults.push({ check: 'Display Name Check', status: 'pass' });
    }
    
    // 3. Homoglyph/Unicode Detection (CRITICAL)
    const homoglyph = detectHomoglyphs(senderEmail);
    if (homoglyph) {
        warnings.push({
            type: 'homoglyph',
            severity: 'critical',
            title: 'Suspicious Characters Detected',
            description: `The email address contains deceptive characters that look like normal letters.`,
            detail: homoglyph
        });
        scanResults.push({ check: 'Character Analysis', status: 'fail' });
    } else {
        scanResults.push({ check: 'Character Analysis', status: 'pass' });
    }
    
    // 4. Lookalike Domain Detection (CRITICAL)
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
        scanResults.push({ check: 'Domain Similarity', status: 'fail' });
    } else {
        scanResults.push({ check: 'Domain Similarity', status: 'pass' });
    }
    
    // 5. Fraud Keywords (CRITICAL)
    const wireKeywords = detectWireFraudKeywords(fullContent);
    if (wireKeywords.length > 0) {
        warnings.push({
            type: 'wire-fraud',
            severity: 'critical',
            title: 'Fraud Keywords Detected',
            description: `This email contains suspicious terms: "${wireKeywords.slice(0, 3).join('", "')}"${wireKeywords.length > 3 ? '...' : ''}`,
            isWireFraud: true
        });
        scanResults.push({ check: 'Fraud Keywords', status: 'fail' });
    } else {
        scanResults.push({ check: 'Fraud Keywords', status: 'pass' });
    }
    
    // 6. Contact Lookalike Detection (CRITICAL)
    const contactLookalike = detectContactLookalike(senderEmail);
    if (contactLookalike) {
        warnings.push({
            type: 'contact-lookalike',
            severity: 'critical',
            title: 'Similar to Known Contact',
            description: `This email is suspiciously similar to someone in your contacts. ${contactLookalike.reason}`,
            senderEmail: contactLookalike.incomingEmail,
            matchedEmail: contactLookalike.matchedContact
        });
        scanResults.push({ check: 'Contact Match', status: 'fail' });
    } else {
        scanResults.push({ check: 'Contact Match', status: 'pass' });
    }
    
    // 7. First-Time Sender Check
    const isFirstTime = !knownSenders.has(senderEmail) && !isTrustedDomain(senderDomain);
    if (isFirstTime) {
        scanResults.push({ check: 'Known Sender', status: 'info', note: 'First-time sender' });
    } else {
        scanResults.push({ check: 'Known Sender', status: 'pass' });
    }
    
    displayResults(warnings, scanResults, isFirstTime);
}

// ============================================
// DETECTION FUNCTIONS
// ============================================
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
        const actualEmail = senderDomain;
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
    
    for (const contactEmail of knownSenders) {
        if (contactEmail === incomingEmail) continue;
        
        const contactParts = parseEmailParts(contactEmail);
        if (!contactParts) continue;
        
        // Same domain, similar username (1-4 chars different)
        if (incomingParts.domain === contactParts.domain) {
            const localDistance = levenshteinDistance(incomingParts.local, contactParts.local);
            if (localDistance > 0 && localDistance <= 4) {
                return {
                    incomingEmail: incomingEmail,
                    matchedContact: contactEmail,
                    reason: `Username is ${localDistance} character${localDistance > 1 ? 's' : ''} different`
                };
            }
        }
        
        // Similar domain (1-2 chars different)
        const domainDistance = levenshteinDistance(incomingParts.domain, contactParts.domain);
        if (domainDistance > 0 && domainDistance <= 2) {
            return {
                incomingEmail: incomingEmail,
                matchedContact: contactEmail,
                reason: `Domain is ${domainDistance} character${domainDistance > 1 ? 's' : ''} different`
            };
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

function displayResults(warnings, scanResults, isFirstTime) {
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
    } else if (isFirstTime) {
        document.body.classList.add('status-info');
        statusBadge.className = 'status-badge info';
        statusIcon.textContent = '👤';
        statusText.textContent = 'First-Time Sender';
    } else {
        document.body.classList.add('status-safe');
        statusBadge.className = 'status-badge safe';
        statusIcon.textContent = '✅';
        statusText.textContent = 'No Issues Detected';
    }
    
    // Display warnings
    const warningsSection = document.getElementById('warnings-section');
    const warningsList = document.getElementById('warnings-list');
    
    if (warnings.length > 0) {
        warningsSection.classList.remove('hidden');
        warningsList.innerHTML = warnings.map(w => {
            let emailHtml = '';
            if (w.senderEmail && w.matchedEmail) {
                const matchLabel = w.type === 'replyto-mismatch' ? 'Replies go to' : 'Similar to';
                emailHtml = `
                    <div class="warning-emails">
                        <div class="warning-email-row">
                            <span class="warning-email-label">Sender:</span>
                            <span class="warning-email-value">${w.senderEmail}</span>
                        </div>
                        <div class="warning-email-row">
                            <span class="warning-email-label">${matchLabel}:</span>
                            <span class="warning-email-value">${w.matchedEmail}</span>
                        </div>
                    </div>
                `;
            } else if (w.detail) {
                emailHtml = `<div class="warning-detail" style="font-size:12px;color:#605e5c;margin-top:8px;">${w.detail}</div>`;
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
    }
    
    // Display scan results
    const scanResultsEl = document.getElementById('scan-results');
    scanResultsEl.innerHTML = scanResults.map(r => `
        <div class="scan-item">
            <span class="scan-check ${r.status === 'pass' ? 'scan-pass' : r.status === 'fail' ? 'scan-fail' : 'scan-info'}">
                ${r.status === 'pass' ? '✓' : r.status === 'fail' ? '✗' : 'ℹ'}
            </span>
            <span>${r.check}${r.note ? ` (${r.note})` : ''}</span>
        </div>
    `).join('');
}

// Make toggleScanDetails globally accessible
window.toggleScanDetails = toggleScanDetails;
