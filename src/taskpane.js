/**
 * Email Fraud Detector - Outlook Web Add-in (v1.1.0)
 * Auto-scans with full panel background color based on threat level
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const CONFIG = {
    // Microsoft Graph API settings
    msalConfig: {
        auth: {
            clientId: '622f0452-d622-45d1-aab3-3a2026389dd3',
            authority: 'https://login.microsoftonline.com/common',
            redirectUri: 'https://journeybrennan22-bot.github.io/outlook-fraud-detector/src/taskpane.html'
        },
        cache: {
            cacheLocation: 'sessionStorage',
            storeAuthStateInCookie: false
        }
    },
    graphScopes: ['Contacts.Read', 'User.Read'],
    
    // Trusted domains for your organization
    trustedDomains: [
        'baynac.com',
        'purelogicescrow.com',
        'journeyinsurance.com',
    ],
    
    // Company keywords to watch for in display names (impersonation detection)
    trustedCompanyKeywords: [
        'microsoft', 'office', 'outlook', 'azure',
        'google', 'gmail', 'apple', 'icloud',
        'amazon', 'aws', 'paypal', 'venmo',
        'chase', 'bank of america', 'wells fargo', 'citi',
        'fidelity', 'schwab', 'vanguard',
        'fedex', 'ups', 'usps',
        'irs', 'social security',
        'pure logic', 'purelogic', 'journey insurance',
    ],
    
    // Wire fraud keywords
    wireKeywords: [
        'wire transfer', 'wire instructions', 'wiring instructions',
        'bank transfer', 'routing number', 'account number',
        'ach transfer', 'swift code', 'iban',
        'updated bank', 'new bank', 'changed bank',
        'payment instructions', 'fund transfer',
        'escrow', 'closing', 'settlement',
        'urgent payment', 'immediate transfer'
    ],
    
    // Levenshtein threshold for lookalike detection
    lookalikeSimilarityThreshold: 0.85,
    
    // Common lookalike domain patterns
    commonLookalikeTLDs: ['.co', '.net', '.org', '.info', '.biz', '.xyz', '.online', '.site']
};

// ============================================================================
// STATE
// ============================================================================

let msalInstance = null;
let userContacts = [];
let knownSenders = new Set();
let contactsLoaded = false;
let currentItemId = null;

// ============================================================================
// INITIALIZATION
// ============================================================================

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        initializeApp();
    }
});

async function initializeApp() {
    // Initialize MSAL
    msalInstance = new msal.PublicClientApplication(CONFIG.msalConfig);
    
    // Set up event listeners
    document.getElementById('retry-btn')?.addEventListener('click', analyzeEmail);
    document.getElementById('rescan-btn')?.addEventListener('click', analyzeEmail);
    
    // Set up collapsible sections
    document.querySelectorAll('.collapsible-header').forEach(header => {
        header.addEventListener('click', () => {
            header.closest('.collapsible').classList.toggle('collapsed');
        });
    });
    
    // Load contacts once at startup
    await loadContactsOnce();
    
    // Auto-scan: Listen for email item changes
    try {
        Office.context.mailbox.addHandlerAsync(
            Office.EventType.ItemChanged,
            onItemChanged,
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log('ItemChanged handler registered successfully');
                    updateAutoScanStatus(true);
                } else {
                    console.log('Failed to register ItemChanged handler:', result.error);
                    updateAutoScanStatus(false);
                }
            }
        );
    } catch (e) {
        console.log('ItemChanged event not supported:', e);
        updateAutoScanStatus(false);
    }
    
    // Initial analysis
    await analyzeEmail();
}

function onItemChanged(eventArgs) {
    console.log('Email changed - auto-scanning...');
    currentItemId = null; // Reset to force re-scan
    setTimeout(() => {
        analyzeEmail();
    }, 100);
}

function updateAutoScanStatus(enabled) {
    const footer = document.querySelector('.footer');
    let statusEl = document.getElementById('auto-scan-status');
    
    if (!statusEl) {
        statusEl = document.createElement('p');
        statusEl.id = 'auto-scan-status';
        statusEl.style.fontSize = '11px';
        statusEl.style.marginTop = '4px';
        footer.insertBefore(statusEl, footer.querySelector('.version'));
    }
    
    if (enabled) {
        statusEl.innerHTML = '🔄 <span style="color: #107c10;">Auto-scan ON</span>';
    } else {
        statusEl.innerHTML = '⏸️ <span style="color: #8a8886;">Manual mode</span>';
    }
}

async function loadContactsOnce() {
    if (contactsLoaded) return;
    
    try {
        await fetchUserContacts();
        contactsLoaded = true;
        console.log('Contacts loaded:', knownSenders.size, 'known senders');
    } catch (e) {
        console.log('Contact loading deferred');
    }
}

// ============================================================================
// MICROSOFT GRAPH API - CONTACTS
// ============================================================================

async function getAccessToken() {
    try {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            const response = await msalInstance.acquireTokenSilent({
                scopes: CONFIG.graphScopes,
                account: accounts[0]
            });
            return response.accessToken;
        }
        
        const response = await msalInstance.acquireTokenPopup({
            scopes: CONFIG.graphScopes
        });
        return response.accessToken;
    } catch (error) {
        console.error('Token acquisition failed:', error);
        return null;
    }
}

async function fetchUserContacts() {
    try {
        const token = await getAccessToken();
        if (!token) {
            console.log('No token available, skipping contact sync');
            return [];
        }
        
        const response = await fetch('https://graph.microsoft.com/v1.0/me/contacts?$select=emailAddresses,displayName&$top=1000', {
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (!response.ok) {
            throw new Error(`Graph API error: ${response.status}`);
        }
        
        const data = await response.json();
        const contacts = [];
        
        data.value.forEach(contact => {
            if (contact.emailAddresses) {
                contact.emailAddresses.forEach(email => {
                    contacts.push({
                        email: email.address.toLowerCase(),
                        name: contact.displayName || ''
                    });
                    knownSenders.add(email.address.toLowerCase());
                });
            }
        });
        
        await fetchRecentPeople(token);
        
        return contacts;
    } catch (error) {
        console.error('Failed to fetch contacts:', error);
        return [];
    }
}

async function fetchRecentPeople(token) {
    try {
        const response = await fetch('https://graph.microsoft.com/v1.0/me/people?$top=100', {
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            data.value.forEach(person => {
                if (person.scoredEmailAddresses) {
                    person.scoredEmailAddresses.forEach(email => {
                        knownSenders.add(email.address.toLowerCase());
                    });
                }
            });
        }
    } catch (error) {
        console.log('People API not available:', error);
    }
}

// ============================================================================
// EMAIL DATA EXTRACTION
// ============================================================================

async function getEmailData() {
    return new Promise((resolve, reject) => {
        const item = Office.context.mailbox.item;
        
        if (!item) {
            reject(new Error('No email item available'));
            return;
        }
        
        const emailData = {
            subject: item.subject || '',
            from: null,
            replyTo: null,
            body: '',
            itemId: item.itemId
        };
        
        if (item.from) {
            emailData.from = {
                displayName: item.from.displayName || '',
                emailAddress: item.from.emailAddress || ''
            };
        }
        
        if (item.replyTo && item.replyTo.length > 0) {
            emailData.replyTo = {
                displayName: item.replyTo[0].displayName || '',
                emailAddress: item.replyTo[0].emailAddress || ''
            };
        }
        
        item.body.getAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                emailData.body = result.value;
            }
            resolve(emailData);
        });
    });
}

// ============================================================================
// DETECTION LOGIC
// ============================================================================

function levenshteinDistance(str1, str2) {
    const m = str1.length;
    const n = str2.length;
    const dp = Array(m + 1).fill(null).map(() => Array(n + 1).fill(0));
    
    for (let i = 0; i <= m; i++) dp[i][0] = i;
    for (let j = 0; j <= n; j++) dp[0][j] = j;
    
    for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
            if (str1[i - 1] === str2[j - 1]) {
                dp[i][j] = dp[i - 1][j - 1];
            } else {
                dp[i][j] = 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
            }
        }
    }
    
    return dp[m][n];
}

function calculateSimilarity(str1, str2) {
    const distance = levenshteinDistance(str1.toLowerCase(), str2.toLowerCase());
    const maxLen = Math.max(str1.length, str2.length);
    return maxLen === 0 ? 1 : 1 - (distance / maxLen);
}

function extractDomain(email) {
    if (!email) return '';
    const parts = email.toLowerCase().split('@');
    return parts.length > 1 ? parts[1] : '';
}

function extractBaseDomain(domain) {
    return domain.replace(/\.(com|net|org|co|io|biz|info|xyz|online|site)$/i, '');
}

function detectHomoglyphs(text) {
    const homoglyphMap = {
        'а': 'a', 'е': 'e', 'і': 'i', 'о': 'o', 'р': 'p', 'с': 'c', 'у': 'y', 'х': 'x',
        'ɑ': 'a', 'ḃ': 'b', 'ċ': 'c', 'ḋ': 'd', 'ė': 'e', 'ḟ': 'f', 'ġ': 'g', 'ḣ': 'h',
        'і': 'i', 'ј': 'j', 'κ': 'k', 'ḷ': 'l', 'ṁ': 'm', 'ṅ': 'n', 'ο': 'o', 'ρ': 'p',
        'ԛ': 'q', 'ṙ': 'r', 'ṡ': 's', 'ṫ': 't', 'υ': 'u', 'ν': 'v', 'ẃ': 'w', 'х': 'x',
        'ỳ': 'y', 'ż': 'z', '0': 'o', '1': 'l', 'ⅰ': 'i', 'ⅼ': 'l', 'ℓ': 'l',
        'ɡ': 'g', 'ɩ': 'i', 'ɪ': 'i', 'ʝ': 'j', 'ĸ': 'k', 'ŀ': 'l', 'ɴ': 'n', 'ɵ': 'o'
    };
    
    const detected = [];
    
    for (const char of text) {
        const code = char.charCodeAt(0);
        if (homoglyphMap[char]) {
            detected.push({
                char: char,
                looksLike: homoglyphMap[char],
                code: code
            });
        } else if (
            (code >= 0x0400 && code <= 0x04FF) ||
            (code >= 0x0370 && code <= 0x03FF) ||
            (code >= 0x2100 && code <= 0x214F) ||
            (code >= 0xFF00 && code <= 0xFFEF)
        ) {
            detected.push({
                char: char,
                looksLike: '?',
                code: code
            });
        }
    }
    
    return detected;
}

function detectLookalikeDomain(senderDomain) {
    const results = [];
    const senderBase = extractBaseDomain(senderDomain);
    
    for (const trustedDomain of CONFIG.trustedDomains) {
        const trustedBase = extractBaseDomain(trustedDomain);
        
        if (senderDomain === trustedDomain) continue;
        
        const similarity = calculateSimilarity(senderBase, trustedBase);
        
        if (similarity >= CONFIG.lookalikeSimilarityThreshold && similarity < 1) {
            results.push({
                senderDomain: senderDomain,
                trustedDomain: trustedDomain,
                similarity: Math.round(similarity * 100)
            });
        }
        
        if (isTyposquatting(senderBase, trustedBase)) {
            results.push({
                senderDomain: senderDomain,
                trustedDomain: trustedDomain,
                similarity: 90,
                type: 'typosquatting'
            });
        }
    }
    
    return results;
}

function isTyposquatting(sender, trusted) {
    for (let i = 0; i < trusted.length - 1; i++) {
        const swapped = trusted.slice(0, i) + trusted[i + 1] + trusted[i] + trusted.slice(i + 2);
        if (sender === swapped) return true;
    }
    
    for (let i = 0; i < trusted.length; i++) {
        const missing = trusted.slice(0, i) + trusted.slice(i + 1);
        if (sender === missing) return true;
    }
    
    for (let i = 0; i < trusted.length; i++) {
        const doubled = trusted.slice(0, i + 1) + trusted[i] + trusted.slice(i + 1);
        if (sender === doubled) return true;
    }
    
    const commonReplacements = {
        'a': ['s', 'q', 'z'], 'b': ['v', 'n', 'g'], 'c': ['x', 'v', 'd'],
        'd': ['s', 'f', 'e'], 'e': ['w', 'r', 'd'], 'f': ['d', 'g', 'r'],
        'g': ['f', 'h', 't'], 'h': ['g', 'j', 'y'], 'i': ['u', 'o', 'k'],
        'j': ['h', 'k', 'u'], 'k': ['j', 'l', 'i'], 'l': ['k', 'o', 'p'],
        'm': ['n', 'j', 'k'], 'n': ['b', 'm', 'h'], 'o': ['i', 'p', 'l'],
        'p': ['o', 'l'], 'q': ['w', 'a'], 'r': ['e', 't', 'f'],
        's': ['a', 'd', 'w'], 't': ['r', 'y', 'g'], 'u': ['y', 'i', 'j'],
        'v': ['c', 'b', 'f'], 'w': ['q', 'e', 's'], 'x': ['z', 'c', 's'],
        'y': ['t', 'u', 'h'], 'z': ['a', 'x', 's']
    };
    
    for (let i = 0; i < trusted.length; i++) {
        const char = trusted[i].toLowerCase();
        if (commonReplacements[char]) {
            for (const replacement of commonReplacements[char]) {
                const replaced = trusted.slice(0, i) + replacement + trusted.slice(i + 1);
                if (sender === replaced) return true;
            }
        }
    }
    
    return false;
}

function detectDisplayNameImpersonation(displayName, senderDomain) {
    if (!displayName) return null;
    
    const lowerName = displayName.toLowerCase();
    const isTrustedDomain = CONFIG.trustedDomains.some(d => senderDomain.includes(d));
    
    if (isTrustedDomain) return null;
    
    for (const keyword of CONFIG.trustedCompanyKeywords) {
        if (lowerName.includes(keyword.toLowerCase())) {
            return {
                keyword: keyword,
                displayName: displayName,
                actualDomain: senderDomain
            };
        }
    }
    
    return null;
}

function detectWireKeywords(body, subject) {
    const text = `${subject} ${body}`.toLowerCase();
    const foundKeywords = [];
    
    for (const keyword of CONFIG.wireKeywords) {
        if (text.includes(keyword.toLowerCase())) {
            foundKeywords.push(keyword);
        }
    }
    
    return foundKeywords;
}

function isFirstTimeSender(email) {
    return !knownSenders.has(email.toLowerCase());
}

/**
 * Check if incoming email is a lookalike of any known contact
 * This catches things like journeybrennan22@gmail.com vs journeybrennan@gmail.com
 */
function detectContactLookalike(incomingEmail) {
    const incoming = parseEmailParts(incomingEmail);
    if (!incoming) return null;
    
    for (const knownEmail of knownSenders) {
        const known = parseEmailParts(knownEmail);
        if (!known) continue;
        
        // Skip if exact match (it's the actual contact)
        if (incoming.full === known.full) continue;
        
        // Calculate distances
        const localDistance = levenshteinDistance(incoming.local, known.local);
        const domainDistance = levenshteinDistance(incoming.domain, known.domain);
        
        // Rule 1: Same domain, username is 1-4 characters different
        if (incoming.domain === known.domain && localDistance > 0 && localDistance <= 4) {
            return {
                incomingEmail: incomingEmail,
                matchedContact: knownEmail,
                reason: `Username is ${localDistance} character${localDistance > 1 ? 's' : ''} different`,
                localDistance: localDistance,
                domainDistance: domainDistance
            };
        }
        
        // Rule 2: Domain is 1-2 characters different (very suspicious)
        if (domainDistance > 0 && domainDistance <= 2) {
            return {
                incomingEmail: incomingEmail,
                matchedContact: knownEmail,
                reason: `Domain is ${domainDistance} character${domainDistance > 1 ? 's' : ''} different`,
                localDistance: localDistance,
                domainDistance: domainDistance
            };
        }
        
        // Rule 3: Same domain name but different TLD (.com vs .co vs .net)
        const incomingBase = incoming.domain.split('.').slice(0, -1).join('.');
        const knownBase = known.domain.split('.').slice(0, -1).join('.');
        const incomingTLD = incoming.domain.split('.').pop();
        const knownTLD = known.domain.split('.').pop();
        
        if (incomingBase === knownBase && incomingTLD !== knownTLD && incoming.local === known.local) {
            return {
                incomingEmail: incomingEmail,
                matchedContact: knownEmail,
                reason: `Same name but different domain extension (.${incomingTLD} vs .${knownTLD})`,
                localDistance: localDistance,
                domainDistance: domainDistance
            };
        }
        
        // Rule 4: Same domain, high similarity (70%+) in local part
        if (incoming.domain === known.domain) {
            const maxLen = Math.max(incoming.local.length, known.local.length);
            const similarity = 1 - (localDistance / maxLen);
            if (similarity >= 0.7 && localDistance > 0 && localDistance > 4) {
                return {
                    incomingEmail: incomingEmail,
                    matchedContact: knownEmail,
                    reason: `Username is ${Math.round(similarity * 100)}% similar`,
                    localDistance: localDistance,
                    domainDistance: domainDistance
                };
            }
        }
    }
    
    return null;
}

/**
 * Parse email into local and domain parts
 */
function parseEmailParts(email) {
    const normalized = email.toLowerCase().trim();
    const parts = normalized.split('@');
    
    if (parts.length !== 2) return null;
    
    return {
        local: parts[0],
        domain: parts[1],
        full: normalized
    };
}

// ============================================================================
// MAIN ANALYSIS
// ============================================================================

async function analyzeEmail() {
    showLoading();
    
    try {
        const emailData = await getEmailData();
        
        if (emailData.itemId === currentItemId) {
            console.log('Same email, using cached results');
            return;
        }
        currentItemId = emailData.itemId;
        
        const warnings = [];
        const scanResults = [];
        
        if (!emailData.from) {
            throw new Error('Could not read email sender information');
        }
        
        const senderEmail = emailData.from.emailAddress.toLowerCase();
        const senderDomain = extractDomain(senderEmail);
        const displayName = emailData.from.displayName;
        
        // 1. Reply-To Mismatch (MEDIUM severity)
        if (emailData.replyTo && emailData.replyTo.emailAddress) {
            const replyToEmail = emailData.replyTo.emailAddress.toLowerCase();
            if (replyToEmail !== senderEmail) {
                warnings.push({
                    type: 'replyto-mismatch',
                    severity: 'medium',
                    title: 'Reply-To Mismatch',
                    description: 'Replies will go to a different address than the sender.',
                    detail: `From: ${senderEmail}\nReply-To: ${replyToEmail}`
                });
                scanResults.push({ check: 'Reply-To Match', status: 'fail' });
            } else {
                scanResults.push({ check: 'Reply-To Match', status: 'pass' });
            }
        } else {
            scanResults.push({ check: 'Reply-To Match', status: 'pass' });
        }
        
        // 2. Display Name Impersonation (CRITICAL)
        const impersonation = detectDisplayNameImpersonation(displayName, senderDomain);
        if (impersonation) {
            warnings.push({
                type: 'impersonation',
                severity: 'critical',
                title: 'Possible Impersonation',
                description: `Display name contains "${impersonation.keyword}" but email is from untrusted domain.`,
                detail: `Name: ${impersonation.displayName}\nDomain: ${impersonation.actualDomain}`
            });
            scanResults.push({ check: 'Display Name Check', status: 'fail' });
        } else {
            scanResults.push({ check: 'Display Name Check', status: 'pass' });
        }
        
        // 3. Unicode/Homoglyph Detection (CRITICAL)
        const homoglyphs = detectHomoglyphs(senderEmail);
        if (homoglyphs.length > 0) {
            warnings.push({
                type: 'homoglyph',
                severity: 'critical',
                title: 'Suspicious Characters Detected',
                description: 'Email address contains characters that look like normal letters but are actually different.',
                detail: homoglyphs.map(h => `"${h.char}" looks like "${h.looksLike}" (code: ${h.code})`).join('\n')
            });
            scanResults.push({ check: 'Character Analysis', status: 'fail' });
        } else {
            scanResults.push({ check: 'Character Analysis', status: 'pass' });
        }
        
        // 4. Lookalike Domain Detection (CRITICAL)
        const lookalikes = detectLookalikeDomain(senderDomain);
        if (lookalikes.length > 0) {
            const match = lookalikes[0];
            warnings.push({
                type: 'lookalike',
                severity: 'critical',
                title: 'Lookalike Domain Detected',
                description: `This domain looks similar to "${match.trustedDomain}" (${match.similarity}% match).`,
                detail: `Sender: ${match.senderDomain}\nLooks like: ${match.trustedDomain}`
            });
            scanResults.push({ check: 'Domain Similarity', status: 'fail' });
        } else {
            scanResults.push({ check: 'Domain Similarity', status: 'pass' });
        }
        
        // 5. Wire Fraud Keywords (CRITICAL)
        const wireKeywords = detectWireKeywords(emailData.body, emailData.subject);
        if (wireKeywords.length > 0) {
            warnings.push({
                type: 'wire-fraud',
                severity: 'critical',
                title: 'KEYWORD DETECTED',
                description: `This email contains terms commonly used in fraud (${wireKeywords.join(', ')}). If payment is requested, verify by phone using a number you search for online - never use a number from this email.`,
                detail: null,
                isWireFraud: true
            });
            scanResults.push({ check: 'Wire Fraud Keywords', status: 'fail' });
        } else {
            scanResults.push({ check: 'Wire Fraud Keywords', status: 'pass' });
        }
        
        // 6. Contact Lookalike Detection (CRITICAL) - compares against saved contacts
        const contactLookalike = detectContactLookalike(senderEmail);
        if (contactLookalike) {
            warnings.push({
                type: 'contact-lookalike',
                severity: 'critical',
                title: 'Similar to Known Contact',
                description: `This email address is suspiciously similar to someone in your contacts. ${contactLookalike.reason}.`,
                detail: `Sender: ${contactLookalike.incomingEmail}\nSimilar to: ${contactLookalike.matchedContact}`
            });
            scanResults.push({ check: 'Contact Match', status: 'fail' });
        } else {
            scanResults.push({ check: 'Contact Match', status: 'pass' });
        }
        
        // 7. First-Time Sender (INFO only)
        const firstTime = isFirstTimeSender(senderEmail);
        if (firstTime) {
            scanResults.push({ check: 'Known Sender', status: 'info', note: 'First-time sender' });
        } else {
            scanResults.push({ check: 'Known Sender', status: 'pass' });
        }
        
        // Display results
        displayResults(warnings, scanResults, firstTime);
        
    } catch (error) {
        console.error('Analysis error:', error);
        showError(error.message);
    }
}

// ============================================================================
// UI RENDERING
// ============================================================================

function showLoading() {
    document.getElementById('loading').classList.remove('hidden');
    document.getElementById('results').classList.add('hidden');
    document.getElementById('error').classList.add('hidden');
    
    // Reset body background
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
    
    // Determine overall status
    const criticalCount = warnings.filter(w => w.severity === 'critical').length;
    const mediumCount = warnings.filter(w => w.severity === 'medium').length;
    
    // Set body background based on status
    document.body.classList.remove('status-critical', 'status-medium', 'status-info', 'status-safe');
    
    const statusBadge = document.getElementById('status-badge');
    
    if (criticalCount > 0) {
        document.body.classList.add('status-critical');
        statusBadge.className = 'status-badge danger';
        statusBadge.querySelector('.status-icon').textContent = '🚨';
        statusBadge.querySelector('.status-text').textContent = `${criticalCount} Critical Warning${criticalCount > 1 ? 's' : ''}`;
    } else if (mediumCount > 0) {
        document.body.classList.add('status-medium');
        statusBadge.className = 'status-badge warning';
        statusBadge.querySelector('.status-icon').textContent = '⚠️';
        statusBadge.querySelector('.status-text').textContent = `${mediumCount} Warning${mediumCount > 1 ? 's' : ''}`;
    } else if (isFirstTime) {
        document.body.classList.add('status-info');
        statusBadge.className = 'status-badge info';
        statusBadge.querySelector('.status-icon').textContent = '👤';
        statusBadge.querySelector('.status-text').textContent = 'First-Time Sender';
    } else {
        document.body.classList.add('status-safe');
        statusBadge.className = 'status-badge safe';
        statusBadge.querySelector('.status-icon').textContent = '✅';
        statusBadge.querySelector('.status-text').textContent = 'No Issues Detected';
    }
    
    // Display warnings
    const warningsSection = document.getElementById('warnings-section');
    const warningsList = document.getElementById('warnings-list');
    
    if (warnings.length > 0) {
        warningsSection.classList.remove('hidden');
        warningsList.innerHTML = warnings.map(w => `
            <div class="warning-item ${w.severity}${w.isWireFraud ? ' wire-fraud' : ''}">
                <div class="warning-title">${w.title}</div>
                <div class="warning-description">${w.description}</div>
                ${w.detail ? `<div class="warning-detail">${w.detail}</div>` : ''}
            </div>
        `).join('');
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
