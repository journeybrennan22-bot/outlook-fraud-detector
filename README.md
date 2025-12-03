# Email Fraud Detector - Outlook Web Add-in

A Microsoft 365 / Outlook Web add-in that detects email fraud, ported from the Gmail Chrome extension.

## Features

1. **Reply-To Mismatch Warning** - Alerts when reply address differs from sender
2. **Display Name Impersonation** - Detects trusted company names from untrusted domains
3. **First-Time Sender Flag** - Highlights new contacts with name + email
4. **Unicode/Homoglyph Detection** - Catches Cyrillic/Greek character spoofing
5. **Lookalike Domain Detection** - Levenshtein distance + typosquatting patterns
6. **Wire Fraud Keywords** - Special warning with verification instructions

## Architecture Difference: Chrome Extension vs Outlook Add-in

| Chrome Extension | Outlook Add-in |
|------------------|----------------|
| manifest.json | manifest.xml |
| content_scripts | Office.js API |
| Google People API | Microsoft Graph API |
| Runs in browser | Runs in Office sandbox |
| Chrome Web Store | Microsoft AppSource / Sideload |

---

## Deployment Options

### Option 1: Sideload for Personal Testing (Easiest)

Best for initial testing before wider deployment.

1. **Host the files** (pick one method):
   
   **Local development:**
   ```bash
   npm install
   npm start
   # Runs on http://localhost:3000
   ```
   
   **Or use GitHub Pages:**
   - Push to GitHub repo
   - Enable Pages in repo settings
   - Update manifest.xml URLs to `https://YOUR_USERNAME.github.io/REPO_NAME/`

2. **Update manifest.xml:**
   - Replace all `YOUR_DOMAIN` with your hosting URL
   - Generate a new GUID for the `<Id>` field: https://www.guidgenerator.com/

3. **Sideload in Outlook Web:**
   - Go to https://outlook.office.com
   - Click the **gear icon** → **View all Outlook settings**
   - Go to **Mail** → **Customize actions** → **Get add-ins**
   - Click **My add-ins** → **Add a custom add-in** → **Add from file**
   - Upload your `manifest.xml`

---

### Option 2: Centralized Deployment (Recommended for Production)

Deploy to all users in your Microsoft 365 organization.

#### Step 1: Register Azure AD App (for Microsoft Graph contacts)

1. Go to https://portal.azure.com
2. Navigate to **Azure Active Directory** → **App registrations** → **New registration**
3. Configure:
   - Name: `Email Fraud Detector`
   - Supported account types: **Accounts in this organizational directory only**
   - Redirect URI: **Single-page application (SPA)** → `https://YOUR_DOMAIN/src/taskpane.html`
4. After creation, note the **Application (client) ID**
5. Go to **API permissions** → **Add a permission**:
   - Microsoft Graph → Delegated permissions
   - Add: `Contacts.Read`, `User.Read`, `People.Read`
   - Click **Grant admin consent**

#### Step 2: Update Configuration

In `manifest.xml`, update:
```xml
<WebApplicationInfo>
  <Id>YOUR_APP_CLIENT_ID</Id>
  <Resource>api://YOUR_DOMAIN/YOUR_APP_CLIENT_ID</Resource>
</WebApplicationInfo>
```

In `taskpane.js`, update:
```javascript
msalConfig: {
    auth: {
        clientId: 'YOUR_APP_CLIENT_ID',
        redirectUri: 'https://YOUR_DOMAIN/src/taskpane.html'
    }
}
```

#### Step 3: Host Files on HTTPS

Options:
- **Azure Static Web Apps** (free tier available)
- **GitHub Pages** (free, easy)
- **Your own server** (must have valid SSL)

#### Step 4: Deploy via Microsoft 365 Admin Center

1. Go to https://admin.microsoft.com
2. Navigate to **Settings** → **Integrated apps** → **Upload custom apps**
3. Choose **Office Add-in**
4. Upload `manifest.xml`
5. Choose deployment scope:
   - **Just me** (testing)
   - **Entire organization**
   - **Specific users/groups**
6. Click **Deploy**

**Note for GoDaddy-managed M365:** Your admin.microsoft.com access should work. If you encounter permission issues, log into GoDaddy's Microsoft 365 dashboard and look for the admin center link there.

---

## Configuration

### Trusted Domains

Edit `taskpane.js` to add your trusted domains:

```javascript
trustedDomains: [
    'baynac.com',
    'purelogicescrow.com',
    'journeyinsurance.com',
    'firstam.com',
    'oldrepublictitle.com',
    // Add title companies, banks, etc.
],
```

### Company Keywords (Impersonation Detection)

```javascript
trustedCompanyKeywords: [
    'pure logic', 'purelogic', 
    'journey insurance',
    'first american', 'old republic',
    'chase', 'wells fargo',
    // Add companies you regularly work with
],
```

### Wire Fraud Keywords

```javascript
wireKeywords: [
    'wire transfer', 'wire instructions',
    'routing number', 'account number',
    'escrow', 'closing', 'settlement',
    // Already includes common escrow/title terms
],
```

---

## Files Structure

```
outlook-email-detector/
├── manifest.xml          # Add-in definition (entry point)
├── package.json          # Development dependencies
├── src/
│   ├── taskpane.html     # Main UI
│   ├── taskpane.css      # Styling
│   ├── taskpane.js       # Detection logic + Graph API
│   └── functions.html    # Background event handlers
└── assets/
    ├── icon-16.png       # Required icon sizes
    ├── icon-32.png
    ├── icon-64.png
    ├── icon-80.png
    └── icon-128.png
```

---

## Testing Checklist

- [ ] Add-in loads in Outlook Web
- [ ] "Fraud Detector" button appears in ribbon
- [ ] Basic email info displays correctly
- [ ] Reply-To mismatch detected (test with email that has different reply-to)
- [ ] Wire keywords detected (test with email containing "wire transfer")
- [ ] First-time sender flag works
- [ ] Contacts sync (check console for Graph API success)
- [ ] Lookalike domain detection (send test from similar domain)

---

## Troubleshooting

### "Add-in failed to load"
- Ensure all URLs in manifest.xml are HTTPS
- Check browser console for specific errors
- Verify manifest with: `npx office-addin-manifest validate manifest.xml`

### "Graph API permission denied"
- Ensure admin consent was granted in Azure AD
- Check that app registration has correct redirect URI
- User may need to sign out and back in

### Add-in not appearing
- Clear browser cache
- Try incognito/private window
- Re-sideload the manifest

### GoDaddy M365 Admin Issues
- Access admin center directly: https://admin.microsoft.com
- Sign in with journey@baynac.com
- If blocked, contact GoDaddy support to verify admin permissions

---

## Future Enhancements

- [ ] Auto-scan on email open (LaunchEvent)
- [ ] Notification badges in email list
- [ ] Sync with centralized threat database
- [ ] Report suspicious emails to admin
- [ ] Outlook desktop support (Windows/Mac)

---

## Support

For issues or feature requests, contact: journey@baynac.com
