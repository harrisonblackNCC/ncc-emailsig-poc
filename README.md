# NCC Email Signature Add-in

Outlook add-in that auto-inserts a branded NCC email signature, pulling each
staff member's name and role directly from their M365 profile.

---

## Project structure

```
ncc-signature-addin/
├── src/
│   ├── taskpane.html     ← Sidebar UI (sign-off picker, ext, insert button)
│   ├── taskpane.js       ← Graph API, RoamingSettings, signature builder
│   ├── launchevent.js    ← Auto-inserts signature on compose open
│   └── commands.html     ← Required Office JS stub
└── manifest/
    └── manifest.xml      ← Add-in manifest (sideload this into Outlook)
```

---

## Setup steps

### 1. Azure AD app registration

1. Go to https://portal.azure.com and sign in with harrison.black@ncc.qld.edu.au
2. Search for **App registrations** → **New registration**
3. Name: `NCC Signature Add-in POC`
4. Supported account types: **Accounts in this organisational directory only**
5. Click **Register**
6. Copy the **Application (client) ID** — you'll need it below

**Set API permissions:**
- Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**
- Add: `User.Read`, `openid`, `profile`, `offline_access`
- Click **Grant admin consent** (or ask IT to do this step)

**Expose an API:**
- Go to **Expose an API** → **Add a scope**
- Application ID URI: `api://YOUR-HOSTED-URL/YOUR_CLIENT_ID`
- Scope name: `access_as_user`
- Who can consent: **Admins and users**
- Save

**Add a client application:**
- Still in Expose an API → **Add a client application**
- Add these IDs (Outlook clients):
  - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Outlook Desktop)
  - `bc59ab01-8403-45c6-8796-ac3ef710b3e4` (Outlook Web)

---

### 2. Host the files

Upload everything in `/src/` to a public HTTPS host. Options:

- **GitHub Pages** (free, easiest): push to a repo, enable Pages
- **Azure Static Web Apps** (free tier): deploy from GitHub
- **Cloudflare Pages** (you already use this!)

---

### 3. Update placeholders

In both `manifest.xml` and `taskpane.js`, replace:
- `YOUR_AZURE_CLIENT_ID` → your actual client ID from step 1
- `YOUR-HOSTED-URL` → your actual hosting URL (e.g. `ncc-sig.pages.dev`)

---

### 4. Sideload into Outlook (POC testing)

**Outlook on the web:**
1. Go to https://outlook.office.com
2. Open a new email compose window
3. Click **...** (more options) → **Get Add-ins**
4. Go to **My add-ins** → **Add a custom add-in** → **Add from file**
5. Upload `manifest/manifest.xml`
6. Close and open a new compose — the signature should auto-insert!

**Outlook Desktop:**
1. Open Outlook → **File** → **Manage Add-ins** (opens browser)
2. Same steps as above

---

### 5. Org-wide deployment (when ready)

Have IT/admin go to:
**Microsoft 365 Admin Centre** → **Settings** → **Integrated Apps** → **Upload custom app**

Upload `manifest.xml` and assign to all users or specific groups.

---

## What it does

- **Auto-inserts** signature when compose window opens (`OnMessageComposeOpened`)
- **Pulls** display name, job title, and email from M365 profile via Graph API
- **Lets staff customise** sign-off phrase and extension via the taskpane sidebar
- **Saves preferences** to M365 `RoamingSettings` (follows them across devices)
- **Locked fields** (name, role, email) cannot be edited — always accurate
