/* NCC Email Signature Client - taskpane.js
   GRAPH via MSAL (NO Office SSO)
*/

const LOGO_URL = "https://www.ncc.qld.edu.au/wp-content/uploads/NCC-Email_600x200.jpg";
let userProfile = { displayName: "", jobTitle: "", mail: "" };

/* ── MSAL CONFIG ─────────────────────────────────────────── */
const msalConfig = {
  auth: {
    clientId: "55e5528d-7efd-4bd5-a437-0d31c68d3542",
    authority: "https://login.microsoftonline.com/YOUR_TENANT_ID",
    redirectUri: "https://localhost"
  },
  cache: {
    cacheLocation: "localStorage"
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

/* ── Office init ─────────────────────────────────────────── */
Office.onReady(async () => {
  await loadProfile();
  loadPreferences();
  wireEvents();
  document.getElementById("loading").style.display = "none";
  document.getElementById("main").style.display = "block";
});

/* ── Load profile via Graph popup auth ───────────────────── */
async function loadProfile() {
  try {
    const account =
      msalInstance.getAllAccounts()[0] ||
      (await msalInstance.loginPopup({ scopes: ["User.Read"] })).account;

    const tokenResult = await msalInstance.acquireTokenSilent({
      account,
      scopes: ["User.Read"]
    });

    const me = await fetch("https://graph.microsoft.com/v1.0/me", {
      headers: { Authorization: "Bearer " + tokenResult.accessToken }
    }).then(r => r.json());

    userProfile.displayName = me.displayName || "";
    userProfile.mail = me.mail || me.userPrincipalName || "";
    userProfile.jobTitle = me.jobTitle || "";
  } catch (err) {
    console.warn("Graph login failed, falling back to mailbox profile", err);
    fallbackMailboxProfile();
  }

  renderProfile();
}

/* ── Fallback ────────────────────────────────────────────── */
function fallbackMailboxProfile() {
  const p = Office.context.mailbox.userProfile;
  userProfile.displayName = p.displayName || "";
  userProfile.mail = p.emailAddress || "";
  userProfile.jobTitle =
    Office.context.roamingSettings.get("jobTitle") || "";
}

/* ── Render profile ──────────────────────────────────────── */
function renderProfile() {
  document.getElementById("display-name").textContent =
    userProfile.displayName || "—";
  document.getElementById("email").textContent =
    userProfile.mail || "—";
  document.getElementById("job-title").textContent =
    userProfile.jobTitle || "—";
}

/* ── Preferences (unchanged) ─────────────────────────────── */
function loadPreferences() {
  const s = Office.context.roamingSettings;
  document.getElementById("title-select").value = s.get("title") || "";
  document.getElementById("ext-input").value = s.get("ext") || "";
  document.getElementById("phone-input").value = s.get("phone") || "";
}

function wireEvents() {
  document.getElementById("btn-action").addEventListener("click", insertSignature);
}

/* ── Insert signature (unchanged logic) ──────────────────── */
function buildSignature() {
  const title = document.getElementById("title-select").value.trim();
  const fullName = title
    ? `${title} ${userProfile.displayName}`
    : userProfile.displayName;

  return `
    <p>Kind regards,<br>
    <strong>${fullName}</strong><br>
    ${userProfile.jobTitle || ""}<br>
    ${userProfile.mail}</p>
  `;
}

function insertSignature() {
  const sig = buildSignature();
  Office.context.mailbox.item.body.setSignatureAsync(sig);
}
