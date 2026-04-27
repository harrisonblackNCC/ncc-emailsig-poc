/* NCC Email Signature Client Add-in - taskpane.js (NAA / nested auth)
 *
 * Profile load flow:
 *   1. MSAL.js initialises with supportsNestedAppAuth: true. Inside
 *      Office, it picks up the SSO token from the host automatically.
 *   2. msal.acquireTokenSilent({ scopes: ['User.Read'] }) — MSAL talks
 *      to Microsoft's auth backend, which performs the on-behalf-of
 *      swap and hands back a Graph token. No backend on our side.
 *   3. We call Graph /me directly with that Graph token to pull
 *      jobTitle (plus displayName + mail as a bonus).
 *   4. If anything fails: fall back to Office.context.mailbox.userProfile
 *      (name + email still work, but no auto jobTitle).
 *
 * Why NAA over a backend:
 *   NAA (Nested App Authentication) is Microsoft's 2024 pattern that
 *   lets Office add-ins do OBO without standing up a server. The SSO
 *   token can't be sent to Graph directly because of audience mismatch,
 *   but MSAL inside Office does the swap server-side via Microsoft's
 *   own auth servers. Trade-off: we need MSAL.js loaded in the page,
 *   and the Outlook host has to be a recent version that supports NAA
 *   (OWA / New Outlook on Windows / Outlook Mobile — Classic Outlook
 *   for Windows does NOT support NAA, so it falls back to mailbox-only).
 */

const LOGO_URL  = "https://www.ncc.qld.edu.au/wp-content/uploads/NCC-Email_600x200.jpg";
const LOGO_TARGET_WIDTH = 430;
const SSO_TIMEOUT_MS = 6000;   // give up on the whole NAA + Graph chain after 6s

// Entra app registration for the add-in's API.
const CLIENT_ID  = "55e5528d-7efd-4bd5-a437-0d31c68d3542";
// Authority targets the NCC tenant. Using the verified domain is
// equivalent to the tenant GUID and keeps this code free of magic IDs.
const AUTHORITY  = "https://login.microsoftonline.com/nambourcc.onmicrosoft.com";

let userProfile = { displayName: "", jobTitle: "", mail: "" };
let logoAspect  = 600 / 200;

/* ── Detect logo aspect ratio for the signature image ───────────────── */
(function detectLogoAspect() {
  try {
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => {
      if (img.naturalWidth && img.naturalHeight) {
        logoAspect = img.naturalWidth / img.naturalHeight;
      }
    };
    img.src = LOGO_URL;
  } catch (e) { /* fall back to default */ }
})();

/* ── Office init ────────────────────────────────────────────────────── */
Office.onReady(async () => {
  await loadProfile();
  loadPreferences();
  wireEvents();
  document.getElementById("loading").style.display = "none";
  document.getElementById("main").style.display    = "block";
});

/* ── Try SSO first, fall back to mailbox.userProfile if anything goes wrong ── */
async function loadProfile() {
  const mailboxFallback = () => {
    const p = Office.context.mailbox.userProfile;
    userProfile.displayName = p.displayName  || "";
    userProfile.mail        = p.emailAddress || "";
    userProfile.jobTitle    = Office.context.roamingSettings.get("jobTitle") || "";
  };

  try {
    const me = await withTimeout(loadProfileViaNAA(), SSO_TIMEOUT_MS);
    userProfile.displayName = me.displayName || "";
    userProfile.mail        = me.mail || me.userPrincipalName || "";
    userProfile.jobTitle    = me.jobTitle
                              || Office.context.roamingSettings.get("jobTitle")
                              || "";

    // Cache the fresh Entra jobTitle so launchevent.js has it next time
    // SSO is slow or unavailable. Best-effort, don't block render on it.
    if (me.jobTitle) {
      try {
        const s = Office.context.roamingSettings;
        s.set("jobTitle", me.jobTitle);
        s.saveAsync(() => {});
      } catch (e) { /* ignore */ }
    }
  } catch (err) {
    console.warn("NAA/Graph chain failed, falling back to mailbox.userProfile:", err);
    mailboxFallback();
  }

  renderProfile();
}

/* Lazy-init MSAL with Nested App Authentication. The
 * supportsNestedAppAuth flag tells MSAL to look for an Office host
 * and use its SSO token as the assertion for OBO, rather than
 * starting an interactive flow.
 */
let _msalInstance = null;
async function getMsalInstance() {
  if (_msalInstance) return _msalInstance;
  if (typeof msal === "undefined" || !msal.PublicClientNext) {
    throw new Error("MSAL Browser not loaded — check the script tag in taskpane.html");
  }
  _msalInstance = await msal.PublicClientNext.createPublicClientApplication({
    auth: {
      clientId:               CLIENT_ID,
      authority:              AUTHORITY,
      supportsNestedAppAuth:  true
    },
    cache: {
      cacheLocation: "localStorage"   // localStorage is recommended for NAA
    },
    system: {
      loggerOptions: {
        logLevel: msal.LogLevel.Warning,
        loggerCallback: (level, message) => {
          if (level <= msal.LogLevel.Warning) console.log("[MSAL]", message);
        }
      }
    }
  });
  return _msalInstance;
}

/* Acquire a Graph token via NAA, then call /me. Throws on any failure
 * so the caller falls back to mailbox.userProfile.
 */
async function loadProfileViaNAA() {
  const pca = await getMsalInstance();

  const tokenResult = await pca.acquireTokenSilent({
    scopes: ["User.Read"]
  });

  if (!tokenResult || !tokenResult.accessToken) {
    throw new Error("MSAL acquireTokenSilent returned no access token");
  }

  const res = await fetch(
    "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle",
    { headers: { Authorization: "Bearer " + tokenResult.accessToken } }
  );
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Graph /me returned ${res.status}: ${text}`);
  }
  return res.json();
}

function renderProfile() {
  document.getElementById("display-name").textContent = userProfile.displayName || "—";
  document.getElementById("email").textContent        = userProfile.mail        || "—";
  document.getElementById("job-title").textContent    = userProfile.jobTitle    || "—";
}

function withTimeout(promise, ms) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error("SSO timeout")), ms))
  ]);
}

/* ── Load saved preferences ─────────────────────────────────────────── */
function loadPreferences() {
  const settings = Office.context.roamingSettings;

  document.getElementById("title-select").value = settings.get("title") || "";

  applySignoffSelection("newSignoff",
    settings.get("newSignoff") || settings.get("signoff") || "Kind regards");
  applySignoffSelection("replySignoff",
    settings.get("replySignoff") || "Thanks");

  document.getElementById("ext-input").value   = settings.get("ext")   || "";
  document.getElementById("phone-input").value = settings.get("phone") || "";
}

function applySignoffSelection(prefix, value) {
  const sel = document.getElementById(prefix + "-select");
  const opt = [...sel.options].find(o => o.value === value);
  const wrap = document.getElementById(prefix === "newSignoff" ? "newCustom-wrap" : "replyCustom-wrap");

  if (opt) {
    sel.value = value;
    wrap.style.display = "none";
  } else {
    sel.value = "custom";
    document.getElementById(prefix + "-custom").value = value;
    wrap.style.display = "block";
  }
}

/* ── Wire events ────────────────────────────────────────────────────── */
function wireEvents() {
  ["newSignoff", "replySignoff"].forEach(prefix => {
    const sel = document.getElementById(prefix + "-select");
    sel.addEventListener("change", () => {
      const wrap = document.getElementById(prefix === "newSignoff" ? "newCustom-wrap" : "replyCustom-wrap");
      wrap.style.display = sel.value === "custom" ? "block" : "none";
    });
  });

  document.getElementById("btn-action").addEventListener("click", saveAndInsert);
}

function resolveSignoff(prefix, fallback) {
  const sel = document.getElementById(prefix + "-select");
  if (sel.value === "custom") {
    return document.getElementById(prefix + "-custom").value.trim() || fallback;
  }
  return sel.value;
}

/* ── Save preferences ───────────────────────────────────────────────── */
function savePreferences() {
  return new Promise((resolve) => {
    const settings = Office.context.roamingSettings;

    const title        = document.getElementById("title-select").value;
    const newSignoff   = resolveSignoff("newSignoff",   "Kind regards");
    const replySignoff = resolveSignoff("replySignoff", "Thanks");
    const ext          = document.getElementById("ext-input").value.trim();
    const phone        = document.getElementById("phone-input").value.trim();

    settings.set("title",        title);
    settings.set("newSignoff",   newSignoff);
    settings.set("replySignoff", replySignoff);
    settings.set("ext",          ext);
    settings.set("phone",        phone);

    // jobTitle is auto-pulled from Entra only — not user-editable. The
    // cached copy in roamingSettings is maintained by launchevent.js and
    // the loadProfile() SSO call so it stays fresh.

    settings.saveAsync((result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded);
    });
  });
}

/* ── Save + insert (single action) ──────────────────────────────────── */
async function saveAndInsert() {
  const btn    = document.getElementById("btn-action");
  const status = document.getElementById("action-status");

  btn.disabled = true;
  status.textContent = "";
  status.className   = "status";

  const saved = await savePreferences();
  if (!saved) {
    btn.disabled = false;
    status.textContent = "Save failed — try again";
    status.className   = "status error";
    return;
  }

  insertSignature(() => { btn.disabled = false; });
}

/* ── Reply vs new message ───────────────────────────────────────────── */
function isReplyContext() {
  try {
    const item = Office.context.mailbox.item;
    if (item.conversationId) return true;
    if (item.subject && typeof item.subject === "string") {
      if (/^(re|fw|fwd)\s*:/i.test(item.subject.trim())) return true;
    }
  } catch (e) { /* ignore */ }
  return false;
}

/* ── Build HTML signature ───────────────────────────────────────────── */
function buildSignature() {
  const settings = Office.context.roamingSettings;

  const title        = document.getElementById("title-select").value.trim() || settings.get("title") || "";
  const jobTitle     = userProfile.jobTitle || settings.get("jobTitle") || "";
  const newSignoff   = resolveSignoff("newSignoff",   settings.get("newSignoff")   || "Kind regards");
  const replySignoff = resolveSignoff("replySignoff", settings.get("replySignoff") || "Thanks");
  const ext          = document.getElementById("ext-input").value.trim()   || settings.get("ext")   || "";
  const phone        = document.getElementById("phone-input").value.trim() || settings.get("phone") || "";

  const signoff = isReplyContext() ? replySignoff : newSignoff;
  const { displayName, mail } = userProfile;
  const fullName = title ? `${title} ${displayName}` : displayName;

  const extLine = ext
    ? ` <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">| Ext: </span></strong><strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#000000;">${ext}</span></strong>`
    : "";

  const telHref = phone ? phone.replace(/[^+\d]/g, "") : "";
  const phoneLine = phone
    ? ` <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">| P: </span></strong><strong><u><a href="tel:${telHref}" style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#000000;text-decoration:underline;">${phone}</a></u></strong>`
    : "";

  // Role paragraph is conditional — skip the line entirely if blank so we
  // don't render an awkward empty paragraph. Kept in sync with launchevent.js.
  const rolePara = jobTitle
    ? `<p style="margin:0pt;line-height:normal;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt;color:#005953;">${jobTitle}</span></strong>
</p>`
    : "";

  // signoff "" = staff picked "None" — skip the sign-off paragraph
  // entirely so we don't render a lone comma above the name.
  const signoffPara = signoff
    ? `<p style="margin:0pt;line-height:normal;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt;">${signoff},</span></strong>
</p>`
    : "";

  return `
${signoffPara}
<p style="margin:0pt;margin-bottom:10pt;line-height:normal;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt;color:#ec3426;">${fullName}</span></strong>
</p>
${rolePara}
<p style="margin:0pt;line-height:normal;font-size:9pt;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">E: </span></strong><strong><u><a href="mailto:${mail}" style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#000000;text-decoration:underline;">${mail}</a></u></strong>${extLine}${phoneLine}
</p>
<p style="margin:0pt;margin-top:4pt;line-height:normal;font-size:9pt;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">Nambour Christian College</span></strong>
</p>
<p style="margin:0pt;line-height:normal;font-size:7pt;background-color:#ffffff;">
  <span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#333333;">2 McKenzie Road, Woombye QLD 4559 | PO Box 500, Nambour QLD 4560</span>
</p>
<p style="margin:0pt;margin-bottom:8pt;line-height:normal;font-size:7pt;background-color:#ffffff;">
  <a href="tel:+61754513333" style="text-decoration:underline;color:#ec3426;"><strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:7pt;color:#ec3426;">(07) 5451 3333</span></strong></a>
  <span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:7pt;color:#333333;"> | </span>
  <a href="mailto:info@ncc.qld.edu.au" style="text-decoration:underline;color:#ec3426;"><strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:7pt;color:#ec3426;">info@ncc.qld.edu.au</span></strong></a>
  <span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:7pt;color:#333333;"> | </span>
  <a href="https://www.ncc.qld.edu.au" style="text-decoration:underline;color:#ec3426;"><strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:7pt;color:#ec3426;">www.ncc.qld.edu.au</span></strong></a>
</p>
<p style="margin:0pt;line-height:normal;font-size:11pt;background-color:#ffffff;">
  <img src="${LOGO_URL}" width="${LOGO_TARGET_WIDTH}" height="${Math.round(LOGO_TARGET_WIDTH / logoAspect)}" alt="Nambour Christian College" style="display:block;border:0;">
</p>
<p style="margin:0pt;line-height:normal;font-size:7pt;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">CRICOS:</span></strong>
  <span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#333333;"> 01461G</span>
</p>
`.trim();
}

/* ── Insert ─────────────────────────────────────────────────────────── */
function insertSignature(done) {
  const status = document.getElementById("action-status");
  const sig    = buildSignature();
  const finish = () => { if (typeof done === "function") done(); };

  Office.context.mailbox.item.body.setSignatureAsync(
    sig,
    { coercionType: Office.CoercionType.Html },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        status.textContent = "\u2713 Saved & inserted";
        status.className   = "status";
        setTimeout(() => { status.textContent = ""; }, 2500);
        finish();
      } else {
        Office.context.mailbox.item.body.prependAsync(
          `<br><br>${sig}`,
          { coercionType: Office.CoercionType.Html },
          (r2) => {
            if (r2.status === Office.AsyncResultStatus.Succeeded) {
              status.textContent = "\u2713 Saved & inserted";
              status.className   = "status";
            } else {
              status.textContent = "Insert failed — try again";
              status.className   = "status error";
            }
            finish();
          }
        );
      }
    }
  );
}
