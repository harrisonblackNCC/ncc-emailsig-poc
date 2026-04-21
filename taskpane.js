/* NCC Email Signature Add-in - taskpane.js */

const CLIENT_ID = "55e5528d-7efd-4bd5-a437-0d31c68d3542";
const LOGO_URL  = "https://www.ncc.qld.edu.au/wp-content/uploads/NCC-Email_600x200.jpg";

let userProfile = { displayName: "", jobTitle: "", mail: "" };

/* ── Office initialisation ─────────────────────────────────────────── */
Office.onReady(async () => {
  await loadProfile();
  loadPreferences();
  wireEvents();
  document.getElementById("loading").style.display = "none";
  document.getElementById("main").style.display    = "block";
});

/* ── Load user profile from Microsoft Graph ─────────────────────────── */
async function loadProfile() {
  try {
    const token = await getAccessToken();
    const res   = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,mail", {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!res.ok) throw new Error("Graph request failed");

    const data = await res.json();
    userProfile.displayName = data.displayName || "";
    userProfile.jobTitle    = data.jobTitle    || "";
    userProfile.mail        = data.mail        || "";

    document.getElementById("display-name").textContent = userProfile.displayName;
    document.getElementById("job-title").textContent    = userProfile.jobTitle;
    document.getElementById("email").textContent        = userProfile.mail;

  } catch (err) {
    console.error("Profile load error:", err);
    document.getElementById("display-name").textContent = "Could not load — check sign-in";
    document.getElementById("job-title").textContent    = "";
    document.getElementById("email").textContent        = "";
  }
}

/* ── Get access token via Office SSO ─────────────────────────────────── */
async function getAccessToken() {
  return new Promise((resolve, reject) => {
    Office.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true }, (result) => {
      if (result.status === "succeeded") {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

/* ── Load saved preferences from RoamingSettings ────────────────────── */
function loadPreferences() {
  const settings = Office.context.roamingSettings;

  const title       = settings.get("title")       || "";
  const newSignoff  = settings.get("newSignoff")  || settings.get("signoff") || "Kind regards"; // back-compat
  const replySignoff= settings.get("replySignoff")|| "Thanks";
  const ext         = settings.get("ext")         || "";
  const phone       = settings.get("phone")       || "";

  document.getElementById("title-select").value = title;

  applySignoffSelection("newSignoff",   newSignoff);
  applySignoffSelection("replySignoff", replySignoff);

  document.getElementById("ext-input").value   = ext;
  document.getElementById("phone-input").value = phone;
}

function applySignoffSelection(prefix, value) {
  const sel = document.getElementById(prefix + "-select");
  const opt = [...sel.options].find(o => o.value === value);

  if (opt) {
    sel.value = value;
    document.getElementById(prefix === "newSignoff" ? "newCustom-wrap" : "replyCustom-wrap").style.display = "none";
  } else {
    sel.value = "custom";
    document.getElementById(prefix + "-custom").value = value;
    document.getElementById(prefix === "newSignoff" ? "newCustom-wrap" : "replyCustom-wrap").style.display = "block";
  }
}

/* ── Wire up UI events ──────────────────────────────────────────────── */
function wireEvents() {
  ["newSignoff", "replySignoff"].forEach(prefix => {
    const sel = document.getElementById(prefix + "-select");
    sel.addEventListener("change", () => {
      const wrap = document.getElementById(prefix === "newSignoff" ? "newCustom-wrap" : "replyCustom-wrap");
      wrap.style.display = sel.value === "custom" ? "block" : "none";
    });
  });

  document.getElementById("btn-save").addEventListener("click", savePreferences);
  document.getElementById("btn-insert").addEventListener("click", insertSignature);
}

/* ── Resolve a sign-off from the UI ─────────────────────────────────── */
function resolveSignoff(prefix, fallback) {
  const sel = document.getElementById(prefix + "-select");
  if (sel.value === "custom") {
    return document.getElementById(prefix + "-custom").value.trim() || fallback;
  }
  return sel.value;
}

/* ── Save preferences to RoamingSettings ────────────────────────────── */
function savePreferences() {
  const btn      = document.getElementById("btn-save");
  const status   = document.getElementById("save-status");
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

  btn.disabled = true;
  status.textContent = "";

  settings.saveAsync((result) => {
    btn.disabled = false;
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      status.textContent = "\u2713 Saved";
      status.className   = "status";
      setTimeout(() => { status.textContent = ""; }, 2500);
    } else {
      status.textContent = "Save failed — try again";
      status.className   = "status error";
    }
  });
}

/* ── Detect reply vs new message ────────────────────────────────────── */
function isReplyContext() {
  try {
    const item = Office.context.mailbox.item;
    // conversationId is present on reply/forward compose sessions
    if (item.conversationId) return true;
    // Subject-based fallback
    if (item.subject && typeof item.subject === "string") {
      if (/^(re|fw|fwd)\s*:/i.test(item.subject.trim())) return true;
    }
  } catch (e) { /* ignore, default to new */ }
  return false;
}

/* ── Build the HTML signature string ────────────────────────────────── */
function buildSignature() {
  const settings = Office.context.roamingSettings;

  const title        = document.getElementById("title-select").value.trim() || settings.get("title") || "";
  const newSignoff   = resolveSignoff("newSignoff",   settings.get("newSignoff")   || "Kind regards");
  const replySignoff = resolveSignoff("replySignoff", settings.get("replySignoff") || "Thanks");
  const ext          = document.getElementById("ext-input").value.trim()   || settings.get("ext")   || "";
  const phone        = document.getElementById("phone-input").value.trim() || settings.get("phone") || "";

  const signoff = isReplyContext() ? replySignoff : newSignoff;

  const { displayName, jobTitle, mail } = userProfile;
  const fullName = title ? `${title} ${displayName}` : displayName;

  const extLine = ext
    ? ` <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">| Ext: </span></strong><strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#000000;">${ext}</span></strong>`
    : "";

  const phoneLine = phone
    ? ` <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">| P: </span></strong><strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#000000;">${phone}</span></strong>`
    : "";

  return `
<p style="margin:0pt;line-height:normal;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt;">${signoff},</span></strong>
</p>
<p style="margin:0pt;margin-bottom:10pt;line-height:normal;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt;color:#ec3426;">${fullName}</span></strong>
</p>
<p style="margin:0pt;line-height:normal;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt;color:#005953;">${jobTitle}</span></strong>
</p>
<p style="margin:0pt;margin-top:4pt;line-height:normal;font-size:9pt;background-color:#ffffff;">
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
  <img src="${LOGO_URL}" width="430" height="143" alt="Nambour Christian College" style="display:block;border:0;">
</p>
<p style="margin:0pt;line-height:normal;font-size:7pt;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">CRICOS:</span></strong>
  <span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#333333;"> 01461G</span>
</p>
`.trim();
}

/* ── Insert signature into compose body ─────────────────────────────── */
function insertSignature() {
  const status = document.getElementById("insert-status");
  const sig    = buildSignature();

  Office.context.mailbox.item.body.setSignatureAsync(
    sig,
    { coercionType: Office.CoercionType.Html },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        status.textContent = "\u2713 Signature inserted";
        status.className   = "status";
        setTimeout(() => { status.textContent = ""; }, 2500);
      } else {
        // Fallback: prepend to body if setSignatureAsync not available
        Office.context.mailbox.item.body.prependAsync(
          `<br><br>${sig}`,
          { coercionType: Office.CoercionType.Html },
          (r2) => {
            if (r2.status === Office.AsyncResultStatus.Succeeded) {
              status.textContent = "\u2713 Signature inserted";
              status.className   = "status";
            } else {
              status.textContent = "Insert failed — try again";
              status.className   = "status error";
            }
          }
        );
      }
    }
  );
}
