/* NCC Email Signature Add-in - taskpane.js */

const CLIENT_ID = "YOUR_AZURE_CLIENT_ID"; // Replace after Azure app registration
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
  const signoff  = settings.get("signoff") || "Kind regards";
  const ext      = settings.get("ext")     || "";

  const sel = document.getElementById("signoff-select");
  const opt = [...sel.options].find(o => o.value === signoff);

  if (opt) {
    sel.value = signoff;
  } else {
    // It's a custom value saved previously
    sel.value = "custom";
    document.getElementById("signoff-custom").value = signoff;
    document.getElementById("custom-wrap").style.display = "block";
  }

  document.getElementById("ext-input").value = ext;
}

/* ── Wire up UI events ──────────────────────────────────────────────── */
function wireEvents() {
  const sel = document.getElementById("signoff-select");
  sel.addEventListener("change", () => {
    document.getElementById("custom-wrap").style.display =
      sel.value === "custom" ? "block" : "none";
  });

  document.getElementById("btn-save").addEventListener("click", savePreferences);
  document.getElementById("btn-insert").addEventListener("click", insertSignature);
}

/* ── Save preferences to RoamingSettings ────────────────────────────── */
function savePreferences() {
  const btn      = document.getElementById("btn-save");
  const status   = document.getElementById("save-status");
  const settings = Office.context.roamingSettings;

  const sel     = document.getElementById("signoff-select");
  const signoff = sel.value === "custom"
    ? document.getElementById("signoff-custom").value.trim() || "Kind regards"
    : sel.value;
  const ext = document.getElementById("ext-input").value.trim();

  settings.set("signoff", signoff);
  settings.set("ext",     ext);

  btn.disabled = true;
  status.textContent = "";

  settings.saveAsync((result) => {
    btn.disabled = false;
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      status.textContent = "✓ Saved";
      status.className   = "status";
      setTimeout(() => { status.textContent = ""; }, 2500);
    } else {
      status.textContent = "Save failed — try again";
      status.className   = "status error";
    }
  });
}

/* ── Build the HTML signature string ────────────────────────────────── */
function buildSignature() {
  const settings = Office.context.roamingSettings;
  const sel      = document.getElementById("signoff-select");
  const signoff  = sel.value === "custom"
    ? document.getElementById("signoff-custom").value.trim() || "Kind regards"
    : sel.value;
  const ext      = document.getElementById("ext-input").value.trim()
                 || settings.get("ext") || "";

  const { displayName, jobTitle, mail } = userProfile;

  const extLine = ext
    ? `<strong><span style="font-family:Helvetica,Arial,sans-serif;color:#005953;">| Ext: </span></strong><strong><span style="font-family:Helvetica,Arial,sans-serif;color:#000000;">${ext}</span></strong>`
    : "";

  return `
<p style="margin-top:0pt;margin-bottom:0pt;line-height:normal;background-color:#ffffff;">
  <strong><span style="font-family:Helvetica,Arial,sans-serif;font-size:12pt;">${signoff},</span></strong>
</p>
<p style="margin-top:0pt;margin-bottom:12pt;line-height:normal;background-color:#ffffff;">
  <strong><span style="font-family:Helvetica,Arial,sans-serif;font-size:12pt;color:#ec3426;">${displayName}</span></strong>
</p>
<p style="margin-top:0pt;margin-bottom:0pt;line-height:normal;background-color:#ffffff;">
  <strong><span style="font-family:Helvetica,Arial,sans-serif;font-size:12pt;color:#005953;">${jobTitle}</span></strong>
</p>
<p style="margin-top:0pt;margin-bottom:0pt;line-height:normal;font-size:10pt;background-color:#ffffff;">
  <strong><span style="font-family:Helvetica,Arial,sans-serif;color:#005953;">E: </span></strong><strong><u><a href="mailto:${mail}" style="font-family:Helvetica,Arial,sans-serif;color:#000000;text-decoration:underline;">${mail}</a></u></strong>
  <strong><span style="font-family:Helvetica,Arial,sans-serif;color:#005953;"> ${extLine}</span></strong>
</p>
<p style="margin-top:0pt;margin-bottom:8pt;">
  <br>
  <strong><span style="font-family:Helvetica,Arial,sans-serif;font-size:10pt;color:#005953;">Nambour Christian College</span></strong><br>
  <span style="font-family:Helvetica,Arial,sans-serif;font-size:8pt;color:#333333;">2 McKenzie Road, Woombye QLD 4559 | PO Box 500, Nambour QLD 4560</span><br>
  <a href="tel:+61754513333" style="text-decoration:underline;color:#ec3426;"><strong><span style="font-family:Helvetica,Arial,sans-serif;font-size:8pt;color:#ec3426;">(07) 5451 3333 </span></strong></a><span style="font-family:Helvetica,Arial,sans-serif;font-size:8pt;">|</span>
  <a href="mailto:info@ncc.qld.edu.au" style="text-decoration:underline;color:#ec3426;"><strong><span style="font-family:Helvetica,Arial,sans-serif;font-size:8pt;color:#ec3426;"> info@ncc.qld.edu.au</span></strong></a>
  <span style="font-family:Helvetica,Arial,sans-serif;font-size:8pt;">|</span><a href="https://www.ncc.qld.edu.au" style="text-decoration:underline;color:#ec3426;"><strong><span style="font-family:Helvetica,Arial,sans-serif;font-size:8pt;color:#ec3426;"> www.ncc.qld.edu.au</span></strong></a>
</p>
<p style="margin-top:0pt;margin-bottom:0pt;line-height:normal;font-size:12pt;background-color:#ffffff;">
  <img src="${LOGO_URL}" width="430" height="143" alt="Nambour Christian College" style="display:block;border:0;">
</p>
<p style="margin-top:0pt;margin-bottom:0pt;line-height:normal;font-size:8pt;background-color:#ffffff;">
  <strong><span style="font-family:Helvetica,Arial,sans-serif;color:#005953;">CRICOS:</span></strong>
  <span style="font-family:Helvetica,Arial,sans-serif;color:#333333;"> 01461G</span>
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
        status.textContent = "✓ Signature inserted";
        status.className   = "status";
        setTimeout(() => { status.textContent = ""; }, 2500);
      } else {
        // Fallback: prepend to body if setSignatureAsync not available
        Office.context.mailbox.item.body.prependAsync(
          `<br><br>${sig}`,
          { coercionType: Office.CoercionType.Html },
          (r2) => {
            if (r2.status === Office.AsyncResultStatus.Succeeded) {
              status.textContent = "✓ Signature inserted";
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
