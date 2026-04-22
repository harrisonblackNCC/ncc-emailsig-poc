/* NCC Signature Add-in - launchevent.js
   Runs automatically when a new compose window opens.
   Auto-inserts the signature without any user action needed. */

const LOGO_URL = "https://www.ncc.qld.edu.au/wp-content/uploads/NCC-Email_600x200.jpg";

async function onMessageComposeOpened(event) {
  try {
    const token   = await getTokenSilent();
    const profile = await fetchProfile(token);
    const prefs   = loadPrefs();
    const sig     = buildSig(profile, prefs);

    Office.context.mailbox.item.body.setSignatureAsync(
      sig,
      { coercionType: Office.CoercionType.Html },
      () => event.completed()
    );
  } catch (err) {
    console.error("Auto-insert failed:", err);
    event.completed();
  }
}

async function getTokenSilent() {
  return new Promise((resolve, reject) => {
    Office.auth.getAccessToken(
      { allowSignInPrompt: false, allowConsentPrompt: false, forMSGraphAccess: true },
      (result) => {
        if (result.status === "succeeded") resolve(result.value);
        else reject(new Error(result.error?.message || "Token error"));
      }
    );
  });
}

async function fetchProfile(token) {
  const res = await fetch(
    "https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,mail",
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (!res.ok) throw new Error("Graph failed");
  return res.json();
}

function loadPrefs() {
  const s = Office.context.roamingSettings;
  return {
    signoff: s.get("signoff") || "Kind regards",
    ext:     s.get("ext")     || ""
  };
}

function buildSig(profile, prefs) {
  const name    = profile.displayName || "";
  const role    = profile.jobTitle    || "";
  const email   = profile.mail        || "";
  const signoff = prefs.signoff;
  const extLine = prefs.ext
    ? `<strong><span style="font-family:Helvetica,Arial,sans-serif;color:#005953;">| Ext: </span></strong><strong><span style="font-family:Helvetica,Arial,sans-serif;color:#000000;">${prefs.ext}</span></strong>`
    : "";

  return `
<p style="margin-top:0pt;margin-bottom:0pt;line-height:normal;">
  <strong><span style="font-family:Helvetica,Arial,sans-serif;font-size:12pt;">${signoff},</span></strong>
</p>
<p style="margin-top:0pt;margin-bottom:12pt;line-height:normal;">
  <strong><span style="font-family:Helvetica,Arial,sans-serif;font-size:12pt;color:#ec3426;">${name}</span></strong>
</p>
<p style="margin-top:0pt;margin-bottom:0pt;line-height:normal;">
  <strong><span style="font-family:Helvetica,Arial,sans-serif;font-size:12pt;color:#005953;">${role}</span></strong>
</p>
<p style="margin-top:0pt;margin-bottom:0pt;line-height:normal;font-size:10pt;">
  <strong><span style="font-family:Helvetica,Arial,sans-serif;color:#005953;">E: </span></strong><strong><u><a href="mailto:${email}" style="font-family:Helvetica,Arial,sans-serif;color:#000000;text-decoration:underline;">${email}</a></u></strong>
  ${extLine}
</p>
<p style="margin-top:0pt;margin-bottom:8pt;">
  <br>
  <strong><span style="font-family:Helvetica,Arial,sans-serif;font-size:10pt;color:#005953;">Nambour Christian College</span></strong><br>
  <span style="font-family:Helvetica,Arial,sans-serif;font-size:8pt;color:#333333;">2 McKenzie Road, Woombye QLD 4559 | PO Box 500, Nambour QLD 4560</span><br>
  <a href="tel:+61754513333" style="text-decoration:underline;color:#ec3426;"><strong><span style="font-family:Helvetica,Arial,sans-serif;font-size:8pt;color:#ec3426;">(07) 5451 3333 </span></strong></a><span style="font-family:Helvetica,Arial,sans-serif;font-size:8pt;">|</span>
  <a href="mailto:info@ncc.qld.edu.au" style="text-decoration:underline;color:#ec3426;"><strong><span style="font-family:Helvetica,Arial,sans-serif;font-size:8pt;color:#ec3426;"> info@ncc.qld.edu.au</span></strong></a>
  <span style="font-family:Helvetica,Arial,sans-serif;font-size:8pt;">|</span><a href="https://www.ncc.qld.edu.au" style="text-decoration:underline;color:#ec3426;"><strong><span style="font-family:Helvetica,Arial,sans-serif;font-size:8pt;color:#ec3426;"> www.ncc.qld.edu.au</span></strong></a>
</p>
<p style="margin-top:0pt;margin-bottom:0pt;line-height:normal;font-size:12pt;">
  <img src="${LOGO_URL}" width="430" height="143" alt="Nambour Christian College" style="display:block;border:0;">
</p>
<p style="margin-top:0pt;margin-bottom:0pt;line-height:normal;font-size:8pt;">
  <strong><span style="font-family:Helvetica,Arial,sans-serif;color:#005953;">CRICOS:</span></strong>
  <span style="font-family:Helvetica,Arial,sans-serif;color:#333333;"> 01461G</span>
</p>
`.trim();
}

// Register the handler
Office.actions.associate("onMessageComposeOpened", onMessageComposeOpened);
