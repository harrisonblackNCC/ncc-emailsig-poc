/* NCC Email Signature Client - launchevent.js
 *
 * Event-based activation: runs automatically the moment a compose window
 * opens and drops the personalised NCC signature in — no taskpane, no
 * click, no UI.
 *
 * Data sourcing (in order of preference):
 *   1. Name + email     → Office.context.mailbox.userProfile (always works, sync)
 *   2. Job title        → SSO → Graph /me (silent, 3s timeout) → fallback to
 *                         RoamingSettings jobTitle → omit line entirely
 *   3. Title / ext /
 *      phone / signoffs → RoamingSettings (saved once via the taskpane)
 *
 * The baseline signature works for a brand-new user who's never opened
 * the taskpane: they get { default signoff + name + role (from SSO) +
 * email + school block + logo }. The role line is the only thing that
 * might be missing, and only if both SSO fails AND they have no jobTitle
 * set in Entra. Everything else always populates.
 *
 * The HTML signature here is kept in lockstep with buildSignature() in
 * taskpane.js so output is identical whether auto-inserted or manual.
 */

const LOGO_URL      = "https://www.ncc.qld.edu.au/wp-content/uploads/NCC-Email_600x200.jpg";
const LOGO_WIDTH    = 430;
const LOGO_HEIGHT   = 143; // 430 / (600/200) rounded
const SSO_TIMEOUT_MS = 3000; // give up on SSO after 3s, fall back silently

/* Register the handler so the manifest's FunctionName resolves. */
Office.onReady(() => {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
});

/* ── Main handler ────────────────────────────────────────────────────── */
async function onNewMessageComposeHandler(event) {
  try {
    const settings = Office.context.roamingSettings;
    const profile  = Office.context.mailbox.userProfile;
    const item     = Office.context.mailbox.item;

    const title        = settings.get("title")        || "";
    const newSignoff   = settings.get("newSignoff")   || settings.get("signoff") || "Kind regards";
    const replySignoff = settings.get("replySignoff") || "Thanks";
    const ext          = settings.get("ext")          || "";
    const phone        = settings.get("phone")        || "";

    const displayName = profile.displayName  || "";
    const mail        = profile.emailAddress || "";

    // Try SSO → Graph for jobTitle. Headless context, so must be silent.
    // Falls back to RoamingSettings, then to empty (and line is omitted).
    const jobTitle = await resolveJobTitle(settings);

    const signoff  = isReplyContext(item) ? replySignoff : newSignoff;
    const fullName = title ? `${title} ${displayName}` : displayName;

    const sig = buildSignature({ signoff, fullName, jobTitle, mail, ext, phone });

    item.body.setSignatureAsync(
      sig,
      { coercionType: Office.CoercionType.Html },
      () => event.completed()
    );
  } catch (err) {
    console.error("Auto-insert failed:", err);
    event.completed();
  }
}

/* ── SSO → Graph /me, with silent fallback ───────────────────────────── */
async function resolveJobTitle(settings) {
  const cached = settings.get("jobTitle") || "";
  try {
    const token = await withTimeout(
      Office.auth.getAccessToken({
        allowSignInPrompt: false,
        allowConsentPrompt: false,
        forMSGraphAccess: true
      }),
      SSO_TIMEOUT_MS
    );
    if (!token) return cached;

    const res = await withTimeout(
      fetch("https://graph.microsoft.com/v1.0/me?$select=jobTitle", {
        headers: { Authorization: `Bearer ${token}` }
      }).then(r => r.ok ? r.json() : null),
      SSO_TIMEOUT_MS
    );

    const fresh = (res && res.jobTitle) ? res.jobTitle : "";
    if (fresh) {
      // Cache it for next time, including for compose windows where SSO
      // might be slower or fail. Best-effort, don't block on the save.
      try {
        settings.set("jobTitle", fresh);
        settings.saveAsync(() => {});
      } catch (e) { /* ignore */ }
      return fresh;
    }
    return cached;
  } catch (err) {
    // Silent fallback — SSO unavailable, not provisioned, consent missing,
    // offline, or just slow. Use whatever's cached (or nothing).
    return cached;
  }
}

/* ── Promise timeout helper ──────────────────────────────────────────── */
function withTimeout(promise, ms) {
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject(new Error("timeout")), ms))
  ]);
}

/* ── Reply vs new message ────────────────────────────────────────────── */
function isReplyContext(item) {
  try {
    if (item.conversationId) return true;
    if (item.subject && typeof item.subject === "string" &&
        /^(re|fw|fwd)\s*:/i.test(item.subject.trim())) {
      return true;
    }
  } catch (e) { /* default to new */ }
  return false;
}

/* ── Signature HTML (kept identical to taskpane.js buildSignature) ──── */
function buildSignature({ signoff, fullName, jobTitle, mail, ext, phone }) {
  const extLine = ext
    ? ` <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">| Ext: </span></strong><strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#000000;">${ext}</span></strong>`
    : "";

  const telHref = phone ? phone.replace(/[^+\d]/g, "") : "";
  const phoneLine = phone
    ? ` <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">| P: </span></strong><strong><u><a href="tel:${telHref}" style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#000000;text-decoration:underline;">${phone}</a></u></strong>`
    : "";

  // Role paragraph is conditional — empty string if jobTitle is blank so
  // we don't render an awkward empty line for users who've never opened
  // the taskpane AND have no jobTitle in Entra.
  const rolePara = jobTitle
    ? `<p style="margin:0pt;line-height:normal;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt;color:#005953;">${jobTitle}</span></strong>
</p>`
    : "";

  return `
<p style="margin:0pt;line-height:normal;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt;">${signoff},</span></strong>
</p>
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
  <img src="${LOGO_URL}" width="${LOGO_WIDTH}" height="${LOGO_HEIGHT}" alt="Nambour Christian College" style="display:block;border:0;">
</p>
<p style="margin:0pt;line-height:normal;font-size:7pt;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">CRICOS:</span></strong>
  <span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#333333;"> 01461G</span>
</p>
`.trim();
}
