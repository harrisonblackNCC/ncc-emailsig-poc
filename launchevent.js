/* NCC Email Signature Client - launchevent.js
 *
 * Event-based activation: runs automatically the moment a compose window
 * opens and drops the personalised NCC signature in — no taskpane, no
 * click, no UI. Pulls name + email from mailbox.userProfile (no SSO
 * needed) and title/role/signoff/ext/phone from RoamingSettings (set
 * once via the taskpane, reused forever).
 *
 * The HTML signature here is kept in lockstep with buildSignature() in
 * taskpane.js so the output is identical whether the add-in inserts
 * automatically or the user clicks the button manually.
 */

const LOGO_URL = "https://www.ncc.qld.edu.au/wp-content/uploads/NCC-Email_600x200.jpg";
const LOGO_WIDTH  = 430;
const LOGO_HEIGHT = 143; // 430 / (600/200) rounded — matches the WordPress-hosted asset

/* Register the handler so the manifest's FunctionName resolves. */
Office.onReady(() => {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
});

/* ── Main handler ────────────────────────────────────────────────────── */
function onNewMessageComposeHandler(event) {
  try {
    const settings = Office.context.roamingSettings;
    const profile  = Office.context.mailbox.userProfile;
    const item     = Office.context.mailbox.item;

    const title        = settings.get("title")        || "";
    const jobTitle     = settings.get("jobTitle")     || "";
    const newSignoff   = settings.get("newSignoff")   || settings.get("signoff") || "Kind regards";
    const replySignoff = settings.get("replySignoff") || "Thanks";
    const ext          = settings.get("ext")          || "";
    const phone        = settings.get("phone")        || "";

    const displayName = profile.displayName  || "";
    const mail        = profile.emailAddress || "";

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
  <img src="${LOGO_URL}" width="${LOGO_WIDTH}" height="${LOGO_HEIGHT}" alt="Nambour Christian College" style="display:block;border:0;">
</p>
<p style="margin:0pt;line-height:normal;font-size:7pt;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">CRICOS:</span></strong>
  <span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#333333;"> 01461G</span>
</p>
`.trim();
}
