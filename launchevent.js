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

const LOGO_WIDTH    = 430;
const SSO_TIMEOUT_MS = 3000; // give up on SSO after 3s, fall back silently

// Org config — must match ORGS in taskpane.js. We can't auto-detect aspect
// here (no DOM in headless context), so the height is computed from a
// stored aspect ratio per org. Update aspect when the asset changes.
const ORGS = {
  ncc: {
    displayName: "Nambour Christian College",
    logoUrl: "https://www.ncc.qld.edu.au/wp-content/uploads/NCC-Email_600x200.jpg",
    aspect: 600 / 200
  },
  group: {
    displayName: "NCC Education Group",
    logoUrl: "https://www.ncc.qld.edu.au/wp-content/uploads/cc118d61-7fab-4089-93fd-6dc007d00674.jpg",
    aspect: 600 / 200
  }
};

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

    // Role logic:
    //  1. If staff have explicitly overridden their role in the taskpane,
    //     respect that — don't override their override.
    //  2. Otherwise try SSO → Graph for the live M365 jobTitle.
    //  3. Fall back to RoamingSettings cache, then to empty.
    const override = (settings.get("jobTitleOverride") || "").trim();
    const jobTitle = override || await resolveJobTitle(settings);

    const signoff  = isReplyContext(item) ? replySignoff : newSignoff;
    const fullName = title ? `${title} ${displayName}` : displayName;

    // Working hours: parse the JSON the taskpane saved + compact for display.
    let wh = null;
    try { wh = JSON.parse(settings.get("workingHours") || "null"); }
    catch (e) { wh = null; }
    const whText = compactWorkingHours(wh);

    // Organisation: defaults to "ncc" if nothing saved or saved key isn't known.
    const savedOrg = settings.get("org");
    const orgKey   = (savedOrg && ORGS[savedOrg]) ? savedOrg : "ncc";

    const sig = buildSignature({ signoff, fullName, jobTitle, mail, ext, phone, whText, orgKey });

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

/* ── Working hours helpers (kept identical to taskpane.js) ──────────── */
// NCC operates Mon-Fri only. Sat/Sun stay out of the loop so any legacy
// saved weekend entries silently drop out of the rendered signature.
const WH_DAYS_LE = [
  { key: "mon", label: "Monday" },
  { key: "tue", label: "Tuesday" },
  { key: "wed", label: "Wednesday" },
  { key: "thu", label: "Thursday" },
  { key: "fri", label: "Friday" }
];

function formatWHTime(t) {
  if (!t || typeof t !== "string" || !t.includes(":")) return t || "";
  const [hStr, mStr] = t.split(":");
  const h = parseInt(hStr, 10);
  const m = parseInt(mStr, 10);
  if (isNaN(h) || isNaN(m)) return t;
  const period = h < 12 ? "am" : "pm";
  const h12 = ((h + 11) % 12) + 1;
  return m === 0 ? `${h12}${period}` : `${h12}:${String(m).padStart(2, "0")}${period}`;
}

function compactWorkingHours(wh) {
  if (!wh || !wh.show || !wh.days) return "";
  const ordered = WH_DAYS_LE.map(d => ({
    key: d.key,
    label: d.label,
    ...((wh.days[d.key]) || { active: false, start: "", end: "" })
  }));
  const groups = [];
  let current = null;
  for (let i = 0; i < ordered.length; i++) {
    const day = ordered[i];
    if (!day.active) { if (current) { groups.push(current); current = null; } continue; }
    if (current && current.endIdx === i - 1 && current.start === day.start && current.end === day.end) {
      current.endLabel = day.label;
      current.endIdx = i;
    } else {
      if (current) groups.push(current);
      current = {
        startLabel: day.label, endLabel: day.label,
        start: day.start, end: day.end,
        startIdx: i, endIdx: i
      };
    }
  }
  if (current) groups.push(current);
  return groups.map(g => {
    const days = (g.startLabel === g.endLabel) ? g.startLabel : `${g.startLabel}-${g.endLabel}`;
    return `${days} ${formatWHTime(g.start)}-${formatWHTime(g.end)}`;
  }).join(", ");
}

/* ── Signature HTML (kept identical to taskpane.js buildSignature) ──── */
function buildSignature({ signoff, fullName, jobTitle, mail, ext, phone, whText, orgKey }) {
  const org = ORGS[orgKey] || ORGS.ncc;
  const LOGO_HEIGHT = Math.round(LOGO_WIDTH / org.aspect);
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
  // Explicit line-height (vs. "normal") tightens the gap to working hours +
  // contact line. Kept in sync with taskpane.js.
  const rolePara = jobTitle
    ? `<p style="margin:0pt;line-height:12pt;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt;color:#005953;">${jobTitle}</span></strong>
</p>`
    : "";

  // Working hours line: italic, smaller than role, NCC green. Sits under
  // role and above contact details. Skipped when blank.
  const whPara = whText
    ? `<p style="margin:0pt;line-height:10pt;background-color:#ffffff;">
  <strong><em><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:9pt;color:#005953;">Working:</span></em></strong><em><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:9pt;color:#000000;"> ${whText}</span></em>
</p>`
    : "";

  // signoff "" = staff picked "None" — skip the sign-off paragraph
  // entirely so we don't render a lone comma above the name.
  const signoffPara = signoff
    ? `<p style="margin:0pt;line-height:normal;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt;color:#000000;">${signoff},</span></strong>
</p>`
    : "";

  return `
${signoffPara}
<p style="margin:0pt;margin-bottom:12pt;line-height:13pt;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt;color:#ec3426;">${fullName}</span></strong>
</p>
${rolePara}
${whPara}
<p style="margin:0pt;line-height:10pt;font-size:9pt;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">E: </span></strong><strong><u><a href="mailto:${mail}" style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#000000;text-decoration:underline;">${mail}</a></u></strong>${extLine}${phoneLine}
</p>
<p style="margin:0pt;margin-top:12pt;line-height:normal;font-size:9pt;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">${org.displayName}</span></strong>
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
  <img src="${org.logoUrl}" width="${LOGO_WIDTH}" height="${LOGO_HEIGHT}" alt="${org.displayName}" style="display:block;border:0;">
</p>
<p style="margin:0pt;line-height:normal;font-size:7pt;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">CRICOS:</span></strong>
  <span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#333333;"> 01461G</span>
</p>
`.trim();
}
