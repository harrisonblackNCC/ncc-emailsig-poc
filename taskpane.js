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

const LOGO_TARGET_WIDTH = 430;
const SSO_TIMEOUT_MS = 6000;   // give up on the whole NAA + Graph chain after 6s

// Organisations the signature can render as. Default = ncc; staff can
// switch via the Organisation dropdown at the top of the taskpane and the
// choice persists per-mailbox via RoamingSettings ("org" key).
const ORGS = {
  ncc: {
    displayName: "Nambour Christian College",
    logoUrl: "https://www.ncc.qld.edu.au/wp-content/uploads/NCC-Email_600x200.jpg",
    aspect: 600 / 200,
    affiliationText: "",       // empty = no affiliation line under role
    showSchoolDetails: true    // college name + address + phone/email/web
  },
  group: {
    displayName: "NCC Education Group",
    logoUrl: "https://www.ncc.qld.edu.au/wp-content/uploads/cc118d61-7fab-4089-93fd-6dc007d00674.jpg",
    aspect: 600 / 150,         // approx — overwritten by detectLogoAspects()
    affiliationText: "NCC Education Group",  // shown immediately under role
    showSchoolDetails: false   // composite logo carries all the school branding
  }
};

// Entra app registration for the add-in's API.
const CLIENT_ID  = "55e5528d-7efd-4bd5-a437-0d31c68d3542";
// Authority targets the NCC tenant. Using the verified domain is
// equivalent to the tenant GUID and keeps this code free of magic IDs.
const AUTHORITY  = "https://login.microsoftonline.com/nambourcc.onmicrosoft.com";

let userProfile = { displayName: "", jobTitle: "", mail: "" };

/* ── Detect aspect ratios for both org logos ────────────────────────── */
// Image aspect ratios are detected from the actual served images so the
// signature renders proportionally regardless of which logo is in use.
// Detected values are cached to RoamingSettings so launchevent.js (which
// runs in a headless context with no DOM/Image API to use) can read the
// real ratio on every compose. crossOrigin is intentionally NOT set —
// the WP server may not send CORS headers, and we don't need pixel
// access, only naturalWidth/naturalHeight which are always readable.
function detectAndCacheAspect(key) {
  try {
    const img = new Image();
    img.onload = () => {
      if (!img.naturalWidth || !img.naturalHeight) return;
      const aspect = img.naturalWidth / img.naturalHeight;
      ORGS[key].aspect = aspect;
      try {
        if (Office && Office.context && Office.context.roamingSettings) {
          Office.context.roamingSettings.set("logoAspect_" + key, aspect);
          Office.context.roamingSettings.saveAsync(() => { /* fire & forget */ });
        }
      } catch (e) { /* roaming settings not ready yet — Office.onReady will retry on next save */ }
    };
    img.src = ORGS[key].logoUrl;
  } catch (e) { /* fall back to default aspect */ }
}

// Read any cached aspects synchronously at boot (faster than waiting for
// Image load) — only safe after Office.onReady, so guarded.
function loadCachedAspects() {
  try {
    Object.keys(ORGS).forEach(key => {
      const cached = Office.context.roamingSettings.get("logoAspect_" + key);
      if (typeof cached === "number" && cached > 0) ORGS[key].aspect = cached;
    });
  } catch (e) { /* settings not ready */ }
}

// Kick off live detection for both orgs immediately.
Object.keys(ORGS).forEach(detectAndCacheAspect);

/* ── Load admin-managed variants ─────────────────────────────────────
   Variants are event-themed signatures (e.g. school musicals) that admin
   maintains via admin.html. They get exported to variants.json and
   committed to the GitHub Pages repo. This taskpane fetches that JSON
   on launch, merges entries into ORGS so they show up in the dropdown,
   and caches them in RoamingSettings so launchevent.js can read them
   on every compose without paying a fetch cost. */
async function loadVariants() {
  try {
    const res = await fetch("variants.json", { cache: "no-cache" });
    if (!res.ok) return [];
    const data = await res.json();
    const list = Array.isArray(data) ? data : (data && Array.isArray(data.variants) ? data.variants : []);
    list.forEach(v => {
      if (!v || !v.key || !v.logoUrl) return;
      const baseKey = v.baseOrg || "ncc";
      const base = ORGS[baseKey] || ORGS.ncc;
      ORGS[v.key] = {
        displayName: v.displayName || base.displayName,
        logoUrl: v.logoUrl,
        aspect: base.aspect,
        affiliationText: base.affiliationText,
        showSchoolDetails: base.showSchoolDetails
      };
      // Detect aspect for the variant's logo too.
      detectAndCacheAspect(v.key);
    });
    // Cache merged variants to RoamingSettings for launchevent.js.
    try {
      Office.context.roamingSettings.set("variantsCache", JSON.stringify(list));
      Office.context.roamingSettings.saveAsync(() => {});
    } catch (e) { /* settings not ready */ }
    return list;
  } catch (e) { return []; }
}

function repopulateOrgSelect() {
  const sel = document.getElementById("org-select");
  if (!sel) return;
  const previousValue = sel.value;
  // Wipe non-base options (everything beyond the first two: ncc, group).
  Array.from(sel.options).forEach(opt => {
    if (opt.value !== "ncc" && opt.value !== "group") opt.remove();
  });
  Object.keys(ORGS).forEach(key => {
    if (key === "ncc" || key === "group") return;
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = ORGS[key].displayName;
    sel.appendChild(opt);
  });
  // Restore previous selection if still valid.
  if (ORGS[previousValue]) sel.value = previousValue;
}

/* ── Office init ────────────────────────────────────────────────────── */
Office.onReady(async () => {
  loadCachedAspects();
  await loadVariants();          // merge admin-published variants into ORGS
  repopulateOrgSelect();         // make them available in the dropdown
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

  // Role: editable, pre-filled with the user's saved override if any,
  // otherwise the latest M365 jobTitle. Don't clobber a value the user
  // is currently typing into.
  const roleInput = document.getElementById("role-input");
  if (roleInput && document.activeElement !== roleInput) {
    const settings = Office.context.roamingSettings;
    const override = settings.get("jobTitleOverride");
    roleInput.value = (override !== undefined && override !== null && override !== "")
      ? override
      : (userProfile.jobTitle || "");
  }
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

  // Organisation: default to "ncc" if nothing saved, or if a saved value
  // doesn't match a known org (defensive against legacy/misspelled keys).
  const savedOrg = settings.get("org");
  const orgSelect = document.getElementById("org-select");
  if (orgSelect) {
    orgSelect.value = (savedOrg && ORGS[savedOrg]) ? savedOrg : "ncc";
  }
  refreshOrgDisclaimer();

  document.getElementById("title-select").value = settings.get("title") || "";

  applySignoffSelection("newSignoff",
    settings.get("newSignoff") || settings.get("signoff") || "Kind regards");
  applySignoffSelection("replySignoff",
    settings.get("replySignoff") || "Thanks");

  document.getElementById("ext-input").value   = settings.get("ext")   || "";
  document.getElementById("phone-input").value = settings.get("phone") || "";

  // Inject the 5am-7pm option list into every wh-time select before we
  // try to apply saved values, otherwise the selects have nothing to pick.
  populateTimeSelects();
  loadWorkingHours();
}

/* ── Working hours ──────────────────────────────────────────────────── */
// NCC operates Mon-Fri only. Time options run 5am-7pm in 30-min steps
// (29 options). Generated at runtime + injected into every wh-time
// <select> so staff literally can't pick anything outside the window.
const WH_DAYS = [
  { key: "mon", label: "Monday" },
  { key: "tue", label: "Tuesday" },
  { key: "wed", label: "Wednesday" },
  { key: "thu", label: "Thursday" },
  { key: "fri", label: "Friday" }
];
const WH_DEFAULT_START = "08:00";
const WH_DEFAULT_END   = "16:00";

function buildTimeOptionList() {
  const out = [];
  for (let h = 5; h <= 19; h++) {
    for (let m = 0; m < 60; m += 15) {
      if (h === 19 && m > 0) break; // hard-stop at 19:00
      const value = `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
      out.push({ value, label: formatWHTime(value) });
    }
  }
  return out;
}

// Hidden input is the value source-of-truth. The visible button shows the
// human-formatted label. setWHTimeValue keeps both in lockstep + marks the
// active option in the popup.
function setWHTimeValue(hiddenInput, value) {
  if (!hiddenInput) return;
  hiddenInput.value = value || "";
  const wrap = hiddenInput.closest(".wh-time-wrap");
  if (!wrap) return;
  const btn = wrap.querySelector(".wh-time-btn");
  if (btn) btn.textContent = value ? formatWHTime(value) : "—";
  wrap.querySelectorAll(".wh-time-popup button").forEach(b => {
    b.classList.toggle("active", b.dataset.val === value);
  });
}

function setWHTimeDisabled(hiddenInput, disabled) {
  if (!hiddenInput) return;
  hiddenInput.disabled = disabled;
  const wrap = hiddenInput.closest(".wh-time-wrap");
  if (!wrap) return;
  const btn = wrap.querySelector(".wh-time-btn");
  if (btn) btn.disabled = disabled;
  const popup = wrap.querySelector(".wh-time-popup");
  if (popup && disabled) popup.classList.remove("open");
}

function populateTimeSelects() {
  const opts = buildTimeOptionList();
  const html = opts.map(o => `<button type="button" data-val="${o.value}">${o.label}</button>`).join("");
  document.querySelectorAll(".wh-time-wrap").forEach(wrap => {
    const popup = wrap.querySelector(".wh-time-popup");
    const hidden = wrap.querySelector("input[type='hidden']");
    if (!popup || !hidden) return;
    popup.innerHTML = html;
    // Clicking an option updates the hidden value + closes the popup.
    popup.querySelectorAll("button").forEach(optBtn => {
      optBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        setWHTimeValue(hidden, optBtn.dataset.val);
        popup.classList.remove("open");
      });
    });
    // Make sure the visible button reflects the current hidden value
    // (defaults already set on the hidden input via the markup).
    setWHTimeValue(hidden, hidden.value || hidden.dataset.default || WH_DEFAULT_START);
  });
  wireWHDropdownToggles();
}

// One delegated listener handles every time-button click + outside click.
let _whDropdownsWired = false;
function wireWHDropdownToggles() {
  if (_whDropdownsWired) return;
  _whDropdownsWired = true;
  document.addEventListener("click", (e) => {
    const btn = e.target.closest(".wh-time-btn");
    if (btn && !btn.disabled) {
      const wrap = btn.closest(".wh-time-wrap");
      const popup = wrap && wrap.querySelector(".wh-time-popup");
      if (!popup) return;
      const wasOpen = popup.classList.contains("open");
      // Close any other open popups first.
      document.querySelectorAll(".wh-time-popup.open").forEach(p => p.classList.remove("open"));
      if (!wasOpen) {
        popup.classList.add("open");
        // Scroll the active option into view so users land near their pick.
        const active = popup.querySelector("button.active");
        if (active) active.scrollIntoView({ block: "nearest" });
      }
      return;
    }
    // Click outside any popup closes them all.
    if (!e.target.closest(".wh-time-popup")) {
      document.querySelectorAll(".wh-time-popup.open").forEach(p => p.classList.remove("open"));
    }
  });
}

// Saved value -> closest valid option. Falls back to default if the saved
// value isn't in our list (e.g. legacy entries from before the 5am-7pm cap).
function pickValidTime(saved, fallback) {
  if (!saved || typeof saved !== "string") return fallback;
  const opts = buildTimeOptionList();
  return opts.some(o => o.value === saved) ? saved : fallback;
}

function loadWorkingHours() {
  const settings = Office.context.roamingSettings;
  let saved;
  try { saved = JSON.parse(settings.get("workingHours") || "null"); }
  catch (e) { saved = null; }

  const showToggle = document.getElementById("wh-show");
  const list = document.getElementById("wh-days-list");
  if (!showToggle || !list) return;

  const showOn = !!(saved && saved.show);
  showToggle.checked = showOn;
  list.style.display = showOn ? "block" : "none";

  WH_DAYS.forEach(d => {
    const dayPrefs = (saved && saved.days && saved.days[d.key]) || {};
    const cb = document.getElementById("wh-" + d.key);
    const startEl = document.getElementById("wh-" + d.key + "-start");
    const endEl   = document.getElementById("wh-" + d.key + "-end");
    if (!cb || !startEl || !endEl) return;
    cb.checked = !!dayPrefs.active;
    setWHTimeValue(startEl, pickValidTime(dayPrefs.start, WH_DEFAULT_START));
    setWHTimeValue(endEl,   pickValidTime(dayPrefs.end,   WH_DEFAULT_END));
    setWHTimeDisabled(startEl, !cb.checked);
    setWHTimeDisabled(endEl,   !cb.checked);
  });
}

function collectWorkingHours() {
  const showToggle = document.getElementById("wh-show");
  const out = { show: !!(showToggle && showToggle.checked), days: {} };
  WH_DAYS.forEach(d => {
    const cb = document.getElementById("wh-" + d.key);
    const startEl = document.getElementById("wh-" + d.key + "-start");
    const endEl   = document.getElementById("wh-" + d.key + "-end");
    out.days[d.key] = {
      active: !!(cb && cb.checked),
      start:  (startEl && startEl.value) || WH_DEFAULT_START,
      end:    (endEl && endEl.value)     || WH_DEFAULT_END
    };
  });
  return out;
}

/* Format "HH:MM" 24h string to friendly 12h form.
   "09:00" -> "9am"   "13:30" -> "1:30pm"   "12:00" -> "12pm"   "00:00" -> "12am" */
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

/* Compact e.g. Tue 9-14 / Wed 9-14 / Thu 9-14 / Fri 9-13:30
   into "Tuesday-Thursday 9am-2pm, Friday 9am-1:30pm".
   Only consecutive days with identical start AND end times group. */
function compactWorkingHours(wh) {
  if (!wh || !wh.show || !wh.days) return "";
  const ordered = WH_DAYS.map(d => ({
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

/* ── Org-disclaimer visibility ──────────────────────────────────────── */
// Only the NCC Education Group signature requires Director-of-Marketing
// sign-off. Other orgs hide the warning entirely.
function refreshOrgDisclaimer() {
  const sel = document.getElementById("org-select");
  const banner = document.getElementById("org-group-disclaimer");
  if (!sel || !banner) return;
  banner.style.display = (sel.value === "group") ? "block" : "none";
}

/* ── Wire events ────────────────────────────────────────────────────── */
function wireEvents() {
  const orgSelect = document.getElementById("org-select");
  if (orgSelect) orgSelect.addEventListener("change", refreshOrgDisclaimer);

  ["newSignoff", "replySignoff"].forEach(prefix => {
    const sel = document.getElementById(prefix + "-select");
    sel.addEventListener("change", () => {
      const wrap = document.getElementById(prefix === "newSignoff" ? "newCustom-wrap" : "replyCustom-wrap");
      wrap.style.display = sel.value === "custom" ? "block" : "none";
    });
  });

  // Working hours: master toggle expands/collapses the day list.
  const whToggle = document.getElementById("wh-show");
  if (whToggle) {
    whToggle.addEventListener("change", () => {
      const list = document.getElementById("wh-days-list");
      if (list) list.style.display = whToggle.checked ? "block" : "none";
    });
  }

  // Each day checkbox enables/disables its time pickers in lockstep.
  WH_DAYS.forEach(d => {
    const cb = document.getElementById("wh-" + d.key);
    if (!cb) return;
    cb.addEventListener("change", () => {
      setWHTimeDisabled(document.getElementById("wh-" + d.key + "-start"), !cb.checked);
      setWHTimeDisabled(document.getElementById("wh-" + d.key + "-end"),   !cb.checked);
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

    const orgSelect = document.getElementById("org-select");
    const org = (orgSelect && ORGS[orgSelect.value]) ? orgSelect.value : "ncc";

    settings.set("org",          org);
    settings.set("title",        title);
    settings.set("newSignoff",   newSignoff);
    settings.set("replySignoff", replySignoff);
    settings.set("ext",          ext);
    settings.set("phone",        phone);

    // Role: user can override the M365 jobTitle. Empty string = "use the
    // M365 default". The cached "jobTitle" key (set by loadProfile + the
    // launchevent SSO call) stays as the upstream source of truth.
    const roleInput = document.getElementById("role-input");
    const roleValue = roleInput ? roleInput.value.trim() : "";
    const upstreamJobTitle = (settings.get("jobTitle") || userProfile.jobTitle || "").trim();
    if (roleValue && roleValue !== upstreamJobTitle) {
      settings.set("jobTitleOverride", roleValue);
    } else {
      settings.set("jobTitleOverride", "");
    }

    // Working hours — stored as JSON so launchevent.js can re-apply
    // the same compaction logic on every compose.
    settings.set("workingHours", JSON.stringify(collectWorkingHours()));

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

  // Role: prefer (1) what's currently in the input, then (2) saved override,
  // then (3) live M365 jobTitle, then (4) cached jobTitle.
  const roleInputEl  = document.getElementById("role-input");
  const liveRole     = roleInputEl ? roleInputEl.value.trim() : "";
  const jobTitle     = liveRole
                       || settings.get("jobTitleOverride")
                       || userProfile.jobTitle
                       || settings.get("jobTitle")
                       || "";
  const newSignoff   = resolveSignoff("newSignoff",   settings.get("newSignoff")   || "Kind regards");
  const replySignoff = resolveSignoff("replySignoff", settings.get("replySignoff") || "Thanks");
  const ext          = document.getElementById("ext-input").value.trim()   || settings.get("ext")   || "";
  const phone        = document.getElementById("phone-input").value.trim() || settings.get("phone") || "";

  const signoff = isReplyContext() ? replySignoff : newSignoff;
  const { displayName, mail } = userProfile;
  const fullName = title ? `${title} ${displayName}` : displayName;

  // Pick the org for this signature. Selected value > saved setting > ncc.
  const orgSelect = document.getElementById("org-select");
  const orgKey    = (orgSelect && ORGS[orgSelect.value])
                    ? orgSelect.value
                    : (ORGS[settings.get("org")] ? settings.get("org") : "ncc");
  const org = ORGS[orgKey];

  const extLine = ext
    ? ` <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">| Ext: </span></strong><strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#000000;">${ext}</span></strong>`
    : "";

  const telHref = phone ? phone.replace(/[^+\d]/g, "") : "";
  const phoneLine = phone
    ? ` <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">| P: </span></strong><strong><u><a href="tel:${telHref}" style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#000000;text-decoration:underline;">${phone}</a></u></strong>`
    : "";

  // Role paragraph is conditional — skip the line entirely if blank so we
  // don't render an awkward empty paragraph. Kept in sync with launchevent.js.
  // Explicit line-height (vs. "normal") tightens the gap to working hours +
  // contact line so the upper block matches the tight bottom block.
  const rolePara = jobTitle
    ? `<p style="margin:0pt;line-height:12pt;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt;color:#005953;">${jobTitle}</span></strong>
</p>`
    : "";

  // Working hours line: italic, smaller than role, NCC green. Sits under
  // role and above contact details. Compaction collapses runs of identical
  // hours so the line stays short.
  const whLive = (typeof collectWorkingHours === "function") ? collectWorkingHours() : null;
  let whSource = whLive;
  if (!whSource) {
    try { whSource = JSON.parse(settings.get("workingHours") || "null"); }
    catch (e) { whSource = null; }
  }
  const whText = compactWorkingHours(whSource);

  // In themes that drop the school-details block, the contact info
  // becomes its own visual block — give it a 12pt gap from the role/
  // affiliation block above. Falls on whPara if shown, else on the
  // contact line below.
  const contactBlockGap = !org.showSchoolDetails ? "margin-top:12pt;" : "";
  const whParaTop  = whText ? contactBlockGap : "";
  const contactTop = whText ? "" : contactBlockGap;

  const whPara = whText
    ? `<p style="margin:0pt;${whParaTop}line-height:10pt;background-color:#ffffff;">
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

  // Affiliation line (e.g. "NCC Education Group") sits immediately under
  // role for orgs whose composite logo replaces the school-details block.
  // Same styling as role so the two read as a single 2-line title block.
  const affiliationPara = org.affiliationText
    ? `<p style="margin:0pt;line-height:12pt;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt;color:#005953;">${org.affiliationText}</span></strong>
</p>`
    : "";

  // School details block: college name + address + phone/email/web links.
  // Skipped for orgs whose composite logo carries that info already.
  const schoolDetailsBlock = org.showSchoolDetails
    ? `<p style="margin:0pt;margin-top:12pt;line-height:normal;font-size:9pt;background-color:#ffffff;">
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
</p>`
    : "";

  // When the school-details block is dropped, push the logo down so there's
  // still a clean gap from the contact line above it.
  const logoMarginTop = org.showSchoolDetails ? "0pt" : "12pt";

  return `
${signoffPara}
<p style="margin:0pt;margin-bottom:12pt;line-height:13pt;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;font-size:11pt;color:#ec3426;">${fullName}</span></strong>
</p>
${rolePara}
${affiliationPara}
${whPara}
<p style="margin:0pt;${contactTop}line-height:10pt;font-size:9pt;background-color:#ffffff;">
  <strong><span style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#005953;">E: </span></strong><strong><u><a href="mailto:${mail}" style="font-family:Aptos,Calibri,Helvetica,Arial,sans-serif;color:#000000;text-decoration:underline;">${mail}</a></u></strong>${extLine}${phoneLine}
</p>
${schoolDetailsBlock}
<p style="margin:0pt;margin-top:${logoMarginTop};line-height:normal;font-size:11pt;background-color:#ffffff;">
  <img src="${org.logoUrl}" width="${LOGO_TARGET_WIDTH}" height="${Math.round(LOGO_TARGET_WIDTH / org.aspect)}" alt="${org.displayName}" style="display:block;border:0;">
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
