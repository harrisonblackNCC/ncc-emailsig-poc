# NCC Email Signature Client

Outlook add-in that auto-inserts a branded NCC email signature on every new
message, reply and forward. Pulls each staff member's name and role straight
from their M365 profile via Nested App Authentication (NAA) — no backend, no
secrets, no per-user setup.

Runs on Outlook on the Web, new Outlook for Windows / Mac, and Outlook for iOS
and Android. Outlook Classic for Windows falls back to a manual ribbon button
since Classic does not support NAA.

---

## Repo layout

This folder is a 1:1 mirror of the `ncc-emailsig-poc` GitHub Pages repo. Push
the whole thing to the repo root.

```
.
├── README.md                ← this file
├── manifest.xml             ← v1.4.0.0 — sideload or Centrally Deploy
├── .gitignore
│
├── index.html               ← landing page at github.io/ncc-emailsig-poc/
├── admin.html               ← marketing/IT signature template editor
│
├── taskpane.html            ← ribbon-button taskpane UI
├── taskpane.js              ← NAA → Graph /me, builds + inserts signature
│
├── launchevent.html         ← headless runtime host for auto-apply
├── launchevent.js           ← OnNewMessageCompose handler
│
├── commands.html            ← required Office FunctionFile stub
├── favicon.ico
├── icon-{16,32,64,80,128}.png
├── illustration.png         ← landing-page background
├── logo.png                 ← NCC shield (used on landing/admin pages)
│
└── docs/
    └── azure-setup.md           ← everything Daniel/IT needs: NAA, SP provisioning, common errors
```

Old manifest backups and superseded drafts live one level up in
`Final Ship/misc/` — local-only, not part of the repo.

The signature image (`NCC-Email_600x200.jpg`) is currently hosted on the
school WordPress site. Self-hosting in this repo is parked — see "Backlog"
below.

---

## How it works

1. Office loads `manifest.xml` and registers the LaunchEvent + ribbon button.
2. On every new compose / reply / forward, Outlook fires
   `OnNewMessageCompose`, which runs `launchevent.js` headlessly.
3. `launchevent.js` reads the user's saved preferences from `RoamingSettings`
   (title, ext, phone, sign-offs), grabs name + email from
   `mailbox.userProfile`, and tries SSO → Graph `/me` for `jobTitle` (3-second
   silent fallback to cached value).
4. It builds the HTML signature and calls `item.body.setSignatureAsync` —
   the signature lands in the compose window before the user starts typing.
5. If the user wants to customise their preferences (extension, direct
   line, sign-off phrase), they hit the **Signature** ribbon button to open
   `taskpane.html`. Customisations save to RoamingSettings and apply on the
   next compose.

`taskpane.js`, `launchevent.js`, and `admin.html` each contain a
`buildSignature()` function that produces identical HTML output. Keep them in
lockstep when changing the template.

---

## Deployment

### Desktop (sideload — for testing on your own mailbox)

1. Push this folder to the `ncc-emailsig-poc` repo, wait ~60s for Pages.
2. Outlook on the Web → Settings → Mail → Customise actions → Get add-ins
   → My add-ins → Custom Add-ins → "Add a custom add-in" → "Add from file"
   → pick `manifest.xml`.
3. Open a new compose. Signature should appear within ~2 seconds. Test reply
   and forward too.

### Mobile + org-wide (Centralized Deployment — required for iOS/Android)

Sideloading does **not** propagate to Outlook mobile. For mobile, Daniel does:

> Microsoft 365 Admin Centre → Settings → Integrated apps → Add-in → Upload
> custom apps → upload `manifest.xml` → assign to user(s).

Propagation to mobile clients takes 6–24 hours. Once it lands, the add-in
appears in Outlook for iOS/Android with no user action required. Test scope:
just yourself first, then a small pilot group, then full rollout.

---

## Azure prerequisites

All already provisioned. Documented in `docs/azure-setup.md`. Quick summary:

- App registration ID: `55e5528d-7efd-4bd5-a437-0d31c68d3542`
- Application ID URI: `api://harrisonblackncc.github.io/55e5528d-...`
- Scope: `access_as_user`
- API permission: `User.Read` (admin-consented)
- SPA redirect URIs:
  - `https://harrisonblackncc.github.io/ncc-emailsig-poc/taskpane.html`
  - `brk-9199bf20-a13f-4107-85dc-02114787ef48://harrisonblackncc.github.io`
  - `brk-multihub://harrisonblackncc.github.io`
- Outlook client IDs pre-authorised on the `access_as_user` scope.

If anything goes weird with auth, `docs/azure-setup.md` has a common-errors
table that should cover ~95% of cases.

---

## Versioning

| Version  | What changed                                                  |
| -------- | ------------------------------------------------------------- |
| 1.0.5.0  | Initial POC — no SSO, hand-typed role                          |
| 1.2.x    | SSO experiments via Application ID URI variants                |
| 1.2.5.0  | Switched to NAA — MSAL.js handles OBO client-side, no backend  |
| 1.3.0.0  | Added LaunchEvent: signature auto-applies on every compose     |
| 1.4.0.0  | Added MobileFormFactor — works on Outlook iOS + Android        |

When bumping, update `<Version>` in `manifest.xml` AND the changelog above.
Old manifests live in `Final Ship/misc/old-manifests/` for rollback reference;
that folder is one level up from this one and is not part of the repo.

---

## Rollback

If a deploy goes sideways:

1. Outlook Web → My add-ins → remove the current version.
2. Pick a known-good manifest from `../misc/old-manifests/` and re-sideload.
   The hosted JS files on Pages are mostly backwards-compatible — older
   manifests will work against the current `taskpane.js` and `launchevent.js`
   because the fallback paths are stable.
3. For a Centrally Deployed rollback, Daniel uploads the older manifest in
   the M365 Admin Centre. Mobile clients pick it up within 6–24h.

---

## Backlog

- **Self-host the signature logo image.** Currently pulled from
  `www.ncc.qld.edu.au/wp-content/uploads/...`. Cleanest fix: drop the PNG
  in this repo at `signature-logo.png` and update `LOGO_URL` in
  `taskpane.js` + `launchevent.js`. Optional v1.5 polish: an upload widget
  in `admin.html` so marketing can swap it without touching code.
- **Render proper mobile icons.** `manifest.xml` currently reuses
  `icon-32.png` and `icon-80.png` for the new 25/48 mobile sizes. They
  scale fine but aren't pixel-perfect — replace when marketing has time.
- **Telemetry / error reporting.** Right now if `launchevent.js` silently
  falls back to the mailbox-only path, we have no visibility. A lightweight
  Application Insights ping would help diagnose org-wide rollout issues.
