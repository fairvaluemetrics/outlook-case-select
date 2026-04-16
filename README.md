# Outlook Case Tagger

A sideloaded Outlook add-in that adds a **case selector** to every compose
window. Picking a case prepends the case name to the subject line in a
configurable format (e.g. `Kandel v Claude: Discovery follow-up`). If the user
tries to send without picking a case, the send is blocked until they either
pick one or explicitly mark the email **"Not case related."**

The case list lives in a plain JSON file you can edit without touching code,
and the dropdown supports alphabetical, most-recently-used, and
manually-pinned-favorites sorting.

---

## Repository layout

```
.
├── manifest.xml                  # Office Add-in manifest (sideload this)
├── assets/                       # Icons referenced by the manifest
├── config/
│   ├── cases.json                # ← edit cases here, no code change needed
│   ├── cases.schema.json         # JSON schema for editor support
│   └── config.json               # tag format, sort defaults, etc.
├── src/
│   ├── taskpane/                 # The pane that lives next to the subject
│   │   ├── taskpane.html
│   │   ├── taskpane.css
│   │   └── taskpane.js
│   └── commands/                 # Headless event handlers
│       ├── commands.html
│       └── commands.js           # OnMessageSend / OnNewMessageCompose
├── package.json                  # Optional dev tooling
└── README.md
```

---

## How it appears in Outlook

Outlook add-ins cannot literally inject HTML into the native subject row, but
the manifest's task pane plus the **`SupportsPinning`** flag gets you as close
as the platform allows:

1. A **Tag case** button appears on the **Message** ribbon of every new compose
   window.
2. The first time you open the pane, click the pushpin (📌) in the pane's
   header. From then on, the pane **opens automatically** alongside every
   compose window and stays docked next to the subject line.
3. The pane mirrors the current subject and lets you choose / change / clear
   the case at any point during composition.
4. On send, the **`OnMessageSend`** handler in `commands.js` blocks the send
   with a Smart Alert if no case has been picked and the user hasn't marked
   the message "Not case related."

> **Note:** The auto-launch `OnNewMessageCompose` event is wired up so the
> add-in code is loaded as soon as composition starts. Pinning the pane is a
> one-time per-mailbox action — Outlook doesn't allow add-ins to force a pane
> open.

---

## Editing the case list (no code changes needed)

Edit `config/cases.json`:

```json
{
    "cases": [
        { "id": "kandel-v-claude", "name": "Kandel v Claude", "aliases": ["kandel"], "order": 1 },
        { "id": "doe-v-acme",      "name": "Doe v Acme",       "aliases": ["acme"],   "order": 2 },
        "Quick & Simple Case Name"
    ]
}
```

Each entry can be either:

* a **plain string** — used as both the display name and (slugified) id, or
* an **object** with:
  * `id` — stable identifier (defaults to slug of `name`)
  * `name` — what shows up in the dropdown and gets prepended to the subject
  * `aliases` — extra strings the search box matches against
  * `order` — used by the **"Manual order"** sort mode

The pane fetches `cases.json` with cache-busting on every load, so saving the
file and reopening the compose window picks up the new list immediately.

---

## Configuring tag format & sorting

Edit `config/config.json`:

```json
{
    "casesFile": "../../config/cases.json",
    "tagFormat": "{case}: {subject}",
    "tagDetectRegex": "^[^:]{1,80}:\\s",
    "defaultSort": "pinned-mru",
    "allowNotCaseRelated": true,
    "mruLimit": 25
}
```

| Key | Meaning |
|-----|---------|
| `casesFile` | Path to the case list. Resolved relative to `src/taskpane/taskpane.html`. Move `cases.json` anywhere reachable by the add-in's host and update this path. |
| `tagFormat` | The string template used when stamping the subject. Must include `{case}` and `{subject}`. Example: `"[{case}] {subject}"` or `"{case} — {subject}"`. |
| `tagDetectRegex` | Optional regex used to strip an unknown legacy tag when the user picks a different case. Set to `""` to disable. |
| `defaultSort` | Initial sort mode: `"pinned-mru"`, `"mru"`, `"alpha"`, or `"manual"`. The user can override this per-mailbox via the dropdown in the pane. |
| `allowNotCaseRelated` | If `false`, the **Not case related** button is hidden and every send must have a real case selected. |
| `mruLimit` | Max number of recently-used cases to remember in the user's roaming settings. |

The user's last sort choice and pinned cases are stored in
`Office.context.roamingSettings`, so they ride along with the mailbox.

---

## Sorting modes

* **Pinned + Most recent** *(default)* — pinned cases first, then by recency,
  then alphabetical for the long tail.
* **Most recently used** — purely by last-used timestamp; cases never used
  fall to the bottom alphabetically.
* **Alphabetical** — by display name.
* **Manual order** — uses the `order` field from `cases.json`. Useful when
  you want a hand-curated list independent of usage.

Pin / unpin a case by clicking the ☆/★ button on its row.

---

## Send-blocking

`commands.js` registers an `OnMessageSend` handler. When Outlook calls it the
handler reads two custom properties off the mail item:

* `caseTagger.caseId` — set when the user picks a case
* `caseTagger.notCaseRelated` — set when the user clicks **Not case related**

If neither is set, the handler calls
`event.completed({ allowEvent: false, ... })`. Outlook shows a Smart Alert
that points the user at the **Tag case** button.

If `OnMessageSend` for some reason can't load (`item.loadCustomPropertiesAsync`
unavailable), the handler **fails open** rather than wedging the send.

---

## Local development & sideloading

You need to host the static files over HTTPS on `https://localhost:3000`. The
included `package.json` wires that up via the official Office tooling:

```bash
npm install
npm run certs                       # one-time: install dev HTTPS certs
npm run validate                    # sanity check the manifest
npm start                           # launches Outlook with the add-in sideloaded
```

`npm start` uses `office-addin-debugging` to (a) start a local HTTPS server
serving the repo root, (b) sideload `manifest.xml` into your desktop Outlook,
and (c) open Outlook for you. Stop it with `npm stop`.

If you'd rather sideload manually:

1. Run any HTTPS static server on `https://localhost:3000` rooted at this
   directory (e.g. `npm run serve:simple` after `npm run certs`).
2. In Outlook on the web → **Get Add-ins** → **My add-ins** → **Add a custom
   add-in** → **Add from file…** → pick `manifest.xml`.
3. Open a new email. The **Tag case** button appears on the **Message** ribbon.

For production deployment, change every `https://localhost:3000` URL in
`manifest.xml` to wherever you're hosting the static files (e.g. an Azure
Static Web App, S3 + CloudFront, or your own server) and re-distribute the
manifest through your tenant's centralized deployment.

---

## What gets stored where

| Data | Stored on | Lifetime |
|------|-----------|----------|
| Selected case for the current draft | Custom property on the mail item | Persists with the draft; cleared when send-blocking is satisfied |
| MRU + pinned + sort preference | `Office.context.roamingSettings` | Per-mailbox, follows the user across devices |
| Case list | `config/cases.json` on the add-in's host | Edited by hand / by deploy |

No data leaves the add-in's host.
