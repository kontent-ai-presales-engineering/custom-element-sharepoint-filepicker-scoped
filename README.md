# SharePoint File Picker (Alternative, Sites.Selected) — Kontent.ai Custom Element

An alternative to [`custom-element-sharepoint-filepicker`](https://github.com/kontent-ai-presales-engineering/custom-element-sharepoint-filepicker)
built for customers who need real, least-privilege, site-level access instead of tenant-wide
`AllSites.Read`.

It lets content editors pick one or more files from a small, explicitly approved list of
SharePoint sites, and stores the selected files' name, URL, author and last-modified date as the
element's value, in exactly the same shape the original element uses, so nothing downstream that
reads this value needs to change.

**What's different from the original element:**

| | Original (`custom-element-sharepoint-filepicker`) | This alternative |
|---|---|---|
| Picker UI | Microsoft's File Picker v8 (`_layouts/15/FilePicker.aspx`) in a second popup window | A small file browser built into the element's own panel, no second popup |
| Permissions required | Graph `Files.Read.All` + `Sites.Read.All`, **and** SharePoint (classic API) `AllSites.Read` + `MyFiles.Read` | Graph `Sites.Selected` only |
| Access scope | Every site in the tenant, plus the signed-in user's own OneDrive | Only the specific site(s) a tenant admin has explicitly granted, one at a time |
| Adding a new site | Nothing to configure, already had access to everything | Admin grants `Sites.Selected` on the new site (one Graph API call), then add it to this element's config |
| OneDrive (personal) picking | Supported | Not supported by default (see "Limitations" below) |

**Why this exists:** Microsoft's File Picker v8 has no concept of "just this one site", its own
permission model is all-or-nothing (`AllSites.*` for every site collection in the tenant, or
`MyFiles.*` for the signed-in user's own OneDrive only). Microsoft Graph's `Sites.Selected`
permission *does* support granting an app access to one specific site and nothing else, but it
only works for direct Graph calls, not for the File Picker v8 popup. So this alternative drops the
Microsoft-hosted picker UI entirely and browses the site(s) directly via Graph instead.

## 1. Register an Azure AD (Entra ID) application

1. In the [Azure Portal](https://portal.azure.com) (or Entra admin center), go to **App
   registrations → New registration**.
2. Set **Redirect URI** (platform: **Single-page application**, not "Web") to the exact URL where
   you will host `index.html`, e.g. `https://your-host.example.com/index.html`.
3. Under **API permissions**, add only:
   - **Microsoft Graph** (delegated) → `Sites.Selected`
   That's the entire permission set. No `AllSites.Read`, no `MyFiles.Read`, no `Sites.Read.All`,
   no `Files.Read.All`.
4. Grant admin consent for `Sites.Selected` (it's a privileged permission, so this step is still
   required even though the scope itself is narrow).
5. Copy the **Application (client) ID** and your **tenant ID or verified domain**
   (e.g. `contoso.onmicrosoft.com`) — both go into the element configuration below.

`Sites.Selected` alone grants the app access to **nothing** until an admin explicitly grants it a
role on a specific site. That's the next step, and it's the one that actually enforces the
least-privilege boundary.

## 2. Grant access to each specific site

This is a one-time step per approved site, done by someone with sufficient rights on that site
collection (a SharePoint/Global admin, or a `Sites.FullControl.All`-privileged caller). It can be
run from [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) while signed
in as an admin, or via any Graph-capable script.

**a. Resolve the site's ID** (read-only, needs an admin token with `Sites.Read.All` just for this
one lookup, or use the SharePoint admin center to find the site ID another way):

```http
GET https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/SalesFirst
```

Note the `id` field in the response, e.g. `contoso.sharepoint.com,xxxxxxxx-xxxx-...,yyyyyyyy-...`.

**b. Grant this app's service principal a role on that site:**

```http
POST https://graph.microsoft.com/v1.0/sites/{site-id-from-step-a}/permissions
Content-Type: application/json

{
  "roles": ["read"],
  "grantedToIdentities": [
    {
      "application": {
        "id": "<the Application (client) ID from section 1>",
        "displayName": "Kontent.ai SharePoint File Picker"
      }
    }
  ]
}
```

Use `"roles": ["read"]` since this element only ever reads files, it never writes back to
SharePoint.

**To add another approved site later**: repeat step 2 for the new site, then add it to the `sites`
array in the element's configuration (section 4). No code change, no new Entra permission, no new
admin consent, just one more per-site grant.

**To revoke access to a site:** delete the permission created in step (b), or remove/deny the
app's admin consent entirely. Either fully and immediately cuts the app off from that site.

## 3. Host the files

Same as the original element: this is a static site, any HTTPS static host works (GitHub Pages,
Azure Static Web Apps, Netlify, Vercel, S3/Blob Storage behind a CDN, etc.). Host the whole
repository root (`index.html`, `css/`, `js/`) and note the public URL of `index.html`; that's your
**Hosted code URL** for the next step. The redirect URI registered in Entra ID (section 1) must
match this URL exactly.

For local development:

```bash
npm start
```

Serves the project at `http://localhost:5500`. Add `http://localhost` as an additional redirect
URI in the app registration if you want to test sign-in locally.

## 4. Add the custom element in Kontent.ai

1. In your content type, add a **Custom element**.
2. Set **Hosted code URL** to the public URL of `index.html` from section 3.
3. In **Custom element configuration (JSON)**, provide:

```json
{
  "clientId": "00000000-0000-0000-0000-000000000000",
  "tenant": "contoso.onmicrosoft.com",
  "sites": [
    { "label": "SalesFirst", "hostname": "contoso.sharepoint.com", "path": "/sites/SalesFirst" }
  ],
  "selectionMode": "multiple",
  "debug": false
}
```

| Field           | Required | Description                                                                                     |
|-----------------|----------|---------------------------------------------------------------------------------------------------|
| `clientId`      | Yes      | The Entra application (client) ID from section 1.                                                |
| `tenant`        | Yes      | Your Entra tenant ID (GUID) or a verified domain, e.g. `"contoso.onmicrosoft.com"`.               |
| `sites`         | Yes      | Array of sites this element may browse. Each needs `hostname` (e.g. `"contoso.sharepoint.com"`), `path` (server-relative, e.g. `"/sites/SalesFirst"`), and an optional `label` shown in the site picker. |
| `selectionMode` | No       | `"multiple"` (default) or `"single"`.                                                            |
| `debug`         | No       | `true` to show a small status line useful for troubleshooting. Default `false`.                  |

If `sites` has exactly one entry, the element skips straight to it. With more than one, editors see
a small site dropdown before browsing. **Every entry must already have a `Sites.Selected` grant
from section 2**, adding a site here without also granting it access in Entra will just show that
editor an access-denied message when they try to open it, never silently fall back to browsing
somewhere else.

The element validates this configuration on load and shows an inline error if it's missing or
malformed, instead of failing silently.

## Stored value

Identical to the original element, so existing content items and any code reading this value keep
working unchanged:

```json
[
  {
    "id": "01ABCDEF...",
    "driveId": "b!AbCdEf...",
    "name": "Report.pdf",
    "url": "https://contoso.sharepoint.com/sites/SalesFirst/Shared Documents/Report.pdf",
    "author": "Jane Doe",
    "lastModified": "2026-01-15T10:30:00Z"
  }
]
```

## Limitations

- **No personal OneDrive picking.** The original element could also pick from the signed-in
  user's own OneDrive (`MyFiles.Read`). This alternative only browses the sites listed in
  `sites`, on purpose, since "any editor's personal OneDrive" is exactly the kind of broad,
  hard-to-audit access this rebuild exists to avoid. A personal OneDrive is technically its own
  site collection, so if a specific user's OneDrive genuinely needs to be added, it can be granted
  `Sites.Selected` access the same way as a team site, just ask a more unusual thing of the admin
  running section 2.
- **No search across sites.** Since the app can only ever see the sites it's been explicitly
  granted, there's no tenant-wide "search all of SharePoint" the way the original picker allowed,
  editors browse folder by folder within an approved site instead.
- **Read-only.** This element only requests the `read` role in section 2(b) and never writes to
  SharePoint. If a future use case needs the element to upload or modify files, that's a
  conscious permission upgrade (`"roles": ["write"]`), not something this version does implicitly.
- **One Kontent.ai custom element instance is shared across whatever `sites` its config lists.**
  If different content types need access to different, non-overlapping sets of sites, use
  separate custom element configurations (they can point at the same hosted code, just with
  different `sites` arrays), rather than granting one shared instance access to every site anyone
  might need.

## Notes for integrators

- **Embedding required**: Kontent.ai always loads custom elements inside a child iframe. If
  `index.html` is opened directly (`window === window.top`), the element shows a plain message
  instead of the picker UI rather than trying to run standalone.
- **Only one popup, not two**: sign-in still uses an MSAL popup (a silent iframe-based login isn't
  reliable for AAD), but the file browser itself is rendered inline inside the element's own
  panel, so there's no second popup window to get blocked the way the original element's file
  picker popup could be.
- **No `alert()`/`confirm()`**: Kontent.ai renders custom elements inside a sandboxed iframe that
  does not support native modal dialogs, so all error and confirmation UI in this element is
  implemented inline (banners) rather than with `alert()`/`confirm()`.
- **Content-Security-Policy**: `index.html` ships a restrictive CSP meta tag, tighter than the
  original element's, since there's no more SharePoint-hosted popup to allow: `connect-src` only
  needs `graph.microsoft.com` and `login.microsoftonline.com`, not `*.sharepoint.com`. If you fork
  this and add other script/style/connect sources, update it accordingly.

## License

MIT — see [LICENSE](LICENSE).
