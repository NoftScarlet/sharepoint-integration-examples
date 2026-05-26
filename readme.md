# SharePoint integration examples

This repository contains two SharePoint Framework (SPFx) samples:

- `sharepoint-web-part`: hosts Apryse WebViewer on a SharePoint page.
- `sharepoint-extension`: adds an **Open in PDFTron** command to SharePoint document libraries and opens the selected file in the WebViewer page.

The deployment is sensitive to SharePoint app catalog permissions, SPFx package versions, and WebViewer static asset versioning. Follow the steps below to avoid mixed-runtime and caching issues.

## Prerequisites

1. Node.js and npm compatible with the SPFx version used by the samples.
   - Tested with Node.js 22.x and npm 10/11.x.
2. SharePoint Framework: `1.21.1`.
3. Gulp CLI:
   ```sh
   npm install --global gulp-cli
   ```
4. Microsoft 365 CLI:
   ```sh
   npm install --global @pnp/cli-microsoft365
   ```
5. PowerShell 7 on macOS.
   - Install it with the official Microsoft instructions: https://learn.microsoft.com/en-us/powershell/scripting/install/install-powershell-on-macos?view=powershell-7.6
   - Other macOS installation methods are deprecated or less reliable for this workflow.
6. PnP PowerShell, if you need to administer SharePoint from PowerShell:
   ```powershell
   Install-Module PnP.PowerShell -Scope CurrentUser
   ```
7. A SharePoint tenant with:
   - A target site, for example `https://<tenant>.sharepoint.com/sites/<site-name>`.
   - A tenant app catalog. In the SharePoint admin center, use **More features** > **Apps** > **App catalog** / **Manage apps**.

## Authentication and permissions

Sign in to Microsoft 365 CLI before running deployment commands:

```sh
m365 login --authType browser
```

If your tenant requires a custom Entra application for CLI/PnP operations, sign in with that app instead:

```sh
m365 login --authType browser --appId <client-id> --tenant <tenant-id>
```

The CLI/PnP app may need delegated SharePoint/Graph permissions, depending on your tenant settings. For app catalog deployment and site administration, we needed permissions equivalent to:

- SharePoint: `AllSites.FullControl`
- Microsoft Graph: `Sites.FullControl.All`, `Directory.Read.All`, `User.Read`

Admin consent may be required.

## Create or verify the target SharePoint site

Create a communication/team site in SharePoint admin center, or verify an existing one:

```powershell
Connect-PnPOnline -Url https://<tenant>-admin.sharepoint.com -Interactive -ClientId <client-id>
Get-PnPTenantSite -Detailed -Filter "Url -like '*/sites/<site-name>'" | Select-Object Url,Title,Template
```

If you need custom script enabled for asset hosting or app behavior, run:

```powershell
Connect-PnPOnline -Url https://<tenant>-admin.sharepoint.com -Interactive -ClientId <client-id>
Set-PnPSite -Identity https://<tenant>.sharepoint.com/sites/<site-name> -NoScriptSite $false
```

## WebViewer static assets

WebViewer loads many runtime files from the `path` configured in the web part. These files must all come from the same `@pdftron/webviewer` package version.

**Do not repeatedly overwrite a generic mutable folder such as `Webviewer/js/lib` during debugging.** SharePoint/browser caching can serve a mix of old and new files, which causes unstable errors such as:

- `Cannot read properties of undefined (reading 'ListStylePresets')`
- 404s for files such as `ui/index-wc.html`, `core/pdf/PDFNetLean.js`, or `ui/chunks/*.js`

Recommended approach:

1. Pin WebViewer to an exact version in `sharepoint-web-part/package.json`, for example:
   ```json
   "@pdftron/webviewer": "11.12.0"
   ```
2. Run `npm install` in `sharepoint-web-part` so `package-lock.json` records the exact version.
3. Upload the complete runtime from:
   ```text
   sharepoint-web-part/node_modules/@pdftron/webviewer/public
   ```
   to a versioned SharePoint folder, for example:
   ```text
   Shared Documents/Webviewer/js/lib-11.12.0
   ```
4. Point the web part `.env` to that immutable versioned folder:
   ```env
   TENANT_ID=<tenant>
   SITE_NAME=<site-name>
   WEBVIEWER_LIB_FOLDER_PATH=Webviewer/js/lib-11.12.0
   FOLDER_URL=Shared Documents
   ```

The WebViewer `path` should include a trailing slash and should continue to use the WebViewer web-component loader (`ui/index-wc.html`). Do not force `uiPath: 'ui/index.html'` in this SPFx scenario; it can make relative asset URLs resolve incorrectly.

### Upload runtime assets with Microsoft 365 CLI

For small uploads, this pattern works:

```sh
web="https://<tenant>.sharepoint.com/sites/<site-name>"
root="sharepoint-web-part/node_modules/@pdftron/webviewer/public"
target="Shared Documents/Webviewer/js/lib-11.12.0"

find "$root" -type f ! -name '*.map' | while IFS= read -r file; do
  rel="${file#$root/}"
  dir="${rel%/*}"
  folder="$target"
  if [ "$dir" != "$rel" ]; then
    folder="$target/$dir"
    m365 spo folder add --webUrl "$web" --parentFolderUrl "${folder%/*}" --name "${folder##*/}" --ensureParentFolders true --output none >/dev/null 2>&1 || true
  fi
  m365 spo file add --webUrl "$web" --folder "$folder" --path "$file" --overwrite true --output none
done
```

For large WebViewer runtimes, direct SharePoint REST upload with retries is more reliable than repeatedly calling `m365 spo file add` or PnP upload. SharePoint can transiently return 503 during large batch uploads, so add retry logic if you automate this step.

After uploading, verify key files exist in the versioned folder:

- `ui/index-wc.html`
- `ui/webviewer-ui.min.js`
- `ui/chunks/chunk.theme-light.js`
- `ui/chunks/theme-light.chunk.css`
- `core/webviewer-core.min.js`
- `core/pdf/PDFNetLean.js`
- `core/pdf/lean/PDFNetCWasm.js`
- `core/pdf/lean/PDFNetCWasm.br.wasm`

Also verify folder counts against the local `public` runtime if possible. A complete WebViewer `11.12.0` runtime has hundreds of non-map files.

## SharePoint web part setup

1. Go to the web part folder:
   ```sh
   cd sharepoint-web-part
   npm install
   ```
2. Create `.env`:
   ```env
   TENANT_ID=<tenant>
   SITE_NAME=<site-name>
   WEBVIEWER_LIB_FOLDER_PATH=Webviewer/js/lib-11.12.0
   FOLDER_URL=Shared Documents
   ```
3. Update `config/serve.json` so `initialPage` points to your SharePoint site/page.
4. For local development:
   ```sh
   gulp trust-dev-cert
   gulp serve
   ```

## Deploy the web part

1. Increment `solution.version` in `sharepoint-web-part/config/package-solution.json` whenever changing `.env` values or web part code.
2. Build and package:
   ```sh
   npx gulp clean
   npx gulp bundle --ship
   npx gulp package-solution --ship
   ```
3. Upload the package to the tenant app catalog:
   ```sh
   m365 spo app add --filePath sharepoint/solution/webviewer.sppkg --overwrite --appCatalogScope tenant
   ```
4. Deploy it:
   ```sh
   m365 spo app deploy --id <app-catalog-item-id> --appCatalogScope tenant --skipFeatureDeployment
   ```
5. Install or upgrade it on the target site:
   ```sh
   m365 spo app install --siteUrl https://<tenant>.sharepoint.com/sites/<site-name> --id <product-id> --appCatalogScope tenant
   # or, after a version bump:
   m365 spo app upgrade --siteUrl https://<tenant>.sharepoint.com/sites/<site-name> --id <product-id> --appCatalogScope tenant
   ```
6. Create or edit a SharePoint page, add the WebViewer web part, and publish the page.

## SharePoint extension setup

1. Go to the extension folder:
   ```sh
   cd sharepoint-extension
   npm install
   ```
2. Create `.env`:
   ```env
   SHAREPOINT_SITE_URL=https://<tenant>.sharepoint.com/sites/<site-name>
   SITE_PAGE=WebViewer.aspx
   ```
3. For local debugging, run `gulp serve` and enable the SPFx debug script when prompted.

## Deploy the extension

1. Build and package:
   ```sh
   npx gulp clean
   npx gulp bundle --ship
   npx gulp package-solution --ship
   ```
2. Upload and deploy the `.sppkg` from `sharepoint-extension/sharepoint/solution` in the tenant app catalog.
3. Install the extension app on the target SharePoint site.
4. Go to a document library. Selecting or right-clicking a single document should show **Open in PDFTron**.

## Troubleshooting

### App does not appear in “Add an app”

- Confirm the tenant app catalog exists.
- Confirm the `.sppkg` was uploaded to **Manage apps** and deployed.
- Confirm the user has permission to install apps on the target site.
- If needed, use Microsoft 365 CLI to install or upgrade the app instead of the UI.

### `AADSTS65002` or Graph consent errors

Avoid unnecessary Graph calls in the web part. The current sample can use `this.context.pageContext.user` for the current SharePoint user instead of `AadHttpClient`/Graph. If you reintroduce Graph APIs, add the required `webApiPermissionRequests` and grant admin consent.

### WebViewer loads partially or throws `ListStylePresets`

This almost always means the WebViewer loader, core, UI, or chunk files are from different versions, or SharePoint/browser cache is serving stale files.

Fix:

1. Pin `@pdftron/webviewer` to an exact version.
2. Upload the complete `node_modules/@pdftron/webviewer/public` runtime to a new versioned folder.
3. Update `.env` to the new folder.
4. Bump the SPFx package version.
5. Rebuild, redeploy, and upgrade the site app.
6. Open the page in a private window or with a cache-busting query string.

### Missing files under `ui/chunks` or `core/pdf`

Upload the full WebViewer `public` folder, not just selected top-level files. WebViewer lazily loads many worker, wasm, CSS, locale, and chunk files after initial startup.

### Bad asset URLs such as `lib../core/...` or `libwebviewer-ui.min.js`

Ensure the WebViewer `path` ends with `/`. Do not force `ui/index.html` for the SPFx web-component flow.

## Notes from a verified deployment

The sample was verified with this pattern:

- WebViewer package pinned to `11.12.0`.
- Full WebViewer runtime uploaded to `Shared Documents/Webviewer/js/lib-11.12.0`.
- Web part `.env` pointed to `Webviewer/js/lib-11.12.0`.
- Web part solution version bumped and deployed as a new package version.
- Site app upgraded after deployment.
