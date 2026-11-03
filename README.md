# Florida Shift Pass Down Excel Add-in

Starter Excel task-pane add-in that:

1. Reads the Shift value from `Shift Pass Down!C3`.
2. Uses Microsoft Entra ID SSO through MSAL Nested App Authentication (NAA).
3. Creates an item in the verified `Florida Shift Passdown Queue` list.
4. Lets the existing Power Automate flow handle the backend work.

## Verified SharePoint target

- Site: `https://cmdmedical.sharepoint.com/Production`
- Site ID: `cmdmedical.sharepoint.com,f11b7983-dd20-44f2-b980-0f666ad3d14a,b73c13bc-3856-411b-be45-63c699a7597b`
- List: `Florida Shift Passdown Queue`
- List ID: `1070c43b-933d-4b0e-ad33-2ee50166ade9`
- Current writable business column: `Title`

## Entra registration

Create a single-tenant app registration and record its Application (client) ID and Directory (tenant) ID.

Under Authentication, add Single-page application redirect URIs:

- `brk-multihub://localhost:3000`
- `https://localhost:3000/taskpane.html`

Under API permissions, add Microsoft Graph delegated permission:

- `Sites.ReadWrite.All`

Grant admin consent if required by tenant policy. Do not create or store a client secret in this browser-based Office add-in.

## Configure

Edit `src/config.js` and replace:

- `REPLACE_WITH_ENTRA_APPLICATION_CLIENT_ID`
- `REPLACE_WITH_ENTRA_TENANT_ID`

Confirm that the Shift cell is `Shift Pass Down!C3`; change `worksheetName` or `shiftCell` if needed.

## Run locally

```bash
npm install
npm start
```

The development server runs at `https://localhost:3000`.

Sideload `manifest.xml` into Excel, open the add-in, and select **Submit Shift Pass Down**.

## Production deployment

Host the `dist` files on an internal HTTPS web host. Replace all `https://localhost:3000` and `brk-multihub://localhost:3000` values in the Entra registration and manifest with the production origin. Then deploy the manifest through Microsoft 365 integrated apps or your established Office Add-in deployment process.

## Important design note

The queue list presently exposes only `Title` as a writable business column. This starter therefore creates a trigger item with a descriptive Title. If the flow expects workbook data in the queue item itself, add the required SharePoint columns first and map them in `readSubmissionFromWorkbook()`.
