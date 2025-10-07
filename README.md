# Outlook Add-in: Send to Workbench

## Overview
This project is a cross-platform Outlook Add-in (Web, Mac, Windows) that adds a "Send to Workbench" button under the "Workbench group" for selected emails. It uses Fluent UI for all user experience components.

## Project Structure
- `src/manifest.xml`: Add-in manifest
- `src/components/SendToWorkbenchButton.tsx`: Fluent UI button
- `src/commands/SendToWorkbenchCommand.ts`: Command logic
- `src/App.tsx`: Main React app
- `public/assets/`: Add-in icons (16x16, 32x32, 80x80 PNG)

## Prerequisites
- [Node.js](https://nodejs.org/) (v16+ recommended)
- [npm](https://www.npmjs.com/)
- [Office Add-in CLI](https://www.npmjs.com/package/office-addin-cli) (optional, for sideloading)
- Outlook (Web, Mac, or Windows)

## Setup Instructions

### 1. Clone the repository
```sh
git clone <repo-url>
cd OutlookAddin-scaffolding
```

### 2. Install dependencies
```sh
npm install
```

### 3. Set up HTTPS for local development
This is required for Outlook Web and Mac.
```sh
npx office-addin-dev-certs install
```
This will generate and trust a local HTTPS certificate.

### 4. Start the development server
```sh
npm start
```
By default, the app runs at `https://localhost:3001`.

### 5. Verify icons and app are served
Open in your browser:
- `https://localhost:3001/index.html`
- `https://localhost:3001/assets/icon-32.png`

If you see a security warning, proceed and trust the certificate.

### 6. Sideload the add-in in Outlook
1. Open Outlook Web (https://outlook.office.com/)
2. Go to **Settings > Manage add-ins** (or **My Add-ins**)
3. Click **Add a custom add-in > Add from file**
4. Select `src/manifest.xml`
5. The add-in will appear in your Outlook ribbon

## Troubleshooting

### Port Already in Use
If you see `EADDRINUSE: address already in use :::3001`:
```sh
npm run kill-ports
```
Then restart the dev server:
```sh
npm start
```

### SSL Certificate Warnings
- If you see `ERR_CERT_AUTHORITY_INVALID`, run:
  ```sh
  npx office-addin-dev-certs uninstall
  npx office-addin-dev-certs install
  ```
- Accept the certificate in your browser when prompted.

### Manifest Errors
- Ensure `src/manifest.xml` uses the correct schema (see sample in repo)
- The `Id` must be a valid GUID and an attribute on `<OfficeApp>`
- All URLs should use `https://localhost:3001` (or your chosen port)
- Icons must exist in `public/assets/` and be valid PNG files

### Sideloading Rejected by Exchange
- Ensure you are using the correct manifest schema (not the old MailApp format)
- Only one `Id` attribute on the root `<OfficeApp>`
- If you can sideload other add-ins but not this one, double-check the manifest structure

## Useful Scripts
- `npm run kill-ports` — Frees up common dev ports (3000, 3001, 8080, 8000)

## Resources
- [Office Add-ins documentation](https://docs.microsoft.com/office/dev/add-ins/)
- [Fluent UI documentation](https://developer.microsoft.com/en-us/fluentui)

---

**If you have issues, check the Troubleshooting section or ask your team lead for help!**

## Key Parts
### Fluent UI Button
Implemented in `
