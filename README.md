# Rocketlane Invoice Access Manager App

A Rocketlane custom app for importing PDF invoices and enforcing role-aware visibility.

This implementation is built around Rocketlane custom app patterns (`window.rliSdk.init`, `client.data.get`) and supports administrator and restricted-user flows in a single UI.

## Core capabilities

- Import PDF invoices directly inside the custom app.
- Extract email addresses from uploaded PDFs using PDF.js.
- Assign a **Project Manager email** to each invoice (must be present in extracted PDF emails).
- Invoice visibility rules:
  - **Admin** users can view all imported invoices.
  - **Collaborator / Expert Advisor (and other non-admin users)** can only view invoices where `projectManagerEmail` matches their Rocketlane sign-in email.
- Admin-only diagnostics tab:
  - Error/warning logs
  - Suggested remediation guidance per log item

## Rocketlane docs references used

- App development process  
  https://developer.rocketlane.com/v1.3/docs/app-development-process
- Configure app index  
  https://developer.rocketlane.com/v1.3/docs/configure-the-app-index-2
- Runtime context reference  
  https://developer.rocketlane.com/v1.3/docs/runtime-context-reference
- Data method  
  https://developer.rocketlane.com/v1.3/docs/data-method-1

## Project files

- `index.js` — Rocketlane manifest/widget config
- `dist/index.html` — app shell
- `dist/styles.css` — app styling
- `dist/app.js` — invoice import, access control, diagnostics logic
- `dist/icon.svg` — app icon
- `scripts/package-app.sh` — zip packaging helper

## How access control works

1. App initializes context from Rocketlane:
   - account
   - user
   - project
2. User email and role are inferred from runtime data.
3. On import:
   - PDF text is parsed
   - emails are extracted
   - selected PM email must exist in extracted list
4. On view:
   - Admin: sees all invoices
   - Non-admin: sees only invoices assigned to their email

## Local preview

Open `dist/index.html` directly in a browser.

You can emulate Rocketlane context with query params:

```text
dist/index.html?account=Acme&project=Onboarding&email=admin@acme.com&role=admin
```

Example restricted-user preview:

```text
dist/index.html?account=Acme&project=Onboarding&email=pm@acme.com&role=collaborator
```

## Packaging for upload

```bash
npm run package:zip
```

Output:

- `artifacts/rocketlane-invoice-access-manager-app.zip`

## Validation

Run syntax checks:

```bash
npm run check:js
```

## Notes

- PDF extraction uses CDN-hosted PDF.js (`cdnjs`).
- In this implementation, invoice data is persisted in browser local storage scoped by account.
