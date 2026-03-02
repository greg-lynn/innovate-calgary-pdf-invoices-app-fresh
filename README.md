# Rocketlane Customer Bulletin Board App

A Rocketlane custom app that provides a Sticky Notes-inspired bulletin board for customer-facing updates on the **Customer Portal home view**.

This app is designed to follow Rocketlane Marketplace extension patterns and can be added via Rocketlane placeholders.

## What this app includes

- Sticky-note style board inspired by Windows Sticky Notes:
  - Color-coded notes
  - Pin/unpin notes
  - Quick note creation and deletion
  - Search and filter
  - Automatic save per board scope
- Gmail-like formatting actions in the editor:
  - **Bold**, *Italic*, Underline
  - **Bulleted lists**
  - **Numbered lists**
  - Clear formatting
- Rocketlane runtime support:
  - Initializes through `window.rliSdk.init({})` when available
  - Reads contextual data from Rocketlane (`account`, `user`, `project`) when available
  - Falls back to local preview mode when run outside Rocketlane

## Rocketlane docs used as implementation reference

The following Rocketlane Marketplace developer docs were reviewed and mapped into this implementation:

- App development process  
  https://developer.rocketlane.com/v1.3/docs/app-development-process
- Configure the app index  
  https://developer.rocketlane.com/v1.3/docs/configure-the-app-index-2
- Include placeholders for app deployment  
  https://developer.rocketlane.com/v1.3/docs/include-placeholders-for-app
- Data method  
  https://developer.rocketlane.com/v1.3/docs/data-method-1
- Interface methods  
  https://developer.rocketlane.com/v1.3/docs/interface-method
- Events method  
  https://developer.rocketlane.com/v1.3/docs/events-method
- Runtime context reference  
  https://developer.rocketlane.com/v1.3/docs/runtime-context-reference

## Sticky Notes reference used for UI behavior

Windows Sticky Notes behavior was used as product inspiration for:

- color note variants
- always-visible pinned notes
- lightweight quick-note workflow
- autosave style editing

Reference page used while designing:

- https://en.wikipedia.org/wiki/Sticky_Notes

## Files

- `index.js` — Rocketlane app manifest (widget config)
- `dist/index.html` — app shell
- `dist/styles.css` — sticky note + editor styles
- `dist/app.js` — app logic and Rocketlane runtime integration
- `dist/icon.svg` — app icon

## Add this app to Rocketlane

1. Create or open a Rocketlane app project (using `rli init` if needed).
2. Copy these files into the project root:
   - `index.js`
   - `dist/` directory
3. Build and deploy:
   - `rli build`
   - `rli deploy`
4. In Rocketlane:
   - Go to Customer Portal Builder.
   - Click **Add Sections**.
   - Select **Customer Bulletin Board** (placeholder: `customer_portal_widget`).

The manifest also includes `left_nav` so teams can open the board from workspace navigation.

## Local preview

Open `dist/index.html` directly in a browser.

To load seeded sample notes:

- `dist/index.html?demo=1`

To emulate scope values in preview:

- `dist/index.html?demo=1&account=Acme%20Corp&project=Customer%20Home`

## Screenshots

Screenshots generated for review:

- `docs/screenshots/bulletin-board-overview.png`
- `docs/screenshots/bulletin-board-lists-and-formatting.png`

## Packaging for upload

This repo includes a script that creates a minimal ZIP bundle (manifest + `dist/`) suitable for uploading/adding as a custom app:

```bash
npm run package:zip
```

Output:

- `artifacts/rocketlane-bulletin-board-app.zip`
