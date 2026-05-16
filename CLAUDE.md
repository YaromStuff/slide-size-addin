# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PowerPoint Office Web Add-in ("כלי שקפים") for one-click slide resizing. Vanilla HTML/CSS/JS — no build step, no npm, no frameworks. Deployed via GitHub Pages.

## Architecture

Everything lives in a single `index.html`. There is no bundler, no server, and no local dev server — open `index.html` directly in a browser to inspect the UI (Office.js will fail to initialize outside PowerPoint, but the layout renders fully).

**Data flow:**
1. `PRESETS` array defines all slide sizes — each preset has `wPts`/`hPts` (API source of truth), `wCm`/`hCm` (display), `w`/`h` (preview aspect ratio only)
2. `cardState` object tracks `isRotated` per preset ID
3. Clicking a card selects it; clicking "החל" → `handleApply()` → `applySlideSize(widthPts, heightPts, name)` → `PowerPoint.run()` context
4. `Office.onReady()` builds all cards dynamically via `buildCard()`

**Units:** `pageSetup.slideWidth` and `pageSetup.slideHeight` use **points** (1 pt = 1/72 inch). Requires PowerPointApi 1.10. Convert cm→pts: `cm / 2.54 * 72`. Standard widescreen = 960 × 540 pts.

## Key Constraints

- **No `alert()` for errors** — use `showError()` which writes to `#status` element
- All logic must run inside `Office.onReady()` — direct PowerPoint API calls outside it will throw
- `manifest.xml` currently points to GitHub Pages URLs — serves PowerPoint Desktop via HTTPS

## Validate manifest.xml

```bash
xmllint --noout manifest.xml
```

## Deployment (Mac — local file hosting)

The manifest points to `file:///Users/yaromganor/Documents/MyAddins/slide-size-addin/`.

1. Copy add-in files to the target folder:
   ```bash
   mkdir -p ~/Documents/MyAddins/slide-size-addin/assets
   cp index.html ~/Documents/MyAddins/slide-size-addin/
   cp assets/icon-32.png ~/Documents/MyAddins/slide-size-addin/assets/
   ```
2. Copy `manifest.xml` to the Office wef folder:
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
   cp manifest.xml ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
   ```
3. Restart PowerPoint → Insert → My Add-ins → **כלי שקפים**

## Adding a New Preset

Add an object to the `PRESETS` array in `index.html`:

```js
{ id: 'unique-id', nameDefault: 'שם', nameRotated: 'שם מסובב', w: 1234, h: 5678, wCm: 12.3, hCm: 45.6, wPts: 349, hPts: 1293, canRotate: true }
```

`wPts`/`hPts` are the values sent to the Office.js API. Use exact integers for standard sizes (e.g., 960/540 for widescreen). For paper sizes derive from mm: `mm / 25.4 * 72`.

Set `canRotate: false` and `nameRotated: null` for square or fixed-orientation formats.
