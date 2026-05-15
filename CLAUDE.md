# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PowerPoint Office Web Add-in ("כלי שקפים") for one-click slide resizing. Vanilla HTML/CSS/JS — no build step, no npm, no frameworks. Deployed via GitHub Pages.

## Architecture

Everything lives in a single `index.html`. There is no bundler, no server, and no local dev server — open `index.html` directly in a browser to inspect the UI (Office.js will fail to initialize outside PowerPoint, but the layout renders fully).

**Data flow:**
1. `PRESETS` array defines all slide sizes (pixels + EMU)
2. `cardState` object tracks `isRotated` per preset ID
3. Clicking a card or rotate button → `applySlideSize(widthEmu, heightEmu, name)` → `PowerPoint.run()` context
4. `Office.onReady()` builds all cards dynamically via `buildCard()`

**EMU conversion:** `pixels × 12700 = EMU` (assumes 96 DPI). All Office.js slide dimensions use EMU.

**Content detection heuristic:** `slides.items.length > 1` = presentation has content → show `confirm()` before resizing.

## Key Constraints

- **No `alert()` for errors** — use `showError()` which writes to `#status` element
- **`confirm()` is allowed** for the destructive resize confirmation only
- All logic must run inside `Office.onReady()` — direct PowerPoint API calls outside it will throw
- `manifest.xml` uses `file:///` URLs — works only on PowerPoint Desktop (Mac), not PowerPoint Web

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
{ id: 'unique-id', nameDefault: 'שם', nameRotated: 'שם מסובב', w: 1234, h: 5678, canRotate: true }
```

Set `canRotate: false` and `nameRotated: null` for square or fixed-orientation formats.
