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

## Gotchas — Office.js PowerPoint API

Hard-won findings from debugging and API research. Read before touching the slide-size code.

### 1. `pageSetup.slideWidth/Height` uses POINTS, not EMU

The API properties `context.presentation.pageSetup.slideWidth` and `slideHeight` take **points** (1 pt = 1/72 inch). Requires **PowerPointApi 1.10**.

EMU is used by older/different Office.js patterns (shapes, etc.) but NOT here. Passing EMU values here would set a ~0.015 mm slide (microscopic) and likely hit minimum-size errors.

Correct conversion: `pts = cm / 2.54 * 72`

### 2. Floating-point cm→pts conversions cause `InvalidArgument`

The standard 16:9 widescreen width is **exactly 960 pts** (13.333… inches). The cm approximation `33.867 cm / 2.54 * 72 = 960.009…` is close but NOT exact, and the API rejects it with `InvalidArgument`.

**Rule:** define `wPts`/`hPts` directly as canonical numbers in each preset. Never derive API values from cm at runtime.

- Standard widescreen: `wPts: 960, hPts: 540`
- Standard portrait/story: `wPts: 540, hPts: 960`
- A4: `wPts: 595.28, hPts: 841.89` (from 210mm/297mm — no clean integer exists)

This is why Story worked (its `wPts` happened to be `540` exactly) while Wide did not (`wPts` computed to `960.009`).

### 3. Access path is `presentation.pageSetup.slideWidth`, not `presentation.slideWidth`

`slideWidth`/`slideHeight` live on the `PageSetup` sub-object, not directly on `Presentation`.

```js
// WRONG — silent no-op or TypeError
context.presentation.slideWidth = 960;

// CORRECT
context.presentation.pageSetup.slideWidth = 960;
context.presentation.pageSetup.slideHeight = 540;
await context.sync();
```

### 4. Check `err.code`, not `err.message`, for error classification

Office.js API errors carry a typed `.code` string (`"InvalidArgument"`, `"AccessDenied"`, `"GeneralException"`). The `.message` is a human-readable description that does not reliably contain the code string.

```js
// WRONG — message text varies, may not include the code word
if (err.message.includes('InvalidArgument')) { ... }

// CORRECT
if (err.code === 'InvalidArgument') { ... }
```

Always check `err.code` first; fall back to `err.message` only as a last resort.

### 5. Manifest must declare `PowerPointApi 1.10` requirement

Without the `<Requirements>` element, Office loads the add-in on ANY PowerPoint version. On versions that don't support 1.10, `pageSetup.slideWidth` is undefined and the sync throws a cryptic error.

```xml
<Requirements>
  <Sets DefaultMinVersion="1.1">
    <Set Name="PowerPointApi" MinVersion="1.10"/>
  </Sets>
</Requirements>
```

### 6. `@types/office-js` vs `@microsoft/office-js` type definition discrepancy

The community-maintained `@types/office-js` package contains `slideWidth`/`slideHeight` type definitions. The official `@microsoft/office-js` package (at the time of research) did not. Both exist at runtime — trust the **actual Office.js runtime** and the official Microsoft docs, not whichever type package happens to be installed.

### 7. Valid range for `slideWidth`/`slideHeight`

PowerPoint enforces these limits (same as the UI dialog):

| Bound | Points | Inches | cm |
|-------|--------|--------|----|
| Minimum | 72 pts | 1 in | 2.54 cm |
| Maximum | 4032 pts | 56 in | 142.24 cm |

The code mirrors these in `PT_MIN` / `PT_MAX`. Values outside this range will throw `InvalidArgument`.
