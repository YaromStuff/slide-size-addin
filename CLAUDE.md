# CLAUDE.md

Guidance for Claude Code when working in this repository.

## Project Overview

PowerPoint Office Web Add-in ("כלי שקפים") for one-click slide resizing. Vanilla HTML/CSS/JS — no build step, no npm, no frameworks. Deployed via GitHub Pages.

## Architecture

Everything lives in a single `index.html`. No bundler, no server — open in a browser to inspect the UI (Office.js fails outside PowerPoint, but the layout renders fully).

**Data flow:**
1. `PRESETS` array defines all slide sizes — each preset has:
   - `wPts`/`hPts` — API values in points (source of truth for `pageSetup.slideWidth/Height`)
   - `w`/`h` — pixel dimensions shown in the card UI and used for preview aspect ratio
   - `canRotate` / `nameRotated` — whether the card shows a rotate button
2. `cardState` object tracks `isRotated` per preset ID
3. Clicking a card selects it; clicking "החל" → `handleApply()` → `applySlideSize(widthPts, heightPts, name)` → `PowerPoint.run()` context
4. `Office.onReady()` builds all cards dynamically via `buildCard()`

**Units:** `pageSetup.slideWidth/Height` use **points** (1 pt = 1/72 inch). Requires PowerPointApi 1.10.

## Key Constraints

- **No `alert()` for errors** — use `showError()` which writes to `#status` element
- All logic must run inside `Office.onReady()` — PowerPoint API calls outside it will throw
- `manifest.xml` points to GitHub Pages URLs — serves PowerPoint Desktop via HTTPS

## Validate manifest.xml

```bash
xmllint --noout manifest.xml
```

## Deployment (GitHub Pages)

Push `index.html` and `assets/` to the `main` branch of the GitHub Pages repo. The manifest already points to the live URL. No build step needed.

## Deployment (Mac — local file hosting)

1. Copy add-in files:
   ```bash
   mkdir -p ~/Documents/MyAddins/slide-size-addin/assets
   cp index.html ~/Documents/MyAddins/slide-size-addin/
   cp assets/icon-32.png ~/Documents/MyAddins/slide-size-addin/assets/
   ```
2. Copy manifest:
   ```bash
   cp manifest.xml ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
   ```
3. Restart PowerPoint → Insert → My Add-ins → **כלי שקפים**

## Adding a New Preset

Add an object to the `PRESETS` array in `index.html`:

```js
{ id: 'unique-id', nameDefault: 'שם', nameRotated: 'שם מסובב', w: 1920, h: 1080, wPts: 960, hPts: 540, canRotate: true }
```

- `wPts`/`hPts` — exact API values in points. Use hardcoded integers, never compute from cm at runtime (see Gotcha #2).
- `w`/`h` — pixel display values shown on the card (also drive the preview shape aspect ratio).
- Set `canRotate: false` and `nameRotated: null` for square or fixed-orientation formats.

---

## Gotchas — Office.js PowerPoint API

Hard-won findings from debugging and API research. **Read before touching slide-size code. Update this section whenever new findings are made.**

### 1. `pageSetup.slideWidth/Height` uses POINTS, not EMU

`context.presentation.pageSetup.slideWidth` and `slideHeight` take **points** (1 pt = 1/72 inch).

EMU is used elsewhere in Office.js (shapes, etc.) but NOT here. Passing EMU values would set a microscopic slide and throw `InvalidArgument`.

Conversions: `pts = cm / 2.54 * 72` · `pts = px * 0.75` (at 96 DPI) · `pts = mm / 25.4 * 72`

### 2. Floating-point conversions cause `InvalidArgument` — always hardcode wPts/hPts

`33.867 cm / 2.54 * 72 = 960.009…` — the API rejects this with `InvalidArgument`. Define `wPts`/`hPts` as exact constants:

| Preset | wPts | hPts | Notes |
|--------|------|------|-------|
| Wide (16:9) | 960 | 540 | Exactly 13.333 in wide |
| Story/Portrait | 540 | 960 | Rotated wide |
| Square | 540 | 540 | |
| A4 | 540 | 780 | PowerPoint's internal A4, NOT ISO (see Gotcha #9) |
| A3 | 756 | 1008 | PowerPoint's internal A3, NOT ISO |

Never derive `wPts`/`hPts` from cm/px at runtime in `handleApply`.

### 3. Access path is `presentation.pageSetup.slideWidth`, not `presentation.slideWidth`

`slideWidth`/`slideHeight` live on the `PageSetup` sub-object:

```js
// WRONG — silent no-op
context.presentation.slideWidth = 960;

// CORRECT
context.presentation.pageSetup.slideWidth = 960;
context.presentation.pageSetup.slideHeight = 540;
await context.sync();
```

### 4. Check `err.code`, not `err.message`, for error classification

`.code` is a reliable typed string (`"InvalidArgument"`, `"AccessDenied"`, `"GeneralException"`). `.message` is locale-dependent and unreliable for branching.

```js
if (err.code === 'InvalidArgument') { ... }      // correct
if (err.message.includes('InvalidArgument')) { } // wrong
```

### 5. Manifest must declare `PowerPointApi 1.10` requirement

Without `<Requirements>`, the add-in loads on any version. On versions without 1.10, `pageSetup.slideWidth` is undefined and sync throws cryptic errors.

```xml
<Requirements>
  <Sets DefaultMinVersion="1.1">
    <Set Name="PowerPointApi" MinVersion="1.10"/>
  </Sets>
</Requirements>
```

### 6. Valid range for `slideWidth`/`slideHeight`

| Bound | Points | Pixels (96 DPI) | Inches |
|-------|--------|-----------------|--------|
| Minimum | 72 pts | 96 px | 1 in |
| Maximum | 4032 pts | 5376 px | 56 in |

Code constants: `PT_MIN = 36` (custom input uses px * 0.75, so 48px min → 36 pts), `PT_MAX = 4032`. Values outside this range throw `InvalidArgument`.

### 7. PowerPointApi 1.10 is NOT available on perpetual/LTSC Office

Only available on:
- **Microsoft 365 subscription** — Windows Version 2601+ (Build 19610+), Mac Version 16.105+
- **Office on the web**

**NOT** available on Office 2019, 2021, 2024 volume-licensed (LTSC) builds. The manifest `<Requirements>` gate hides the add-in entirely on unsupported versions. Users who report the add-in not appearing in their list likely have a perpetual license — this cannot be fixed.

### 8. PowerPoint "paper" sizes differ from ISO paper dimensions

PowerPoint's built-in A4/A3 presets use non-standard dimensions. Always verify from PowerPoint's own slide size dialog:

| Preset | PowerPoint (actual) | ISO paper |
|--------|---------------------|-----------|
| A4 | 19.05 × 27.52 cm → **540 × 780 pts** | 21 × 29.7 cm → 595 × 842 pts |
| A3 | 26.67 × 35.56 cm → **756 × 1008 pts** | 29.7 × 42 cm → 842 × 1191 pts |

Using ISO dimensions causes `InvalidArgument`. Always read from PowerPoint's dialog for any paper/print preset.

### 9. Error debugging: `debugInfo` and `context.trace()`

`OfficeExtension.Error` exposes:
- `error.debugInfo` — serialize with `JSON.stringify(error.debugInfo)` for structured logging
- `error.traceMessages` — populated from `context.trace('label')` calls before `sync()`; identifies which queued command failed

```js
context.trace('before pageSetup write');
pageSetup.slideWidth = wPts;
pageSetup.slideHeight = hPts;
await context.sync();
// on failure: error.traceMessages includes 'before pageSetup write'
```

Use `error instanceof OfficeExtension.Error` to distinguish API errors from plain JS exceptions.

### 10. Performance: one `sync()`, one `pageSetup` reference

Assign `const pageSetup = context.presentation.pageSetup` once, set both width and height, then call one `context.sync()`. Minimize proxy object creation and `sync()` calls — each `sync()` is a round-trip, especially costly on Office on the web.

### 11. `resid` attribute values in manifest.xml must be ≤ 32 characters

Longer `resid` strings are silently truncated or cause manifest load errors. Keep all resource IDs short.
