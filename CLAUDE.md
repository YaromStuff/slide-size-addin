# CLAUDE.md

Guidance for Claude Code when working in this repository.

## Project Overview

PowerPoint Office Web Add-in ("כלי שקפים") for one-click slide resizing. Vanilla HTML/CSS/JS — no build step, no npm, no frameworks. Deployed via GitHub Pages.

## Architecture

Everything lives in a single `index.html`. No bundler, no server — open in a browser to inspect the UI (Office.js fails outside PowerPoint, but the layout renders fully).

**Data flow:**
1. `PRESETS` array defines all slide sizes — each preset has:
   - `wPts`/`hPts` — API values in points (sole source of truth). **No separate `w`/`h` fields.** px labels, cm labels, and preview aspect ratio are all derived at render time from `wPts`/`hPts`.
   - `canRotate` / `nameRotated` — whether the card shows a rotate button
2. `unit` (`'px'` | `'cm'`) — global toggle, defaults to `'px'`. `setUnit()` re-renders all cards and the custom section.
3. `currentSlidePts = { w, h }` — populated on load by reading `pageSetup.slideWidth/Height`; refreshed after every successful apply. Drives the custom-input pre-fill.
4. `cardState` object tracks `isRotated` per preset ID.
5. Clicking a card selects it; clicking "החל" → `handleApply()` → `applySlideSize(widthPts, heightPts, name)` → `PowerPoint.run()` context with a 3-second `Promise.race` timeout.
6. `Office.onReady()` (async) builds all cards, wires the unit toggle, reads the current slide size, then wires remaining event listeners.

**Units:** `pageSetup.slideWidth/Height` use **points** (1 pt = 1/72 inch). **Points ≠ pixels** — see Gotcha #1. Requires PowerPointApi 1.10.

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
{ id: 'unique-id', nameDefault: 'שם', nameRotated: 'שם מסובב', wPts: 960, hPts: 540, canRotate: true }
```

- `wPts`/`hPts` — exact API values in points. Use hardcoded integers, never compute from cm at runtime (see Gotcha #2). **Single source of truth** — the card's px label, cm label, and preview aspect ratio are all derived from these (see Gotcha #12). Do NOT add separate `w`/`h` display-pixel fields.
- Set `canRotate: false` and `nameRotated: null` for square or fixed-orientation formats.

---

## Gotchas — Office.js PowerPoint API

Hard-won findings from debugging and API research. **Read before touching slide-size code. Update this section whenever new findings are made.**

### 1. `pageSetup.slideWidth/Height` uses POINTS, not EMU — and points ≠ pixels

`context.presentation.pageSetup.slideWidth` and `slideHeight` take **points** (1 pt = 1/72 inch).

EMU is used elsewhere in Office.js (shapes, etc.) but NOT here. Passing EMU values would set a microscopic slide and throw `InvalidArgument`.

**Points and pixels are different units.** A point is a physical unit (1/72 inch). A pixel has no fixed physical size — it depends on DPI. At 96 DPI (web/Windows standard):

```
96 px = 1 inch = 72 pt
→ 1 pt = 1.3333 px
→ 1 px = 0.75 pt
```

So the *same slide* described in different units:

| Unit | Wide preset |
|------|------------|
| Points (stored by PowerPoint) | 1440 × 810 |
| Pixels at 96 DPI (add-in label) | 1920 × 1080 |
| cm (physical, DPI-independent) | 50.80 × 28.58 |

**cm and inches are the most trustworthy labels** — they are DPI-independent and match what PowerPoint's Slide Size dialog shows. The px label assumes 96 DPI export; if you export a PNG from PowerPoint at a different resolution, the pixel count will differ from what the add-in displays.

Conversions: `pts = cm / 2.54 * 72` · `pts = px * 0.75` (at 96 DPI) · `pts = mm / 25.4 * 72`

### 2. Floating-point conversions cause `InvalidArgument` — always hardcode wPts/hPts

`33.867 cm / 2.54 * 72 = 960.009…` — the API rejects this with `InvalidArgument`. Define `wPts`/`hPts` as exact constants:

| Preset | wPts | hPts | Notes |
|--------|------|------|-------|
| Wide (16:9) | 1440 | 810 | Full-HD pixel canvas: 1920×1080 px (20 × 11.25 in). Not PowerPoint's default 13.33 in / 960×540 — see Gotcha #12. |
| Story/Portrait | 810 | 1440 | Rotated wide: 1080×1920 px |
| Square | 810 | 810 | 1080×1080 px |
| A4 | 540 | 780 | PowerPoint's internal A4, NOT ISO (see Gotcha #9) |
| A3 | 756 | 1008 | PowerPoint's internal A3, NOT ISO |

Never derive `wPts`/`hPts` from cm/px at runtime in `handleApply`.

### 3. Access path is `presentation.pageSetup.slideWidth`, not `presentation.slideWidth`

`slideWidth`/`slideHeight` live on the `PageSetup` sub-object:

```js
// WRONG — silent no-op
context.presentation.slideWidth = 1440;

// CORRECT
context.presentation.pageSetup.slideWidth = 1440;
context.presentation.pageSetup.slideHeight = 810;
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

Code constants: `PT_MIN = 36`, `PT_MAX = 4032`. Values outside this range throw `InvalidArgument`.

**Watch out:** the API minimum is **72 pts** (96 px / 1 in), but `PT_MIN` is set to 36 (= 48px × 0.75). Any custom input between 48–95 px will pass the in-code guard but get rejected by the API with `InvalidArgument`. If you touch the custom input validation, raise `PT_MIN` to 72 to match the API floor.

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

### 12. All UI dimension labels must derive from `wPts`/`hPts` — never store separate display-px

Storing nominal "design pixels" (e.g. `w: 1920, h: 1080` for widescreen) alongside `wPts: 960, hPts: 540` makes the px and cm labels disagree: the card showed `1920 × 1080 px` but cm mode showed `33.9 × 19.1 cm`, and `1920 px ≠ 33.9 cm` (1920 px = 50.8 cm). It also broke the live "current size" indicator (which reads real pts → 1280 × 720 px) and made same-ratio applies look like no-ops, since the slide was already 960×540 but the card advertised a different-looking 1920×1080.

For non-proportional presets it also distorted the preview: A4's ISO design ratio (2480×3508 = 0.707) differs from PowerPoint's actual A4 (540×780 = 0.692).

**Rule:** `wPts`/`hPts` is the single source of truth. Derive on the fly:
- px label: `ptsToPx(pts) = Math.round(pts / 0.75)`
- cm label: `ptsToCm(pts) = pts * 2.54 / 72`
- preview aspect ratio: `calcPreview(wPts, hPts)`

So every label stays internally consistent: e.g. wide (`1440×810 pt`) shows `1920 × 1080 px` / `50.80 × 28.58 cm`, and `1920 px × 0.75 = 1440 pt`, `1440 pt × 2.54/72 = 50.8 cm` all agree.

cm labels are displayed to **2 decimals** (`.toFixed(2)`) — this matches PowerPoint's own Slide Size dialog exactly (e.g. A4 = `19.05 × 27.52 cm`). One decimal (`19.1 × 27.5`) drifts up to 0.05 cm from PowerPoint's value and breaks trust. px labels need no decimals: every preset's points divide cleanly by 0.75, so px is exact. All preset point values are also whole EMUs (12700 EMU/pt), so PowerPoint stores them with zero quantization — the slide is exactly the size shown.

Caveat: a **custom px** entry snaps to the nearest whole point (`Math.round(px * 0.75)`), so it can shift by up to ~1px on round-trip (point granularity ≈ 1.33 px). Preset px values are all exact because they were chosen to map to integer points.

### 13. Refresh the live "current size" indicator after every successful apply

`applySlideSize` must update `currentSlidePts = { w: widthPts, h: heightPts }` and call `updateCustomSection()` on success. Otherwise the pre-filled custom inputs keep showing the slide's size from page load, not the size just applied.
