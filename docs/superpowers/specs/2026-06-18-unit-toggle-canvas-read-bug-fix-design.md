# Design: Unit Toggle, Canvas Read, Apply Bug Fix

**Date:** 2026-06-18  
**Scope:** index.html only — no new files, no build step

---

## Features

### 1. px / cm Unit Toggle

A small `px | cm` pill toggle in the header (right-aligned next to the title). One global `let unit = 'px'` variable drives all dimension display throughout the UI.

**What switches when unit changes:**
- Preset card dimension labels: px shows `${w} × ${h} px`; cm shows `${(wPts * 2.54 / 72).toFixed(1)} × ${(hPts * 2.54 / 72).toFixed(1)} cm`
- Custom input section labels: "רוחב (פיקסל)" / "גובה (פיקסל)" ↔ "רוחב (ס״מ)" / "גובה (ס״מ)"
- Custom input placeholders update to match unit
- Custom input pre-filled values (from live canvas read) re-convert to new unit

**cm conversion rules:**
- Display (preset cards): `pts * 2.54 / 72` — read-only, no floating-point risk
- Custom input write path (cm → API): `Math.round(cm / 2.54 * 72)` — rounding eliminates InvalidArgument risk
- px write path unchanged: `Math.round(px * 0.75)`

**Persistence:** Not persisted. Defaults to `px` on each add-in open.

---

### 2. Current Canvas Pre-fill

On `Office.onReady`, after building preset cards, make one `PowerPoint.run` call:

```js
const pageSetup = context.presentation.pageSetup;
pageSetup.load('slideWidth,slideHeight');
await context.sync();
// convert pts to active unit and populate customWidth / customHeight inputs
```

- Success: custom inputs show the live slide dimensions in the active unit (actual values, not placeholders).
- Failure (API unavailable, timeout): inputs remain empty with placeholder text. No error shown to user — silent degradation.
- Unit toggle: when user switches unit, re-convert the stored pts value and update the inputs in real time. The raw pts are stored in a module-level variable `currentSlidePts = { w, h }`.

---

### 3. Apply Bug Fix — Snappy, No Pre-read

**Root cause (most likely):** When the slide is already at the target dimensions, `context.sync()` hangs or silently no-ops, so `showSuccess` never fires and nothing visible happens.

**Fix:**
- Fire `pageSetup.slideWidth/Height` + one `context.sync()` immediately — no pre-read, one round-trip only.
- Instant visual feedback: button shows a spinner the moment the user clicks.
- `Promise.race` against a **3-second timeout**. If the API doesn't resolve in 3s, show `"שגיאה: הפעולה לא הושלמה — נסה שוב"` and re-enable the button.
- On success: show `✓ name`, re-enable button.
- On error: show the error message, re-enable button.

No pre-check, no extra sync — the 3s timeout is the only safety net.

---

### 4. Code Simplification

After all features are implemented and verified working, run a `/simplify` pass over the full `index.html`. Goal: remove redundancy, flatten unnecessary nesting, improve naming — without changing any behavior.

---

## Out of Scope

- Saving custom sizes as persistent preset cards (deferred)
- Unit preference persistence across sessions
- Any new files beyond `index.html`
