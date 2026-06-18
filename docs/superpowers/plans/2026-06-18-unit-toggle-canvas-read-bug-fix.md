# Slide Master — Unit Toggle, Canvas Read, Apply Fix Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a global px/cm unit toggle, pre-fill custom size inputs with the live slide dimensions, and harden the apply action with instant spinner feedback and a 3-second timeout.

**Architecture:** All changes are confined to `index.html`. Two module-level variables (`unit`, `currentSlidePts`) drive unit display and canvas pre-fill. `Office.onReady` becomes async to support the canvas read. `applySlideSize` uses `Promise.race` against a 3-second timeout. A unit toggle pill is added to the header. `updateCardVisual` becomes the single source of truth for all card label rendering.

**Tech Stack:** Vanilla JS, Office.js PowerPointApi 1.10, no build step, no frameworks.

## Global Constraints
- No new files — all changes in `index.html` only
- No floating-point pts passed to the API — always `Math.round()`
- No `alert()` — use `showError()` / `showSuccess()`
- All PowerPoint API calls inside `Office.onReady` or a `PowerPoint.run` context
- RTL layout (`dir="rtl"`) must be preserved
- Requires PowerPointApi 1.10 (existing manifest requirement)

---

### Task 1: Harden `applySlideSize` — spinner + 3-second timeout

**Files:**
- Modify: `index.html` — `applySlideSize` function + CSS `@keyframes spin` + `.apply-btn[aria-busy="true"]::after`

**Interfaces:**
- Produces: `applySlideSize(widthPts, heightPts, presetName): Promise<void>` — same public signature, now with instant spinner and timeout protection

- [ ] **Step 1: Add spinner CSS**

In the `<style>` block, after the existing `.apply-btn:disabled, .apply-btn[aria-busy="true"] { ... }` rule, add:

```css
@keyframes spin {
  to { transform: rotate(360deg); }
}

.apply-btn[aria-busy="true"]::after {
  content: '';
  display: inline-block;
  width: 12px;
  height: 12px;
  border: 2px solid rgba(255, 255, 255, 0.35);
  border-top-color: #fff;
  border-radius: 50%;
  animation: spin 0.6s linear infinite;
  vertical-align: middle;
  margin-inline-start: 8px;
}
```

- [ ] **Step 2: Replace `applySlideSize` with the timeout-guarded version**

Find and replace the entire `async function applySlideSize` block:

```js
const APPLY_TIMEOUT_MS = 3000;

async function applySlideSize(widthPts, heightPts, presetName) {
  const btn = document.getElementById('applyCustomBtn');
  btn.disabled = true;
  btn.setAttribute('aria-busy', 'true');

  const applyPromise = PowerPoint.run(async (context) => {
    const pageSetup = context.presentation.pageSetup;
    pageSetup.slideWidth = widthPts;
    pageSetup.slideHeight = heightPts;
    await context.sync();
  });

  const timeoutPromise = new Promise((_, reject) =>
    setTimeout(() => reject(new Error('TIMEOUT')), APPLY_TIMEOUT_MS)
  );

  try {
    await Promise.race([applyPromise, timeoutPromise]);
    showSuccess(`✓ ${presetName}`);
  } catch (err) {
    if (err.message === 'TIMEOUT') {
      showError('שגיאה: הפעולה לא הושלמה — נסה שוב');
    } else {
      const code = err && err.code ? err.code : '';
      const msg = err && err.message ? err.message : String(err);
      if (code === 'AccessDenied' || code === 'GeneralException' || msg.includes('AccessDenied')) {
        showError('שגיאה: אין גישה לשינוי גודל השקף');
      } else if (code === 'InvalidArgument' || msg.includes('InvalidArgument')) {
        showError('שגיאה: ערכי גודל לא תקינים');
      } else {
        showError(`שגיאה: ${msg || code}`);
      }
    }
  } finally {
    btn.disabled = false;
    btn.removeAttribute('aria-busy');
  }
}
```

- [ ] **Step 3: Verify manually in PowerPoint**

Load the add-in. Click any preset card then "החל". Confirm:
- Spinner appears immediately on the button
- Success message appears after the slide resizes
- Button re-enables after completion

- [ ] **Step 4: Commit**

```bash
git add index.html
git commit -m "fix: harden applySlideSize with 3s timeout and loading spinner"
```

---

### Task 2: Add unit state, conversion helpers, and unit-aware rendering

**Files:**
- Modify: `index.html` — top of `<script>` block (new variables + functions) + `updateCardVisual` + `buildCard` + `handleApply`

**Interfaces:**
- Produces:
  - `let unit = 'px'` — `'px'` | `'cm'`
  - `let currentSlidePts = { w: 0, h: 0 }` — live canvas dimensions in points
  - `ptsToCm(pts: number): number` — `pts * 2.54 / 72`
  - `cmToPts(cm: number): number` — `Math.round(cm / 2.54 * 72)`
  - `pxToPts(px: number): number` — `Math.round(px * 0.75)`
  - `updateAllCardVisuals(): void`
  - `updateCustomSection(): void`

- [ ] **Step 1: Add module-level state and conversion helpers**

After the `const PREVIEW_MAX = 60;` line, add:

```js
let unit = 'px';
let currentSlidePts = { w: 0, h: 0 };

function ptsToCm(pts) { return pts * 2.54 / 72; }
function cmToPts(cm) { return Math.round(cm / 2.54 * 72); }
function pxToPts(px) { return Math.round(px * 0.75); }
```

- [ ] **Step 2: Update `buildCard` to delegate label rendering to `updateCardVisual`**

In `buildCard`, the top line `const { dw, dh } = calcPreview(card.w, card.h);` and the shape's inline style/dimension set are redundant once `updateCardVisual` is called at the end. Simplify:

Find in `buildCard`:
```js
  const previewArea = document.createElement('div');
  previewArea.className = 'card-preview-area';
  const shape = document.createElement('div');
  shape.className = 'preview-shape';
  shape.style.width = dw + 'px';
  shape.style.height = dh + 'px';
  previewArea.appendChild(shape);
  el.appendChild(previewArea);

  // מידע טקסטואלי
  const info = document.createElement('div');
  info.className = 'card-info';
  info.innerHTML = `
    <span class="card-name">${card.nameDefault}</span>
    <span class="card-dimensions">${card.w} × ${card.h} פיקסל</span>
  `;
  el.appendChild(info);
```

Replace with:
```js
  const previewArea = document.createElement('div');
  previewArea.className = 'card-preview-area';
  const shape = document.createElement('div');
  shape.className = 'preview-shape';
  previewArea.appendChild(shape);
  el.appendChild(previewArea);

  const info = document.createElement('div');
  info.className = 'card-info';
  info.innerHTML = '<span class="card-name"></span><span class="card-dimensions"></span>';
  el.appendChild(info);
```

Then remove the `const { dw, dh } = calcPreview(card.w, card.h);` line at the top of `buildCard` (no longer needed here).

And at the very end of `buildCard`, before `return el;`, add:
```js
  updateCardVisual(card, el);
  return el;
```

- [ ] **Step 3: Update `updateCardVisual` to use `unit` and correct pts for rotated state**

Replace the entire `updateCardVisual` function:

```js
function updateCardVisual(card, el) {
  const rotated = cardState[card.id];
  const wPx = rotated ? card.h : card.w;
  const hPx = rotated ? card.w : card.h;
  const wPts = rotated ? card.hPts : card.wPts;
  const hPts = rotated ? card.wPts : card.hPts;
  const name = rotated ? card.nameRotated : card.nameDefault;
  const { dw, dh } = calcPreview(wPx, hPx);

  el.querySelector('.card-name').textContent = name;
  const dimText = unit === 'cm'
    ? `${ptsToCm(wPts).toFixed(1)} × ${ptsToCm(hPts).toFixed(1)} ס״מ`
    : `${wPx} × ${hPx} פיקסל`;
  el.querySelector('.card-dimensions').textContent = dimText;
  el.setAttribute('aria-label', `${name} — ${dimText}`);

  const shape = el.querySelector('.preview-shape');
  shape.style.width = dw + 'px';
  shape.style.height = dh + 'px';
}
```

- [ ] **Step 4: Add `updateAllCardVisuals` and `updateCustomSection`**

After `updateCardVisual`, add:

```js
function updateAllCardVisuals() {
  PRESETS.forEach(card => {
    const el = document.querySelector(`.preset-card[data-id="${card.id}"]`);
    if (el) updateCardVisual(card, el);
  });
}

function updateCustomSection() {
  const wLabel = document.querySelector('label[for="customWidth"]');
  const hLabel = document.querySelector('label[for="customHeight"]');
  const wInput = document.getElementById('customWidth');
  const hInput = document.getElementById('customHeight');
  const isCm = unit === 'cm';

  wLabel.textContent = isCm ? 'רוחב (ס״מ)' : 'רוחב (פיקסל)';
  hLabel.textContent = isCm ? 'גובה (ס״מ)' : 'גובה (פיקסל)';
  wInput.placeholder = isCm ? '33.9' : '1920';
  hInput.placeholder = isCm ? '19.1' : '1080';
  wInput.step = isCm ? '0.1' : '1';

  if (currentSlidePts.w && currentSlidePts.h) {
    wInput.value = isCm
      ? ptsToCm(currentSlidePts.w).toFixed(1)
      : String(Math.round(currentSlidePts.w / 0.75));
    hInput.value = isCm
      ? ptsToCm(currentSlidePts.h).toFixed(1)
      : String(Math.round(currentSlidePts.h / 0.75));
  }
}
```

- [ ] **Step 5: Update `handleApply` to use unit-aware conversion**

In `handleApply`, replace the entire custom-input branch (after the `if (selectedPresetId)` block returns):

Find:
```js
  const wPx = parseFloat(document.getElementById('customWidth').value);
  const hPx = parseFloat(document.getElementById('customHeight').value);

  if (!wPx || !hPx || isNaN(wPx) || isNaN(hPx) || wPx <= 0 || hPx <= 0) {
    showError('בחר גודל מהרשימה או הזן פיקסלים בשדות למטה');
    return;
  }

  const wPts = wPx * 0.75;
  const hPts = hPx * 0.75;

  if (wPts < PT_MIN || hPts < PT_MIN) {
    showError(`גודל מינימלי: 48 פיקסל`);
    return;
  }
  if (wPts > PT_MAX || hPts > PT_MAX) {
    showError(`גודל מקסימלי: 5,376 פיקסל`);
    return;
  }

  applySlideSize(wPts, hPts, `${wPx} × ${hPx} פיקסל`);
```

Replace with:
```js
  const raw1 = parseFloat(document.getElementById('customWidth').value);
  const raw2 = parseFloat(document.getElementById('customHeight').value);

  if (!raw1 || !raw2 || isNaN(raw1) || isNaN(raw2) || raw1 <= 0 || raw2 <= 0) {
    showError('בחר גודל מהרשימה או הזן מידות בשדות למטה');
    return;
  }

  const wPts = unit === 'cm' ? cmToPts(raw1) : pxToPts(raw1);
  const hPts = unit === 'cm' ? cmToPts(raw2) : pxToPts(raw2);

  if (wPts < PT_MIN || hPts < PT_MIN) {
    showError(unit === 'cm' ? 'גודל מינימלי: 1.3 ס״מ' : 'גודל מינימלי: 48 פיקסל');
    return;
  }
  if (wPts > PT_MAX || hPts > PT_MAX) {
    showError(unit === 'cm' ? 'גודל מקסימלי: 142.2 ס״מ' : 'גודל מקסימלי: 5,376 פיקסל');
    return;
  }

  const label = unit === 'cm'
    ? `${raw1.toFixed(1)} × ${raw2.toFixed(1)} ס״מ`
    : `${Math.round(raw1)} × ${Math.round(raw2)} פיקסל`;
  applySlideSize(wPts, hPts, label);
```

- [ ] **Step 6: Verify in browser (no PowerPoint needed)**

Open `index.html` in a browser. Confirm no console errors. The layout looks identical to before (`unit` defaults to `'px'`, cards show px labels).

- [ ] **Step 7: Commit**

```bash
git add index.html
git commit -m "feat: add unit state, conversion helpers, and unit-aware card/input rendering"
```

---

### Task 3: Unit toggle UI

**Files:**
- Modify: `index.html` — `<header>` HTML + CSS + `setUnit` function + `Office.onReady` wiring

**Interfaces:**
- Consumes: `unit`, `updateAllCardVisuals()`, `updateCustomSection()` (Task 2)
- Produces: `setUnit(newUnit: 'px' | 'cm'): void`

- [ ] **Step 1: Make `header` a flex row**

Find:
```css
header {
  padding-bottom: 12px;
  margin-bottom: 14px;
  border-bottom: 1px solid var(--gray-200);
}
```

Replace with:
```css
header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding-bottom: 12px;
  margin-bottom: 14px;
  border-bottom: 1px solid var(--gray-200);
}
```

- [ ] **Step 2: Add unit toggle CSS**

After the updated `header` rule, add:

```css
.unit-toggle {
  display: flex;
  background: var(--gray-200);
  border-radius: 20px;
  padding: 2px;
}

.unit-btn {
  font-family: 'Heebo', sans-serif;
  font-size: 11px;
  font-weight: 500;
  padding: 3px 10px;
  border: none;
  border-radius: 18px;
  background: transparent;
  color: var(--gray-600);
  cursor: pointer;
  transition: background 0.15s, color 0.15s, box-shadow 0.15s;
  line-height: 1.4;
}

.unit-btn.active {
  background: #ffffff;
  color: var(--accent);
  font-weight: 600;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.12);
}

.unit-btn:focus-visible {
  outline: 2px solid var(--accent);
  outline-offset: 1px;
}
```

- [ ] **Step 3: Add toggle HTML to the header**

Find:
```html
<header>
  <h1>Slide Master</h1>
</header>
```

Replace with:
```html
<header>
  <h1>Slide Master</h1>
  <div class="unit-toggle" role="group" aria-label="יחידת מידה">
    <button class="unit-btn active" id="unitPx" aria-pressed="true">px</button>
    <button class="unit-btn" id="unitCm" aria-pressed="false">ס״מ</button>
  </div>
</header>
```

- [ ] **Step 4: Add `setUnit` function**

After `updateCustomSection`, add:

```js
function setUnit(newUnit) {
  unit = newUnit;
  document.getElementById('unitPx').classList.toggle('active', unit === 'px');
  document.getElementById('unitPx').setAttribute('aria-pressed', String(unit === 'px'));
  document.getElementById('unitCm').classList.toggle('active', unit === 'cm');
  document.getElementById('unitCm').setAttribute('aria-pressed', String(unit === 'cm'));
  updateAllCardVisuals();
  updateCustomSection();
}
```

- [ ] **Step 5: Wire up toggle buttons in `Office.onReady`**

After `PRESETS.forEach(card => { grid.appendChild(buildCard(card)); });`, add:

```js
document.getElementById('unitPx').addEventListener('click', () => setUnit('px'));
document.getElementById('unitCm').addEventListener('click', () => setUnit('cm'));
```

- [ ] **Step 6: Verify manually in PowerPoint**

Load the add-in. Confirm:
- Toggle pill appears in the header, right-aligned next to the title
- Clicking "ס״מ" switches ALL preset card labels:
  - מצגת רחבה → `33.9 × 19.1 ס״מ`
  - ריבוע → `19.1 × 19.1 ס״מ`
  - סטורי → `19.1 × 33.9 ס״מ`
  - A4 לאורך → `19.1 × 27.5 ס״מ`
  - A3 לאורך → `26.7 × 35.6 ס״מ`
- Custom input labels switch from "(פיקסל)" to "(ס״מ)"
- Clicking "px" reverts all labels back
- Rotate button on "מצגת רחבה": after rotating in cm mode, label shows `19.1 × 33.9 ס״מ`

- [ ] **Step 7: Commit**

```bash
git add index.html
git commit -m "feat: add px/cm unit toggle pill to header"
```

---

### Task 4: Current canvas read and pre-fill

**Files:**
- Modify: `index.html` — `Office.onReady` callback (make async + add canvas read)

**Interfaces:**
- Consumes: `currentSlidePts`, `updateCustomSection()` (Task 2)
- Produces: `currentSlidePts` populated with the live slide dimensions on load

- [ ] **Step 1: Make `Office.onReady` callback async**

Find:
```js
Office.onReady((info) => {
```

Replace with:
```js
Office.onReady(async (info) => {
```

- [ ] **Step 2: Add canvas read after card building**

After the unit toggle wiring (`unitPx`/`unitCm` event listeners) and before the `applyCustomBtn` listener, add:

```js
try {
  await PowerPoint.run(async (context) => {
    const pageSetup = context.presentation.pageSetup;
    pageSetup.load('slideWidth,slideHeight');
    await context.sync();
    currentSlidePts = { w: pageSetup.slideWidth, h: pageSetup.slideHeight };
    updateCustomSection();
  });
} catch (_) {
  // silent degradation — inputs keep placeholder text
}
```

- [ ] **Step 3: Verify manually in PowerPoint**

Load the add-in in a widescreen presentation (13.33" × 7.5" = 960 × 540 pts). Confirm:
- In px mode: custom width shows `1280`, height shows `720`
- Switch to cm: width shows `33.9`, height shows `19.1`
- Switch back to px: shows `1280` and `720`
- Typing in the inputs still deselects any selected preset card (existing behavior unchanged)
- In a browser (outside PowerPoint): inputs stay empty with placeholder text — no error shown

- [ ] **Step 4: Commit**

```bash
git add index.html
git commit -m "feat: read current slide size on load and pre-fill custom inputs"
```

---

### Task 5: Simplify pass

**Files:**
- Modify: `index.html`

- [ ] **Step 1: Run `/simplify`**

Invoke the `simplify` skill on `index.html`. The skill reviews for redundancy, unnecessary nesting, over-verbose naming — without changing any behavior.

- [ ] **Step 2: Verify no regressions**

After simplification, manually test in PowerPoint:
- Unit toggle switches all labels (px and cm modes)
- Canvas pre-fill works on load
- Preset cards select and apply with spinner + success message
- Custom input in px mode applies correctly
- Custom input in cm mode applies correctly
- Out-of-range values show the correct unit-specific error message
- Apply spinner appears and disappears on completion

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "refactor: simplify index.html after feature additions"
```
