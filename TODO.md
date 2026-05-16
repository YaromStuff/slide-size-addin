# Layout Master — Task List

---

## Done

- [x] All preset buttons now work (Wide, Square, Story, A4, A3)
- [x] Fixed floating-point precision bug that caused Wide and others to fail (960.009 pts → exact 960 pts)
- [x] Fixed A4 and A3 precision (using full-precision mm → pts conversion: `mm / 25.4 * 72`)
- [x] Fixed error handling — now reads the real error type from the API, not a guess from the message text
- [x] Added PowerPoint API version requirement to manifest.xml so the add-in won't load on unsupported versions
- [x] Changed all "cm" labels in the UI to "פיקסל" (pixel in Hebrew)
- [x] Fixed the custom size input — now accepts pixel values and converts them correctly to the API unit (points)
- [x] Documented all Office.js API gotchas in CLAUDE.md so future AI sessions don't repeat the same mistakes

---

## To Do

### 1. Update the pixel numbers shown on each preset card

Right now the cards show the old cm values with a "פיקסל" label — which is wrong. Each preset needs real pixel dimensions decided and entered.

For each preset below, decide the width × height in pixels and update the `wCm` and `hCm` fields in the `PRESETS` array in `index.html` (these control what's shown on the card):

| Preset | Field name | Current value (wrong) | Correct pixel value (TBD) |
|--------|------------|----------------------|---------------------------|
| Wide | wCm / hCm | 33.87 × 19.05 | e.g., 1920 × 1080 |
| Square | wCm / hCm | 19.05 × 19.05 | e.g., 1080 × 1080 |
| Story | wCm / hCm | 19.05 × 33.87 | e.g., 1080 × 1920 |
| A4 | wCm / hCm | 21 × 29.7 | TBD |
| A3 | wCm / hCm | 29.7 × 42 | TBD |

> Note: the `wPts` / `hPts` values (the actual slide dimensions sent to PowerPoint) stay the same — only the display numbers change.

---

### 2. Update the success message to show the correct pixel values

After fixing task 1, the success message ("✓ מצגת רחבה — 33.87 × 19.05 פיקסל") should show the correct pixel numbers. This means switching the success message to use `preset.wCm` and `preset.hCm` (which will hold pixel values after task 1 is done) instead of the current cm values.

No code change needed beyond task 1 — once the preset display values are updated, the message will automatically show the right numbers.

---

### 3. Test A4 and A3 in PowerPoint after today's fix

The A4/A3 precision fix (task: "Fix A4/A3 pt precision") was just applied. These presets need to be tested live in PowerPoint to confirm they now work correctly. If they still fail, the fallback is to try rounding the pt values to integers (595 × 842 for A4, 842 × 1191 for A3).

---

### 4. Test the custom pixel input end-to-end

Test entering pixel values like 1920 × 1080 and 1280 × 720 in the custom section and applying them. Verify:
- The slide actually changes to the right size in PowerPoint
- The success message shows the entered pixel values
- Edge cases: entering 47 px (below minimum), entering 6000 px (above maximum)

---

### 5. Add more preset sizes (optional / future)

Consider adding common social media and design sizes as preset cards:

| Name | Width × Height (px) | Notes |
|------|---------------------|-------|
| YouTube Thumbnail | 1280 × 720 | 16:9 |
| LinkedIn Post | 1200 × 627 | Landscape |
| LinkedIn Story | 1080 × 1920 | Portrait |
| Twitter/X Post | 1600 × 900 | 16:9 |
| Facebook Post | 1200 × 630 | Landscape |
| Instagram Square | 1080 × 1080 | 1:1 |

---

### 6. Consider adding a "current size" indicator

Show the user what size their presentation currently is when the panel opens. This would help them know what they're changing from.

---

## Notes

- The "custom pixel input" converts pixels to points using: `pts = px × 0.75` (assumes 96 DPI screen)
- PowerPoint's slide dimensions are always stored in points internally (1 pt = 1/72 inch)
- The API requires PowerPointApi 1.10 — this means PowerPoint for Mac/Windows updated after ~2022
- Full technical details and API gotchas are in `CLAUDE.md` under the "Gotchas" section
