# Slide HTML Generation Guidelines

These rules apply to every HTML slide you generate.  
Each slide is rendered by a **headless Chromium browser at a fixed 1280 × 720 px viewport**  
and captured as a screenshot. The screenshot becomes a slide in a PowerPoint file.  
**What Chromium sees is exactly what ends up in the deck — nothing more, nothing less.**

---

## 1. Canvas Contract (Non-Negotiable and Must Comply)

Every slide must be exactly **1280 × 720 px** — no more, no less.

```css
/* Required on every slide — copy this exactly */
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
  width: 1280px;
  height: 720px;
  overflow: hidden;   /* hard crop — nothing outside this box is captured */
}
```

**Why `overflow: hidden` matters:**  
The screenshot clips at exactly 1280 × 720. Any content that overflows is silently cut off.  
If your content might overflow, shrink font sizes or reduce padding — never remove `overflow: hidden`.

---

## 2. Self-Contained HTML Only

The slide HTML must work with **zero network requests**.  
The renderer blocks all external fetches (images, fonts, scripts, stylesheets).

| Resource type | Rule |
|---|---|
| CSS | Inline inside `<style>` tags only |
| JavaScript | Inline inside `<script>` tags only, or omit entirely |
| Images | Base64 `data:` URIs only — no `http://`, no `src="./file.png"` |
| Fonts | System fonts only (see approved list below) — no Google Fonts, no `@font-face` with URLs |
| Icons | Inline SVG only — no icon font CDNs (FontAwesome, Material Icons, etc.) |

**Approved system fonts** (guaranteed present in Chromium sandbox):

```
Arial, Arial Black, Calibri, Cambria, Consolas, Courier New,
Georgia, Impact, Palatino, Tahoma, Times New Roman, Trebuchet MS, Verdana
```

---

## 3. No Scroll, No Interaction

Slides are static screenshots. Anything that requires user interaction is wasted effort.

- **No scroll** — if content doesn't fit at 1280 × 720, cut it, don't scroll it
- **No hover states** — `:hover` CSS will never trigger
- **No animations / transitions** — the screenshot fires after 300 ms; multi-second animations will be captured mid-frame or not at all
- **No `position: fixed`** — fixed positioning is relative to the viewport, which is fine, but test that it lands where you expect at exactly 1280 × 720
- **No `vh` / `vw` units in deeply nested elements** — use `px` for precision; `vh`/`vw` on `body` itself is fine

---

## 4. Typography Rules

| Element | Size range | Notes |
|---|---|---|
| Slide title | 36 – 52 px | Bold; Georgia or Arial Black recommended |
| Section header | 20 – 28 px | Bold |
| Body text | 14 – 18 px | Regular weight |
| Caption / label | 10 – 13 px | Muted color; all-caps + letter-spacing works well |

- **Left-align body text.** Center only titles and full-bleed hero slides.
- **Minimum contrast ratio 4.5:1** between text and background.
- **Do not use `rem` or `em`** unless you have explicitly set a `font-size` on `html`/`body`. Use `px` for reliability.

---

## 5. Layout Patterns That Work

Use these proven patterns. They all fit within 1280 × 720 without overflow risk.

### Hero / Title slide
```
Full-bleed background color or gradient
Centered text block (flex column, justify-content: center, align-items: center)
Eyebrow label → H1 → divider → subtitle
```

### Two-column
```
display: flex on body
Left panel: fixed width ~320–420px, dark background, label + title
Right panel: flex:1, light background, content
```

### Icon + text rows (up to 4 rows)
```
Each row: flex row, icon circle (40px) + text block (bold header + description)
Gap between rows: 20–28px
```

### Stat callout
```
Large numbers (60–80px) with small uppercase labels below
3 stats max in a horizontal flex row
```

### 2×2 card grid
```
display: grid; grid-template-columns: 1fr 1fr; gap: 24px;
Each card: padding 24–32px, border-radius, background color
```

**Maximum content density per slide:**  
4 bullet points, or 4 icon rows, or 3 stat callouts, or a 2×2 grid.  
If you have more content, split it across multiple slides.

---

## 6. What Will Break (Common Failures)

| Mistake | What happens | Fix |
|---|---|---|
| Google Fonts `<link>` | Font loads as fallback serif/sans | Use system fonts |
| `<img src="https://...">` | Broken image or blank space | Convert to base64 data URI |
| Content taller than 720px | Silently cut off at bottom | Reduce content or split slides |
| `font-size` in `rem` without base set | Unpredictable size | Use `px` |
| CSS animation > 300ms | Captured mid-animation | Remove animations |
| `overflow: auto` or `scroll` on body | Creates scrollbar, clips content | Always `overflow: hidden` |
| FontAwesome / Material Icons CDN | Icons render as blank squares | Use inline SVG |
| `position: absolute` without `position: relative` parent | Element lands in wrong place | Always set parent positioning context |

---

## 7. Minimal Valid Slide Template

Copy this as your starting point for every slide:

```html
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    width: 1280px;
    height: 720px;
    overflow: hidden;
    font-family: Calibri, Arial, sans-serif;
    background: #FFFFFF;
    /* Add display: flex here for centering layouts */
  }

  /* Your styles below */
</style>
</head>
<body>
  <!-- Your slide content here -->
</body>
</html>
```

---

## 8. Quick Self-Check Before Finalizing a Slide

Before returning HTML to the user, verify:

- [ ] `body` has `width: 1280px`, `height: 720px`, `overflow: hidden`
- [ ] No external URLs in `src`, `href`, or `url()` — only `data:` URIs or none
- [ ] No Google Fonts or CDN icon libraries
- [ ] All fonts are from the approved system font list
- [ ] No CSS animations or transitions
- [ ] Content fits visually within 720px height (count your rows/elements)
- [ ] Text contrast is sufficient on the background color
- [ ] No `<form>`, `<input>`, `<button>` (interaction elements are pointless in screenshots)