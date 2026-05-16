# Investment Utopia — Visual Redesign · Handoff Spec

This is a **purely aesthetic** refresh of the existing site. All routes, fields, inputs, outputs, scoring criteria, FMP data calls, DCF math, scenario logic, and rendered insights are **unchanged**. Implement this against the existing markup; this doc tells you what to change visually and where.

**Source mockups (pixel reference):**
- `screens/home.html` → maps to `index.html` (Generate Report)
- `screens/dashboard.html` → maps to `dashboard.html`
- `screens/dcf.html` → maps to `dcf.html`
- `screens/news.html` → maps to `news.html`
- `index.html` → side-by-side canvas of all four

---

## 1 · The big idea

Today's site has a strong identity — institutional dark + Playfair/Plex + gold accent — but the hierarchy is flat: gold appears on eyebrows, ticker pills, primary CTAs, headlines, AND active states all at once, so nothing reads as "the most important thing on screen." The DCF page is the densest screen and currently the weakest — 7.5pt mono labels, blue input cells fighting the gold palette, sensitivity tables as undifferentiated grids.

The redesign keeps the identity and fixes those four issues:

1. **Restrict gold** to true accents — primary CTAs, the active nav underline, eyebrows, computed key results (WACC, target price). It is *not* used on every interactive element.
2. **Add a real type scale** (px-based, 9→44) — replaces the ad-hoc 7.5pt / 8pt / 9pt mono labels.
3. **Add a 5-step surface elevation** (was 3) — gives panels, hover states, and inset rows distinct levels.
4. **Calmer "edit" affordance** — input cells in DCF use a soft blue (`--edit`) instead of competing with gold. The convention is now: **gold = computed/derived/key result · blue = editable input · neutral = static data.**

---

## 2 · Tokens (drop-in)

The full token file is `tokens.css`. Replace the existing `:root` block in your global CSS with it. Highlights:

| Group | Old | New |
|---|---|---|
| Background | `#07090e` | `#06080d` |
| Surfaces | 3 levels (`--surface`, `--surface-2`, `--surface-3`) | 5 levels (`--surface-1` … `--surface-4`) + `--bg` |
| Borders | one token | three (`--border-1/2/3`) |
| Gold | `#c9a84c` (single) | `--gold #d4b15a` + `--gold-soft`, `--gold-bright`, `--gold-dim`, `--gold-glow` |
| Ink | 3 levels | 5 levels (`--ink-1` … `--ink-5`) |
| Edit (input cells) | `#3b82f6`-ish blue | `--edit #6ea8ff` + `--edit-dim`, `--edit-bg` |
| Type sizes | pt-based, ad-hoc | px scale: `--t-display 44 / --t-h1 30 / --t-h2 22 / --t-h3 16 / --t-body 14 / --t-small 12 / --t-eyebrow 10 / --t-micro 9` |
| Spacing | none | 8-pt scale `--s-1`…`--s-20` |
| Radii | none | `--r-sm 3 / --r-md 6 / --r-lg 10 / --r-xl 14` |
| Shadows | none | `--shadow-1/2/3` + `--glow-gold` |
| Motion | none | `--dur-fast 120ms / --dur-base 220ms / --dur-slow 400ms` + `--ease`, `--ease-out` |

Fonts unchanged — keep the existing `<link>` to Playfair Display + IBM Plex Sans + IBM Plex Mono.

---

## 3 · Universal patterns (all pages)

### Top nav
- Sticky, 56px tall, `background: rgba(6,8,13,0.85); backdrop-filter: blur(16px)`.
- Logo is mono 11px / 3px tracking, the dot in `Investment.Utopia` is gold.
- Links are mono 10px / 1.5px tracking. Active link gets a 2px gold underline 19px below baseline (uses `::after`, not a border).
- Right side has a small "Live · FMP" status with a pulsing green dot.

### Page header pattern
```
[mono 10px gold eyebrow, 3px tracking]
[Playfair 30px page title]
                                              [ghost action] [ghost action]
─────────────────────────────────── (1px border-1 divider)
```

### Buttons
- `.btn` base: mono 10px, 1.5px tracking, uppercase, 9×16 padding, 6px radius.
- `.btn-primary` — gold fill, dark text (`#1a1408`). One per screen, max.
- `.btn-ghost` — transparent, `--border-2`, `--ink-3` text. Hover: lifts to `--ink-1` + `--border-3`.
- `.btn-pos` — green-ringed for save/run actions.
- `.btn-sm` — 6×12 / 9px font for toolbar/secondary use.

### Panels
- `.panel` — `--surface-1` bg, 1px `--border-2`, 10px radius.
- `.panel-head` — 14×20 padding, 1px `--border-1` divider. Title is mono 10px gold, optional right-aligned hint in mono 9px ink-4.

### Tables (the workhorse)
- TH: mono 9px / 1px tracking / uppercase / `--ink-4`, padded 12×10/14, bottom 1px `--border-2`.
- TD: mono 12px, padded 8–10×10/14, bottom 1px `--border-1`.
- Numeric cells right-aligned. Label cells use `--font-body` 12px and stay left-aligned.
- Hover row: `background: rgba(255,255,255,0.012)` (very subtle).
- Section divider rows: mono 9px / 2px tracking / uppercase / `--ink-4`, `background: rgba(255,255,255,0.015)`.

### Status colors (all surfaces share these conventions)
- **Buy / positive / good** → `--pos` text on `--pos-bg`, ring `rgba(52,211,153,0.2)`.
- **Hold / warning** → `--warn` text on `--warn-bg`, ring `rgba(251,191,36,0.2)`.
- **Sell / negative / bad** → `--neg` text on `--neg-bg`, ring `rgba(248,113,113,0.2)`.

---

## 4 · Page-by-page redlines

### 4.1 — Home / Generate (`screens/home.html`)

**Hero**
- Replace the existing hero with a centered block: 100px top padding, gold eyebrow → Playfair 56px title with one italic gold word (`<em>Built to hold.</em>`) → 15px ink-3 sub.
- Add a subtle radial gold glow behind the hero (`radial-gradient(ellipse, var(--gold-glow), transparent 60%)`, 60% opacity, behind text).

**Generate panel** (the form)
- Card: `--surface-1`, `--border-2`, 14px radius, 28px padding, `--shadow-3`. Max-width 720px, centered.
- Fields use mono 9px / 2px tracking uppercase labels above mono 14px inputs (uppercase, 3px letter-spacing for the ticker).
- Optional fields get a `· optional` hint in `--ink-5`.
- Single primary CTA: `Generate Report →`, full-width, gold.
- Below CTA: mono 10px ink-4 meta line — "Generates Excel model + HTML report · ~30–60 seconds".

**Stats strip**
- Full-width band between hero and content. 4 cells, dividers in `--border-1`. Each cell: Playfair 32px number (with gold accent character) + mono 10px / 2px tracking label.

**Coverage cards**
- 2-up grid (collapse to 1 below 880px). Each card: `--surface-1`, `--border-1`, 10px radius, 28px padding.
- Layout: header row (Playfair 20px company name + mono 10px gold ticker pill) ↔ score ring on right.
- **Score ring** — 64×64 SVG, 26r, stroke 5px, color matches verdict. Center: mono 16px score / mono 8px `/100`.
- **Verdict pill** — pos/warn/neg semantic colors (matches "Good Business at Fair Price" / "Report Pending" / etc).
- **Rating row** — three rating chips (S&P · Moody's · Fitch). Each chip has a tiny `ag` agency label on top of the rating. Use `aaa` (full-pos), `aa` (soft-pos), `nr` (neutral) variants.
- **Price row** — 3-column footer with `Current · Target · Upside`, divided by `--border-1`. Upside in `--pos`/`--neg`.
- Hover: top edge gets a 1px gold-to-transparent gradient bar; card lifts 3px.
- "In Progress" cards drop to 50% opacity with a `pending-tag` badge instead of the ring.

**Methodology section**
- 4 numbered pillars (`01`…`04`). Mono 28px ink-5 number → Playfair 17px title → 12.5px ink-3 body.
- Below: a horizontal "Score Scale" legend bar that fades neg → warn → pos with tick labels (`0 — Avoid`, `50 — Hold`, `65 — Buy`, `75+ — High Conviction`).

### 4.2 — Dashboard (`screens/dashboard.html`)

The current dashboard is information-rich but visually monotonous. The redesign keeps every column but adds rhythm:

- **Filter bar** — mono pill chips for sector / score range / rating / verdict, with a search input on the right.
- **Top KPI strip** — 4 callouts: Total Coverage, Average Score, High-Conviction count, Pending count. Same panel pattern as DCF verdict.
- **Coverage table** — every row uses the same score-ring component as the home cards (smaller, 36px). Verdict pill in the same row. Rating chips inline. Price/Target/Upside numeric columns right-aligned in mono 12px.
- Sortable columns get a tiny chevron on hover.

### 4.3 — DCF Calculator (`screens/dcf.html`) — **the focus**

This is the biggest change. Sections in order:

1. **Page header** — eyebrow "Valuation Workbench" + Playfair 30px "DCF Calculator" + ghost actions (Export Excel · Share Scenario).

2. **Command bar** — single panel, 3 columns: ticker input (with Load button) · company info (mark + name + meta line) · live price + day change. Replaces the scattered ticker bar.

3. **Quick picks row** — mono 10px pill buttons for recent tickers. Active state uses `--gold-dim` bg.

4. **Scenario strip** — Base / Bull / Bear / Analyst Cons. chips with colored dots. New Scenario chip with dashed border. Save/Reset on the right.

5. **Verdict band** (NEW — biggest single addition) — 4-cell horizontal panel at the top of the fold:
   - Cell 1: Current Price (neutral)
   - Cell 2: Gordon Growth target (gold value, upside pill, Excel-parity badge)
   - Cell 3: Exit Multiple target (gold value, upside pill, implied multiple)
   - Cell 4: WACC + TV % of EV
   This makes the answer the headline. The cells use `--surface-2` for the two key results, `--surface-1` for the others.

6. **Two-column main** — 1.5fr / 1fr split:
   - **Left:** Projection Assumptions table — historical FY-3/-2/-1 in muted ink-4, projected FY+1…+5 in gold-tinted headers, all driver rows with editable cells. Section dividers ("Growth & Profitability", "Cash Flow Drivers"). All `<input>`s use the new `--edit` blue convention.
   - **Right:** WACC Build panel + Terminal Value panel stacked. WACC computed row pinned in gold. Mid-year discounting checkbox at the bottom.

7. **DCF Output table** — full-width. UFCF row gets the gold-tinted highlight band. Terminal value rows get the blue-tinted band. Summary rows (EV, Equity Value, etc.) get a heavier top border and `--surface-2` bg.

8. **Sensitivity grid** — two side-by-side heatmaps (GGM and Exit Multiple). 5×5 cells, color from neg (low) → pos (high). The current scenario cell gets a 2px gold outline + gold ◆ corner marker.

9. **Scenario Comparison table** — Base / Bull / Bear / Analyst Cons. as columns, drivers as rows. Active scenario column gets `--gold-dim` background and gold text on its cells.

### 4.4 — News (`screens/news.html`)

- **Page header** — same pattern, eyebrow "Market Pulse".
- **Filter pills** — All / My Coverage / Sector tags.
- **Article list** — editorial layout. Each article: mono source label + timestamp on top → Playfair 22px headline → 14px ink-2 dek → footer row with related ticker chips. 1px `--border-1` divider between articles.
- One **featured article** at the top spans full width with a colored block placeholder and a longer dek.

---

## 5 · Implementation checklist for Claude Code

In rough order:

1. **Drop in `tokens.css`** — replace existing `:root` vars site-wide. Search-and-replace any direct hex colors that match old token values.
2. **Search & replace pt sizes** — replace `font-size: 7.5pt` / `8pt` / `9pt` in DCF with the new `--t-micro` / `--t-eyebrow` / `--t-small` tokens.
3. **Update nav** — adopt sticky + blur + gold underline pattern across all routes.
4. **Add `.btn` system** — replace ad-hoc button styles with the four variants.
5. **Add `.panel` + `.panel-head` system** — wrap existing sections.
6. **DCF**: insert verdict band, restructure into command bar / scenario strip / two-column main / output / sensitivity / compare. **All Jinja template variables stay the same** — this is a pure markup/CSS swap.
7. **DCF inputs**: change `<input type="number">` cells to use `.proj-input` / `.wacc-input` classes (blue edit affordance). The form names, ids, and computed handlers do not change.
8. **Sensitivity tables** — apply the heatmap classes (`.cell-low` / `.cell-mid-low` / `.cell-mid` / `.cell-mid-high` / `.cell-high`) based on the existing computed price-vs-current ratio. Add `.cell-current` to the active scenario cell.
9. **Home cards** — replace the existing coverage card with the score-ring + verdict-pill + ratings-row + price-row pattern.
10. **Methodology** — replace whatever exists with the 4-pillar grid + score legend.
11. **Audit**: any remaining direct gold usage outside primary CTA / nav-active / eyebrow / computed-key-result should be retoned to `--ink-1` or `--ink-2`.

### Files NOT to touch
- All Python/Flask routing
- All FMP API calls and the data dict shape passed to templates
- Any computed values, scoring math, DCF formulae, WACC formula
- `<input name="…">` attributes (form posts must continue working)
- Excel generation pipeline

---

## 6 · Quick visual debt index

These specific things to delete from the current CSS:
- Any `font-size` in `pt`. Replace with px tokens.
- The `#3b82f6` (or similar) blue used on input cells — switch to `--edit`.
- Any usage of gold (`#c9a84c`) on hover states for non-primary controls — switch to `--ink-1` + `--border-3` ghost-hover.
- The plain striped sensitivity tables — replace with heatmap classes.
- Any standalone "ticker pill" gold backgrounds outside the verdict and quick-picks.

---

That's the whole spec. The four mockup HTMLs are pixel-accurate; treat them as the visual source of truth. If anything in the existing app doesn't have a corresponding pattern in this doc, default to the closest panel/table/badge convention from the same family.
