# Claude Code — Apply this redesign

Hand this whole `handoff/` folder (or this repo) to Claude Code with the prompt below. The redesign is **purely aesthetic** — every route, form field, FMP call, DCF formula, scoring criterion, and rendered insight stays the same.

---

## Prompt to paste into Claude Code

> I want you to apply a visual redesign to my Flask app at `<path/to/repo>`. The mockups and spec live in this folder I'm pasting in (or attaching). Treat it as the visual source of truth.
>
> **Files in the handoff folder you'll use:**
> - `iu-redesign.css` — single drop-in stylesheet. Contains all design tokens, baseline rules, and per-page styles. **This is your CSS.**
> - `HANDOFF.md` — full design spec with redlines, what changed and why, page-by-page guide.
> - `screens/home.html`, `screens/dashboard.html`, `screens/dcf.html`, `screens/news.html` — pixel-accurate static mockups. Use them as the visual target.
> - `compare.html` — side-by-side before/after viewer (`uploads/` are the originals).
>
> **What to do:**
> 1. Copy `iu-redesign.css` into `static/` (or wherever your Flask static dir is).
> 2. Link it in the base template **after** any existing stylesheet so it overrides cleanly:
>    ```html
>    <link rel="stylesheet" href="{{ url_for('static', filename='iu-redesign.css') }}">
>    ```
> 3. Add `class="iu-redesign"` to the `<body>` of every page so the namespaced rules apply.
> 4. Update each Jinja template's markup to match the structure of its corresponding mockup in `screens/`. Keep all `{% %}` and `{{ }}` blocks, all `name=` and `id=` attributes on form inputs, all route URLs, all data dict keys. Only the surrounding HTML structure and class names change.
> 5. Top nav: rename the wrapper to `<nav class="iu-nav">`, with `.nav-logo`, `.nav-links`, and `.active` on the current page link. The dot in `Investment.Utopia` should be a `<span>` so it can take the gold accent.
> 6. **DCF page** is the biggest change. The new structure:
>    - Page header (eyebrow + title)
>    - `.cmd-bar` (ticker input + company info + price)
>    - `.qp-bar` (recent tickers — populate from session/localStorage)
>    - `.scenario-strip` (Base/Bull/Bear chips)
>    - `.verdict` 4-cell band (current price, GGM target, Exit-mult target, WACC) — **bind to existing computed vars**
>    - `.layout` 2-col: projection table (left, all `<input>`s become `.proj-input`) + WACC/TV stack (right)
>    - `.output-table` for DCF output, with `.ufcf` / `.tv` / `.sum` row classes
>    - `.sens-grid` 2-up sensitivity heatmap. Apply `.cell-low/.cell-mid-low/.cell-mid/.cell-mid-high/.cell-high` classes server-side based on the existing computed price-vs-current ratio (or in JS after render). Add `.cell-current` to the cell matching the current scenario.
>    - `.compare-table` for scenario comparison
> 7. **Home cards**: each coverage company gets a card with `.card-co`, `.card-tk`, `.ring` (SVG, color from verdict), `.verdict` pill, `.ratings` row of `.rating-tag.aaa/.aa/.nr` chips, `.price-row` 3-col footer, `.card-cta`.
> 8. **Dashboard table**: keep all columns, but use the same score-ring + verdict-pill + rating-chip components from the home card.
> 9. **News list**: each article gets mono source+timestamp + Playfair headline + dek + ticker chips.
>
> **What NOT to touch:**
> - Any Python in `app.py` / route handlers
> - Any FMP API calls
> - DCF math, WACC formula, scoring math
> - Excel report generation
> - Form `name=` attributes (form posts must keep working)
> - Any computed variable names passed to templates
>
> When you're done, run the app locally and visually compare against the mockups in `screens/`.

---

## Quick file map

```
handoff/
├── iu-redesign.css         ← drop into static/, link in base template
├── HANDOFF.md              ← full design spec
├── compare.html            ← before/after viewer (open in browser)
└── README.md               ← this file
screens/
├── home.html               ← target visual for index.html
├── dashboard.html          ← target for dashboard.html
├── dcf.html                ← target for dcf.html
└── news.html               ← target for news.html
uploads/                    ← snapshot of the live site (the "before")
tokens.css                  ← (already inlined in iu-redesign.css)
```

## Sanity-check checklist

- [ ] `<body class="iu-redesign">` on every page
- [ ] All `{% url_for %}` and route URLs unchanged
- [ ] `<input name="ticker">`, `<input name="rating">`, `<select name="business_clarity">`, etc. — all preserved
- [ ] DCF projection inputs still POST/AJAX to the existing handler
- [ ] WACC / DCF / sensitivity values come from the same Jinja vars as before
- [ ] Coverage cards on home read from the same coverage list var
- [ ] News articles render from the same news list var
