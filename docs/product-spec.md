# Product spec — Growth Curve Editor

Standalone, browser-only tool for editing growth driver rows in Excel financial models.

## Positioning

> Pick a growth story, set a few numbers, get a realistic curve — drag until it matches your conviction — download Excel.

Not a full FP&A tool. Owns **one driver lane** (users, signups, MRR flow, GMV, etc.).

## Distribution

| Channel | Role |
|---------|------|
| **Webapp** (static deploy) | Primary — anyone with a browser |
| **GitHub** (public repo) | Source, trust, issues, example profiles |
| **Buy Me a Coffee** | Optional tip — no paywall |

## User journey

```
Upload .xlsx → Map range(s) → Curve wizard → Knot editor → Download .xlsx
```

Demo mode: skip upload, 36-month blank horizon, export blocked until real file loaded.

## Workbook mapping

User provides per scenario:

- Sheet name
- A1 range (horizontal row), e.g. `C8:AN8`
- Optional label (Base, Upside, …)

Horizon = range width (number of columns). Knot months = evenly spaced across horizon (default 6).

## Series semantics

- **Stored in Excel:** monthly flow (signups, new users, etc.)
- **Edited in UI:** cumulative (knots on smooth curve)
- **Math:** monotonic PCHIP → derive integer monthly flow

## Curve types (v0.1)

### Milestone-led (default)

Inputs: up to 3 `(month, cumulative value)` pairs.

### Linear

Inputs: start cumulative, end cumulative @ horizon.

### S-curve

Inputs: early, mid (~45% of final), final @ horizon.

### Hockey stick

Inputs: flat through month N, level after flat, final @ horizon.

### Delayed launch

Inputs: launch month, final @ horizon.

All types compile to `knotMonths` + `knotCumulative` → shared PCHIP pipeline.

## Infer heuristics (from existing row)

Applied to cumulative series `C[1..T]`:

1. **Delayed launch** — first quarter sum / final < 5%
2. **S-curve** — middle-third slope > first and last third
3. **Hockey stick** — last-third slope > 4× first-third slope
4. **Linear** — R² vs straight line > 0.92
5. **Else milestone** — sample at 33%, 66%, 100% of horizon

User can override inferred type in wizard.

## Privacy

- No server upload
- Processing via SheetJS in browser
- State in memory only (no localStorage of workbook in v0.1)

## Out of scope (v0.1)

- Full model / formula graph
- Google Sheets
- Vertical ranges
- Growth % mode (MoM/YoY input)
- Profile JSON save/load
- Community template library

## Future (v0.2+)

- Profile YAML save/load
- Sheet grid range picker
- Duplicate scenario with scaled targets
- Compound % growth type
- Example profiles (`examples/ableten-3yr.yaml`)
- Optional Python local-server for live file sync

## Coffee / monetization

- Single BMC/Ko-fi link in header + post-download
- No feature gating
- GitHub Sponsors optional in README
