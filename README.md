# Growth Curve Editor

Visual growth curve editor for **Excel financial models**. Map your driver row, pick a curve shape, drag knots to refine, download an updated workbook.

**Your file never leaves your browser** — all processing runs locally with [SheetJS](https://sheetjs.com/).

## Quick start

```bash
cd growth-curve-editor
npm install
npm run dev
```

Open **http://127.0.0.1:5173**

Build for production:

```bash
npm run build
npm run preview
```

Deploy the `dist/` folder to any static host (Vercel, Netlify, Cloudflare Pages, GitHub Pages).

### GitHub Pages (recommended)

This repo includes a GitHub Actions workflow (`.github/workflows/deploy.yml`) that builds and deploys on every push to `main`.

1. **Create the repo** on GitHub (public): `growth-curve-editor`
2. **Push** this project:
   ```bash
   git remote add origin git@github.com:annerjiao/growth-curve-editor.git
   git push -u origin main
   ```
3. **Enable Pages**: Repo → **Settings** → **Pages** → Source: **GitHub Actions**
4. After the first workflow run, the site is live at:
   `https://annerjiao.github.io/growth-curve-editor/`

### Custom domain (HTTPS)

GitHub Pages includes free TLS for custom domains.

1. Add a file `public/CNAME` containing your domain (one line), e.g.:
   ```
   growth.ableten.xyz
   ```
2. In repo **Settings → Pages → Custom domain**, enter the same domain.
3. At your DNS provider, add:
   - **CNAME** `growth` → `annerjiao.github.io`  
     (or **A** records to GitHub Pages IPs if using apex domain)
4. Set `VITE_BASE_PATH: /` in `.github/workflows/deploy.yml` (replace the `/repo-name/` line) so assets load correctly at the domain root.
5. Push — GitHub will provision HTTPS automatically (can take up to 24h for DNS).

### Other hosts

`npm run build` outputs static files in `dist/`. Any CDN/static host works; the app has no backend.

## User flow

1. **Upload** — `.xlsx` workbook (or demo mode without a file)
2. **Map** — sheet name + cell range per scenario (e.g. `C8:AN8` for 36 monthly values)
3. **Curve** — choose type (milestone, linear, S-curve, hockey stick, delayed launch); adjust anchors; live preview
4. **Edit** — drag knots on cumulative curve; monthly flow derived via PCHIP
5. **Export** — download updated `.xlsx`

## Curve types

| Type | Best for |
|------|----------|
| Milestone-led | “500 @ M12, 5k @ M24, 20k @ M36” |
| Linear | Steady ramp |
| S-curve | Slow → fast → taper |
| Hockey stick | Flat early, sharp lift |
| Delayed launch | Pre-launch flat, then ramp |

On upload, the app **infers** a suggested curve type from your existing data.

## Claude milestone parsing (optional)

**Your users never deploy anything.** You deploy a tiny Cloudflare Worker once; visitors get AI parsing automatically. Your Anthropic key stays on the worker — it never ships in the JavaScript bundle.

```
Visitor's browser  →  your Worker (secret key)  →  Anthropic Claude Sonnet 4.5
```

Without the worker URL baked into the build, parsing still works locally (regex heuristics).

### One-time setup (you only)

```bash
cd worker
npm install
npx wrangler login
npx wrangler secret put ANTHROPIC_API_KEY
npx wrangler deploy
```

Copy the worker URL, then:

1. Add your live site to `ALLOWED_ORIGINS` in `worker/wrangler.toml`
2. GitHub → **Settings → Actions → Variables** → `VITE_MILESTONE_PARSE_URL` = worker URL
3. Push to `main` (rebuilds GitHub Pages with the worker URL embedded)

`RATE_LIMIT_PER_HOUR` (default 40) limits abuse of your shared key.

### Local dev with Claude

Terminal 1 — worker:

```bash
cd worker && npm install && cp .dev.vars.example .dev.vars
# paste your key into .dev.vars
npm run dev
```

Terminal 2 — app:

```bash
echo 'VITE_MILESTONE_PARSE_URL=http://localhost:8787' > .env.local
npm run dev
```

Without the worker URL, parsing still works **locally in the browser** (regex heuristics).

## Configuration

Links in `index.html`:

- **Buy Me a Coffee** — `https://buymeacoffee.com/tzqdgkjmx`
- **GitHub** — update if your repo URL differs

## Project structure

```
growth-curve-editor/
├── index.html
├── css/styles.css
├── js/
│   ├── app.js              # Step flow + UI
│   ├── curve-math.js       # PCHIP interpolation
│   ├── curve-generators.js # Curve types + infer
│   ├── xlsx-io.js          # Excel read/write
│   └── editor.js           # Canvas knot editor
├── docs/product-spec.md
└── package.json
```

## Scope (v0.1)

- Horizontal ranges (one row, one value per period)
- Monthly flow values; edit via cumulative knots
- Multiple optional scenarios
- Browser-only — no backend

**Not in v0.1:** formula parsing, Google Sheets, vertical ranges, % MoM mode.

## License

MIT
