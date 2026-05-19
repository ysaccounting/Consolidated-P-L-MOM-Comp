# Consolidated P&L Builder

A Flask web app that combines two monthly QuickBooks P&L exports into a single side-by-side Excel report with variance analysis.

## Features

- Upload two monthly P&L `.xlsx` files
- Select month labels and year
- Generates a formatted Excel workbook with:
  - All 14 entities side-by-side (Eliminations excluded)
  - Total column at the end
  - Mar / Apr / $ Chg / % Chg columns per entity
  - Gross Profit Margin row with formulas
  - Color-coded rows (section headers, subtotals, Gross Profit, NOI, Net Income)
  - Gray spacer columns between entities
  - Auto-sized columns

## Local Development

```bash
# Clone the repo
git clone https://github.com/YOUR_ORG/pl-builder.git
cd pl-builder

# Create virtual environment
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run
python app.py
```

Then open http://localhost:5000

## Deploy to Railway

### Option A — Railway Dashboard (easiest)

1. Push this repo to GitHub
2. Go to [railway.app](https://railway.app) → **New Project** → **Deploy from GitHub repo**
3. Select this repository
4. Railway auto-detects Python via Nixpacks and uses the `Procfile`
5. Your app is live — Railway provides a public URL

### Option B — Railway CLI

```bash
npm install -g @railway/cli
railway login
railway init
railway up
```

## Environment Variables

No required environment variables. Railway automatically provides `$PORT`.

Optional:
| Variable | Default | Description |
|----------|---------|-------------|
| `PORT` | `5000` | Port for local dev (Railway sets this automatically) |

## File Structure

```
pl-builder/
├── app.py           # Flask routes
├── pl_builder.py    # Excel generation logic
├── templates/
│   └── index.html   # Upload UI
├── requirements.txt
├── Procfile         # Railway / Heroku process definition
├── railway.toml     # Railway deployment config
└── .gitignore
```
