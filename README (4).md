# GMIC Competitor Price Tracker

**Green Mountains Insect Co. — Burlington, VT**

Automated competitor price intelligence for the dubia roach market.
Scrapes US and Canadian competitors, compares against GMIC pricing,
calculates per-roach costs, tracks CAD/USD live, and outputs a
Power BI-ready Excel file with full price history.

---

## Repo Structure

```
gmic-price-tracker/
├── gmic_price_scraper.py     # Main scraper script
├── requirements.txt          # Python dependencies
├── dashboard/
│   └── gmic_dashboard.html   # Standalone interactive dashboard
├── data/
│   └── .gitkeep              # Excel output goes here (gitignored)
├── .gitignore
└── README.md
```

---

## Quick Start

**1. Clone the repo**
```bash
git clone https://github.com/YOUR_USERNAME/gmic-price-tracker.git
cd gmic-price-tracker
```

**2. Install dependencies**
```bash
py -m pip install requests beautifulsoup4 pandas openpyxl lxml
```

**3. Run the scraper**
```bash
py gmic_price_scraper.py
```

**4. Open the output**
- Excel: `data/gmic_price_analysis.xlsx`
- Dashboard: open `dashboard/gmic_dashboard.html` in any browser

---

## What It Tracks

### Competitors (23 total)

| Region | Type | Competitors |
|---|---|---|
| US | Live Feeders | Dubia.com, Dubia Roach Depot, The Critter Depot, Reptilian Arts, TopFlight Dubia, Dubia Deli, ABDragons, Dubia Roach Broker, Backwater Reptiles, Chewy |
| US | Freeze-Dried | Fluker's, Chewy, Walmart, Petco, PetSmart, Amazon |
| US | Subscription | Dubia.com Sub, Dubia Roach Depot Sub |
| Canada | Freeze-Dried | Amazon.ca, PetSmart Canada, Big Al's |
| Canada | Live/FD | Pisces Pros, Josh's Frogs |

### Data Points Per Run

| Column | Description |
|---|---|
| Price (Native) | Price in original currency |
| Currency | USD or CAD |
| Price (USD Est.) | CAD converted to USD at live rate |
| Per-Roach (USD) | Apples-to-apples cost per roach |
| GMIC Price (USD) | Your equivalent price |
| GMIC Per-Roach | Your per-roach cost |
| vs GMIC |  Competitive /  GMIC Higher /  GMIC Lower |
| Free Ship At | Free shipping threshold |
| Subscription | Subscribe & save discount if detected |
| In Stock | Stock status — flags opportunities |
| CAD/USD Rate | Live rate used for conversion |
| Rate Source | "live" or "fallback" |
| Rate Fetched At | Timestamp of rate fetch |

---

## Live CAD/USD Rate

The script fetches the live CAD/USD rate from
[open.er-api.com](https://open.er-api.com) on every run —
free, no API key required. Rate is logged with every data row
so historical comparisons remain accurate.

If the API is unreachable, falls back to `0.73` automatically.

---

## Excel Output — 4 Sheets

| Sheet | Contents |
|---|---|
| **Competitor Prices** | All scraped price points, color-coded vs GMIC |
| **GMIC Baseline** | Your current prices with per-roach breakdown |
| **Summary** | KPIs, competitive status counts, rate used |
| **Price History** | Every run appended — builds trend data over time |

---

## Power BI Integration

1. Open **Power BI Desktop**
2. **Get Data → Excel Workbook** → select `data/gmic_price_analysis.xlsx`
3. Load all 4 sheets
4. Recommended visuals:
   - **Bar chart**: Per-roach price by competitor (filtered by size)
   - **Line chart**: Price history over time (use Price History sheet)
   - **Table**: Competitor Prices with vs GMIC color coding
   - **Card**: Count of  GMIC Higher alerts
   - **Slicer**: Filter by Country, Product Type, Size

---

## Standalone Dashboard

Open `dashboard/gmic_dashboard.html` in any browser for an
interactive summary dashboard — no Power BI required.
Works offline. Upload the Excel file directly in the browser.

---

## Configuration

Edit the top of `gmic_price_scraper.py` to customize:

```python
# Update your prices here when they change
GMIC_PRICES = { ... }

# Add new competitors here
COMPETITORS = [ ... ]

# Fallback rate if API is unreachable
CAD_TO_USD_FALLBACK = 0.73
```

---

## Recommended Run Schedule

| Frequency | Why |
|---|---|
| Weekly | Track seasonal price shifts |
| Before a price change | Benchmark before adjusting |
| After a competitor launches | Capture new pricing immediately |
| Monthly | Power BI trend report |

---

## GMIC Baseline Prices (as of Jan 2027)

### Live Feeders
| Size | Count | Price | Per-Roach |
|---|---|---|---|
| Small / Nymphs | 50–100 ct | $12.99 | ~$0.17 |
| Medium | 50–100 ct | $17.99 | ~$0.24 |
| Large | 25–50 ct | $22.99 | ~$0.62 |
| Adult Males | 50 ct | $27.99 | ~$0.56 |

### Freeze-Dried
| Product | Weight | Price | Per-Roach |
|---|---|---|---|
| Standard Jar | 1.3 oz | $13.99 | ~$0.30 |
| Premium Pouch | 2 oz | $18.99 | ~$0.27 |
| Large Pouch | 4 oz | $32.99 | ~$0.24 |

---

## Contact

**Green Mountains Insect Co.**
Burlington, Vermont
[greenmountainsinsect.com](https://greenmountainsinsect.com)
contact@greenmountainsinsect.com
wholesale@greenmountainsinsect.com

---

*Vermont-raised. Harvest Right processed. Shipped with care, every time.*
