# Rise TVP — Trade Route Optimizer

Extracts commodity data from city/EasyDock screenshots via OCR, then analyses
trade opportunities and outputs a structured Excel workbook.

---

## How screenshots are parsed

### City images (`CityName_1.png`, `CityName_2.png`, …)

Every city screen always shows **exactly 7 commodity rows in a fixed order**:

| Row | Commodity |
|-----|-----------|
| 1 | Rare/Precious |
| 2 | Foodstuffs |
| 3 | Natural Materials |
| 4 | Fuel Ore |
| 5 | Consumer Goods |
| 6 | Fabricated Material |
| 7 | Refined Fuel |

The script uses **positional parsing** — each row is mapped directly to its
commodity by index, not by trying to read the OCR name. This makes extraction
robust even when Tesseract mis-reads a character.

Each row yields five numbers in column order:

| Col | Field |
|-----|-------|
| 0 | Quantity MT (currently in city) |
| 1 | Reserve MT (locked, not tradeable) |
| 2 | Selling CR/MT (price to **buy from** this city) |
| 3 | Buying CR/MT (price this city will **pay**) |
| 4 | Maximum MT (maximum the city will accept) |

The **Totals** line (e.g. `Totals  64,345 / 65,535 MT  97,339,730 CR`) is also
parsed and surfaced in the Cities sheet as a summary table showing MT used, MT
free, and CR currently in the city.

### EasyDock image (`EasyDock_1.png`)

EasyDock commodities are not in a fixed order, so name-prefix + positional
column parsing is used. Only rows with a non-zero `Selling CR` are kept.

### Skipped images

Files whose base name starts with `tvi` (case-insensitive) are automatically
skipped — these are newspaper/header assets, not city data.

---

## Requirements

### 1. Tesseract OCR

Download and install from: https://github.com/UB-Mannheim/tesseract/wiki

On Windows the script auto-detects the default install locations. On Linux/macOS
it uses the `tesseract` command from PATH (`sudo apt install tesseract-ocr`).

### 2. Python packages

```bash
pip install -r requirements.txt
```

Core dependencies: `pytesseract`, `Pillow`, `openpyxl`

The optional `reportlab` and `PyMuPDF` packages are only needed if you run
`generate_news_pdf.py` separately.

---

## Usage

Place screenshots in the `images/` folder, then run:

```bash
# Windows
run_trade.bat

# Linux / macOS
bash ln_run_trade.sh
```

You will be prompted to choose a mode:

| Mode | Description |
|------|-------------|
| **1 — Regular Trade** | All opportunities from all cities |
| **2 — City-Specific** | All opportunities + a filtered sheet from your origin |
| **3 — Trade Route** | Multi-hop chained routes (A → B → C → origin) |

Then answer the prompts for ship, containers, rental, origin city, and budget.

You can also pass arguments directly:

```bash
python ocr_to_excel.py --ship "E-10 Saint" --origin Kansas --mode regular
```

---

## Output — `final_trade.xlsx`

| Sheet | Contents |
|-------|----------|
| **Config** | Selected ship, capacity, origin, budget, rental info |
| **Cities** | All city commodity rows + Totals summary table |
| **EasyDock** | EasyDock commodity rows |
| **Opportunities** | Every profitable A→B route, graded A–D, with ROI colour |
| **From CityName** | *(City mode only)* Opportunities originating from your city |
| **Trade Routes CityName** | *(Route mode only)* Multi-hop routes ordered by total profit |
| **MACRO** | Province-wide summary: top cities, best sellers/buyers per commodity |

### Opportunity grades

| Grade | Meaning |
|-------|---------|
| **A** | Top 25 % of profitable trips by CR earned |
| **B** | 25–50 % |
| **C** | 50–75 % |
| **D** | Bottom 25 % |

### ROI colour scale (Trip ROI column)

| Colour | ROI |
|--------|-----|
| 🟢 Dark green | > 100 % |
| 🟢 Light green | 50–100 % |
| 🟡 Yellow | 20–50 % |
| 🟠 Orange | 0–20 % |
| 🔴 Red | negative |

### Cities Totals summary

At the bottom of the **Cities** sheet a summary table shows, per city:

- **MT Used / MT Total** — storage fill level (highlighted red if > 90 % full)
- **MT Free** — remaining capacity
- **CR in City** — money currently held in the city treasury

---

## Ships

### AIR AND SPACE

| Ship | Base MT | Max containers | Container MT | Max capacity |
|------|---------|---------------|--------------|-------------|
| E-10 Saint | 7 | 6 | 17 | 109 MT |
| E-11 Saint | 7 | 6 | 17 | 109 MT |
| P-13 Prowler | 1 | 0 | — | 1 MT |
| W-6 Manx | 7 | 3 | 17 | 58 MT |

### ONLY AIR

| Ship | Base MT | Max containers | Container MT | Max capacity |
|------|---------|---------------|--------------|-------------|
| A-4 Wanderer | 1 | 1 | 17 | 18 MT |
| T-19 Stratomaster | 1 | 1 | 17 | 18 MT |

Capacity formula: `Base MT + (containers carried × 17 MT)`

---

## Flight times

Travel time = **5 min takeoff + flight + 10 min landing**.

Known flight times (flight minutes only, not including takeoff/landing):

**From Delois Spot:**
Alphaville 60 · Comstock 55 · Deadwood 60 · Ederar 60 · Erie 60 ·
Freedom 150 · Gettysburg 60 · Kansas 150 · Lancaster 120 · Pimli 35 ·
SovietUnion 60 · Terrazul 60 · Sharney 1 60 · Sharney 2 120 · Sharney 3 180

**From Kansas:**
Alphaville 35 · Comstock 30 · Deadwood 25 · Ederar 20 · Erie 45 ·
Freedom 15 · Gettysburg 40 · Lancaster 50 · Pimli 10 ·
SovietUnion 65 · Terrazul 60 · Sharney 1 30 · Sharney 2 60 · Sharney 3 90

For city pairs not in the table, the script estimates via Haversine coordinates
(if available) or scaled map-pixel distance.

---

## Rental cost

One in-game day = **14 hours**.  
If you rent a ship, the Config sheet shows the minimum profit per hour and per
day needed to cover the rental cost. The Opportunities sheet gains a
**Covers Rental?** column (green = yes, red = no).

---

## News PDF (`generate_news_pdf.py`)

The newspaper generator is a **separate script** and is no longer called
automatically by `ocr_to_excel.py`. Run it independently when needed:

```bash
python generate_news_pdf.py --workbook final_trade.xlsx
```

Output is written to `TVI_Output/The_Vieneo_Index.pdf` and each page is also
exported as a PNG image in the same folder.

Optional arguments:

```
--workbook PATH      Source workbook (default: final_trade.xlsx)
--output PATH        PDF output path
--header-image PATH  Custom masthead image
--image-output-dir   Folder for per-page PNG exports
```
