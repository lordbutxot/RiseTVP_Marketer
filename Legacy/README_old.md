# OCR to Excel Automation for Rise the Video Province

This project automates the extraction of text from resource/commodity images for cities and EasyDock in the game Rise the Video Province.

## Image Format
- **Cities**: Files named as `"CityName_1.png"` (e.g. `"Paris_1.png"`).
- **EasyDock**: `"EasyDock_1.png"`.
- The script automatically detects if the image is from a city or EasyDock and generates a separate sheet for each prefix.

## Requirements
1. **Install Tesseract OCR**:
   - Download and install Tesseract from: https://github.com/UB-Mannheim/tesseract/wiki
   - Make sure it's in the PATH or adjust the path in the script (`pytesseract.pytesseract.tesseract_cmd`).

2. **Python Environment**:
   - The virtual environment is already configured.
   - Installed packages: `pytesseract`, `openpyxl`, `pillow`.

## Usage
1. Place your images in the `images/` folder.
2. Run the script:
   ```
   cd "e:\PYTHON\RiseTVP\Script Trade"
   run_trade.bat
   ```
   The BAT file will prompt you to choose:
   - 1. Regular Trade: Analyzes all trade opportunities from all cities.
   - 2. City-Specific Opportunities: Analyzes all trade opportunities AND adds a sheet with opportunities starting from a specific city.

3. Follow the prompts to select ship, containers, rental, origin (if applicable), and budget.
4. The script will generate `final_trade.xlsx` with the sheets Config, Cities, EasyDock, Opportunities (always included), Opportunities from CityName (only if city mode chosen), Trade Routes, Profit Trips and ROI Analysis (if there is budget).

**With budget:** The ROI analysis is filtered to show only opportunities that fit within the budget.

The times include:
- 5 minutes takeoff
- Flight time between cities (see table below)
- 10 minutes landing

**From Delois Spot (flight time in minutes):**
- Alphaville: 60 | Comstock: 55 | Deadwood: 60 | Ederar: 60
- Erie: 60 | Freedom: 150 | Gettysburg: 60 | Kansas: 150
- Lancaster: 120 | Pimli: 35 | SovietUnion: 60 | Terrazul: 60
- Sharney 1: 60 | Sharney 2: 120 | Sharney 3: 180

**From Kansas (flight time in minutes):**
- Alphaville: 35 | Comstock: 30 | Deadwood: 25 | Ederar: 20
- Erie: 45 | Freedom: 15 | Gettysburg: 40 | Lancaster: 50
- Pimli: 10 | SovietUnion: 65 | Terrazul: 60
- Sharney 1: 30 | Sharney 2: 60 | Sharney 3: 90

## Chained Trips Recommendation

The system automatically identifies and recommends the best trip chains (A → B → C):
- After selling a commodity at point B, buy another in B and fly to C
- Chains are ordered by **efficiency (CR/hour)**
- Useful to maximize time and profit without returning to origin
- Top 10 are shown when generating the Excel

## ROI Analysis

ROI (Return on Investment) appears in the **Opportunities** sheet as a column with color coding:
- **ROI = (Total Profit / Total Cost) × 100%**

**Color Scale:**
- 🟢 **Dark Green**: ROI > 100% (Excellent return)
- 🟢 **Light Green**: ROI 50-100% (Very good)
- 🟡 **Yellow**: ROI 20-50% (Acceptable)
- 🟠 **Orange**: ROI < 20% (Low return)

**Example:** If you buy 10 MT at 100 CR/MT (cost 1000 CR) and earn 500 CR, the ROI is 50% (Light Green)

If you use `--budget`, opportunities that fit within your initial budget are filtered.

## Output Files

The script generates `final_trade.xlsx` with the following sheets:
- **Config**: Selected ship, capacity, origin, status (rented/purchased)
- **Cities**: Extracted data from cities (quantity, price, etc.)
- **EasyDock**: Extracted data from EasyDock
- **Opportunities**: Single-stop routes with flight times and **ROI (%)**
- **Trade Routes**: Routes grouped by destination with efficiency
- **Profit Trips**: Chained trip routes (A → B → C) ordered by efficiency (CR/hour)


## Ships and Rental Costs

All ships can be rented. Known rental costs:

**AIR AND SPACE:**
- E-10 Saint: Unknown rental cost
- E-11 Saint: Unknown rental cost
- P-13 Prowler: Unknown rental cost
- W-6 Manx: Unknown rental cost

**ONLY AIR:**
- A-4 Wanderer: Unknown rental cost (previously 1359 CR/day)
- T-19 Stratomaster: Unknown rental cost (previously 2264 CR/day)

If you choose to rent a ship with unknown cost, the script will ask you to enter the rental cost per day.

One day in the game = 14 hours.

If you choose to rent, the Config sheet will show the minimum profit needed per trip to cover the daily rental cost. The Opportunities sheet will have a "Covers Rental?" column indicating which trades generate enough profit to cover the rental.

## Run from Windows
You can use the `run_trade.bat` file to easily launch the script from Windows:
```bat
cd /d "e:\PYTHON\RiseTVP\Script Trade"
run_trade.bat
```
Optionally you can pass a ship and images folder:
```bat
run_trade.bat --ship "A-4 Wanderer" --images images --output final_trade.xlsx
```

## Config Sheet
- **Purpose**: Allows selecting the ship used for trade routes calculations.
- **Content**:
  - Selected Ship: The chosen ship (initially the one selected when running).
  - Capacity: Capacity in MT of the selected ship.
  - Available Ships: List of all ships with category and capacity.
  - Dropdown: In cell B3 there is a dropdown menu with all available ships. Change the selection and run the script again to recalculate with the new ship.

## Available Ships
The script includes ship selection with calculated capacities:

**AIR AND SPACE:**
- E-10 Saint: 109 MT (7 + 6x17)
- E-11 Saint: 109 MT (7 + 6x17)
- P-13 Prowler: 1 MT
- W-6 Manx: 58 MT (7 + 3x17)

**ONLY AIR:**
- A-4 Wanderer: 18 MT (1 + 1x17)
- T-19 Stratomaster: 18 MT (1 + 1x17)

The capacity is used to calculate profit per trip and trips needed in Trade Routes.
- **Cities Sheet**: Contains all cities with columns: Location, Category, Commodity Type, Quantity MT, Reserve MT, Selling CR/MT, Buying CR/MT, Maximum MT. Only includes rows with Quantity MT > 0 and Selling CR/MT > 0. Format: bold headers, alternating colored rows.
- **EasyDock Sheet**: Contains EasyDock with columns: Location, Category, Name, MT, Buying MT, Buying CR, Selling MT, Selling CR. Only includes rows with Selling CR > 0. Similar format.
- **Opportunities Sheet**: Lists trade opportunities by category, with columns: Category, Source, Source selling CR/MT, Destination, Destination buying CR/MT, Profit per MT, Max Qty, Total Profit. Ordered by total profit descending. High profits (>10,000) highlighted in yellow.
- **Trade Routes Sheet**: Proposes optimized routes grouped by source-destination, with commodities to transport, total qty, total profit, profit per trip (based on ship capacity) and trips needed.

## Opportunities Sheet
- Additionally, the script creates an `Opportunities` sheet in `output.xlsx`.
- It lists routes with the cheapest `selling` at point A and most expensive `buying` at point B.
- Columns: Category, Source, Source selling CR/MT, Destination, Destination buying CR/MT, Profit CR/MT.

## How the script interprets
- **City format**: detects columns like `Category`, `Commodity Type`, `Quantity MT`, `Reserve MT`, `Selling CR/MT`, `Buying CR/MT`, `Maximum MT`.
- **EasyDock format**: detects columns like `Category`, `Name`, `MT`, `Buying MT`, `Buying CR`, `Selling MT`, `Selling CR`.
- Commodity categories are assigned in fixed order: 1. Rare/Precious, 2. Foodstuffs, 3. Natural Materials, 4. Fuel Ore, 5. Consumer Goods, 6. Fabricated Material, 7. Refined Fuel.
- If it doesn't recognize a table format, it falls back to `Key: Value` parsing.

## Adjustments and improvements
- If you need another image format, modify `parse_text_to_data` in `ocr_to_excel.py`.
- To calculate route/trade opportunities, add post-analysis using the data in `output.xlsx` or extend the script to generate results.

## Notes
- Images must have clear and legible text for good OCR.
- If there are errors, check the Tesseract installation and photo quality.