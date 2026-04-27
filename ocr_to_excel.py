import argparse
import itertools
import math
import os
import platform
import re
import time
import pytesseract
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.pagebreak import Break as PageBreak

# ─────────────────────────────────────────────────────────────────────────────
# TESSERACT — auto-configure based on OS
# ─────────────────────────────────────────────────────────────────────────────
if platform.system() == "Windows":
    _TESSERACT_CANDIDATES = [
        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
    ]
    for _candidate in _TESSERACT_CANDIDATES:
        if os.path.isfile(_candidate):
            pytesseract.pytesseract.tesseract_cmd = _candidate
            break
# On Linux/macOS pytesseract finds tesseract automatically via PATH.

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

COMMODITY_CATEGORIES = [
    "Rare/Precious",
    "Foodstuffs",
    "Natural Materials",
    "Fuel Ore",
    "Consumer Goods",
    "Fabricated Material",
    "Refined Fuel"
]

SHIPS = {
    "AIR AND SPACE": {
        "E-10 Saint":   {"cargo_base": 7, "max_containers": 6, "container_mt": 17, "rental_cost_per_day": None},
        "E-11 Saint":   {"cargo_base": 7, "max_containers": 6, "container_mt": 17, "rental_cost_per_day": None},
        "P-13 Prowler": {"cargo_base": 1, "max_containers": 0, "container_mt": 17, "rental_cost_per_day": None},
        "W-6 Manx":     {"cargo_base": 7, "max_containers": 3, "container_mt": 17, "rental_cost_per_day": None},
    },
    "ONLY AIR": {
        "A-4 Wanderer":      {"cargo_base": 1, "max_containers": 1, "container_mt": 17, "rental_cost_per_day": None},
        "T-19 Stratomaster": {"cargo_base": 1, "max_containers": 1, "container_mt": 17, "rental_cost_per_day": None},
    }
}

CITIES = {
    "Alphaville":   {"x": 420, "y": 180},
    "Comstock":     {"x": 380, "y": 210},
    "Deadwood":     {"x": 340, "y": 230},
    "Ederar":       {"x": 300, "y": 200},
    "Erie":         {"x": 460, "y": 250},
    "Freedom":      {"x": 200, "y": 320},
    "Gettysburg":   {"x": 400, "y": 290},
    "Kansas":       {"x": 250, "y": 300},
    "Lancaster":    {"x": 350, "y": 350},
    "Pimli":        {"x": 270, "y": 280},
    "SovietUnion":  {"x": 500, "y": 150},
    "Terrazul":     {"x": 480, "y": 320},
    "Sharney 1":    {"x": 320, "y": 400},
    "Sharney 2":    {"x": 360, "y": 450},
    "Sharney 3":    {"x": 400, "y": 500},
    "Delois Spot":  {"x": 310, "y": 260},
}

CITY_LIST = sorted(CITIES.keys())

FLIGHT_TIMES_EXPLICIT = {
    ("Delois Spot", "Alphaville"): 60,
    ("Delois Spot", "Comstock"): 55,
    ("Delois Spot", "Deadwood"): 60,
    ("Delois Spot", "Ederar"): 60,
    ("Delois Spot", "Erie"): 60,
    ("Delois Spot", "Freedom"): 150,
    ("Delois Spot", "Gettysburg"): 60,
    ("Delois Spot", "Kansas"): 150,
    ("Delois Spot", "Lancaster"): 120,
    ("Delois Spot", "Pimli"): 35,
    ("Delois Spot", "Sharney 1"): 60,
    ("Delois Spot", "Sharney 2"): 120,
    ("Delois Spot", "Sharney 3"): 180,
    ("Delois Spot", "SovietUnion"): 60,
    ("Delois Spot", "Terrazul"): 60,
    ("Kansas", "Alphaville"): 35,
    ("Kansas", "Comstock"): 30,
    ("Kansas", "Deadwood"): 25,
    ("Kansas", "Ederar"): 20,
    ("Kansas", "Erie"): 45,
    ("Kansas", "Freedom"): 15,
    ("Kansas", "Gettysburg"): 40,
    ("Kansas", "Lancaster"): 50,
    ("Kansas", "Pimli"): 10,
    ("Kansas", "Sharney 1"): 30,
    ("Kansas", "Sharney 2"): 60,
    ("Kansas", "Sharney 3"): 90,
    ("Kansas", "SovietUnion"): 65,
    ("Kansas", "Terrazul"): 60,
}

TAKEOFF_TIME = 5
LANDING_TIME = 10

_REF_PAIRS = [
    (("Kansas", dst), t)
    for (src, dst), t in FLIGHT_TIMES_EXPLICIT.items()
    if src == "Kansas" and dst in CITIES and "Kansas" in CITIES
]
_SPEEDS = []
for (src, dst), t in _REF_PAIRS:
    dx = CITIES[src]["x"] - CITIES[dst]["x"]
    dy = CITIES[src]["y"] - CITIES[dst]["y"]
    dist = math.sqrt(dx * dx + dy * dy)
    if t > 0:
        _SPEEDS.append(dist / t)
SPEED_UNITS_PER_MIN = sum(_SPEEDS) / len(_SPEEDS) if _SPEEDS else 3.0

CITY_COLORS = {
    'Alphaville': 'FFC000',
    'Comstock': '00B0F0',
    'Deadwood': '92D050',
    'Ederar': '7030A0',
    'Erie': '00B050',
    'Freedom': 'C00000',
    'Gettysburg': 'FFFF00',
    'Kansas': '0066CC',
    'Lancaster': 'FF00FF',
    'Pimli': '00B0F0',
    'SovietUnion': '7030A0',
    'Terrazul': '00B050',
    'Sharney 1': 'FFC000',
    'Sharney 2': '92D050',
    'Sharney 3': '0070C0',
    'Delois Spot': 'C00000',
}

COMMODITY_COLORS = {
    "Rare/Precious":       "FFD700",
    "Foodstuffs":          "32CD32",
    "Natural Materials":   "1E90FF",
    "Fuel Ore":            "708090",
    "Consumer Goods":      "DC143C",
    "Fabricated Material": "8A2BE2",
    "Refined Fuel":        "FFA500"
}

GRADE_STYLES = {
    "A": {"fill": "00B050", "font": "FFFFFF"},
    "B": {"fill": "70AD47", "font": "FFFFFF"},
    "C": {"fill": "FFD700", "font": "000000"},
    "D": {"fill": "FFA500", "font": "FFFFFF"},
}

# ─────────────────────────────────────────────────────────────────────────────
# FILE FILTER
# ─────────────────────────────────────────────────────────────────────────────

def is_ignored_source(base_name: str) -> bool:
    """
    Returns True for files that should NOT be treated as trade city sources.
    Currently ignores anything whose base name starts with 'TVI_' (case-insensitive).
    Add more patterns here as needed.
    """
    upper = base_name.upper()
    return upper.startswith("TVI_") or upper == "TVI"

# ─────────────────────────────────────────────────────────────────────────────
# FLIGHT / DISTANCE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _coord_distance(city_a: str, city_b: str) -> float:
    a = CITIES.get(city_a)
    b = CITIES.get(city_b)
    if not a or not b:
        return 0.0
    dx = a["x"] - b["x"]
    dy = a["y"] - b["y"]
    return math.sqrt(dx * dx + dy * dy)


def get_flight_time(origin: str, destination: str) -> int:
    """Total flight time including takeoff and landing (minutes)."""
    if origin == destination:
        return 0
    flight_mins = (
        FLIGHT_TIMES_EXPLICIT.get((origin, destination))
        or FLIGHT_TIMES_EXPLICIT.get((destination, origin))
    )
    if flight_mins is None:
        dist = _coord_distance(origin, destination)
        flight_mins = round(dist / SPEED_UNITS_PER_MIN) if dist > 0 else 60
    return TAKEOFF_TIME + flight_mins + LANDING_TIME


def calculate_ship_capacity(ship_name: str, containers_used: int = None) -> int:
    for category, ships in SHIPS.items():
        if ship_name in ships:
            d = ships[ship_name]
            n = containers_used if containers_used is not None else d["max_containers"]
            return d["cargo_base"] + n * d["container_mt"]
    return 0

# ─────────────────────────────────────────────────────────────────────────────
# OCR
# ─────────────────────────────────────────────────────────────────────────────

def extract_text_from_image(image_path: str) -> str:
    try:
        img = Image.open(image_path)
        custom_config = r'--oem 3 --psm 4'
        text = pytesseract.image_to_string(img, config=custom_config)
        return text.strip()
    except Exception as e:
        print(f"Error processing {image_path}: {e}")
        return ""

# ─────────────────────────────────────────────────────────────────────────────
# PARSING
# ─────────────────────────────────────────────────────────────────────────────

def normalize_number(token):
    if token is None: return 0
    if isinstance(token, (int, float)): return int(token)
    text = re.sub(r'[^\d]', '', str(token))
    try:
        return int(text) if text else 0
    except:
        return 0


def is_total_line(line: str) -> bool:
    lower = line.lower()
    return any(w in lower for w in ['totals', 'total', 'refresh', 'cancel',
                                     'population', 'fees', 'ports staffed',
                                     'mt free', 'cr free'])


def detect_layout(lines):
    joined = ' '.join(lines).lower()
    if any(k in joined for k in ['commodity type', 'reserve mt', 'selling cr/mt',
                                  'buying cr/mt', 'maximum mt']):
        return 'city'
    if any(k in joined for k in ['commodities', 'buying cr', 'selling cr']) and 'name' in joined:
        return 'easydock'
    if any(k in joined for k in ['buying mt', 'selling mt', 'buying cr', 'selling cr']):
        return 'easydock'
    return 'simple'


def parse_table_rows(text_or_lines, min_numbers=3, layout='city'):
    rows = []
    lines = text_or_lines if isinstance(text_or_lines, list) else text_or_lines.split('\n')

    for line in lines:
        line = line.strip()
        if not line: continue

        tokens = line.split()
        numbers_in_line = []
        name_parts = []

        for t in tokens:
            if any(char.isdigit() for char in t):
                numbers_in_line.append(t)
            else:
                name_parts.append(t)

        if len(numbers_in_line) >= 3:
            stock = normalize_number(numbers_in_line[-1])
            sell  = normalize_number(numbers_in_line[-2])
            buy   = normalize_number(numbers_in_line[-3])

            name = " ".join(name_parts)
            name = re.sub(r'[^a-zA-Z\s]', '', name).strip()

            if len(name) < 3 or name.upper() in [c.upper() for c in COMMODITY_CATEGORIES]:
                continue

            rows.append([name, buy, sell, stock])

    return rows


def infer_category_from_text(raw_text: str) -> str:
    if not raw_text or not isinstance(raw_text, str):
        return ''
    text = re.sub(r'^[^A-Za-z]*', '', raw_text).strip()
    for cat in COMMODITY_CATEGORIES:
        if cat.lower() in text.lower():
            return cat
    return text


def infer_commodity_type(raw_text: str) -> str:
    if not raw_text or not isinstance(raw_text, str):
        return ''
    text = re.sub(r'^[^A-Za-z]*', '', raw_text).strip()
    for cat in COMMODITY_CATEGORIES:
        if cat.lower() in text.lower():
            return cat
    return text


def parse_text_to_data(text: str) -> dict:
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    layout = detect_layout(lines)

    if layout == 'city':
        header = ['Category', 'Commodity Type', 'Quantity MT', 'Reserve MT',
                  'Selling CR/MT', 'Buying CR/MT', 'Maximum MT']
        rows = parse_table_rows(lines, min_numbers=5, layout='city')
        for row in rows:
            raw_name = row[0]
            category = infer_category_from_text(raw_name)
            commodity = infer_commodity_type(raw_name)
            row[0] = category
            row.insert(1, commodity)
    elif layout == 'easydock':
        header = ['Category', 'Name', 'MT', 'Buying MT', 'Buying CR',
                  'Selling MT', 'Selling CR']
        rows = parse_table_rows(lines, min_numbers=5, layout='easydock')
        for row in rows:
            raw_name = row[0]
            category = infer_category_from_text(raw_name)
            commodity = infer_commodity_type(raw_name)
            row[0] = category
            row.insert(1, commodity)
    else:
        header = ['Recurso/Commodity', 'Valor']
        rows = []
        for line in lines:
            if ':' in line:
                key, value = line.split(':', 1)
                rows.append([key.strip(), value.strip()])
            else:
                rows.append([line])

    return {'header': header, 'rows': rows}


def sanitize_sheet_name(name: str) -> str:
    invalid = ['\\', '/', '?', '*', '[', ']', ':']
    safe = ''.join('-' if ch in invalid else ch for ch in name)
    return safe[:31]


def normalize_commodity_key(category: str, name: str = None) -> str:
    if category and isinstance(category, str) and category.strip():
        return category.strip()
    if name and isinstance(name, str):
        cleaned = re.sub(r'^[^A-Za-z]*', '', name).strip()
        for cat in COMMODITY_CATEGORIES:
            if cat.lower() in cleaned.lower():
                return cat
        return cleaned
    return ''

# ─────────────────────────────────────────────────────────────────────────────
# STYLE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _contrast_text_color(hex_color: str) -> str:
    hex_color = hex_color.strip().lstrip('#')
    if len(hex_color) != 6:
        return '000000'
    try:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
    except ValueError:
        return '000000'
    brightness = (r * 299 + g * 587 + b * 114) / 1000
    return '000000' if brightness > 160 else 'FFFFFF'


def get_city_fill(city: str):
    color = CITY_COLORS.get(city, 'D9D9D9')
    return PatternFill(start_color=color, end_color=color, fill_type='solid')


def _style_header_row(cells):
    for cell in cells:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin")
        )


def _style_section_title(cells):
    for cell in cells:
        cell.font = Font(bold=True, size=14, color="000000")
        cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin")
        )


def _style_subsection_title(cells, color="E6E6FA"):
    for cell in cells:
        cell.font = Font(bold=True, size=12, color="000000")
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin")
        )


def _style_commodity_separator(cells, color):
    for cell in cells:
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        cell.border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thick"), bottom=Side(style="thick")
        )


def _auto_size_columns(ws, min_width=10, max_width=50):
    from openpyxl.cell.cell import MergedCell
    for col_cells in ws.columns:
        if isinstance(col_cells[0], MergedCell):
            continue
        max_len = max(
            (len(str(cell.value)) if cell.value is not None and not isinstance(cell, MergedCell) else 0)
            for cell in col_cells
        )
        ws.column_dimensions[col_cells[0].column_letter].width = min(
            max_width, max(min_width, max_len + 2)
        )


def _apply_roi_color(cell, roi_pct):
    font_white = Font(color="FFFFFF", bold=True)
    font_black = Font(color="000000", bold=True)
    if roi_pct > 100:
        cell.fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
        cell.font = font_white
    elif roi_pct >= 50:
        cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
        cell.font = font_white
    elif roi_pct >= 20:
        cell.fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
        cell.font = font_black
    elif roi_pct >= 0:
        cell.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
        cell.font = font_white
    else:
        cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        cell.font = font_white


def _add_page_break(ws):
    """Insert a manual page break after the current last row."""
    row_num = ws.max_row  # break goes AFTER this row (before max_row+1)
    ws.row_breaks.append(PageBreak(id=row_num, man=True))


# ─────────────────────────────────────────────────────────────────────────────
# CATALOG & OPPORTUNITIES
# ─────────────────────────────────────────────────────────────────────────────

def build_trade_catalog(data_dict: dict) -> dict:
    catalog = {}
    for sheet_name in ['Cities', 'EasyDock']:
        if sheet_name not in data_dict:
            continue
        header = [h.lower() for h in data_dict[sheet_name]['header']]
        rows = data_dict[sheet_name]['rows']
        for row in rows:
            if len(row) < 3:
                continue
            location = str(row[0]).strip()
            category = str(row[1]).strip()
            raw_name = str(row[2]).strip() if len(row) > 2 else ''
            commodity_key = normalize_commodity_key(category, raw_name)
            if location not in catalog:
                catalog[location] = {}
            item_map = catalog[location]

            selling = buying = None
            quantity = reserve = max_accept = sell_capacity = buy_capacity = 0

            if sheet_name == 'Cities' and 'selling cr/mt' in header:
                si = header.index('selling cr/mt')
                bi = header.index('buying cr/mt')
                qi = header.index('quantity mt')
                ri = header.index('reserve mt')
                mi = header.index('maximum mt')
                selling    = normalize_number(row[si]) if len(row) > si else None
                buying     = normalize_number(row[bi]) if len(row) > bi else None
                quantity   = normalize_number(row[qi]) if len(row) > qi else 0
                reserve    = normalize_number(row[ri]) if len(row) > ri else 0
                max_accept = normalize_number(row[mi]) if len(row) > mi else 0
                sell_capacity = max(quantity - reserve, 0)
                buy_capacity  = max_accept

            elif sheet_name == 'EasyDock' and 'selling cr' in header:
                si  = header.index('selling cr')
                bi  = header.index('buying cr')
                qi  = header.index('mt')
                bci = header.index('buying mt')
                sci = header.index('selling mt')
                selling      = normalize_number(row[si])  if len(row) > si  else None
                buying       = normalize_number(row[bi])  if len(row) > bi  else None
                quantity     = normalize_number(row[qi])  if len(row) > qi  else 0
                buy_capacity = normalize_number(row[bci]) if len(row) > bci else 0
                sell_capacity= normalize_number(row[sci]) if len(row) > sci else 0

            item_map[commodity_key] = {
                'selling': selling,
                'buying': buying,
                'quantity': quantity,
                'reserve': reserve,
                'max_accept': max_accept,
                'sell_capacity': sell_capacity,
                'buy_capacity': buy_capacity,
            }
    return catalog


def find_trade_opportunities(catalog: dict) -> list:
    opportunities = []
    locations = list(catalog.keys())
    for src in locations:
        for dst in locations:
            if src == dst:
                continue
            for commodity, src_data in catalog[src].items():
                if commodity not in catalog[dst]:
                    continue
                dst_data = catalog[dst][commodity]
                if src_data['selling'] is None or dst_data['buying'] is None:
                    continue
                source_available     = src_data.get('sell_capacity', 0) or 0
                destination_capacity = dst_data.get('buy_capacity', 0) or 0
                profit_per_mt = dst_data['buying'] - src_data['selling']
                if profit_per_mt <= 0:
                    continue
                max_qty = (
                    min(source_available, destination_capacity)
                    if source_available > 0 and destination_capacity > 0
                    else 0
                )
                total_profit = profit_per_mt * max_qty if max_qty > 0 else 0
                opportunities.append({
                    'commodity': commodity,
                    'source': src,
                    'source_selling': src_data['selling'],
                    'source_available': source_available,
                    'destination': dst,
                    'destination_buying': dst_data['buying'],
                    'destination_capacity': destination_capacity,
                    'profit_per_mt': profit_per_mt,
                    'max_qty': max_qty,
                    'total_profit': total_profit,
                })
    opportunities.sort(key=lambda x: x['total_profit'], reverse=True)
    return opportunities

# ─────────────────────────────────────────────────────────────────────────────
# GRADING
# ─────────────────────────────────────────────────────────────────────────────

def assign_grades(
    opportunities: list,
    ship_capacity: int = 109,
    budget: int = None,
    is_rental: bool = False,
    rental_cost_per_day: int = None
) -> list:
    for op in opportunities:
        price         = op["source_selling"] or 1
        max_by_ship   = ship_capacity
        max_by_source = op["source_available"]
        max_by_dest   = op["destination_capacity"]
        max_by_budget = int(budget / price) if budget else float("inf")

        qty = min(max_by_ship, max_by_source, max_by_dest, max_by_budget)
        qty = max(qty, 0)

        cost_trip   = qty * price
        profit_trip = qty * op["profit_per_mt"]

        rental_cost_per_trip = 0
        if is_rental and rental_cost_per_day:
            flight_time = get_flight_time(op['source'], op['destination'])
            total_trip_time = flight_time + 15
            rental_cost_per_trip = (total_trip_time / 60) * (rental_cost_per_day / 14)
            cost_trip += rental_cost_per_trip

        roi_trip = (profit_trip / cost_trip * 100) if cost_trip > 0 else 0

        op["_qty_trip"]    = qty
        op["_cost_trip"]   = cost_trip
        op["_profit_trip"] = profit_trip
        op["_roi"]         = roi_trip
        op["_rental_cost"] = rental_cost_per_trip
        if is_rental and rental_cost_per_day:
            op["covers_rental"] = profit_trip >= rental_cost_per_trip
        op["_affordable"] = qty > 0 and profit_trip > 0

    viable        = [op for op in opportunities if op["_affordable"]]
    viable_sorted = sorted(viable, key=lambda x: x["_profit_trip"], reverse=True)
    n = len(viable_sorted)

    rank_map = {}
    for idx, op in enumerate(viable_sorted):
        pct = idx / n if n > 0 else 1
        if pct < 0.25:
            rank_map[id(op)] = "A"
        elif pct < 0.50:
            rank_map[id(op)] = "B"
        elif pct < 0.75:
            rank_map[id(op)] = "C"
        else:
            rank_map[id(op)] = "D"

    for op in opportunities:
        op["grade"] = rank_map.get(id(op), "D")

    return opportunities

# ─────────────────────────────────────────────────────────────────────────────
# TRADE ROUTE ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def find_best_leg(
    catalog: dict,
    src: str,
    dst: str,
    allowed_commodities: list,
    ship_capacity: int,
    budget: float,
) -> dict | None:
    best = None
    src_data = catalog.get(src, {})
    dst_data = catalog.get(dst, {})

    for commodity in allowed_commodities:
        if commodity not in src_data or commodity not in dst_data:
            continue
        s = src_data[commodity]
        d = dst_data[commodity]
        if s['selling'] is None or d['buying'] is None:
            continue
        profit_per_mt = d['buying'] - s['selling']
        if profit_per_mt <= 0:
            continue

        available = s.get('sell_capacity', 0) or 0
        capacity  = d.get('buy_capacity', 0) or 0
        if available <= 0 or capacity <= 0:
            continue

        max_by_budget = int(budget / s['selling']) if budget and s['selling'] > 0 else float("inf")
        qty = min(ship_capacity, available, capacity, max_by_budget)
        qty = max(qty, 0)
        if qty <= 0:
            continue

        cost   = qty * s['selling']
        profit = qty * profit_per_mt
        flight = get_flight_time(src, dst)
        roi    = (profit / cost * 100) if cost > 0 else 0

        if best is None or profit > best['profit']:
            best = {
                'commodity':    commodity,
                'src':          src,
                'dst':          dst,
                'buy_price':    s['selling'],
                'sell_price':   d['buying'],
                'profit_per_mt':profit_per_mt,
                'qty':          qty,
                'cost':         cost,
                'profit':       profit,
                'flight_min':   flight,
                'roi':          roi,
            }
    return best


def compute_trade_routes(
    catalog: dict,
    origin: str,
    allowed_commodities: list,
    ship_capacity: int,
    budget: float,
    max_hops: int = 5,
    top_n: int = 20,
) -> list:
    locations = [loc for loc in catalog.keys() if loc != origin]
    all_routes = []

    for n_stops in range(1, max_hops + 1):
        for stops in itertools.permutations(locations, n_stops):
            city_sequence = [origin] + list(stops) + [origin]
            legs = []
            total_profit  = 0
            total_cost    = 0
            total_time    = 0
            remaining_budget = budget if budget else None
            valid = True

            for i in range(len(city_sequence) - 1):
                src = city_sequence[i]
                dst = city_sequence[i + 1]
                leg = find_best_leg(
                    catalog, src, dst, allowed_commodities,
                    ship_capacity,
                    remaining_budget,
                )
                if leg is None:
                    if i == len(city_sequence) - 2:
                        legs.append({
                            'commodity':    '— empty —',
                            'src':          src,
                            'dst':          dst,
                            'buy_price':    0,
                            'sell_price':   0,
                            'profit_per_mt':0,
                            'qty':          0,
                            'cost':         0,
                            'profit':       0,
                            'flight_min':   get_flight_time(src, dst),
                            'roi':          0,
                        })
                        total_time += get_flight_time(src, dst)
                    else:
                        valid = False
                        break
                else:
                    legs.append(leg)
                    total_profit += leg['profit']
                    total_cost   += leg['cost']
                    total_time   += leg['flight_min']
                    if remaining_budget is not None:
                        remaining_budget = remaining_budget - leg['cost'] + leg['profit']

            if not valid or total_profit <= 0:
                continue

            roi = (total_profit / total_cost * 100) if total_cost > 0 else 0
            cr_per_hour = (total_profit / total_time * 60) if total_time > 0 else 0

            all_routes.append({
                'stops':       city_sequence,
                'legs':        legs,
                'n_hops':      len(legs),
                'total_profit':total_profit,
                'total_cost':  total_cost,
                'total_time':  total_time,
                'roi':         roi,
                'cr_per_hour': cr_per_hour,
            })

    seen = set()
    unique_routes = []
    for r in sorted(all_routes, key=lambda x: x['total_profit'], reverse=True):
        key = tuple(r['stops'])
        if key not in seen:
            seen.add(key)
            unique_routes.append(r)
        if len(unique_routes) >= top_n:
            break

    return unique_routes

# ─────────────────────────────────────────────────────────────────────────────
# TRADE ROUTE SHEET WRITER
# ─────────────────────────────────────────────────────────────────────────────

def _write_trade_routes_sheet(ws, routes: list, origin: str):
    font_white = Font(color="FFFFFF", bold=True)
    alt_fill   = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")

    if not routes:
        ws.append(['No profitable routes found from', origin])
        return

    for rank, route in enumerate(routes, 1):
        # Page break before every route except the first
        if rank > 1:
            _add_page_break(ws)

        summary_label = (
            f"Route #{rank}  |  "
            f"{' → '.join(route['stops'])}  |  "
            f"Profit: {route['total_profit']:,.0f} CR  |  "
            f"ROI: {route['roi']:.1f}%  |  "
            f"Time: {route['total_time']} min  |  "
            f"Efficiency: {route['cr_per_hour']:,.0f} CR/h"
        )
        ws.append([summary_label])
        title_row = ws.max_row
        ws.merge_cells(
            start_row=title_row, start_column=1,
            end_row=title_row, end_column=10
        )
        title_cell = ws.cell(row=title_row, column=1)
        title_cell.font  = Font(bold=True, size=12, color="FFFFFF")
        title_cell.fill  = get_city_fill(origin)
        title_cell.alignment = Alignment(horizontal="left", vertical="center")

        leg_cols = [
            'Hop', 'From', 'To', 'Commodity',
            'Buy (CR/MT)', 'Sell (CR/MT)', 'Profit/MT',
            'MT', 'Leg Cost (CR)', 'Leg Profit (CR)', 'Flight (min)', 'Leg ROI (%)'
        ]
        ws.append(leg_cols)
        _style_header_row(ws[ws.max_row])

        for hop_i, leg in enumerate(route['legs'], 1):
            row_data = [
                hop_i,
                leg['src'],
                leg['dst'],
                leg['commodity'],
                leg['buy_price'],
                leg['sell_price'],
                leg['profit_per_mt'],
                leg['qty'],
                leg['cost'],
                leg['profit'],
                leg['flight_min'],
                round(leg['roi'], 2),
            ]
            ws.append(row_data)
            cur = ws.max_row

            src_cell = ws.cell(row=cur, column=2)
            src_fill = get_city_fill(leg['src'])
            src_cell.fill = src_fill
            src_cell.font = Font(color=_contrast_text_color(src_fill.start_color.rgb), bold=True)

            dst_cell = ws.cell(row=cur, column=3)
            dst_fill = get_city_fill(leg['dst'])
            dst_cell.fill = dst_fill
            dst_cell.font = Font(color=_contrast_text_color(dst_fill.start_color.rgb), bold=True)

            roi_cell = ws.cell(row=cur, column=12)
            _apply_roi_color(roi_cell, leg['roi'])

            if hop_i % 2 == 0:
                for col in [1, 4, 5, 6, 7, 8, 9, 10, 11]:
                    ws.cell(row=cur, column=col).fill = alt_fill

        ws.append([
            '', '', '', 'TOTAL',
            '', '', '',
            '',
            route['total_cost'],
            route['total_profit'],
            route['total_time'],
            round(route['roi'], 2),
        ])
        tot_row = ws.max_row
        for col in range(1, 13):
            c = ws.cell(row=tot_row, column=col)
            c.font = Font(bold=True)
            c.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        _apply_roi_color(ws.cell(row=tot_row, column=12), route['roi'])

        ws.append([])

    ws.freeze_panes = "A2"
    _auto_size_columns(ws)

# ─────────────────────────────────────────────────────────────────────────────
# MACRO DATA
# ─────────────────────────────────────────────────────────────────────────────

def calculate_macro_data(catalog: dict, opportunities: list) -> dict:
    macro = {
        'city_summary': [],
        'commodity_best_sellers': {},
        'commodity_best_buyers': {},
        'top_cities_by_profit': [],
        'cities_by_lucrative_commodities': []
    }

    city_profits     = {}
    city_commodities = {}
    city_lucrative   = {}

    for op in opportunities:
        src       = op['source']
        profit    = op['total_profit']
        commodity = op['commodity']

        if src not in city_profits:
            city_profits[src]     = 0
            city_commodities[src] = set()
            city_lucrative[src]   = set()

        city_profits[src]     += profit
        city_commodities[src].add(commodity)
        if profit > 0:
            city_lucrative[src].add(commodity)

    for city in catalog:
        macro['city_summary'].append({
            'city':            city,
            'total_profit':    city_profits.get(city, 0),
            'num_commodities': len(city_commodities.get(city, set())),
            'num_lucrative':   len(city_lucrative.get(city, set())),
        })

    macro['city_summary'].sort(key=lambda x: x['total_profit'], reverse=True)
    macro['top_cities_by_profit'] = macro['city_summary'][:10]
    macro['cities_by_lucrative_commodities'] = sorted(
        macro['city_summary'], key=lambda x: x['num_lucrative'], reverse=True
    )[:10]

    for commodity in COMMODITY_CATEGORIES:
        sellers, buyers = [], []
        for city, data in catalog.items():
            if commodity in data:
                item = data[commodity]
                if item['selling'] is not None and item['sell_capacity'] > 0:
                    sellers.append({'city': city, 'price': item['selling'], 'capacity': item['sell_capacity']})
                if item['buying'] is not None and item['buy_capacity'] > 0:
                    buyers.append({'city': city, 'price': item['buying'], 'capacity': item['buy_capacity']})
        sellers.sort(key=lambda x: x['price'])
        buyers.sort(key=lambda x: x['price'], reverse=True)
        macro['commodity_best_sellers'][commodity] = sellers[:5]
        macro['commodity_best_buyers'][commodity]  = buyers[:5]

    return macro

# ─────────────────────────────────────────────────────────────────────────────
# OPPORTUNITIES SHEET HELPER
# ─────────────────────────────────────────────────────────────────────────────

def _build_opp_columns(is_rental, rental_cost_per_day):
    cols = [
        'Grade', 'Commodity', 'Source', 'Buy Price (CR/MT)',
        'Destination', 'Sell Price (CR/MT)', 'Profit/MT (CR)',
        'Src Stock (MT)', 'Dst Capacity (MT)', 'MT loaded per trip',
        'Trip Cost (CR)', 'Trip Profit (CR)',
    ]
    if is_rental and rental_cost_per_day:
        cols.append('Rental Cost (CR)')
    cols.append('Trip ROI (%)')
    if is_rental and rental_cost_per_day:
        cols.append('Covers Rental?')
    return cols


def _write_opportunities_sheet(
    ws, opportunities, columns, ship_capacity,
    is_rental, rental_cost_per_day, origin,
    GRADE_COL, ROI_COL
):
    font_white = Font(color="FFFFFF", bold=True)
    font_black = Font(color="000000", bold=True)
    green_fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
    red_fill   = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    ws.append(columns)
    _style_header_row(ws[1])

    for op in opportunities:
        grade       = op['grade']
        qty_trip    = op['_qty_trip']
        cost_trip   = op['_cost_trip']
        profit_trip = op['_profit_trip']
        roi_pct     = op['_roi']

        row = [
            grade,
            op['commodity'],
            op['source'],
            op['source_selling'],
            op['destination'],
            op['destination_buying'],
            op['profit_per_mt'],
            op['source_available'],
            op['destination_capacity'],
            qty_trip,
            cost_trip,
            profit_trip,
        ]
        if is_rental and rental_cost_per_day:
            row.append(round(op['_rental_cost'], 2))
        row.append(round(roi_pct, 2))
        if is_rental and rental_cost_per_day:
            row.append("YES" if op.get('covers_rental', False) else "NO")

        ws.append(row)
        cur_row = ws.max_row

        gs = GRADE_STYLES[grade]
        gc = ws.cell(row=cur_row, column=GRADE_COL)
        gc.fill = PatternFill(start_color=gs["fill"], end_color=gs["fill"], fill_type="solid")
        gc.font = Font(color=gs["font"], bold=True)
        gc.alignment = Alignment(horizontal="center")

        source_cell = ws.cell(row=cur_row, column=3)
        source_fill = get_city_fill(op['source'])
        source_cell.fill = source_fill
        source_cell.font = Font(color=_contrast_text_color(source_fill.start_color.rgb), bold=True)

        roi_cell = ws.cell(row=cur_row, column=ROI_COL)
        _apply_roi_color(roi_cell, roi_pct)

        if is_rental and rental_cost_per_day:
            rc = ws.cell(row=cur_row, column=len(columns))
            if rc.value == "YES":
                rc.fill = green_fill; rc.font = font_white
            else:
                rc.fill = red_fill;   rc.font = font_white

    ws.freeze_panes = "B2"
    _auto_size_columns(ws)

# ─────────────────────────────────────────────────────────────────────────────
# SAVE TO EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def save_to_excel(
    data_dict,
    selected_ship,
    ship_capacity,
    output_file='final_trade.xlsx',
    is_rental=False,
    rental_cost_per_day=None,
    origin="Delois Spot",
    budget=None,
    containers_used=None,
    mode='regular',
    trade_route_params=None,
):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    alt_fill  = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
    alt_fill2 = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")

    # ── CONFIG ────────────────────────────────────────────────────────────────
    ws_cfg = wb.create_sheet(title='Config')
    ws_cfg.append(['Parameter', 'Value'])
    _style_header_row(ws_cfg[1])

    ship_data     = None
    ship_category = ''
    for cat, ships in SHIPS.items():
        if selected_ship in ships:
            ship_data     = ships[selected_ship]
            ship_category = cat
            break

    max_c   = ship_data["max_containers"] if ship_data else 0
    c_used  = containers_used if containers_used is not None else max_c
    c_mt    = ship_data["container_mt"] if ship_data else 17
    cargo_b = ship_data["cargo_base"] if ship_data else 0

    ws_cfg.append(['Ship',                   selected_ship])
    ws_cfg.append(['Category',               ship_category])
    ws_cfg.append(['Base cargo (fixed MT)',  cargo_b])
    ws_cfg.append(['Containers carried',     f'{c_used} / {max_c} max'])
    ws_cfg.append(['MT per container',       c_mt])
    ws_cfg.append(['Total Capacity (MT)',    f'{cargo_b} + {c_used} x {c_mt} = {ship_capacity}'])
    ws_cfg.append(['Origin',                 origin])
    ws_cfg.append(['OS',                     platform.system()])
    ws_cfg.append(['Status',                 'Rented' if is_rental else 'Purchased'])
    if is_rental and rental_cost_per_day:
        ws_cfg.append(['Rental cost/day',            f'{rental_cost_per_day} CR/day (14 h)'])
        ws_cfg.append(['Min profit/day to cover rental', f'{rental_cost_per_day} CR'])
        ws_cfg.append(['Min profit/hour to cover rental', f'{rental_cost_per_day / 14:.2f} CR/h'])
    if budget:
        ws_cfg.append(['Initial Budget', f'{budget} CR'])
    if trade_route_params:
        ws_cfg.append(['Mode',               'Trade Route'])
        ws_cfg.append(['Max hops',           trade_route_params.get('max_hops', 5)])
        ws_cfg.append(['Allowed commodities',', '.join(trade_route_params.get('allowed_commodities', ['ALL']))])

    _auto_size_columns(ws_cfg)

    # ── CITIES / EASYDOCK ─────────────────────────────────────────────────────
    for sheet_name in ['Cities', 'EasyDock']:
        if sheet_name in data_dict and data_dict[sheet_name]['rows']:
            ws = wb.create_sheet(title=sheet_name)
            ws.append(data_dict[sheet_name]['header'])
            _style_header_row(ws[1])
            for i, row in enumerate(data_dict[sheet_name]['rows'], start=2):
                ws.append(row)
                if i % 2 == 0:
                    for cell in ws[i]:
                        cell.fill = alt_fill
            _auto_size_columns(ws)

    # ── CATALOG & OPPORTUNITIES ───────────────────────────────────────────────
    catalog           = build_trade_catalog(data_dict)
    all_opportunities = find_trade_opportunities(catalog)

    if all_opportunities:
        all_opportunities = assign_grades(
            all_opportunities,
            ship_capacity=ship_capacity,
            budget=budget,
            is_rental=is_rental,
            rental_cost_per_day=rental_cost_per_day,
        )
        all_opportunities.sort(key=lambda x: x["_profit_trip"], reverse=True)

        columns   = _build_opp_columns(is_rental, rental_cost_per_day)
        GRADE_COL = 1
        ROI_COL   = columns.index('Trip ROI (%)') + 1

        ws_op = wb.create_sheet(title='Opportunities')
        _write_opportunities_sheet(
            ws_op, all_opportunities, columns, ship_capacity,
            is_rental, rental_cost_per_day, origin, GRADE_COL, ROI_COL,
        )

        if mode == 'city':
            city_opps = [op for op in all_opportunities if op['source'] == origin]
            if city_opps:
                ws_city = wb.create_sheet(title=sanitize_sheet_name(f'From {origin}'))
                _write_opportunities_sheet(
                    ws_city, city_opps, columns, ship_capacity,
                    is_rental, rental_cost_per_day, origin, GRADE_COL, ROI_COL,
                )

    # ── TRADE ROUTES ──────────────────────────────────────────────────────────
    if mode == 'route' and trade_route_params:
        allowed  = trade_route_params.get('allowed_commodities', COMMODITY_CATEGORIES)
        max_hops = trade_route_params.get('max_hops', 3)

        print(f"\n  Computing trade routes from {origin} (max {max_hops} hops)...")
        print(f"  Commodities: {', '.join(allowed)}")
        print(f"  This may take a moment...")

        routes = compute_trade_routes(
            catalog=catalog,
            origin=origin,
            allowed_commodities=allowed,
            ship_capacity=ship_capacity,
            budget=budget,
            max_hops=max_hops,
            top_n=20,
        )

        sheet_title = sanitize_sheet_name(f'Trade Routes {origin}')
        ws_routes   = wb.create_sheet(title=sheet_title)
        _write_trade_routes_sheet(ws_routes, routes, origin)

        print(f"\n  Found {len(routes)} profitable routes.")
        if routes:
            best = routes[0]
            print(f"  Best route: {' → '.join(best['stops'])}")
            print(f"  Profit: {best['total_profit']:,.0f} CR | ROI: {best['roi']:.1f}% | Time: {best['total_time']} min")

    # ── MACRO ─────────────────────────────────────────────────────────────────
    opps_for_macro = all_opportunities if all_opportunities else []
    macro_data     = calculate_macro_data(catalog, opps_for_macro)
    ws_macro       = wb.create_sheet(title='MACRO')

    # Enable sheet-level print settings so page breaks are respected
    ws_macro.page_setup.fitToPage = False
    ws_macro.sheet_properties.pageSetUpPr.fitToPage = False

    # ── Section 1: Summary by City ────────────────────────────────────────────
    ws_macro.append(['SUMMARY BY CITY'])
    _style_section_title(ws_macro[1])
    ws_macro.append(['City', 'Total Potential Profit (CR)', 'Num Commodities', 'Num Lucrative Commodities'])
    _style_header_row(ws_macro[2])
    for i, item in enumerate(macro_data['city_summary'], start=3):
        ws_macro.append([item['city'], item['total_profit'], item['num_commodities'], item['num_lucrative']])
        if i % 2 == 0:
            for cell in ws_macro[i]: cell.fill = alt_fill2

    # ── Section 2: Top 10 by profit ───────────────────────────────────────────
    ws_macro.append([])
    _add_page_break(ws_macro)
    ws_macro.append(['TOP 10 CITIES BY TOTAL PROFIT'])
    _style_section_title(ws_macro[ws_macro.max_row])
    ws_macro.append(['City', 'Total Profit (CR)', 'Num Lucrative Commodities'])
    _style_header_row(ws_macro[ws_macro.max_row])
    row_start = ws_macro.max_row + 1
    for i, item in enumerate(macro_data['top_cities_by_profit'], start=row_start):
        ws_macro.append([item['city'], item['total_profit'], item['num_lucrative']])
        if i % 2 == 0:
            for cell in ws_macro[i]: cell.fill = alt_fill2

    # ── Section 3: Top 10 by lucrative commodities ────────────────────────────
    ws_macro.append([])
    _add_page_break(ws_macro)
    ws_macro.append(['TOP 10 CITIES BY NUMBER OF LUCRATIVE COMMODITIES'])
    _style_section_title(ws_macro[ws_macro.max_row])
    ws_macro.append(['City', 'Num Lucrative Commodities', 'Total Profit (CR)'])
    _style_header_row(ws_macro[ws_macro.max_row])
    row_start = ws_macro.max_row + 1
    for i, item in enumerate(macro_data['cities_by_lucrative_commodities'], start=row_start):
        ws_macro.append([item['city'], item['num_lucrative'], item['total_profit']])
        if i % 2 == 0:
            for cell in ws_macro[i]: cell.fill = alt_fill2

    # ── Section 4: Best sellers/buyers by commodity ───────────────────────────
    ws_macro.append([])
    _add_page_break(ws_macro)
    ws_macro.append(['BEST SELLERS AND BUYERS BY COMMODITY'])
    _style_section_title(ws_macro[ws_macro.max_row])

    for commodity in COMMODITY_CATEGORIES:
        color = COMMODITY_COLORS.get(commodity, "E6E6FA")

        # Each commodity block starts on its own page
        ws_macro.append([])
        _add_page_break(ws_macro)
        ws_macro.append([f'--- {commodity.upper()} ---'])
        _style_commodity_separator(ws_macro[ws_macro.max_row], color)

        ws_macro.append([])
        ws_macro.append([f'{commodity} - Best Sellers (lowest price)'])
        _style_subsection_title(ws_macro[ws_macro.max_row], color)
        ws_macro.append(['City', 'Selling Price (CR/MT)', 'Capacity (MT)'])
        _style_header_row(ws_macro[ws_macro.max_row])
        row_start = ws_macro.max_row + 1
        for i, seller in enumerate(macro_data['commodity_best_sellers'].get(commodity, []), start=row_start):
            ws_macro.append([seller['city'], seller['price'], seller['capacity']])
            if i % 2 == 0:
                for cell in ws_macro[i]: cell.fill = alt_fill2

        ws_macro.append([])
        ws_macro.append([f'{commodity} - Best Buyers (highest price)'])
        _style_subsection_title(ws_macro[ws_macro.max_row], color)
        ws_macro.append(['City', 'Buying Price (CR/MT)', 'Capacity (MT)'])
        _style_header_row(ws_macro[ws_macro.max_row])
        row_start = ws_macro.max_row + 1
        for i, buyer in enumerate(macro_data['commodity_best_buyers'].get(commodity, []), start=row_start):
            ws_macro.append([buyer['city'], buyer['price'], buyer['capacity']])
            if i % 2 == 0:
                for cell in ws_macro[i]: cell.fill = alt_fill2

    _auto_size_columns(ws_macro)

    # ── SAVE ──────────────────────────────────────────────────────────────────
    try:
        wb.save(output_file)
        print(f"\n✔  Saved: {output_file}")
        print(f"   Ship: {selected_ship} | Capacity: {ship_capacity} MT | Origin: {origin}")
    except PermissionError:
        ts   = int(time.time())
        base, ext = os.path.splitext(output_file)
        alt  = f"{base}_{ts}{ext}"
        wb.save(alt)
        print(f"⚠  '{output_file}' was in use → saved as '{alt}'")
    except Exception as e:
        print(f"✗  Error saving Excel: {e}")
        raise

# ─────────────────────────────────────────────────────────────────────────────
# PROMPT HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _prompt_city(prompt_text: str) -> str:
    print(f"\n{prompt_text}")
    for i, city in enumerate(CITY_LIST, 1):
        print(f"  {i:>2}. {city}")
    while True:
        try:
            choice = int(input("City number: ").strip())
            if 1 <= choice <= len(CITY_LIST):
                return CITY_LIST[choice - 1]
            print(f"   Choose between 1 and {len(CITY_LIST)}.")
        except ValueError:
            print("   Enter a valid number.")


def _prompt_commodities() -> list:
    print("\n  Select commodities to trade (you can choose multiple):")
    print("   0. ALL commodities")
    for i, cat in enumerate(COMMODITY_CATEGORIES, 1):
        print(f"  {i:>2}. {cat}")
    print("\n  Enter numbers separated by commas (e.g. 1,3,5) or 0 for all:")
    while True:
        raw = input("  Choice: ").strip()
        if raw == '0':
            return list(COMMODITY_CATEGORIES)
        parts = [p.strip() for p in raw.split(',')]
        try:
            indices = [int(p) for p in parts if p]
            selected = []
            valid = True
            for idx in indices:
                if 1 <= idx <= len(COMMODITY_CATEGORIES):
                    selected.append(COMMODITY_CATEGORIES[idx - 1])
                else:
                    print(f"  Invalid option: {idx}")
                    valid = False
                    break
            if valid and selected:
                print(f"  Selected: {', '.join(selected)}")
                return selected
        except ValueError:
            pass
        print("  Enter valid numbers separated by commas, or 0 for all.")


def _prompt_max_hops() -> int:
    print("\n  How many stops maximum (not counting return to origin)?")
    print("  (2 = A→B→C→A, 3 = A→B→C→D→A, up to 5)")
    print("  ⚠  More stops = much longer calculation time.")
    while True:
        try:
            n = int(input("  Max stops (2-5): ").strip())
            if 2 <= n <= 5:
                return n
            print("  Enter a value between 2 and 5.")
        except ValueError:
            print("  Enter a valid number.")

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main(image_folder, selected_ship=None, output_file='final_trade.xlsx',
         budget=None, containers_used=None, origin=None, mode='regular'):

    print("\n====================================")
    print("  RISE TVP — ship configuration")
    print(f"  OS detected: {platform.system()}")
    print("====================================")

    # ── 1. Category ───────────────────────────────────────────────────────────
    print("\nCategory:")
    print("  1. AIR AND SPACE")
    print("  2. ONLY AIR")
    while True:
        try:
            cat_choice = int(input("Category number: ").strip())
            if cat_choice in (1, 2):
                category = ["AIR AND SPACE", "ONLY AIR"][cat_choice - 1]
                break
            print("  Choose 1 or 2.")
        except ValueError:
            print("  Enter a valid number.")

    # ── 2. Ship model ─────────────────────────────────────────────────────────
    ships     = SHIPS[category]
    ship_list = list(ships.keys())
    print(f"\nAvailable models in {category}:")
    for i, s in enumerate(ship_list, 1):
        d = ships[s]
        cap_max = d["cargo_base"] + d["max_containers"] * d["container_mt"]
        if d["max_containers"] > 0:
            print(f"  {i}. {s}  —  {d['cargo_base']} MT base + up to {d['max_containers']} containers x {d['container_mt']} MT  (max {cap_max} MT)")
        else:
            print(f"  {i}. {s}  —  {d['cargo_base']} MT (no containers)")

    while True:
        try:
            s_choice = int(input("Model number: ").strip())
            if 1 <= s_choice <= len(ship_list):
                selected_ship = ship_list[s_choice - 1]
                break
            print(f"  Choose between 1 and {len(ship_list)}.")
        except ValueError:
            print("  Enter a valid number.")

    ship_data = ships[selected_ship]
    max_c     = ship_data["max_containers"]

    # ── 3. Containers ─────────────────────────────────────────────────────────
    if containers_used is None:
        if max_c == 0:
            print(f"\n  {selected_ship} has no container slots. Fixed cargo: {ship_data['cargo_base']} MT.")
            containers_used = 0
        else:
            print(f"\n  {selected_ship} supports up to {max_c} cargo container(s).")
            print(f"  Each container adds {ship_data['container_mt']} MT additional.")
            while True:
                try:
                    n = int(input(f"  How many containers are you carrying today? (0 to {max_c}): ").strip())
                    if 0 <= n <= max_c:
                        containers_used = n
                        break
                    print(f"  Enter a value between 0 and {max_c}.")
                except ValueError:
                    print("  Enter a valid number.")

    ship_capacity = ship_data["cargo_base"] + containers_used * ship_data["container_mt"]
    print(f"  Calculated capacity: {ship_data['cargo_base']} + {containers_used} x {ship_data['container_mt']} = {ship_capacity} MT")

    ship_rental_cost = ship_data.get("rental_cost_per_day")

    # ── 4. Rental or purchase ─────────────────────────────────────────────────
    is_rental = False
    print(f"\n  Do you want to rent the {selected_ship}?")
    while True:
        r = input("  Rented (A) or Purchased (C)? ").strip().upper()
        if r == 'A':
            is_rental = True
            if ship_rental_cost is None:
                while True:
                    try:
                        ship_rental_cost = int(input("  Enter rental cost per day (CR): ").strip())
                        break
                    except ValueError:
                        print("  Please enter a valid number.")
            break
        elif r == 'C':
            break
        print("  Enter A or C.")

    # ── 5. Origin ─────────────────────────────────────────────────────────────
    if origin is None:
        origin = _prompt_city("SELECT ORIGIN CITY:")

    print(f"\n  Ship      : {selected_ship} ({ship_capacity} MT) — {'Rented' if is_rental else 'Purchased'}")
    print(f"  Containers: {containers_used}/{max_c}")
    print(f"  Origin    : {origin}")

    # ── 6. Budget ─────────────────────────────────────────────────────────────
    if budget is None:
        print("\n  Available budget in CR?")
        print("  (Used to calculate A/B/C/D grade and constrain route purchases)")
        b_input = input("  Budget CR (Enter to skip): ").strip()
        if b_input:
            try:
                budget = int(b_input)
            except ValueError:
                print("  Invalid value, skipped.")
    if budget:
        print(f"  Budget: {budget} CR")

    # ── 7. Trade Route extra prompts ──────────────────────────────────────────
    trade_route_params = None
    if mode == 'route':
        allowed_commodities = _prompt_commodities()
        max_hops            = _prompt_max_hops()
        trade_route_params  = {
            'allowed_commodities': allowed_commodities,
            'max_hops':            max_hops,
        }

    # ── Process images ────────────────────────────────────────────────────────
    if not os.path.exists(image_folder):
        print(f"\n✗  The folder '{image_folder}' does not exist.")
        return

    data_dict = {
        'Cities': {
            'header': ['Location', 'Category', 'Commodity Type', 'Quantity MT',
                       'Reserve MT', 'Selling CR/MT', 'Buying CR/MT', 'Maximum MT'],
            'rows': []
        },
        'EasyDock': {
            'header': ['Location', 'Category', 'Name', 'MT', 'Buying MT',
                       'Buying CR', 'Selling MT', 'Selling CR'],
            'rows': []
        }
    }

    for filename in sorted(os.listdir(image_folder)):
        if not filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff')):
            continue

        # Strip the trailing _N suffix to get the source name
        base_name = filename.rsplit('_', 1)[0]

        # ── FILTER: skip TVI_ files and any other ignored prefixes ────────────
        if is_ignored_source(base_name):
            print(f"  ⏭  Skipped (ignored source): {filename}")
            continue

        sheet_key  = 'EasyDock' if base_name.lower() == 'easydock' else 'Cities'
        image_path = os.path.join(image_folder, filename)
        text = extract_text_from_image(image_path)
        if text:
            parsed = parse_text_to_data(text)
            for row in parsed['rows']:
                row.insert(0, base_name)
            data_dict[sheet_key]['rows'].extend(parsed['rows'])
            print(f"  ✔ {filename} → {sheet_key}")
        else:
            print(f"  ✗ No text extracted: {filename}")

    if data_dict['Cities']['rows'] or data_dict['EasyDock']['rows']:
        catalog       = build_trade_catalog(data_dict)
        opportunities = find_trade_opportunities(catalog)

        print(f"\n{'='*60}")
        print(f"  Opportunities found: {len(opportunities)}")
        print(f"{'='*60}")

        save_to_excel(
            data_dict,
            selected_ship,
            ship_capacity,
            output_file=output_file,
            is_rental=is_rental,
            rental_cost_per_day=ship_rental_cost,
            origin=origin,
            budget=budget,
            containers_used=containers_used,
            mode=mode,
            trade_route_params=trade_route_params,
        )
    else:
        print("\n  No data found to save.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Extract data from images and generate trade route analysis'
    )
    parser.add_argument('--images',     default='images',           help='Image folder')
    parser.add_argument('--ship',       default=None,               help='Ship to use')
    parser.add_argument('--output',     default='final_trade.xlsx', help='Output Excel file')
    parser.add_argument('--budget',     default=None, type=int,     help='Initial budget in CR')
    parser.add_argument('--containers', default=None, type=int,     help='Number of cargo containers carried')
    parser.add_argument('--origin',     default=None,               help='Origin city')
    parser.add_argument('--mode',       default='regular',          help='Mode: regular, city, or route')
    args = parser.parse_args()

    main(
        args.images,
        selected_ship=args.ship,
        output_file=args.output,
        budget=args.budget,
        containers_used=args.containers,
        origin=args.origin,
        mode=args.mode,
    )
