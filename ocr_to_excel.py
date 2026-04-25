import argparse
import math
import os
import re
import time
import pytesseract
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

# Configurar la ruta a Tesseract (ajusta si es necesario)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Categorías de commodities en orden
COMMODITY_CATEGORIES = [
    "Rare/Precious",
    "Foodstuffs",
    "Natural Materials",
    "Fuel Ore",
    "Consumer Goods",
    "Fabricated Material",
    "Refined Fuel"
]

# Ships data
# cargo_base : MT fijos del ship (siempre disponibles)
# max_containers : numero maximo de cargo containers que admite el modelo
# container_mt : MT que aporta cada container (siempre 17)
# rental_cost_per_day : CR/dia si es alquilable, None si solo compra
SHIPS = {
    "AIR AND SPACE": {
        "E-10 Saint":   {"cargo_base": 7, "max_containers": 6, "container_mt": 17, "rental_cost_per_day": None},
        "E-11 Saint":   {"cargo_base": 7, "max_containers": 6, "container_mt": 17, "rental_cost_per_day": None},
        "P-13 Prowler": {"cargo_base": 1, "max_containers": 0, "container_mt": 17, "rental_cost_per_day": None},
        "W-6 Manx":     {"cargo_base": 7, "max_containers": 3, "container_mt": 17, "rental_cost_per_day": None},
    },
    "ONLY AIR": {
        "A-4 Wanderer":      {"cargo_base": 1, "max_containers": 1, "container_mt": 17, "rental_cost_per_day": 1359},
        "T-19 Stratomaster": {"cargo_base": 1, "max_containers": 1, "container_mt": 17, "rental_cost_per_day": 2264},
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# CITIES — coordenadas aproximadas (unidades de juego, no geo-reales).
# Se usan para calcular distancias relativas entre ciudades cuando no existe
# una entrada exacta en FLIGHT_TIMES.
# ─────────────────────────────────────────────────────────────────────────────
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

# Listado ordenado para mostrar en el menú de selección de origen
CITY_LIST = sorted(CITIES.keys())

# Flight times explícitos (origen → destino, en minutos de vuelo puro).
# Si no existe una entrada, se estima a partir de las coordenadas.
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

# Tiempos fijos de maniobra (minutos)
TAKEOFF_TIME = 5
LANDING_TIME = 10

# Velocidad de referencia para estimar tiempos desconocidos.
# Se calcula sobre los datos explícitos de Kansas (origen bien poblado).
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


def _coord_distance(city_a: str, city_b: str) -> float:
    a = CITIES.get(city_a)
    b = CITIES.get(city_b)
    if not a or not b:
        return 0.0
    dx = a["x"] - b["x"]
    dy = a["y"] - b["y"]
    return math.sqrt(dx * dx + dy * dy)


def get_flight_time(origin: str, destination: str) -> int:
    """Tiempo total de vuelo incluyendo despegue y aterrizaje (minutos)."""
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
    """
    Calcula la capacidad total en MT del ship.
      capacidad = cargo_base + containers_used * container_mt
    Si containers_used es None se asume el maximo del modelo.
    """
    for category, ships in SHIPS.items():
        if ship_name in ships:
            d = ships[ship_name]
            n = containers_used if containers_used is not None else d["max_containers"]
            return d["cargo_base"] + n * d["container_mt"]
    return 0


def extract_text_from_image(image_path: str) -> str:
    try:
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img)
        return text.strip()
    except Exception as e:
        print(f"Error procesando {image_path}: {e}")
        return ""


def normalize_number(token):
    if token is None:
        return None
    if isinstance(token, (int, float)):
        return int(token)
    text = str(token).replace(',', '').strip()
    clean = re.sub(r'[^0-9.]', '', text)
    if clean == '':
        return None
    try:
        return int(float(clean))
    except ValueError:
        return None


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


def parse_table_rows(lines, min_numbers=5, layout='city'):
    rows = []
    for line in lines:
        if is_total_line(line):
            continue
        if line.lower().startswith(('commodity type', 'name', 'type',
                                    'population', 'fees')):
            continue
        numbers = re.findall(r'[\d,]+', line)
        if len(numbers) < min_numbers:
            continue
        first_number = re.search(r'[\d,]+', line)
        if not first_number:
            continue
        name = line[:first_number.start()].strip()
        if not name:
            continue
        parsed_numbers = [normalize_number(n) for n in numbers[:min_numbers]]
        row = [name] + parsed_numbers
        if layout == 'city':
            qty = parsed_numbers[0] if len(parsed_numbers) > 0 else None
            sell = parsed_numbers[2] if len(parsed_numbers) > 2 else None
            if qty and qty > 0 and sell and sell > 0:
                rows.append(row)
        elif layout == 'easydock':
            sell_cr = parsed_numbers[4] if len(parsed_numbers) > 4 else None
            if sell_cr and sell_cr > 0:
                rows.append(row)
    return rows


def parse_text_to_data(text: str) -> dict:
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    layout = detect_layout(lines)

    if layout == 'city':
        header = ['Category', 'Commodity Type', 'Quantity MT', 'Reserve MT',
                  'Selling CR/MT', 'Buying CR/MT', 'Maximum MT']
        rows = parse_table_rows(lines, min_numbers=5, layout='city')
        for i, row in enumerate(rows):
            cat = COMMODITY_CATEGORIES[i] if i < len(COMMODITY_CATEGORIES) else ''
            row.insert(0, cat)
    elif layout == 'easydock':
        header = ['Category', 'Name', 'MT', 'Buying MT', 'Buying CR',
                  'Selling MT', 'Selling CR']
        rows = parse_table_rows(lines, min_numbers=5, layout='easydock')
        for i, row in enumerate(rows):
            cat = COMMODITY_CATEGORIES[i] if i < len(COMMODITY_CATEGORIES) else ''
            row.insert(0, cat)
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
            commodity_key = str(row[2]).strip() if len(row) > 2 else category
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
                selling = normalize_number(row[si]) if len(row) > si else None
                buying  = normalize_number(row[bi]) if len(row) > bi else None
                quantity = normalize_number(row[qi]) if len(row) > qi else 0
                reserve  = normalize_number(row[ri]) if len(row) > ri else 0
                max_accept = normalize_number(row[mi]) if len(row) > mi else 0
                sell_capacity = max(quantity - reserve, 0)
                buy_capacity  = max_accept

            elif sheet_name == 'EasyDock' and 'selling cr' in header:
                si = header.index('selling cr')
                bi = header.index('buying cr')
                qi = header.index('mt')
                bci = header.index('buying mt')
                sci = header.index('selling mt')
                selling = normalize_number(row[si]) if len(row) > si else None
                buying  = normalize_number(row[bi]) if len(row) > bi else None
                quantity = normalize_number(row[qi]) if len(row) > qi else 0
                buy_capacity  = normalize_number(row[bci]) if len(row) > bci else 0
                sell_capacity = normalize_number(row[sci]) if len(row) > sci else 0

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
                source_available = src_data.get('sell_capacity', 0) or 0
                destination_capacity = dst_data.get('buy_capacity', 0) or 0
                profit_per_mt = dst_data['buying'] - src_data['selling']
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

def assign_grades(opportunities: list, ship_capacity: int = 109, budget: int = None) -> list:
    """
    Calcula, para cada oportunidad, cuantas MT puedes comprar con el presupuesto
    y la capacidad del ship, luego rankea por profit_per_trip resultante (A>B>C>D).

    qty_trip  = min(source_available, destination_capacity, ship_capacity,
                    floor(budget / source_selling) si hay budget)
    cost_trip = qty_trip * source_selling
    profit_trip = qty_trip * profit_per_mt
    roi_trip  = profit_trip / cost_trip * 100

    El grade se asigna por cuartil de profit_trip entre todas las oportunidades
    con profit_trip > 0. Las rutas sin stock o sin margen quedan en D.
    """
    for op in opportunities:
        price = op["source_selling"] or 1
        max_by_ship   = ship_capacity
        max_by_source = op["source_available"]
        max_by_dest   = op["destination_capacity"]
        max_by_budget = int(budget / price) if budget else float("inf")

        qty = min(max_by_ship, max_by_source, max_by_dest, max_by_budget)
        qty = max(qty, 0)

        cost_trip   = qty * price
        profit_trip = qty * op["profit_per_mt"]
        roi_trip    = (profit_trip / cost_trip * 100) if cost_trip > 0 else 0

        op["_qty_trip"]    = qty
        op["_cost_trip"]   = cost_trip
        op["_profit_trip"] = profit_trip
        op["_roi"]         = roi_trip
        # _affordable: el viaje tiene sentido (hay algo que comprar y hay margen)
        op["_affordable"]  = qty > 0 and profit_trip > 0

    viable = [op for op in opportunities if op["_affordable"]]
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


# Grade fills: A=verde oscuro, B=verde claro, C=amarillo, D=naranja/rojo
GRADE_STYLES = {
    "A": {"fill": "00B050", "font": "FFFFFF"},
    "B": {"fill": "70AD47", "font": "FFFFFF"},
    "C": {"fill": "FFD700", "font": "000000"},
    "D": {"fill": "FFA500", "font": "FFFFFF"},
}


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

def _style_header_row(cells):
    for cell in cells:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin")
        )


def _auto_size_columns(ws, min_width=10, max_width=50):
    for col_cells in ws.columns:
        max_len = max(
            (len(str(cell.value)) if cell.value is not None else 0)
            for cell in col_cells
        )
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_width, max(min_width, max_len + 2))


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
):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    alt_fill        = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
    high_profit_fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")

    # ── Hoja CONFIG ───────────────────────────────────────────────────────────
    ws_cfg = wb.create_sheet(title='Config')
    ws_cfg.append(['Parámetro', 'Valor'])
    _style_header_row(ws_cfg[1])

    ship_data = None
    ship_category = ''
    for cat, ships in SHIPS.items():
        if selected_ship in ships:
            ship_data = ships[selected_ship]
            ship_category = cat
            break

    max_c   = ship_data["max_containers"] if ship_data else 0
    c_used  = containers_used if containers_used is not None else max_c
    c_mt    = ship_data["container_mt"] if ship_data else 17
    cargo_b = ship_data["cargo_base"] if ship_data else 0

    ws_cfg.append(['Ship', selected_ship])
    ws_cfg.append(['Categoria', ship_category])
    ws_cfg.append(['Cargo base (MT fijos)', cargo_b])
    ws_cfg.append(['Containers llevados', f'{c_used} / {max_c} max'])
    ws_cfg.append(['MT por container', c_mt])
    ws_cfg.append(['Capacidad total (MT)', f'{cargo_b} + {c_used} x {c_mt} = {ship_capacity}'])
    ws_cfg.append(['Origen', origin])
    ws_cfg.append(['Estado', 'Alquilado' if is_rental else 'Comprado'])
    if is_rental and rental_cost_per_day:
        ws_cfg.append(['Coste alquiler/dia', f'{rental_cost_per_day} CR/dia (14 h)'])
        ws_cfg.append(['Profit minimo/dia para cubrir alquiler', f'{rental_cost_per_day} CR'])
        ws_cfg.append(['Profit minimo/hora para cubrir alquiler', f'{rental_cost_per_day / 14:.2f} CR/h'])
    if budget:
        ws_cfg.append(['Presupuesto inicial', f'{budget} CR'])

    _auto_size_columns(ws_cfg)

    # ── Hoja CITIES ───────────────────────────────────────────────────────────
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

    # ── Hoja OPPORTUNITIES ────────────────────────────────────────────────────
    catalog = build_trade_catalog(data_dict)
    opportunities = find_trade_opportunities(catalog)

    if opportunities:
        # Asignar grades con logica real de viaje
        opportunities = assign_grades(opportunities, ship_capacity=ship_capacity, budget=budget)

        # Ordenar por profit de viaje real descendente (A primero)
        opportunities.sort(key=lambda x: x["_profit_trip"], reverse=True)

        ws_op = wb.create_sheet(title='Opportunities')

        columns = [
            'Grade',
            'Commodity',
            'Source',
            'Buy Price (CR/MT)',
            'Destination',
            'Sell Price (CR/MT)',
            'Profit/MT (CR)',
            'Src Stock (MT)',
            'Dst Capacity (MT)',
            'MT cargados en viaje',
            'Cost del viaje (CR)',
            'Profit del viaje (CR)',
            'ROI viaje (%)',
        ]
        if is_rental and rental_cost_per_day:
            columns.append('Cubre Alquiler?')

        ws_op.append(columns)
        _style_header_row(ws_op[1])

        font_white = Font(color="FFFFFF", bold=True)
        font_black = Font(color="000000", bold=True)
        green_fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
        red_fill   = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        GRADE_COL  = 1
        ROI_COL    = 13   # columna ROI viaje (%)

        for op in opportunities:
            grade        = op['grade']
            qty_trip     = op['_qty_trip']
            cost_trip    = op['_cost_trip']
            profit_trip  = op['_profit_trip']
            roi_pct      = op['_roi']

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
                round(roi_pct, 2),
            ]
            if is_rental and rental_cost_per_day:
                row.append("SI" if profit_trip >= rental_cost_per_day else "NO")

            ws_op.append(row)
            cur_row = ws_op.max_row

            # ── Color columna Grade ───────────────────────────────────────────
            gs = GRADE_STYLES[grade]
            gc = ws_op.cell(row=cur_row, column=GRADE_COL)
            gc.fill = PatternFill(start_color=gs["fill"], end_color=gs["fill"], fill_type="solid")
            gc.font = Font(color=gs["font"], bold=True)
            gc.alignment = Alignment(horizontal="center")

            # ── Color columna ROI ─────────────────────────────────────────────
            roi_cell = ws_op.cell(row=cur_row, column=ROI_COL)
            if roi_pct > 100:
                roi_cell.fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
                roi_cell.font = font_white
            elif roi_pct >= 50:
                roi_cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                roi_cell.font = font_white
            elif roi_pct >= 20:
                roi_cell.fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
                roi_cell.font = font_black
            elif roi_pct >= 0:
                roi_cell.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
                roi_cell.font = font_white
            else:
                roi_cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                roi_cell.font = font_white

            # ── Color columna Cubre Alquiler? ─────────────────────────────────
            if is_rental and rental_cost_per_day:
                rc = ws_op.cell(row=cur_row, column=len(columns))
                if rc.value == "SI":
                    rc.fill = green_fill; rc.font = font_white
                else:
                    rc.fill = red_fill;   rc.font = font_white

        ws_op.freeze_panes = "B2"
        _auto_size_columns(ws_op)

    # ── Guardar ───────────────────────────────────────────────────────────────
    try:
        wb.save(output_file)
        print(f"\n✔  Guardado: {output_file}")
        print(f"   Ship: {selected_ship} | Capacidad: {ship_capacity} MT | Origen: {origin}")
    except PermissionError:
        ts = int(time.time())
        base, ext = os.path.splitext(output_file)
        alt = f"{base}_{ts}{ext}"
        wb.save(alt)
        print(f"⚠  '{output_file}' estaba en uso → guardado como '{alt}'")
    except Exception as e:
        print(f"✗  Error guardando Excel: {e}")
        raise


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def _prompt_city(prompt_text: str) -> str:
    """Muestra la lista completa de ciudades y devuelve la elegida."""
    print(f"\n{prompt_text}")
    for i, city in enumerate(CITY_LIST, 1):
        print(f"  {i:>2}. {city}")
    while True:
        try:
            choice = int(input("Número de ciudad: ").strip())
            if 1 <= choice <= len(CITY_LIST):
                return CITY_LIST[choice - 1]
            print(f"   Elige entre 1 y {len(CITY_LIST)}.")
        except ValueError:
            print("   Introduce un número válido.")


def main(image_folder, selected_ship=None, output_file='final_trade.xlsx',
         budget=None, containers_used=None, origin=None):

    # ── 1. Seleccion de categoria ─────────────────────────────────────────────
    print("\n====================================")
    print("  RISE TVP — configuracion de nave")
    print("====================================")
    print("\nCategoria:")
    print("  1. AIR AND SPACE")
    print("  2. ONLY AIR")
    while True:
        try:
            cat_choice = int(input("Numero de categoria: ").strip())
            if cat_choice in (1, 2):
                category = ["AIR AND SPACE", "ONLY AIR"][cat_choice - 1]
                break
            print("  Elige 1 o 2.")
        except ValueError:
            print("  Introduce un numero valido.")

    # ── 2. Seleccion de modelo ────────────────────────────────────────────────
    ships = SHIPS[category]
    ship_list = list(ships.keys())
    print(f"\nModelos disponibles en {category}:")
    for i, s in enumerate(ship_list, 1):
        d = ships[s]
        cap_max = d["cargo_base"] + d["max_containers"] * d["container_mt"]
        if d["max_containers"] > 0:
            print(f"  {i}. {s}  —  {d['cargo_base']} MT base + hasta {d['max_containers']} containers x {d['container_mt']} MT  (max {cap_max} MT)")
        else:
            print(f"  {i}. {s}  —  {d['cargo_base']} MT (sin containers)")

    while True:
        try:
            s_choice = int(input("Numero de modelo: ").strip())
            if 1 <= s_choice <= len(ship_list):
                selected_ship = ship_list[s_choice - 1]
                break
            print(f"  Elige entre 1 y {len(ship_list)}.")
        except ValueError:
            print("  Introduce un numero valido.")

    ship_data = ships[selected_ship]
    max_c = ship_data["max_containers"]

    # ── 3. Numero de cargo containers llevados ────────────────────────────────
    if containers_used is None:
        if max_c == 0:
            print(f"\n  {selected_ship} no tiene slots de container. Cargo fijo: {ship_data['cargo_base']} MT.")
            containers_used = 0
        else:
            print(f"\n  {selected_ship} admite hasta {max_c} cargo container(s).")
            print(f"  Cada container aporta {ship_data['container_mt']} MT adicionales.")
            while True:
                try:
                    n = int(input(f"  Cuantos containers llevas hoy? (0 a {max_c}): ").strip())
                    if 0 <= n <= max_c:
                        containers_used = n
                        break
                    print(f"  Introduce un valor entre 0 y {max_c}.")
                except ValueError:
                    print("  Introduce un numero valido.")

    ship_capacity = ship_data["cargo_base"] + containers_used * ship_data["container_mt"]
    print(f"  Capacidad calculada: {ship_data['cargo_base']} + {containers_used} x {ship_data['container_mt']} = {ship_capacity} MT")

    ship_category_key = category
    ship_rental_cost = ship_data.get("rental_cost_per_day")

    # ── 4. Seleccion de origen ────────────────────────────────────────────────
    if origin is None:
        origin = _prompt_city("Ciudad de ORIGEN:")

    # ── 5. Alquiler o compra ──────────────────────────────────────────────────
    is_rental = False
    if ship_rental_cost is not None:
        print(f"\n  {selected_ship} puede alquilarse ({ship_rental_cost} CR/dia) o comprarse.")
        while True:
            r = input("  Alquilado (A) o Comprado (C)? ").strip().upper()
            if r == 'A':
                is_rental = True
                break
            elif r == 'C':
                break
            print("  Introduce A o C.")
    else:
        print(f"\n  {selected_ship} solo puede comprarse.")

    print(f"\n  Ship    : {selected_ship} ({ship_capacity} MT) — {'Alquilado' if is_rental else 'Comprado'}")
    print(f"  Containers: {containers_used}/{max_c}")
    print(f"  Origen  : {origin}")

    # ── 6. Presupuesto (siempre al final del flujo) ───────────────────────────
    if budget is None:
        print("\n  Presupuesto disponible en CR?")
        print("  (Se usara para calcular el grade A/B/C/D de cada oportunidad)")
        b_input = input("  Presupuesto CR (Enter para omitir): ").strip()
        if b_input:
            try:
                budget = int(b_input)
            except ValueError:
                print("  Valor no valido, se omite.")
    if budget:
        print(f"  Presupuesto: {budget} CR")

    # ── Procesar imágenes ─────────────────────────────────────────────────────
    if not os.path.exists(image_folder):
        print(f"\n✗  La carpeta '{image_folder}' no existe.")
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
        base_name = filename.rsplit('_', 1)[0]
        sheet_key = 'EasyDock' if base_name.lower() == 'easydock' else 'Cities'
        image_path = os.path.join(image_folder, filename)
        text = extract_text_from_image(image_path)
        if text:
            parsed = parse_text_to_data(text)
            for row in parsed['rows']:
                row.insert(0, base_name)
            data_dict[sheet_key]['rows'].extend(parsed['rows'])
            print(f"  ✔ {filename} → {sheet_key}")
        else:
            print(f"  ✗ Sin texto: {filename}")

    if data_dict['Cities']['rows'] or data_dict['EasyDock']['rows']:
        catalog = build_trade_catalog(data_dict)
        opportunities = find_trade_opportunities(catalog)

        print(f"\n{'='*60}")
        print(f"  Oportunidades encontradas: {len(opportunities)}")
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
        )
    else:
        print("\n  No se encontraron datos para guardar.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Extrae datos de imágenes y genera análisis de rutas comerciales'
    )
    parser.add_argument('--images',     default='images',          help='Carpeta de imágenes')
    parser.add_argument('--ship',       default=None,              help='Ship a usar')
    parser.add_argument('--output',     default='final_trade.xlsx', help='Archivo Excel de salida')
    parser.add_argument('--budget',     default=None, type=int,    help='Presupuesto inicial en CR')
    parser.add_argument('--containers', default=None, type=int,    dest='containers',
                        help='Numero de cargo containers llevados en este vuelo')
    parser.add_argument('--origin',     default=None,              help='Ciudad de origen')
    args = parser.parse_args()

    main(
        args.images,
        selected_ship=args.ship,
        output_file=args.output,
        budget=args.budget,
        containers_used=args.containers,
        origin=args.origin,
    )
