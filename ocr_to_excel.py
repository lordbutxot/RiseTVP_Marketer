import argparse
import os
import re
import time
import pytesseract
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

# Configurar la ruta a Tesseract (ajusta si es necesario)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Ruta por defecto en Windows

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
SHIPS = {
    "AIR AND SPACE": {
        "E-10 Saint": {"cargo": 7, "containers": 6, "container_mt": 17, "rental_cost_per_day": None},
        "E-11 Saint": {"cargo": 7, "containers": 6, "container_mt": 17, "rental_cost_per_day": None},
        "P-13 Prowler": {"cargo": 1, "containers": 0, "container_mt": 17, "rental_cost_per_day": None},
        "W-6 Manx": {"cargo": 7, "containers": 3, "container_mt": 17, "rental_cost_per_day": None},
    },
    "ONLY AIR": {
        "A-4 Wanderer": {"cargo": 1, "containers": 1, "container_mt": 17, "rental_cost_per_day": 1359},
        "T-19 Stratomaster": {"cargo": 1, "containers": 1, "container_mt": 17, "rental_cost_per_day": 2264},
    }
}

# Flight times in minutes between cities (flight time only, not including takeoff/landing)
FLIGHT_TIMES = {
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

# Fixed times in minutes
TAKEOFF_TIME = 5  # Despegue
LANDING_TIME = 10  # Aterrizaje

def get_flight_time(origin, destination):
    """Get total flight time including takeoff and landing."""
    if origin == destination:
        return 0
    flight_mins = FLIGHT_TIMES.get((origin, destination)) or FLIGHT_TIMES.get((destination, origin))
    if flight_mins is None:
        return 60  # Default fallback
    return TAKEOFF_TIME + flight_mins + LANDING_TIME

def calculate_ship_capacity(ship_name):
    for category, ships in SHIPS.items():
        if ship_name in ships:
            data = ships[ship_name]
            return data["cargo"] + data["containers"] * data["container_mt"]
    return 0

def extract_text_from_image(image_path):
    """Extrae texto de una imagen usando OCR."""
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
        number = float(clean)
        return int(number)
    except ValueError:
        return None


def is_total_line(line):
    lower = line.lower()
    return any(word in lower for word in ['totals', 'total', 'refresh', 'cancel', 'population', 'fees', 'ports staffed', 'mt free', 'cr free'])


def detect_layout(lines):
    joined = ' '.join(lines).lower()
    if any(keyword in joined for keyword in ['commodity type', 'reserve mt', 'selling cr/mt', 'buying cr/mt', 'maximum mt']):
        return 'city'
    if any(keyword in joined for keyword in ['commodities', 'buying cr', 'selling cr']) and 'name' in joined:
        return 'easydock'
    if any(keyword in joined for keyword in ['buying mt', 'selling mt', 'buying cr', 'selling cr']):
        return 'easydock'
    return 'simple'


def parse_table_rows(lines, min_numbers=5, layout='city'):
    rows = []
    for line in lines:
        if is_total_line(line):
            continue
        if line.lower().startswith(('commodity type', 'name', 'type', 'population', 'fees')):
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

        parsed_numbers = [normalize_number(num) for num in numbers[:min_numbers]]
        row = [name] + parsed_numbers

        # Filtrar por stock y precio
        if layout == 'city':
            quantity_mt = parsed_numbers[0] if len(parsed_numbers) > 0 else None
            selling_cr_mt = parsed_numbers[2] if len(parsed_numbers) > 2 else None
            if quantity_mt and quantity_mt > 0 and selling_cr_mt and selling_cr_mt > 0:
                rows.append(row)
        elif layout == 'easydock':
            selling_cr = parsed_numbers[4] if len(parsed_numbers) > 4 else None
            if selling_cr and selling_cr > 0:
                rows.append(row)

    return rows


def parse_text_to_data(text):
    """Parsea el texto extraído en datos estructurados, detectando tablas de ciudad o EasyDock."""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    layout = detect_layout(lines)

    if layout == 'city':
        header = ['Category', 'Commodity Type', 'Quantity MT', 'Reserve MT', 'Selling CR/MT', 'Buying CR/MT', 'Maximum MT']
        rows = parse_table_rows(lines, min_numbers=5, layout='city')
        for i, row in enumerate(rows):
            if i < len(COMMODITY_CATEGORIES):
                row.insert(0, COMMODITY_CATEGORIES[i])
    elif layout == 'easydock':
        header = ['Category', 'Name', 'MT', 'Buying MT', 'Buying CR', 'Selling MT', 'Selling CR']
        rows = parse_table_rows(lines, min_numbers=5, layout='easydock')
        for i, row in enumerate(rows):
            if i < len(COMMODITY_CATEGORIES):
                row.insert(0, COMMODITY_CATEGORIES[i])
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


def sanitize_sheet_name(name):
    invalid = ['\\', '/', '?', '*', '[', ']', ':']
    safe_name = ''.join('-' if ch in invalid else ch for ch in name)
    return safe_name[:31]


def build_trade_catalog(data_dict):
    catalog = {}
    for sheet_name in ['Cities', 'EasyDock']:
        if sheet_name not in data_dict:
            continue
        header = [h.lower() for h in data_dict[sheet_name]['header']]
        rows = data_dict[sheet_name]['rows']
        location_map = {}
        for row in rows:
            if len(row) < 3:
                continue
            location = str(row[0]).strip()
            category = str(row[1]).strip()
            commodity_key = str(row[2]).strip() if len(row) > 2 else category
            if location not in location_map:
                location_map[location] = {}
            item_map = location_map[location]

            selling = None
            buying = None
            quantity = 0
            reserve = 0
            max_accept = 0
            sell_capacity = 0
            buy_capacity = 0
            if sheet_name == 'Cities' and 'selling cr/mt' in header and 'buying cr/mt' in header:
                selling_idx = header.index('selling cr/mt')
                buying_idx = header.index('buying cr/mt')
                quantity_idx = header.index('quantity mt')
                reserve_idx = header.index('reserve mt')
                max_idx = header.index('maximum mt')
                selling = normalize_number(row[selling_idx]) if len(row) > selling_idx else None
                buying = normalize_number(row[buying_idx]) if len(row) > buying_idx else None
                quantity = normalize_number(row[quantity_idx]) if len(row) > quantity_idx else 0
                reserve = normalize_number(row[reserve_idx]) if len(row) > reserve_idx else 0
                max_accept = normalize_number(row[max_idx]) if len(row) > max_idx else 0
                sell_capacity = max(quantity - reserve, 0)
                buy_capacity = max_accept
            elif sheet_name == 'EasyDock' and 'selling cr' in header and 'buying cr' in header:
                selling_idx = header.index('selling cr')
                buying_idx = header.index('buying cr')
                quantity_idx = header.index('mt')
                buying_capacity_idx = header.index('buying mt')
                selling_capacity_idx = header.index('selling mt')
                selling = normalize_number(row[selling_idx]) if len(row) > selling_idx else None
                buying = normalize_number(row[buying_idx]) if len(row) > buying_idx else None
                quantity = normalize_number(row[quantity_idx]) if len(row) > quantity_idx else 0
                buy_capacity = normalize_number(row[buying_capacity_idx]) if len(row) > buying_capacity_idx else 0
                sell_capacity = normalize_number(row[selling_capacity_idx]) if len(row) > selling_capacity_idx else 0

            item_map[commodity_key] = {
                'selling': selling,
                'buying': buying,
                'quantity': quantity,
                'reserve': reserve,
                'max_accept': max_accept,
                'sell_capacity': sell_capacity,
                'buy_capacity': buy_capacity
            }
        catalog.update(location_map)
    return catalog


def find_trade_opportunities(catalog):
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
                # Remove profit filtering - show ALL possible trades
                # if source_available <= 0 or destination_capacity <= 0:
                #     continue
                # if src_data['selling'] < dst_data['buying']:
                profit_per_mt = dst_data['buying'] - src_data['selling']
                max_qty = min(source_available, destination_capacity) if source_available > 0 and destination_capacity > 0 else 0
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
    opportunities.sort(key=lambda item: item['total_profit'], reverse=True)
    return opportunities


def recommend_travel_chains(opportunities, origin, ship_capacity):
    """
    Recomienda cadenas de viajes óptimas (A→B→C) basadas en oportunidades.
    """
    chains = []
    
    # Crear un índice de oportunidades por destino para búsqueda rápida
    ops_by_destination = {}
    for op in opportunities:
        dest = op['destination']
        if dest not in ops_by_destination:
            ops_by_destination[dest] = []
        ops_by_destination[dest].append(op)
    
    # Para cada oportunidad A→B, buscar encadenamientos B→C
    for first_op in opportunities:
        destination_b = first_op['destination']
        if destination_b in ops_by_destination:
            for second_op in ops_by_destination[destination_b]:
                destination_c = second_op['destination']
                
                # Evitar ciclos cortos
                if destination_c == first_op['source']:
                    continue
                
                # Calcular tiempos y ganancias
                time_a_to_b = get_flight_time(origin, destination_b)
                time_b_to_c = get_flight_time(destination_b, destination_c)
                total_time = time_a_to_b + time_b_to_c  # En minutos
                
                trip_qty_1 = min(first_op['max_qty'], ship_capacity)
                trip_qty_2 = min(second_op['max_qty'], ship_capacity)
                
                profit_1 = trip_qty_1 * first_op['profit_per_mt']
                profit_2 = trip_qty_2 * second_op['profit_per_mt']
                total_profit = profit_1 + profit_2
                
                chains.append({
                    'route': f"{first_op['source']} → {destination_b} → {destination_c}",
                    'commodities': f"{first_op['commodity']} / {second_op['commodity']}",
                    'total_time_mins': total_time,
                    'profit_1': profit_1,
                    'profit_2': profit_2,
                    'total_profit': total_profit,
                    'efficiency': total_profit / (total_time / 60) if total_time > 0 else 0  # CR/hora
                })
    
    # Ordenar por eficiencia (ganancia por hora)
    chains.sort(key=lambda x: x['efficiency'], reverse=True)
    return chains[:10]  # Retornar top 10


def calculate_roi_opportunities(opportunities, budget=None):
    """
    Calcula ROI para oportunidades filtrando por presupuesto.
    ROI = (Ganancia / Costo Total) × 100%
    """
    roi_list = []
    
    for op in opportunities:
        if not budget:  # Si no hay presupuesto, mostrar todas
            roi_list.append({
                'route': f"{op['source']} → {op['destination']}",
                'commodity': op['commodity'],
                'qty_to_buy': op['source_available'],
                'buy_price_per_mt': op['source_selling'],
                'cost_to_buy': op['source_available'] * op['source_selling'],
                'profit': op['total_profit'],
                'roi_percent': (op['total_profit'] / (op['source_available'] * op['source_selling']) * 100) if (op['source_available'] * op['source_selling']) > 0 else 0
            })
        else:
            # Filtrar por presupuesto
            cost = op['source_available'] * op['source_selling']
            if cost <= budget:
                roi_list.append({
                    'route': f"{op['source']} → {op['destination']}",
                    'commodity': op['commodity'],
                    'qty_to_buy': op['source_available'],
                    'buy_price_per_mt': op['source_selling'],
                    'cost_to_buy': cost,
                    'profit': op['total_profit'],
                    'roi_percent': (op['total_profit'] / cost * 100) if cost > 0 else 0
                })
    
    # Ordenar por ROI descendente
    roi_list.sort(key=lambda x: x['roi_percent'], reverse=True)
    return roi_list


def save_to_excel(data_dict, selected_ship, ship_capacity, output_file='final_trade.xlsx', is_rental=False, rental_cost_per_day=None, origin="Delois Spot", chains=None, budget=None, roi_list=None):
    """Guarda el diccionario de datos en un archivo Excel con múltiples hojas."""
    if chains is None:
        chains = []
    if roi_list is None:
        roi_list = []
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_font = Font(bold=True)
    alt_fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
    high_profit_fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")

    def style_header_row(cells):
        for cell in cells:
            cell.font = header_font
            cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    def auto_size_columns(ws, min_width=10, max_width=50):
        for column_cells in ws.columns:
            max_length = max((len(str(cell.value)) if cell.value is not None else 0) for cell in column_cells)
            adjusted_width = min(max_width, max(min_width, max_length + 2))
            ws.column_dimensions[column_cells[0].column_letter].width = adjusted_width

    # Hoja Config
    ws_config = wb.create_sheet(title='Config')
    ws_config.append(['Ship Selection'])
    ws_config.append([])
    ws_config.append(['Selected Ship:', selected_ship])
    ws_config.append(['Capacity:', f'{ship_capacity} MT'])
    ws_config.append(['Origin:', origin])
    ws_config.append(['Status:', 'Alquilado' if is_rental else 'Comprado'])
    if is_rental and rental_cost_per_day:
        ws_config.append(['Rental Cost per Day:', f'{rental_cost_per_day} CR/día (14 horas)'])
        daily_profit_needed = rental_cost_per_day
        hourly_profit_needed = rental_cost_per_day / 14
        ws_config.append(['Profit needed per trip to break even:', f'{daily_profit_needed} CR/día'])
        ws_config.append(['Profit needed per hour to break even:', f'{hourly_profit_needed:.2f} CR/hora'])
    ws_config.append([])
    ws_config.append(['Available Ships:'])
    ws_config.append(['Category', 'Ship', 'Capacity MT'])
    style_header_row(ws_config[ws_config.max_row])
    ship_list = []
    for category, ships in SHIPS.items():
        for ship in ships:
            capacity = calculate_ship_capacity(ship)
            ws_config.append([category, ship, capacity])
            ship_list.append(ship)

    # Agregar dropdown en la celda de selección
    dv = DataValidation(type="list", formula1=f'"{",".join(ship_list)}"', allow_blank=False)
    ws_config.add_data_validation(dv)
    dv.add('B3')  # Celda B3 para Selected Ship
    ws_config['B3'] = selected_ship  # Valor inicial

    # Nota para el usuario
    ws_config.append([])
    ws_config.append(['Nota: Cambia la selección y ejecuta el script de nuevo para recalcular con el nuevo ship.'])
    auto_size_columns(ws_config)

    for sheet_name in ['Cities', 'EasyDock']:
        if sheet_name in data_dict and data_dict[sheet_name]['rows']:
            ws = wb.create_sheet(title=sheet_name)
            ws.append([sheet_name])
            ws.append([])
            header_row = data_dict[sheet_name]['header']
            ws.append(header_row)
            style_header_row(ws[ws.max_row])
            for i, row in enumerate(data_dict[sheet_name]['rows'], start=ws.max_row):
                ws.append(row)
                if i % 2 == 0:
                    for cell in ws[ws.max_row]:
                        cell.fill = alt_fill
            auto_size_columns(ws)

    catalog = build_trade_catalog(data_dict)
    opportunities = find_trade_opportunities(catalog)
    if opportunities:
        ws = wb.create_sheet(title='Opportunities')
        columns = ['Category', 'Source', 'Source selling CR/MT', 'Source available MT', 'Destination', 'Destination buying CR/MT', 'Destination capacity MT', 'Max Qty', f'Profit per Trip ({selected_ship})', 'Flight Time (min)', 'Total Profit', 'ROI (%)']
        if is_rental and rental_cost_per_day:
            columns.append('Covers Rental?')
        ws.append(columns)
        style_header_row(ws[ws.max_row])
        
        green_fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
        green_font = Font(color="FFFFFF", bold=True)
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        red_font = Font(color="FFFFFF", bold=True)
        
        for op in opportunities:
            trip_qty = min(op['max_qty'], ship_capacity)
            profit_per_trip = trip_qty * op['profit_per_mt'] if op.get('profit_per_mt') is not None else 0
            flight_time = get_flight_time(origin, op['destination'])
            
            # Calcular costo de compra y ROI
            cost_to_buy = op['source_available'] * op['source_selling']
            roi_percent = (op['total_profit'] / cost_to_buy * 100) if cost_to_buy > 0 else 0
            
            row = [
                op['commodity'],
                op['source'],
                op['source_selling'],
                op['source_available'],
                op['destination'],
                op['destination_buying'],
                op['destination_capacity'],
                op['max_qty'],
                profit_per_trip,
                flight_time,
                op['total_profit'],
                f"{roi_percent:.2f}",
            ]
            if is_rental and rental_cost_per_day:
                covers_rental = "YES" if profit_per_trip >= rental_cost_per_day else "NO"
                row.append(covers_rental)
            ws.append(row)
            
            # Calcular costo de compra y ROI
            cost_to_buy = op['source_available'] * op['source_selling']
            roi_percent = (op['total_profit'] / cost_to_buy * 100) if cost_to_buy > 0 else 0
            
            # Color coding for ROI
            roi_cell = ws[ws.max_row][11]  # Columna ROI (índice 11)
            if roi_percent > 100:
                roi_cell.fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")  # Verde oscuro
                roi_cell.font = Font(color="FFFFFF", bold=True)
            elif roi_percent >= 50:
                roi_cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")  # Verde claro
                roi_cell.font = Font(color="FFFFFF", bold=True)
            elif roi_percent >= 20:
                roi_cell.fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")  # Amarillo
                roi_cell.font = Font(color="000000", bold=True)
            elif roi_percent >= 0:
                roi_cell.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")  # Naranja
                roi_cell.font = Font(color="FFFFFF", bold=True)
            else:
                roi_cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # Rojo
                roi_cell.font = Font(color="FFFFFF", bold=True)
            
            # Resaltar profits altos
            if op['total_profit'] > 10000:  # Ejemplo threshold
                for i, cell in enumerate(ws[ws.max_row]):
                    if i != 11:  # No sobrescribir el color del ROI
                        cell.fill = high_profit_fill
            
            # Colorear columna Covers Rental?
            if is_rental and rental_cost_per_day:
                last_cell = ws[ws.max_row][-1]
                if last_cell.value == "YES":
                    last_cell.fill = green_fill
                    last_cell.font = green_font
                else:
                    last_cell.fill = red_fill
                    last_cell.font = red_font
        auto_size_columns(ws)

        # Hoja Trade Routes: proponer rutas optimizadas
        ws_routes = wb.create_sheet(title='Trade Routes')
        ws_routes.append(['Route', 'Commodity', 'Total Qty', 'Total Profit', f'Profit per Trip ({selected_ship})', 'Flight Time (min)', 'Trips Needed'])
        style_header_row(ws_routes[ws_routes.max_row])

        # Agrupar por source-destination-commodity para obtener profit por viaje real
        routes = {}
        for op in opportunities:
            key = (op['source'], op['destination'], op['commodity'])
            if key not in routes:
                routes[key] = {'total_qty': 0, 'total_profit': 0, 'profit_per_mt': op['profit_per_mt']}
            routes[key]['total_qty'] += op['max_qty']
            routes[key]['total_profit'] += op['total_profit']

        for (source, destination, commodity), data in sorted(routes.items(), key=lambda x: x[1]['total_profit'], reverse=True):
            profit_per_mt = data['profit_per_mt']
            trip_qty = min(data['total_qty'], ship_capacity)
            profit_per_trip = trip_qty * profit_per_mt
            flight_time = get_flight_time(origin, destination)
            trips_needed = max(1, (data['total_qty'] + ship_capacity - 1) // ship_capacity) if ship_capacity > 0 else 1
            ws_routes.append([
                f"{source} -> {destination}",
                commodity,
                data['total_qty'],
                data['total_profit'],
                profit_per_trip,
                flight_time,
                trips_needed
            ])
        auto_size_columns(ws_routes)

    # Hoja Profit Trips: cadenas de viajes recomendadas
    if chains:
        ws_profit = wb.create_sheet(title='Profit Trips')
        ws_profit.append(['Route', 'Commodities', 'Flight Time (min)', 'Profit Trip 1 (CR)', 'Profit Trip 2 (CR)', 'Total Profit (CR)', 'Efficiency (CR/hour)'])
        style_header_row(ws_profit[ws_profit.max_row])
        
        for chain in chains:
            ws_profit.append([
                chain['route'],
                chain['commodities'],
                chain['total_time_mins'],
                f"{chain['profit_1']:.0f}",
                f"{chain['profit_2']:.0f}",
                f"{chain['total_profit']:.0f}",
                f"{chain['efficiency']:.0f}"
            ])
            
            # Resaltar filas con alta eficiencia
            if chain['efficiency'] > 5000:  # Umbral de eficiencia
                for cell in ws_profit[ws_profit.max_row]:
                    cell.fill = high_profit_fill
        
        auto_size_columns(ws_profit)

    try:
        wb.save(output_file)
        print(f"Datos guardados en {output_file} con ship {selected_ship} (capacidad {ship_capacity} MT)")
    except PermissionError:
        timestamp = int(time.time())
        base, ext = os.path.splitext(output_file)
        alternate_file = f"{base}_{timestamp}{ext}"
        wb.save(alternate_file)
        print(f"Advertencia: no se pudo guardar '{output_file}' porque el archivo está en uso. Guardado como '{alternate_file}'.")
    except Exception as e:
        print(f"Error guardando el Excel: {e}")
        raise

def main(image_folder, selected_ship=None, output_file='final_trade.xlsx', budget=None):
    """Procesa todas las imágenes en la carpeta y guarda en Excel."""
    if selected_ship is None:
        print("Selecciona categoría:")
        print("1. AIR AND SPACE")
        print("2. ONLY AIR")
        while True:
            try:
                category_choice = int(input("Ingresa el número de la categoría: ").strip())
                if category_choice == 1:
                    category = "AIR AND SPACE"
                    ships = SHIPS[category]
                    break
                elif category_choice == 2:
                    category = "ONLY AIR"
                    ships = SHIPS[category]
                    break
                else:
                    print("Opción inválida. Elige 1 o 2.")
            except ValueError:
                print("Por favor ingresa un número válido.")

        print(f"\nShips en {category}:")
        ship_list = list(ships.keys())
        for i, ship in enumerate(ship_list, 1):
            capacity = calculate_ship_capacity(ship)
            print(f"{i}. {ship}: {capacity} MT")

        while True:
            try:
                ship_choice = int(input("Ingresa el número del ship: ").strip())
                if 1 <= ship_choice <= len(ship_list):
                    selected_ship = ship_list[ship_choice - 1]
                    break
                else:
                    print(f"Número inválido. Elige entre 1 y {len(ship_list)}.")
            except ValueError:
                print("Por favor ingresa un número válido.")

        # Seleccionar origen
        print("\nSelecciona tu punto de origen:")
        print("1. Delois Spot")
        print("2. Kansas")
        while True:
            try:
                origin_choice = int(input("Ingresa el número del origen: ").strip())
                if origin_choice == 1:
                    origin = "Delois Spot"
                    break
                elif origin_choice == 2:
                    origin = "Kansas"
                    break
                else:
                    print("Opción inválida. Elige 1 o 2.")
            except ValueError:
                print("Por favor ingresa un número válido.")
    else:
        origin = "Delois Spot"  # Default origin if ship is passed via args

    ship_capacity = calculate_ship_capacity(selected_ship)
    ship_rental_cost = SHIPS[next(cat for cat, ships in SHIPS.items() if selected_ship in ships)][selected_ship].get("rental_cost_per_day")
    
    # Preguntar si el ship es comprado o alquilado
    is_rental = False
    if ship_rental_cost is not None:
        print(f"\n{selected_ship} puede ser alquilado ({ship_rental_cost} CR/día) o comprado.")
        while True:
            rental_choice = input("¿Alquilado (A) o Comprado (C)? ").strip().upper()
            if rental_choice == 'A':
                is_rental = True
                break
            elif rental_choice == 'C':
                is_rental = False
                break
            else:
                print("Por favor ingresa 'A' o 'C'.")
    else:
        print(f"\n{selected_ship} solo puede ser comprado (no tiene opción de alquiler).")
        is_rental = False

    print(f"Selected ship: {selected_ship} with capacity {ship_capacity} MT ({'Alquilado' if is_rental else 'Comprado'})")
    print(f"Origin: {origin}")
    
    # Pregunta de presupuesto si no se proporciona vía argumento
    if budget is None:
        print("\n¿Tienes un presupuesto inicial para calcular ROI?")
        budget_input = input("Presupuesto en CR (dejar en blanco para omitir): ").strip()
        if budget_input:
            try:
                budget = int(budget_input)
                print(f"Presupuesto establecido: {budget} CR")
            except ValueError:
                print("Presupuesto no válido, continuando sin presupuesto.")
                budget = None

    if not os.path.exists(image_folder):
        print(f"La carpeta {image_folder} no existe.")
        return

    data_dict = {'Cities': {'header': ['Location', 'Category', 'Commodity Type', 'Quantity MT', 'Reserve MT', 'Selling CR/MT', 'Buying CR/MT', 'Maximum MT'], 'rows': []},
                 'EasyDock': {'header': ['Location', 'Category', 'Name', 'MT', 'Buying MT', 'Buying CR', 'Selling MT', 'Selling CR'], 'rows': []}}

    for filename in os.listdir(image_folder):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.tiff')):
            base_name = filename.rsplit('_', 1)[0]  # Ej. "Paris_1.png" -> "Paris"
            if base_name.lower() == 'easydock':
                sheet_key = 'EasyDock'
            else:
                sheet_key = 'Cities'
            image_path = os.path.join(image_folder, filename)
            text = extract_text_from_image(image_path)
            if text:
                parsed = parse_text_to_data(text)
                # Agregar location a cada row
                for row in parsed['rows']:
                    row.insert(0, base_name)
                data_dict[sheet_key]['rows'].extend(parsed['rows'])
                print(f"Procesado: {filename} -> Agregado a: {sheet_key}")
            else:
                print(f"No se pudo extraer texto de: {filename}")

    if data_dict['Cities']['rows'] or data_dict['EasyDock']['rows']:
        # Generar oportunidades y cadenas recomendadas
        catalog = build_trade_catalog(data_dict)
        
        # Debug: Print all locations found
        print(f"\nDEBUG: Locations found in catalog: {list(catalog.keys())}")
        for location, commodities in catalog.items():
            print(f"DEBUG: {location} has {len(commodities)} commodities")
            for commodity, data in commodities.items():
                print(f"  {commodity}: sell={data['selling']}, buy={data['buying']}, sell_cap={data['sell_capacity']}, buy_cap={data['buy_capacity']}")
        
        opportunities = find_trade_opportunities(catalog)
        chains = []
        roi_list = []
        
        if opportunities:
            print(f"\n{'='*70}")
            print(f"Total de oportunidades encontradas: {len(opportunities)}")
            print(f"{'='*70}\n")
            
            # Mostrar recomendaciones de viajes encadenados
            chains = recommend_travel_chains(opportunities, origin, ship_capacity)
            if chains:
                print("TOP 10 CADENAS DE VIAJES RECOMENDADAS (A → B → C):")
                print("-" * 70)
                for i, chain in enumerate(chains, 1):
                    hours = chain['total_time_mins'] / 60
                    print(f"\n{i}. {chain['route']}")
                    print(f"   Commodities: {chain['commodities']}")
                    print(f"   Tiempo total: {chain['total_time_mins']} min ({hours:.1f} horas)")
                    print(f"   Ganancia viaje 1: {chain['profit_1']:.0f} CR")
                    print(f"   Ganancia viaje 2: {chain['profit_2']:.0f} CR")
                    print(f"   Ganancia total: {chain['total_profit']:.0f} CR")
                    print(f"   Eficiencia: {chain['efficiency']:.0f} CR/hora")
                print("\n" + "-" * 70)
            
            # Calcular y mostrar ROI
            roi_list = calculate_roi_opportunities(opportunities, budget=budget)
            if roi_list:
                print("\nTOP ROI OPORTUNIDADES:")
                if budget:
                    print(f"Presupuesto: {budget} CR")
                print("-" * 70)
                for i, roi in enumerate(roi_list[:10], 1):  # Mostrar top 10
                    if budget and roi['cost_to_buy'] > budget:
                        continue  # Saltar si no cabe en presupuesto
                    print(f"\n{i}. {roi['route']}")
                    print(f"   Commodity: {roi['commodity']}")
                    print(f"   Qty to buy: {roi['qty_to_buy']} MT @ {roi['buy_price_per_mt']} CR/MT")
                    print(f"   Cost: {roi['cost_to_buy']:.0f} CR")
                    print(f"   Profit: {roi['profit']:.0f} CR")
                    print(f"   ROI: {roi['roi_percent']:.2f}%")
                print("\n" + "-" * 70)
        
        # Guardar a Excel
        save_to_excel(data_dict, selected_ship, ship_capacity, output_file=output_file, is_rental=is_rental, rental_cost_per_day=ship_rental_cost, origin=origin, chains=chains, budget=budget, roi_list=roi_list)
    else:
        print("No se encontraron datos para guardar.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Extrae datos de imágenes y guarda resultados estáticos en Excel')
    parser.add_argument('--images', default='images', help='Carpeta de imágenes a procesar')
    parser.add_argument('--ship', default=None, help='Ship a usar para cálculo de rutas')
    parser.add_argument('--output', default='final_trade.xlsx', help='Archivo Excel de salida')
    parser.add_argument('--budget', default=None, type=int, help='Presupuesto inicial en CR para calcular ROI')
    args = parser.parse_args()
    main(args.images, selected_ship=args.ship, output_file=args.output, budget=args.budget)