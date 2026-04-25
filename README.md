# Automatización OCR a Excel para Rise the Video Province

Este proyecto automatiza la extracción de texto de imágenes de recursos/commodities para ciudades y EasyDock del juego Rise the Video Province.

## Formato de las Imágenes
- **Ciudades**: Archivos nombrados como `"NombreCiudad_1.png"` (ej. `"Paris_1.png"`).
- **EasyDock**: `"EasyDock_1.png"`.
- El script detecta automáticamente si la imagen es de ciudad o de EasyDock y genera una hoja separada por cada prefijo.

## Requisitos
1. **Instalar Tesseract OCR**:
   - Descarga e instala Tesseract desde: https://github.com/UB-Mannheim/tesseract/wiki
   - Asegúrate de que esté en el PATH o ajusta la ruta en el script (`pytesseract.pytesseract.tesseract_cmd`).

2. **Entorno Python**:
   - El entorno virtual ya está configurado.
   - Paquetes instalados: `pytesseract`, `openpyxl`, `pillow`.

## Uso
1. Coloca tus imágenes en la carpeta `images/`.
2. Ejecuta el script:
   ```
   cd "e:\PYTHON\RiseTVP\Script Trade"
   & "e:/PYTHON/RiseTVP/Script Trade/.venv/Scripts/python.exe" ocr_to_excel.py
   ```
   O con presupuesto:
   ```
   & "e:/PYTHON/RiseTVP/Script Trade/.venv/Scripts/python.exe" ocr_to_excel.py --budget 50000
   ```
3. El script pedirá seleccionar:
   - Categoría (1 para AIR AND SPACE, 2 para ONLY AIR)
   - Ship dentro de esa categoría
   - Punto de origen (1 para Delois Spot, 2 para Kansas)
   - Si el ship puede ser alquilado, preguntará si deseas alquilarlo o comprarlo
4. Luego generará `final_trade.xlsx` con las hojas Config, Cities, EasyDock, Opportunities, Trade Routes, Profit Trips y ROI Analysis (si hay presupuesto).

**Con presupuesto:** Se filtra el análisis ROI para mostrar solo oportunidades que quepan en el presupuesto.

Los tiempos incluyen:
- 5 minutos de despegue
- Tiempo de vuelo entre ciudades (ver tabla abajo)
- 10 minutos de aterrizaje

**Desde Delois Spot (tiempo de vuelo en minutos):**
- Alphaville: 60 | Comstock: 55 | Deadwood: 60 | Ederar: 60
- Erie: 60 | Freedom: 150 | Gettysburg: 60 | Kansas: 150
- Lancaster: 120 | Pimli: 35 | SovietUnion: 60 | Terrazul: 60
- Sharney 1: 60 | Sharney 2: 120 | Sharney 3: 180

**Desde Kansas (tiempo de vuelo en minutos):**
- Alphaville: 35 | Comstock: 30 | Deadwood: 25 | Ederar: 20
- Erie: 45 | Freedom: 15 | Gettysburg: 40 | Lancaster: 50
- Pimli: 10 | SovietUnion: 65 | Terrazul: 60
- Sharney 1: 30 | Sharney 2: 60 | Sharney 3: 90

## Recomendación de Viajes Encadenados

El sistema automáticamente identifica y recomienda las mejores cadenas de viajes (A → B → C):
- Después de vender una mercancía en punto B, compra otra en B y vuela a C
- Las cadenas se ordenan por **eficiencia (CR/hora)** 
- Útil para maximizar tiempo y ganancia sin volver al origen
- Se muestran los top 10 cuando generas el Excel

## Análisis ROI

El ROI (Return on Investment) aparece en la hoja **Opportunities** como una columna con código de colores:
- **ROI = (Ganancia Total / Costo Total) × 100%**

**Escala de colores:**
- 🟢 **Verde Oscuro**: ROI > 100% (Excelente retorno)
- 🟢 **Verde Claro**: ROI 50-100% (Muy bueno)
- 🟡 **Amarillo**: ROI 20-50% (Aceptable)
- 🟠 **Naranja**: ROI < 20% (Bajo retorno)

**Ejemplo:** Si compras 10 MT a 100 CR/MT (costo 1000 CR) y ganas 500 CR, el ROI es 50% (Verde Claro)

Si usas `--budget`, se filtran las oportunidades que quepan en tu presupuesto inicial.

## Archivos de salida

El script genera `final_trade.xlsx` con las siguientes hojas:
- **Config**: Ship seleccionado, capacidad, origen, estado (alquilado/comprado)
- **Cities**: Datos extraídos de ciudades (cantidad, precio, etc.)
- **EasyDock**: Datos extraídos de EasyDock
- **Opportunities**: Rutas de una sola parada con tiempos de vuelo y **ROI (%)**
- **Trade Routes**: Rutas agrupadas por destino con eficiencia
- **Profit Trips**: Cadenas de viajes encadenadas (A → B → C) ordenadas por eficiencia (CR/hora)


## Naves y Costos de Alquiler

**AIR AND SPACE:**
- E-10 Saint: No rentable (comprado solo)
- E-11 Saint: No rentable (comprado solo)
- P-13 Prowler: No rentable (comprado solo)
- W-6 Manx: No rentable (comprado solo)

**ONLY AIR:**
- A-4 Wanderer: 1359 CR/día (alquilable o comprado)
- T-19 Stratomaster: 2264 CR/día (alquilable o comprado)

Un día en el juego = 14 horas.

Si eliges alquilar, la hoja Config mostrará el profit mínimo necesario por viaje para cubrir el costo de alquiler diario. La hoja Opportunities tendrá una columna "Covers Rental?" indicando qué trades generan suficiente profit para cubrir el alquiler.

## Ejecutar desde Windows
Puedes usar el archivo `run_trade.bat` para lanzar el script fácilmente desde Windows:
```bat
cd /d "e:\PYTHON\RiseTVP\Script Trade"
run_trade.bat
```
Opcionalmente puedes pasar un ship y carpeta de imágenes:
```bat
run_trade.bat --ship "A-4 Wanderer" --images images --output final_trade.xlsx
```

## Hoja Config
- **Propósito**: Permite seleccionar el ship usado para cálculos de trade routes.
- **Contenido**:
  - Selected Ship: El ship elegido (inicialmente el seleccionado al ejecutar).
  - Capacity: Capacidad en MT del ship seleccionado.
  - Available Ships: Lista de todos los ships con categoría y capacidad.
  - Dropdown: En la celda B3 hay un menú desplegable con todos los ships disponibles. Cambia la selección y ejecuta el script de nuevo para recalcular con el nuevo ship.

## Ships Disponibles
El script incluye selección de ships con capacidades calculadas:

**AIR AND SPACE:**
- E-10 Saint: 109 MT (7 + 6x17)
- E-11 Saint: 109 MT (7 + 6x17)
- P-13 Prowler: 1 MT
- W-6 Manx: 58 MT (7 + 3x17)

**ONLY AIR:**
- A-4 Wanderer: 18 MT (1 + 1x17)
- T-19 Stratomaster: 18 MT (1 + 1x17)

La capacidad se usa para calcular profit por viaje y viajes necesarios en Trade Routes.
- **Hoja Cities**: Contiene todas las ciudades con columnas: Location, Category, Commodity Type, Quantity MT, Reserve MT, Selling CR/MT, Buying CR/MT, Maximum MT. Solo incluye filas con Quantity MT > 0 y Selling CR/MT > 0. Formato: encabezados en negrita, filas alternas coloreadas.
- **Hoja EasyDock**: Contiene EasyDock con columnas: Location, Category, Name, MT, Buying MT, Buying CR, Selling MT, Selling CR. Solo incluye filas con Selling CR > 0. Formato similar.
- **Hoja Opportunities**: Lista oportunidades de comercio por categoría, con columnas: Category, Source, Source selling CR/MT, Destination, Destination buying CR/MT, Profit per MT, Max Qty, Total Profit. Ordenado por total profit descendente. Profits altos (>10,000) resaltados en amarillo.
- **Hoja Trade Routes**: Propone rutas optimizadas agrupadas por source-destination, con commodities a transportar, total qty, total profit, profit por viaje (basado en capacidad del ship) y viajes necesarios.

## Hoja de Oportunidades
- Además, el script crea una hoja `Opportunities` en `output.xlsx`.
- En ella se listan las rutas con `selling` más barato en un punto A y `buying` más caro en un punto B.
- Columnas: Category, Source, Source selling CR/MT, Destination, Destination buying CR/MT, Profit CR/MT.

## Cómo interpreta el script
- **Formato ciudad**: detecta columnas como `Category`, `Commodity Type`, `Quantity MT`, `Reserve MT`, `Selling CR/MT`, `Buying CR/MT`, `Maximum MT`.
- **Formato EasyDock**: detecta columnas como `Category`, `Name`, `MT`, `Buying MT`, `Buying CR`, `Selling MT`, `Selling CR`.
- Las categorías de commodities se asignan en orden fijo: 1. Rare/Precious, 2. Foodstuffs, 3. Natural Materials, 4. Fuel Ore, 5. Consumer Goods, 6. Fabricated Material, 7. Refined Fuel.
- Si no reconoce un formato de tabla, cae en un parseo de `Clave: Valor`.

## Ajustes y mejoras
- Si necesitas otro formato de imagen, modifica `parse_text_to_data` en `ocr_to_excel.py`.
- Para calcular oportunidades de trayectos/comercios, añade análisis posterior usando los datos en `output.xlsx` o extiende el script para generar resultados.

## Notas
- Las imágenes deben tener texto claro y legible para un buen OCR.
- Si hay errores, revisa la instalación de Tesseract y la calidad de las fotos.