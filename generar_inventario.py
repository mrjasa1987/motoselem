import json, os, sys

try:
    import openpyxl
except ImportError:
    print("Instalando openpyxl...")
    os.system(f"{sys.executable} -m pip install openpyxl -q")
    import openpyxl

# Buscar archivo Excel
PATHS = [
    r"C:\Users\Gerencia\Downloads\inv_15_abril_520_pm.XLSX",
    r"C:\Users\Gerencia\Downloads\inv 15 abril 520 pm.XLSX",
    r"C:\Users\Gerencia\Documents\inv_15_abril_520_pm.XLSX",
    r"C:\Users\Gerencia\Documents\inv 15 abril 520 pm.XLSX",
    r"C:\Users\Gerencia\Desktop\inv_15_abril_520_pm.XLSX",
    r"C:\Users\Gerencia\Desktop\inv 15 abril 520 pm.XLSX",
]

excel_path = None
for p in PATHS:
    if os.path.exists(p):
        excel_path = p
        break

if not excel_path:
    print("ERROR: No se encontró el archivo Excel")
    sys.exit(1)

print(f"Leyendo: {excel_path}")

wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
ws = wb.active

# Códigos cortos de sucursales (columnas F-T)
SUC_KEYS = ['OBR','SAN','CAU','C71','FCO','KAN','PAL','MEL','MXN','MIR','MOT','MUL','PEN','PRO','VAL']

productos = []
for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    if not row or not row[0]:
        continue

    folio = row[0]
    codigo = str(row[1] or '').strip()
    descripcion = str(row[2] or '').strip()
    marca = str(row[3] or '').strip()
    precio = row[4] or 0

    try:
        precio = round(float(precio), 2)
    except (ValueError, TypeError):
        precio = 0

    # Stock por sucursal (columnas F=5 a T=19)
    stock = {}
    total_stock = 0
    for j, key in enumerate(SUC_KEYS):
        val = row[5 + j] if len(row) > 5 + j else 0
        try:
            val = int(val or 0)
        except (ValueError, TypeError):
            val = 0
        if val > 0:
            stock[key] = val
        total_stock += val

    if total_stock <= 0:
        continue

    productos.append({
        "f": folio,
        "c": codigo,
        "d": descripcion,
        "m": marca,
        "p": precio,
        "s": stock
    })

wb.close()

# Crear carpeta data
os.makedirs(os.path.join(os.path.dirname(__file__), "data"), exist_ok=True)

output_path = os.path.join(os.path.dirname(__file__), "data", "inventario.json")
with open(output_path, "w", encoding="utf-8") as f:
    json.dump(productos, f, ensure_ascii=False, separators=(',', ':'))

size = os.path.getsize(output_path)
size_str = f"{size/1024/1024:.1f} MB" if size > 1024*1024 else f"{size/1024:.0f} KB"

print(f"Total productos con stock: {len(productos):,}")
print(f"Tamaño del archivo: {size_str}")
print("inventario.json generado exitosamente")
