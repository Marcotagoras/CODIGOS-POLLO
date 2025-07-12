import re
import pandas as pd
from pathlib import Path
import pdfplumber

# === CONFIGURACIÓN ===
CARPETA = Path(r"C:\Users\Dell\OneDrive\GS COMERCIO\1 FINANZAS\2025\ESTADO DE CUENTA\Estados_liberados")

# === EXPRESIONES REGULARES ===
re_moneda = re.compile(r"MONEDA:\s+(SOLES|D[ÓO]LARES)", re.IGNORECASE)
re_mes_anio = re.compile(r"ESTADO DE CUENTA NEGOCIOS\s+Mes:\s+([A-Za-z]+)\s+(\d{4})", re.IGNORECASE)
re_linea = re.compile(
    r"(\d{2}/\d{2})\s+\d{2}/\d{2}\s+(.+?)\s+(-?[\d,]+\.\d{2})\s+(-?[\d,]+\.\d{2})",
    re.MULTILINE
)

# === FUNCIÓN DE CATEGORIZACIÓN ===
def categorizar(descripcion: str) -> str:
    desc = descripcion.upper()
    if "SUNAT" in desc:
        return "SUNAT"
    elif "ITF" in desc:
        return "ITF"
    elif "TRANSFERENCIA" in desc:
        return "TRANSFERENCIA"
    elif "ABONO" in desc or "DEPOSITO" in desc:
        return "ABONO"
    elif "RETIRO" in desc or "CARGO" in desc:
        return "RETIRO"
    else:
        return "OTROS"

# === CONTENEDORES POR MONEDA ===
movimientos = {"SOLES": [], "DOLARES": []}

# === PROCESAMIENTO DE CADA PDF ===
for archivo in CARPETA.glob("*.pdf"):
    with pdfplumber.open(archivo) as pdf:
        texto = "\n".join([page.extract_text() or '' for page in pdf.pages])

    moneda_match = re_moneda.search(texto)
    moneda = moneda_match.group(1).upper().replace("Ó", "O") if moneda_match else "DESCONOCIDO"

    if moneda not in movimientos:
        print(f"⚠️  Moneda no reconocida en: {archivo.name}")
        continue

    # Detectar mes y año del estado
    mes_anio_match = re_mes_anio.search(texto)
    mes_año = f"{mes_anio_match.group(1).capitalize()} {mes_anio_match.group(2)}" if mes_anio_match else ""

    matches = list(re_linea.finditer(texto))
    print(f"📄 {archivo.name} — {moneda} — {len(matches)} movimientos encontrados")

    for match in matches:
        fecha, descripcion, monto, saldo = match.groups()
        categoria = categorizar(descripcion)
        movimientos[moneda].append({
            "archivo": archivo.name,
            "mes_año": mes_año,
            "fecha": fecha,
            "descripcion": descripcion.strip(),
            "monto": float(monto.replace(",", "")),
            "saldo": float(saldo.replace(",", "")),
            "categoria": categoria
        })

# === EXPORTACIÓN A EXCEL CON HOJAS ===
output_path = CARPETA / "movimientos_bancarios.xlsx"
with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
    for moneda, registros in movimientos.items():
        if registros:
            df = pd.DataFrame(registros)
            df.to_excel(writer, sheet_name=moneda, index=False)
            print(f"✅ Hoja '{moneda}' exportada con {len(df)} movimientos")
        else:
            print(f"ℹ️  No se encontraron movimientos en {moneda}")

print(f"\n📦 Archivo generado: {output_path}")
