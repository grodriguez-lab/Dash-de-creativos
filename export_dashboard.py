#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Exporta Monitoreo_Creativos_Meta.xlsx (ya lleno con fill_monitoreo.py y RECALCULADO
con recalc.py de la skill xlsx) a un data.json liviano para el dashboard estático
(index.html en GitHub Pages).

Uso:
  python export_dashboard.py "<ruta>/Monitoreo_Creativos_Meta.xlsx" data.json
  python export_dashboard.py "<ruta>/Monitoreo_Creativos_Meta.xlsx" data.json --fecha "17 jul 2026"

Requisitos:
- El archivo YA debe estar recalculado (data_only=True necesita los valores en caché;
  si no corriste recalc.py después de fill_monitoreo.py, las columnas de fórmula
  como CPI, ESTADO, Acción saldrán vacías en el JSON).
- No modifica el xlsx. Solo lee.
"""
import sys
import json
import argparse
from datetime import date, datetime
import openpyxl

SHEET_NAME = "Monitoreo Piezas"
HEADER_ROW = 4
FIRST_DATA_ROW = 5

# Columnas que nos interesan y su clave en el JSON (deben calzar EXACTO con los
# encabezados del Excel; si el usuario renombra una columna, ajustar este mapa).
COLS = {
    "Pieza ID": "pieza_id",
    "Ruta": "ruta",
    "Conjunto": "conjunto",
    "Formato": "formato",
    "Ángulo / Tema": "angulo",
    "Fecha inicio test": "fecha_inicio",
    "Cohorte": "cohorte",
    "Días en test": "dias_en_test",
    "Gasto ($)": "gasto",
    "Impresiones": "impresiones",
    "Clics": "clics",
    "Installs": "installs",
    "Views 3s": "views3s",
    "KYC compl. (opc.)": "kyc",
    "CPI": "cpi",
    "CTR": "ctr",
    "Hook": "hook",
    "CPI ref": "cpi_ref",
    "CPI vs ref": "cpi_vs_ref",
    "ESTADO": "estado",
    "Acción": "accion",
    "Notas": "notas",
}

# Mapeo de texto de ESTADO -> categoría normalizada + color de semáforo.
# Cubre variantes con o sin emoji, mayúsculas/minúsculas.
def clasificar_estado(estado_raw):
    if not estado_raw:
        return {"categoria": "sin_dato", "color": "gray", "label": "Sin dato"}
    s = str(estado_raw).lower()
    if "escal" in s:
        return {"categoria": "escalar", "color": "green", "label": "Escalar"}
    if "apag" in s:
        return {"categoria": "apagar", "color": "red", "label": "Apagar"}
    if "iter" in s:
        return {"categoria": "iterar", "color": "orange", "label": "Iterar"}
    if "manten" in s:
        return {"categoria": "mantener", "color": "yellow", "label": "Mantener"}
    return {"categoria": "otro", "color": "gray", "label": str(estado_raw)}


def to_native(v):
    """Convierte tipos de openpyxl (date, datetime) a algo serializable en JSON."""
    if v is None:
        return None
    if isinstance(v, (date, datetime)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def col_map(ws, header_row=HEADER_ROW):
    m = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v:
            m[str(v).strip()] = c
    return m


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", help="Ruta al Monitoreo_Creativos_Meta.xlsx ya recalculado")
    ap.add_argument("out_json", help="Ruta de salida, normalmente data.json")
    ap.add_argument("--fecha", default=None,
                     help='Etiqueta de fecha a mostrar en el dashboard, ej. "17 jul 2026". '
                          "Por defecto usa la fecha de hoy.")
    args = ap.parse_args()

    # data_only=True para leer VALORES calculados de las fórmulas (CPI, ESTADO, etc).
    # Esto requiere que el archivo haya sido recalculado con LibreOffice (recalc.py).
    wb = openpyxl.load_workbook(args.xlsx, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        print(f"ERROR: no encontré la hoja '{SHEET_NAME}'. Hojas disponibles: {wb.sheetnames}")
        sys.exit(1)
    ws = wb[SHEET_NAME]
    cm = col_map(ws)

    faltantes = [h for h in COLS if h not in cm]
    if faltantes:
        print(f"AVISO: no encontré estas columnas en el encabezado (fila {HEADER_ROW}): "
              f"{faltantes}. Se omiten en el export.")

    piezas = []
    r = FIRST_DATA_ROW
    empty_streak = 0
    while empty_streak < 15 and r <= ws.max_row:
        pid_col = cm.get("Pieza ID")
        pid_val = ws.cell(row=r, column=pid_col).value if pid_col else None
        if pid_val is None or str(pid_val).strip() == "":
            empty_streak += 1
            r += 1
            continue
        empty_streak = 0
        row = {}
        for header, key in COLS.items():
            if header not in cm:
                continue
            row[key] = to_native(ws.cell(row=r, column=cm[header]).value)
        estado_info = clasificar_estado(row.get("estado"))
        row["estado_categoria"] = estado_info["categoria"]
        row["estado_color"] = estado_info["color"]
        row["estado_label"] = estado_info["label"]
        piezas.append(row)
        r += 1

    # Resumen agregado (lo recalculamos aquí en vez de parsear la hoja 'Resumen
    # Viernes' en texto libre, porque las categorías ya están en la columna ESTADO).
    resumen = {"escalar": 0, "mantener": 0, "iterar": 0, "apagar": 0, "sin_dato": 0, "otro": 0}
    gasto_total = 0
    installs_total = 0
    nuevas_q = 0
    for p in piezas:
        resumen[p["estado_categoria"]] = resumen.get(p["estado_categoria"], 0) + 1
        gasto_total += p.get("gasto") or 0
        installs_total += p.get("installs") or 0
        cohorte = str(p.get("cohorte") or "")
        if "nueva" in cohorte.lower() or "🆕" in cohorte:
            nuevas_q += 1

    # Top ganadores (escalar, ordenados por menor CPI vs ref = mejor eficiencia relativa)
    escalar = [p for p in piezas if p["estado_categoria"] == "escalar"]
    apagar = [p for p in piezas if p["estado_categoria"] == "apagar"]

    def cpi_vs_ref_num(p):
        v = p.get("cpi_vs_ref")
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    ganadores = sorted(escalar, key=cpi_vs_ref_num)[:5]
    perdedores = sorted(apagar, key=cpi_vs_ref_num, reverse=True)[:5]

    data = {
        "generado": args.fecha or date.today().strftime("%d %b %Y"),
        "total_piezas": len(piezas),
        "gasto_total": round(gasto_total, 2),
        "installs_total": installs_total,
        "nuevas_q": nuevas_q,
        "resumen": resumen,
        "ganadores": [{"pieza_id": p["pieza_id"], "ruta": p.get("ruta"),
                        "cpi": p.get("cpi"), "cpi_vs_ref": p.get("cpi_vs_ref")} for p in ganadores],
        "perdedores": [{"pieza_id": p["pieza_id"], "ruta": p.get("ruta"),
                         "cpi": p.get("cpi"), "cpi_vs_ref": p.get("cpi_vs_ref")} for p in perdedores],
        "piezas": piezas,
    }

    with open(args.out_json, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"OK: {len(piezas)} piezas exportadas a {args.out_json}")
    print(f"Resumen: {resumen}")


if __name__ == "__main__":
    main()
