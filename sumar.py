"""Agrupa las facturas validadas por trimestre y tipo de IVA, y escribe
un Excel de sumatorios por cliente.

Piso 5: sin llamadas a la API, corre gratis. Lee rebudes/validadas/
(GASTOS) y apartados/ingressos_validadas/ (INGRESOS) de cada cliente --
son fixtures, nunca se tocan. Cada run reescribe el Excel entero: no
hay nada que idempotizar, es gratis y siempre tiene que reflejar el
ultimo estado de las validadas, no el de la vez anterior.

Piso 5B: viste el Excel -- bloque de titulo por hoja, colores suaves,
formato de moneda, bordes en TOTAL/retenciones, anchos de columna y
ajuste de texto en Motivos. La estructura de datos no cambia.
"""

import csv
import json
import os
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

TIPOS_IVA = [0, 4, 5, 10, 12, 21]
EXTENSIONES_ORIGINAL = (".pdf", ".jpg", ".jpeg", ".png")

FORMATO_MONEDA = '#,##0.00" €"'

FUENTE_TITULO = Font(bold=True, size=14)
FUENTE_SUBTITULO = Font(size=11)
ESTILO_ENCABEZADO = Font(bold=True)
ESTILO_ENLACE = Font(color="0563C1", underline="single")

RELLENO_BLOQUE = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
RELLENO_AVISO = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
BORDE_SUPERIOR = Border(top=Side(style="thin"))
AJUSTE_TEXTO = Alignment(wrap_text=True, vertical="top")

GENERADO_EL = datetime.now().strftime("%d/%m/%Y %H:%M")


def leer_clientes():
    with open("clientes/clientes.csv") as f:
        return list(csv.DictReader(f))


def cargar_validadas(carpeta):
    """Lee todos los .json de una carpeta de validadas. Si la carpeta
    no existe (el cliente no tiene ese flujo todavia), devuelve vacio."""
    facturas = []
    if not os.path.isdir(carpeta):
        return facturas
    for nombre in sorted(os.listdir(carpeta)):
        if not nombre.lower().endswith(".json"):
            continue
        with open(os.path.join(carpeta, nombre)) as f:
            facturas.append((nombre, json.load(f)))
    return facturas


def trimestre_de(fecha):
    """'2026-04-21' -> '2T'. Si la fecha falta o no se puede leer, None."""
    if not fecha:
        return None
    try:
        mes = int(fecha[5:7])
    except (ValueError, IndexError):
        return None
    return f"{(mes - 1) // 3 + 1}T"


def encontrar_original(carpeta_origen, nombre_json):
    base = os.path.splitext(nombre_json)[0]
    for ext in EXTENSIONES_ORIGINAL:
        ruta = os.path.join(carpeta_origen, base + ext)
        if os.path.exists(ruta):
            return ruta
    return None


def escribir_titulo(ws, nombre_cliente, nif_cliente):
    ws.cell(row=1, column=1, value=f"SUMATORIOS {nombre_cliente}").font = FUENTE_TITULO
    ws.cell(row=2, column=1, value=f"NIF {nif_cliente}").font = FUENTE_SUBTITULO
    ws.cell(row=3, column=1, value="Ejercici 2026").font = FUENTE_SUBTITULO
    ws.cell(row=4, column=1, value=f"Generat el {GENERADO_EL}").font = FUENTE_SUBTITULO
    return 6  # fila 5 en blanco, el resto empieza en la 6


def sumar_bloque(facturas):
    """Suma un bloque (gastos o ingresos) de un trimestre: base/cuota
    por tipo de IVA, total y retencion de las OK, y aparte la lista de
    las que hay que revisar."""
    sumas = {tipo: {"base": 0.0, "cuota": 0.0} for tipo in TIPOS_IVA}
    sumas["otros"] = {"base": 0.0, "cuota": 0.0}
    total_ok = 0.0
    retencion_ok = 0.0
    revisar = []

    for nombre, datos in facturas:
        if datos.get("estado") != "OK":
            revisar.append((nombre, datos))
            continue
        for linea in datos.get("lineas_iva") or []:
            tipo = linea.get("tipo_iva")
            clave = tipo if tipo in TIPOS_IVA else "otros"
            sumas[clave]["base"] += linea.get("base") or 0
            sumas[clave]["cuota"] += linea.get("cuota") or 0
        total_ok += datos.get("total") or 0
        retencion_ok += datos.get("retencion_cuota") or 0

    return sumas, total_ok, retencion_ok, revisar


def escribir_bloque(ws, fila, titulo, sumas, total_ok, con_retencion, retencion_ok):
    for col in range(1, 5):
        celda = ws.cell(row=fila, column=col)
        celda.font = ESTILO_ENCABEZADO
        celda.fill = RELLENO_BLOQUE
    ws.cell(row=fila, column=1, value=titulo)
    fila += 1

    for col, texto in enumerate(["Tipo IVA", "Base", "Cuota", "Total"], start=1):
        ws.cell(row=fila, column=col, value=texto).font = ESTILO_ENCABEZADO
    fila += 1

    tipos_a_escribir = list(TIPOS_IVA)
    if sumas["otros"]["base"] or sumas["otros"]["cuota"]:
        tipos_a_escribir.append("otros")

    suma_base = 0.0
    suma_cuota = 0.0
    for tipo in tipos_a_escribir:
        base = sumas[tipo]["base"]
        cuota = sumas[tipo]["cuota"]
        etiqueta_tipo = "OTROS" if tipo == "otros" else tipo
        ws.cell(row=fila, column=1, value=etiqueta_tipo)
        celda_base = ws.cell(row=fila, column=2, value=round(base, 2))
        celda_base.number_format = FORMATO_MONEDA
        celda_cuota = ws.cell(row=fila, column=3, value=round(cuota, 2))
        celda_cuota.number_format = FORMATO_MONEDA
        suma_base += base
        suma_cuota += cuota
        fila += 1

    for col in range(1, 5):
        ws.cell(row=fila, column=col).border = BORDE_SUPERIOR
    ws.cell(row=fila, column=1, value="TOTAL").font = ESTILO_ENCABEZADO
    celda = ws.cell(row=fila, column=2, value=round(suma_base, 2))
    celda.font = ESTILO_ENCABEZADO
    celda.number_format = FORMATO_MONEDA
    celda = ws.cell(row=fila, column=3, value=round(suma_cuota, 2))
    celda.font = ESTILO_ENCABEZADO
    celda.number_format = FORMATO_MONEDA
    celda = ws.cell(row=fila, column=4, value=round(total_ok, 2))
    celda.font = ESTILO_ENCABEZADO
    celda.number_format = FORMATO_MONEDA
    fila += 1

    if con_retencion:
        for col in range(1, 5):
            ws.cell(row=fila, column=col).border = BORDE_SUPERIOR
        ws.cell(row=fila, column=1, value="Σ RETENCIONES").font = ESTILO_ENCABEZADO
        celda = ws.cell(row=fila, column=4, value=round(retencion_ok, 2))
        celda.font = ESTILO_ENCABEZADO
        celda.number_format = FORMATO_MONEDA
        fila += 1

    return fila + 1  # una fila en blanco antes del siguiente bloque


def escribir_pendientes(ws, fila, pendientes, carpeta_cliente):
    for col in range(1, 4):
        celda = ws.cell(row=fila, column=col)
        celda.font = ESTILO_ENCABEZADO
        celda.fill = RELLENO_AVISO
    ws.cell(row=fila, column=1, value="PENDIENTE DE REVISIÓN")
    fila += 1
    for col, texto in enumerate(["Archivo", "Tipo", "Motivos"], start=1):
        ws.cell(row=fila, column=col, value=texto).font = ESTILO_ENCABEZADO
    fila += 1

    for nombre, datos, tipo_bloque, carpeta_original in pendientes:
        celda = ws.cell(row=fila, column=1, value=nombre)
        ruta_original = encontrar_original(carpeta_original, nombre)
        if ruta_original:
            ruta_relativa = os.path.relpath(ruta_original, carpeta_cliente)
            celda.hyperlink = ruta_relativa
            celda.font = ESTILO_ENLACE
        else:
            print(f"AVISO: no se encontró el original de {nombre}")
        ws.cell(row=fila, column=2, value=tipo_bloque)
        celda_motivos = ws.cell(row=fila, column=3, value="; ".join(datos.get("motivos") or []))
        celda_motivos.alignment = AJUSTE_TEXTO
        fila += 1

    return fila + 1


for fila_cliente in leer_clientes():
    carpeta = fila_cliente["carpeta"]
    carpeta_cliente = f"clientes/{carpeta}"

    gastos = cargar_validadas(f"{carpeta_cliente}/rebudes/validadas")
    ingresos = cargar_validadas(f"{carpeta_cliente}/apartados/ingressos_validadas")

    if not gastos and not ingresos:
        continue

    # Agrupar por trimestre (o SIN FECHA si una REVISAR no trae fecha_factura)
    trimestres = {}
    for nombre, datos in gastos:
        t = trimestre_de(datos.get("fecha_factura")) or "SIN FECHA"
        trimestres.setdefault(t, {"gastos": [], "ingresos": []})["gastos"].append((nombre, datos))
    for nombre, datos in ingresos:
        t = trimestre_de(datos.get("fecha_factura")) or "SIN FECHA"
        trimestres.setdefault(t, {"gastos": [], "ingresos": []})["ingresos"].append((nombre, datos))

    wb = Workbook()
    wb.remove(wb.active)

    orden = sorted(t for t in trimestres if t != "SIN FECHA") + (
        ["SIN FECHA"] if "SIN FECHA" in trimestres else []
    )

    for trimestre in orden:
        datos_trimestre = trimestres[trimestre]
        ws = wb.create_sheet(trimestre)
        ws.column_dimensions["A"].width = 50
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 60
        ws.column_dimensions["D"].width = 16

        fila = escribir_titulo(ws, fila_cliente["nombre"], fila_cliente["nif"])
        pendientes = []

        if trimestre != "SIN FECHA":
            sumas_g, total_g, _, revisar_g = sumar_bloque(datos_trimestre["gastos"])
            fila = escribir_bloque(ws, fila, "GASTOS", sumas_g, total_g, con_retencion=False, retencion_ok=0)

            sumas_i, total_i, retencion_i, revisar_i = sumar_bloque(datos_trimestre["ingresos"])
            fila = escribir_bloque(ws, fila, "INGRESOS", sumas_i, total_i, con_retencion=True, retencion_ok=retencion_i)

            n_ok_g = len(datos_trimestre["gastos"]) - len(revisar_g)
            n_ok_i = len(datos_trimestre["ingresos"]) - len(revisar_i)
            print(
                f"{carpeta} / {trimestre}: GASTOS total={total_g:.2f} ({n_ok_g} OK, {len(revisar_g)} a revisar) | "
                f"INGRESOS total={total_i:.2f} ({n_ok_i} OK, {len(revisar_i)} a revisar)"
            )
        else:
            # Sin fecha no hay nada que sumar por trimestre -- solo se lista para revisar
            _, _, _, revisar_g = sumar_bloque(datos_trimestre["gastos"])
            _, _, _, revisar_i = sumar_bloque(datos_trimestre["ingresos"])
            print(f"{carpeta} / SIN FECHA: {len(revisar_g) + len(revisar_i)} a revisar (sin fecha_factura)")

        for nombre, datos in revisar_g:
            pendientes.append((nombre, datos, "GASTO", f"{carpeta_cliente}/rebudes/entrada"))
        for nombre, datos in revisar_i:
            pendientes.append((nombre, datos, "INGRESO", f"{carpeta_cliente}/apartados/ingressos"))

        if pendientes:
            escribir_pendientes(ws, fila, pendientes, carpeta_cliente)

    ruta_excel = f"{carpeta_cliente}/sumatorios_2026.xlsx"
    wb.save(ruta_excel)
    print(f"Escrito: {ruta_excel}")
