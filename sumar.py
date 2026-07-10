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

Piso 6A: hoja final "AVISOS" por cliente (regla 9 de CLAUDE.md --
nada muere en silencio tampoco en el enrutado): albaranes apartados,
paginas de ruido leidas de los manifiestos de trocear.py, y avisos de
consistencia (mismo proveedor, mas de un tipo de IVA). Solo informa,
no cambia ningun estado OK/REVISAR.

Piso 6B: encontrar_original tambien busca en subcarpetas por proveedor
(clientes con facturas de compra organizadas asi, ej. davinstal), y el
origen de las emitidas es personalizable por cliente.
"""

import csv
import json
import os
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

TIPOS_IVA = [0, 4, 5, 10, 12, 21]
EXTENSIONES_ORIGINAL = (".pdf", ".jpg", ".jpeg", ".png")

# Mismo criterio que en extraer_todas.py: algunos clientes ya tenian su
# carpeta de facturas emitidas organizada antes de este pipeline, asi
# que el enlace del original tiene que apuntar ahi en vez de al
# apartados/ingressos/ por defecto.
RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS = {"davinstal": "Emeses/davinstal"}

SUBCARPETAS_RESERVADAS = {"extraidas", "validadas", "procesadas", "lotes_escaneados", "lotes_procesados"}

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
    """Busca el original directamente en carpeta_origen, y si no esta
    ahi, en cualquier subcarpeta hermana no reservada del pipeline --
    algunos clientes (davinstal) organizan sus facturas de compra por
    proveedor (rebudes/biosca/, rebudes/SALTOKI/...) en vez de dejarlas
    sueltas en rebudes/entrada/."""
    base = os.path.splitext(nombre_json)[0]
    for ext in EXTENSIONES_ORIGINAL:
        ruta = os.path.join(carpeta_origen, base + ext)
        if os.path.exists(ruta):
            return ruta
    if os.path.isdir(carpeta_origen):
        for nombre_sub in sorted(os.listdir(carpeta_origen)):
            ruta_sub = os.path.join(carpeta_origen, nombre_sub)
            if not os.path.isdir(ruta_sub) or nombre_sub.lower() in SUBCARPETAS_RESERVADAS:
                continue
            for ext in EXTENSIONES_ORIGINAL:
                ruta = os.path.join(ruta_sub, base + ext)
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


def cargar_manifiestos(carpeta_lotes_procesados):
    """Lee todos los *_manifiesto.json de lotes_procesados/. Devuelve
    lista de (nombre_lote_pdf, documentos). Los lotes procesados antes
    del piso 6A no tienen manifiesto -- quedan fuera, asumido."""
    manifiestos = []
    if not os.path.isdir(carpeta_lotes_procesados):
        return manifiestos
    for nombre in sorted(os.listdir(carpeta_lotes_procesados)):
        if not nombre.endswith("_manifiesto.json"):
            continue
        nombre_lote_pdf = nombre[: -len("_manifiesto.json")] + ".pdf"
        with open(os.path.join(carpeta_lotes_procesados, nombre)) as f:
            documentos = json.load(f)
        manifiestos.append((nombre_lote_pdf, documentos))
    return manifiestos


def avisos_consistencia(facturas):
    """Agrupa por nombre de proveedor (no por nif_proveedor: un NIF mal
    leido por OCR varia de una factura a otra del mismo proveedor real,
    mientras que el nombre suele salir igual -- ver el caso del Celler).
    Devuelve una entrada por proveedor que use mas de un tipo de IVA
    distinto entre sus facturas -- OK y REVISAR juntas, es solo
    informativo, no cambia ningun estado."""
    por_proveedor = {}
    for nombre, datos in facturas:
        proveedor = datos.get("proveedor")
        if proveedor is None:
            continue
        entrada = por_proveedor.setdefault(proveedor, {"nifs": set(), "tipos": set(), "archivos": []})
        nif = datos.get("nif_proveedor")
        if nif is not None:
            entrada["nifs"].add(nif)
        for linea in datos.get("lineas_iva") or []:
            tipo = linea.get("tipo_iva")
            if tipo is not None:
                entrada["tipos"].add(tipo)
        entrada["archivos"].append(nombre)

    avisos = []
    for proveedor, info in sorted(por_proveedor.items()):
        if len(info["tipos"]) > 1:
            avisos.append((proveedor, sorted(info["nifs"]), sorted(info["tipos"]), info["archivos"]))
    return avisos


def escribir_avisos(ws, fila, carpeta_cliente, gastos, ingresos):
    # a) DOCUMENTOS APARTADOS -- albaranes que trocear.py aparta y nadie mas procesa
    for col in range(1, 3):
        celda = ws.cell(row=fila, column=col)
        celda.font = ESTILO_ENCABEZADO
        celda.fill = RELLENO_BLOQUE
    ws.cell(row=fila, column=1, value="DOCUMENTOS APARTADOS")
    fila += 1
    for col, texto in enumerate(["Archivo", "Justificación"], start=1):
        ws.cell(row=fila, column=col, value=texto).font = ESTILO_ENCABEZADO
    fila += 1

    carpeta_albarans = f"{carpeta_cliente}/apartados/albarans"
    nombres_albaran = []
    if os.path.isdir(carpeta_albarans):
        nombres_albaran = sorted(
            n for n in os.listdir(carpeta_albarans) if n.lower().endswith(EXTENSIONES_ORIGINAL)
        )

    if not nombres_albaran:
        ws.cell(row=fila, column=1, value="sin documentos apartados")
        fila += 1
    else:
        for nombre in nombres_albaran:
            celda = ws.cell(row=fila, column=1, value=nombre)
            ruta_relativa = os.path.relpath(os.path.join(carpeta_albarans, nombre), carpeta_cliente)
            celda.hyperlink = ruta_relativa
            celda.font = ESTILO_ENLACE
            celda_justificacion = ws.cell(
                row=fila,
                column=2,
                value="Albarà: no se contabiliza — la factura posterior agrupa sus albarans. Verificar que esa factura llegó",
            )
            celda_justificacion.alignment = AJUSTE_TEXTO
            fila += 1
    fila += 1

    # b) PAGINAS DESCARTADAS COMO RUIDO -- leidas de los manifiestos de trocear.py
    for col in range(1, 4):
        celda = ws.cell(row=fila, column=col)
        celda.font = ESTILO_ENCABEZADO
        celda.fill = RELLENO_BLOQUE
    ws.cell(row=fila, column=1, value="PÁGINAS DESCARTADAS COMO RUIDO")
    fila += 1
    for col, texto in enumerate(["Lote", "Páginas", "Pista"], start=1):
        ws.cell(row=fila, column=col, value=texto).font = ESTILO_ENCABEZADO
    fila += 1

    manifiestos = cargar_manifiestos(f"{carpeta_cliente}/rebudes/lotes_procesados")
    filas_ruido = [
        (nombre_lote, doc)
        for nombre_lote, documentos in manifiestos
        for doc in documentos
        if doc.get("tipo") == "ruido"
    ]

    if not manifiestos:
        ws.cell(row=fila, column=1, value="sin datos de lotes anteriores")
        fila += 1
    elif not filas_ruido:
        ws.cell(row=fila, column=1, value="sin páginas de ruido")
        fila += 1
    else:
        for nombre_lote, doc in filas_ruido:
            celda = ws.cell(row=fila, column=1, value=nombre_lote)
            ruta_lote = os.path.join(carpeta_cliente, "rebudes/lotes_procesados", nombre_lote)
            if os.path.exists(ruta_lote):
                celda.hyperlink = os.path.relpath(ruta_lote, carpeta_cliente)
                celda.font = ESTILO_ENLACE
            ws.cell(row=fila, column=2, value=f"p{doc['pagina_inicio']}-{doc['pagina_fin']}")
            ws.cell(row=fila, column=3, value=doc.get("emisor_pista"))
            fila += 1
    fila += 1

    # c) AVISOS DE CONSISTENCIA -- mismo proveedor, mas de un tipo de IVA
    for col in range(1, 6):
        celda = ws.cell(row=fila, column=col)
        celda.font = ESTILO_ENCABEZADO
        celda.fill = RELLENO_BLOQUE
    ws.cell(row=fila, column=1, value="AVISOS DE CONSISTENCIA")
    fila += 1
    for col, texto in enumerate(["Flujo", "Proveedor", "NIF", "Tipos de IVA", "Archivos afectados"], start=1):
        ws.cell(row=fila, column=col, value=texto).font = ESTILO_ENCABEZADO
    fila += 1

    filas_consistencia = [("GASTO", a) for a in avisos_consistencia(gastos)] + [
        ("INGRESO", a) for a in avisos_consistencia(ingresos)
    ]

    if not filas_consistencia:
        ws.cell(row=fila, column=1, value="sin avisos de consistencia")
        fila += 1
    else:
        for flujo, (proveedor, nifs, tipos, archivos) in filas_consistencia:
            ws.cell(row=fila, column=1, value=flujo)
            ws.cell(row=fila, column=2, value=proveedor)
            ws.cell(row=fila, column=3, value=" / ".join(nifs))
            ws.cell(row=fila, column=4, value="{" + ", ".join(str(t) for t in tipos) + "}")
            celda_archivos = ws.cell(row=fila, column=5, value="; ".join(archivos))
            celda_archivos.alignment = AJUSTE_TEXTO
            fila += 1

    return len(nombres_albaran), len(filas_ruido), len(filas_consistencia), bool(manifiestos)


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

        origen_ingressos = RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS.get(carpeta, "apartados/ingressos")
        for nombre, datos in revisar_g:
            pendientes.append((nombre, datos, "GASTO", f"{carpeta_cliente}/rebudes"))
        for nombre, datos in revisar_i:
            pendientes.append((nombre, datos, "INGRESO", f"{carpeta_cliente}/{origen_ingressos}"))

        if pendientes:
            escribir_pendientes(ws, fila, pendientes, carpeta_cliente)

    ws_avisos = wb.create_sheet("AVISOS")
    ws_avisos.column_dimensions["A"].width = 55
    ws_avisos.column_dimensions["B"].width = 70
    ws_avisos.column_dimensions["C"].width = 20
    ws_avisos.column_dimensions["D"].width = 20
    ws_avisos.column_dimensions["E"].width = 70

    fila = escribir_titulo(ws_avisos, fila_cliente["nombre"], fila_cliente["nif"])
    n_albaranes, n_ruido, n_consistencia, hay_manifiestos = escribir_avisos(
        ws_avisos, fila, carpeta_cliente, gastos, ingresos
    )
    detalle_ruido = f"{n_ruido} páginas de ruido" if hay_manifiestos else "sin datos de lotes anteriores"
    print(f"{carpeta} / AVISOS: {n_albaranes} documentos apartados, {detalle_ruido}, {n_consistencia} avisos de consistencia")

    ruta_excel = f"{carpeta_cliente}/sumatorios_2026.xlsx"
    wb.save(ruta_excel)
    print(f"Escrito: {ruta_excel}")
