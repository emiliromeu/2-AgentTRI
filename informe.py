"""Genera un informe HTML de auditoria por cliente: hoja de
comparacion para rellenar a mano, y una tarjeta completa por CADA
factura (OK y REVISAR), plegada por trimestre y luego por flujo.

Piso 7: sin llamadas a la API, corre gratis. Lee las mismas validadas
que sumar.py -- funciones duplicadas a proposito (mismo criterio del
proyecto desde el piso 2: nada de modulo compartido entre scripts).
Se regenera entero en cada run, igual que sumar.py.

Piso 7B: modo auditoria total -- todas las facturas tienen tarjeta
(antes solo las REVISAR), organizadas en <details>/<summary> plegables
por trimestre y flujo (sin JavaScript), con una hoja de comparacion
arriba para que el departamento anote el calculo manual al lado.
"""

import base64
import csv
import html
import json
import os
from datetime import datetime
from urllib.parse import quote

EXTENSIONES_ORIGINAL = (".pdf", ".jpg", ".jpeg", ".png")
EXTENSIONES_IMAGEN = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png"}

RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS = {"davinstal": "Emeses/davinstal"}
SUBCARPETAS_RESERVADAS = {"extraidas", "validadas", "procesadas", "lotes_escaneados", "lotes_procesados"}

GENERADO_EL = datetime.now().strftime("%d/%m/%Y %H:%M")


def leer_clientes():
    with open("clientes/clientes.csv") as f:
        return list(csv.DictReader(f))


def cargar_validadas(carpeta):
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
    if not fecha:
        return None
    try:
        mes = int(fecha[5:7])
    except (ValueError, IndexError):
        return None
    return f"{(mes - 1) // 3 + 1}T"


def encontrar_original(carpeta_origen, nombre_json):
    """Igual que en sumar.py: busca directamente, y si no, en
    subcarpetas hermanas no reservadas (proveedores organizados en
    subcarpetas, ej. davinstal)."""
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


def sumar_bloque(facturas):
    """Igual que en sumar.py: total de las OK, y aparte la lista de
    las que hay que revisar (las REVISAR no entran en el total)."""
    total_ok = 0.0
    revisar = []
    for nombre, datos in facturas:
        if datos.get("estado") != "OK":
            revisar.append((nombre, datos))
            continue
        total_ok += datos.get("total") or 0
    return total_ok, revisar


def ruta_relativa_html(ruta_original, carpeta_cliente):
    """Ruta relativa a la carpeta del cliente, con espacios/comas
    escapados para que el navegador la resuelva bien."""
    relativa = os.path.relpath(ruta_original, carpeta_cliente)
    return quote(relativa)


def esc(valor):
    """Escapa texto para HTML; None se convierte en cadena vacia."""
    return html.escape(str(valor)) if valor is not None else ""


def tarjeta_factura(nombre, datos, tipo_bloque, carpeta_original, carpeta_cliente):
    """Ficha completa de UNA factura, OK o REVISAR. Izquierda: todos
    los campos extraidos. Derecha: imagen incrustada (jpg/png) o
    enlace grande (pdf)."""
    ruta_original = encontrar_original(carpeta_original, nombre)
    extension = os.path.splitext(ruta_original)[1].lower() if ruta_original else None

    if ruta_original and extension in EXTENSIONES_IMAGEN:
        with open(ruta_original, "rb") as f:
            b64 = base64.standard_b64encode(f.read()).decode("utf-8")
        lado_derecho = f'<img loading="lazy" src="data:{EXTENSIONES_IMAGEN[extension]};base64,{b64}" alt="original">'
    elif ruta_original:
        href = ruta_relativa_html(ruta_original, carpeta_cliente)
        lado_derecho = f'<a class="btn-abrir" href="{href}" target="_blank">Abrir original ↗</a>'
    else:
        print(f"AVISO: no se encontró el original de {nombre}")
        lado_derecho = '<p class="sin-original">Original no encontrado</p>'

    lineas_html = "".join(
        f"<li>Base {esc(l.get('base'))} € × {esc(l.get('tipo_iva'))}% = {esc(l.get('cuota'))} €</li>"
        for l in (datos.get("lineas_iva") or [])
    )

    retencion_cuota = datos.get("retencion_cuota") or 0
    retencion_html = (
        f"<p>Retención: {esc(datos.get('retencion_pct'))}% = {esc(retencion_cuota)} €</p>"
        if retencion_cuota else ""
    )

    estado = datos.get("estado")
    motivos = datos.get("motivos") or []
    motivos_html = (
        "<p><strong>Motivos:</strong></p><ul>" + "".join(f"<li>{esc(m)}</li>" for m in motivos) + "</ul>"
        if motivos else ""
    )

    clase_estado = "revisar" if estado == "REVISAR" else "ok"

    return f"""
    <div class="tarjeta {clase_estado}">
      <div class="tarjeta-izq">
        <span class="etiqueta-tipo etiqueta-{tipo_bloque.lower()}">{tipo_bloque}</span>
        <span class="etiqueta-estado">{esc(estado)}</span>
        <h3>{esc(datos.get("proveedor"))}</h3>
        <p>NIF: {esc(datos.get("nif_proveedor"))} · Factura: {esc(datos.get("num_factura"))} · Fecha: {esc(datos.get("fecha_factura"))}</p>
        <ul class="lineas-iva">{lineas_html}</ul>
        <p>Total: {esc(datos.get("total"))} €</p>
        {retencion_html}
        <p class="archivo">{esc(nombre)}</p>
        {motivos_html}
      </div>
      <div class="tarjeta-der">{lado_derecho}</div>
    </div>"""


TIPO_BLOQUE_SINGULAR = {"GASTOS": "GASTO", "INGRESOS": "INGRESO"}


def seccion_flujo(titulo, facturas, carpeta_original, carpeta_cliente):
    total_ok, revisar = sumar_bloque(facturas)
    tipo_bloque = TIPO_BLOQUE_SINGULAR[titulo]
    tarjetas = "".join(
        tarjeta_factura(nombre, datos, tipo_bloque, carpeta_original, carpeta_cliente)
        for nombre, datos in facturas
    )
    resumen = f"{titulo} — {len(facturas)} facturas, {len(revisar)} a revisar — {total_ok:.2f} €"
    return f"""
    <details>
      <summary>{esc(resumen)}</summary>
      {tarjetas if tarjetas else "<p>Sin facturas.</p>"}
    </details>""", total_ok


def seccion_trimestre(trimestre, datos_t, carpeta_cliente, origen_gastos, origen_ingressos):
    html_gastos, total_g = seccion_flujo("GASTOS", datos_t["gastos"], origen_gastos, carpeta_cliente)
    html_ingresos, total_i = seccion_flujo("INGRESOS", datos_t["ingresos"], origen_ingressos, carpeta_cliente)
    n_total = len(datos_t["gastos"]) + len(datos_t["ingresos"])
    resumen = f"{trimestre} — {n_total} facturas ({total_g:.2f} € gastos, {total_i:.2f} € ingresos)"
    return f"""
    <details>
      <summary class="resumen-trimestre">{esc(resumen)}</summary>
      {html_gastos}
      {html_ingresos}
    </details>""", total_g, total_i


def tabla_comparacion(filas):
    """filas: lista de (trimestre, flujo, total_agente)."""
    filas_html = "".join(
        f"""<tr><td>{esc(t)}</td><td>{esc(flujo)}</td><td class="num">{total:.2f}</td>
            <td class="rellenar"></td><td class="rellenar"></td></tr>"""
        for t, flujo, total in filas
    )
    return f"""
    <table class="comparacion">
      <thead><tr><th>Trimestre</th><th>Flujo</th><th>Agente (€)</th>
      <th>Cálculo manual (€)</th><th>Diferencia</th></tr></thead>
      <tbody>{filas_html}</tbody>
    </table>"""


ESTILO = """
body { font-family: -apple-system, "Segoe UI", Roboto, sans-serif; max-width: 1100px;
       margin: 2rem auto; padding: 0 1rem; color: #1a1a1a; line-height: 1.5; }
h1 { margin-bottom: 0.2rem; }
.subtitulo { color: #555; margin-top: 0; }
h2 { border-bottom: 2px solid #D9E1F2; padding-bottom: 0.3rem; margin-top: 2.5rem; }

table.comparacion { border-collapse: collapse; width: 100%; margin: 1rem 0 2rem 0; }
table.comparacion th, table.comparacion td { border: 1px solid #bbb; padding: 0.6rem 0.8rem; }
table.comparacion th { background: #D9E1F2; text-align: left; }
table.comparacion td.num { text-align: right; }
table.comparacion td.rellenar { min-width: 120px; }

details { margin-bottom: 0.5rem; }
details details { margin: 0.6rem 0 0.6rem 1.2rem; }
summary { cursor: pointer; padding: 0.6rem 1rem; background: #D9E1F2; border-radius: 6px;
          font-weight: bold; }
summary.resumen-trimestre { background: #b9c9e6; font-size: 1.05rem; }

.tarjeta { display: flex; gap: 1.5rem; border-radius: 8px; padding: 1rem 1.5rem;
           margin: 0.8rem 0 0.8rem 1rem; border: 1px solid #ddd; }
.tarjeta.revisar { background: #FCE4D6; border-color: #e8b48f; }
.tarjeta-izq { flex: 1; }
.tarjeta-der { flex: 1; display: flex; align-items: center; justify-content: center; }
.tarjeta-der img { max-width: 100%; max-height: 500px; border-radius: 4px; }
.etiqueta-tipo, .etiqueta-estado { display: inline-block; padding: 0.15rem 0.6rem; border-radius: 4px;
                 font-size: 0.8rem; font-weight: bold; background: #fff; margin-right: 0.3rem; }
.lineas-iva { margin: 0.3rem 0; padding-left: 1.2rem; }
.archivo { font-family: monospace; font-size: 0.8rem; color: #555; }
.btn-abrir { display: inline-block; background: #0563C1; color: white; padding: 0.8rem 1.5rem;
             border-radius: 6px; text-decoration: none; font-weight: bold; }
.sin-original { color: #a33; font-style: italic; }

@media (prefers-color-scheme: dark) {
  body { background: #1a1a1a; color: #eee; }
  summary, table.comparacion th { background: #2a3a52; }
  summary.resumen-trimestre { background: #35496b; }
  .tarjeta { border-color: #444; }
  .etiqueta-tipo, .etiqueta-estado { background: #333; }
  .archivo { color: #aaa; }
}

@media print {
  details > * { display: block !important; }
  summary::-webkit-details-marker { display: none; }
  .tarjeta { page-break-inside: avoid; }
}
"""


for fila_cliente in leer_clientes():
    carpeta = fila_cliente["carpeta"]
    carpeta_cliente = f"clientes/{carpeta}"

    gastos = cargar_validadas(f"{carpeta_cliente}/rebudes/validadas")
    ingresos = cargar_validadas(f"{carpeta_cliente}/apartados/ingressos_validadas")

    if not gastos and not ingresos:
        continue

    origen_gastos = f"{carpeta_cliente}/rebudes"
    origen_ingressos = f"{carpeta_cliente}/{RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS.get(carpeta, 'apartados/ingressos')}"

    trimestres = {}
    for nombre, datos in gastos:
        t = trimestre_de(datos.get("fecha_factura")) or "SIN FECHA"
        trimestres.setdefault(t, {"gastos": [], "ingresos": []})["gastos"].append((nombre, datos))
    for nombre, datos in ingresos:
        t = trimestre_de(datos.get("fecha_factura")) or "SIN FECHA"
        trimestres.setdefault(t, {"gastos": [], "ingresos": []})["ingresos"].append((nombre, datos))

    orden = sorted(t for t in trimestres if t != "SIN FECHA") + (
        ["SIN FECHA"] if "SIN FECHA" in trimestres else []
    )

    filas_comparacion = []
    secciones_html = []
    n_tarjetas = 0

    for trimestre in orden:
        datos_t = trimestres[trimestre]
        n_tarjetas += len(datos_t["gastos"]) + len(datos_t["ingresos"])

        if trimestre == "SIN FECHA":
            # Sin fecha no hay trimestre que comparar -- se listan igual, aparte
            html_gastos, total_g = seccion_flujo("GASTOS", datos_t["gastos"], origen_gastos, carpeta_cliente)
            html_ingresos, total_i = seccion_flujo("INGRESOS", datos_t["ingresos"], origen_ingressos, carpeta_cliente)
            secciones_html.append(f"""
            <details>
              <summary class="resumen-trimestre">SIN FECHA — {len(datos_t["gastos"]) + len(datos_t["ingresos"])} facturas</summary>
              {html_gastos}
              {html_ingresos}
            </details>""")
            continue

        html_trimestre, total_g, total_i = seccion_trimestre(
            trimestre, datos_t, carpeta_cliente, origen_gastos, origen_ingressos
        )
        secciones_html.append(html_trimestre)
        filas_comparacion.append((trimestre, "GASTOS", total_g))
        filas_comparacion.append((trimestre, "INGRESOS", total_i))

    html_final = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>Informe {esc(fila_cliente["nombre"])}</title>
<style>{ESTILO}</style>
</head>
<body>
  <h1>{esc(fila_cliente["nombre"])}</h1>
  <p class="subtitulo">NIF {esc(fila_cliente["nif"])} · Generado el {GENERADO_EL}</p>

  <h2>Hoja de comparación</h2>
  {tabla_comparacion(filas_comparacion)}

  <h2>Facturas por trimestre</h2>
  {"".join(secciones_html)}
</body>
</html>"""

    ruta_html = f"{carpeta_cliente}/informe_2026.html"
    with open(ruta_html, "w") as f:
        f.write(html_final)

    print(f"{carpeta} / informe: {n_tarjetas} tarjetas")
    print(f"Escrito: {ruta_html}")
