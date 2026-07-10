"""Genera un informe HTML de auditoria per client: panell de
conciliacio, full de comparacio per omplir a ma, i una targeta completa
per CADA factura (OK i REVISAR), plegada per trimestre i despres per
flux.

Piso 7: sin llamadas a la API, corre gratis. Lee las mismas validadas
que sumar.py -- funciones duplicadas a proposito (mismo criterio del
proyecto desde el piso 2: nada de modulo compartido entre scripts).
Se regenera entero en cada run, igual que sumar.py.

Piso 7B: modo auditoria total -- todas las facturas tienen tarjeta
(antes solo las REVISAR), organizadas en <details>/<summary> plegables
por trimestre y flujo (sin JavaScript), con una hoja de comparacion
arriba para que el departamento anote el calculo manual al lado.

Piso 8: tot el text visible passa a catala (els motius de validar.py
es tradueixen aqui). RESULTAT IVA per trimestre (calcul derivat).
Verd suau per OK, taronja per REVISAR. Observacions de lectura a les
targetes. Panell CONCILIACIO (nomes lectura de disc). Seccio Avisos
(la mateixa info que la fulla AVISOS de l'Excel, amb enllacos). I
clientes/index.html com a portada. Cap total canvia.
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

# Igual que en sumar.py: traduccion de los motivos de validar.py por
# sustitucion de las palabras fijas (validar.py no se toca).
TRADUCCIONES_MOTIVO = [
    ("campo obligatorio vacío:", "camp obligatori buit:"),
    ("línea", "línia"),
    ("de IVA con campo vacío", "d'IVA amb camp buit"),
    ("pero cuota indica", "però la quota indica"),
    ("total no cuadra: bases+cuotas=", "el total no quadra: bases+quotes="),
    (", total indica ", ", el total indica "),
    ("nif_receptor no coincide: esperado", "el nif_receptor no coincideix: s'esperava"),
    (", encontrado ", ", s'ha trobat "),
    ("factura duplicada: mismo proveedor+num_factura que", "factura duplicada: mateix proveïdor+núm_factura que"),
    ("retención con cuota > 0: el llibre no tiene columna para representarla",
     "retenció amb quota > 0: el llibre no té columna per representar-la"),
    ("el cliente no aparece ni como emisor ni como receptor",
     "el client no apareix ni com a emissor ni com a receptor"),
]


def traducir_motivo(motivo):
    for es, ca in TRADUCCIONES_MOTIVO:
        motivo = motivo.replace(es, ca)
    return motivo


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


def listar_archivos_rebudes(carpeta_rebudes):
    """Igual que en extraer_todas.py: cuenta entrada/ y cualquier
    subcarpeta hermana no reservada (proveedores organizados por
    subcarpeta, ej. davinstal). Solo lectura -- para el panel CONCILIACIO."""
    rutas = []
    if not os.path.isdir(carpeta_rebudes):
        return rutas
    for nombre in sorted(os.listdir(carpeta_rebudes)):
        ruta = os.path.join(carpeta_rebudes, nombre)
        if not os.path.isdir(ruta) or nombre.lower() in SUBCARPETAS_RESERVADAS:
            continue
        for nombre_archivo in sorted(os.listdir(ruta)):
            if nombre_archivo.lower().endswith(EXTENSIONES_ORIGINAL):
                rutas.append(os.path.join(ruta, nombre_archivo))
    return rutas


def contar_json(carpeta):
    if not os.path.isdir(carpeta):
        return 0
    return len([f for f in os.listdir(carpeta) if f.lower().endswith(".json")])


def cargar_manifiestos(carpeta_lotes_procesados):
    """Igual que en sumar.py."""
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
    """Igual que en sumar.py: agrupa por nombre de proveedor (no por
    nif_proveedor, que puede variar por OCR), un aviso por proveedor
    con mas de un tipo de IVA distinto entre sus facturas."""
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


def sumar_bloque(facturas):
    """Igual que en sumar.py: total de las OK, y aparte la lista de
    las que hay que revisar (las REVISAR no entran en el total)."""
    total_ok = 0.0
    cuota_ok = 0.0
    revisar = []
    for nombre, datos in facturas:
        if datos.get("estado") != "OK":
            revisar.append((nombre, datos))
            continue
        total_ok += datos.get("total") or 0
        for linea in datos.get("lineas_iva") or []:
            cuota_ok += linea.get("cuota") or 0
    return total_ok, cuota_ok, revisar


def ruta_relativa_html(ruta_original, carpeta_cliente):
    """Ruta relativa a la carpeta del cliente, con espacios/comas
    escapados para que el navegador la resuelva bien."""
    relativa = os.path.relpath(ruta_original, carpeta_cliente)
    return quote(relativa)


def esc(valor):
    """Escapa texto para HTML; None se convierte en cadena vacia."""
    return html.escape(str(valor)) if valor is not None else ""


def tarjeta_factura(nombre, datos, tipo_bloque, carpeta_original, carpeta_cliente):
    """Fitxa completa d'UNA factura, OK o REVISAR. Esquerra: tots els
    camps extrets. Dreta: imatge incrustada (jpg/png) o enllaç gran (pdf)."""
    ruta_original = encontrar_original(carpeta_original, nombre)
    extension = os.path.splitext(ruta_original)[1].lower() if ruta_original else None

    if ruta_original and extension in EXTENSIONES_IMAGEN:
        with open(ruta_original, "rb") as f:
            b64 = base64.standard_b64encode(f.read()).decode("utf-8")
        lado_derecho = f'<img loading="lazy" src="data:{EXTENSIONES_IMAGEN[extension]};base64,{b64}" alt="original">'
    elif ruta_original:
        href = ruta_relativa_html(ruta_original, carpeta_cliente)
        lado_derecho = f'<a class="btn-abrir" href="{href}" target="_blank">Obrir original ↗</a>'
    else:
        print(f"AVISO: no se encontró el original de {nombre}")
        lado_derecho = '<p class="sin-original">Original no trobat</p>'

    lineas_html = "".join(
        f"<li>Base {esc(l.get('base'))} € × {esc(l.get('tipo_iva'))}% = {esc(l.get('cuota'))} €</li>"
        for l in (datos.get("lineas_iva") or [])
    )

    retencion_cuota = datos.get("retencion_cuota") or 0
    retencion_html = (
        f"<p>Retenció: {esc(datos.get('retencion_pct'))}% = {esc(retencion_cuota)} €</p>"
        if retencion_cuota else ""
    )

    estado = datos.get("estado")
    motivos = [traducir_motivo(m) for m in (datos.get("motivos") or [])]
    motivos_html = (
        "<p><strong>Motius:</strong></p><ul>" + "".join(f"<li>{esc(m)}</li>" for m in motivos) + "</ul>"
        if motivos else ""
    )

    observaciones = datos.get("observaciones")
    observaciones_html = (
        f"<p><strong>Observacions de lectura:</strong> {esc(observaciones)}</p>" if observaciones else ""
    )

    clase_estado = "revisar" if estado == "REVISAR" else "ok"

    return f"""
    <div class="tarjeta {clase_estado}">
      <div class="tarjeta-izq">
        <span class="etiqueta-tipo etiqueta-{tipo_bloque.lower()}">{tipo_bloque}</span>
        <span class="etiqueta-estado">{esc(estado)}</span>
        <h3>{esc(datos.get("proveedor"))}</h3>
        <p>NIF: {esc(datos.get("nif_proveedor"))} · Factura: {esc(datos.get("num_factura"))} · Data: {esc(datos.get("fecha_factura"))}</p>
        <ul class="lineas-iva">{lineas_html}</ul>
        <p>Total: {esc(datos.get("total"))} €</p>
        {retencion_html}
        <p class="archivo">{esc(nombre)}</p>
        {motivos_html}
        {observaciones_html}
      </div>
      <div class="tarjeta-der">{lado_derecho}</div>
    </div>"""


TIPO_BLOQUE_SINGULAR = {"DESPESES": "DESPESA", "INGRESSOS": "INGRÉS"}


def seccion_flujo(titulo, facturas, carpeta_original, carpeta_cliente):
    total_ok, cuota_ok, revisar = sumar_bloque(facturas)
    tipo_bloque = TIPO_BLOQUE_SINGULAR[titulo]
    tarjetas = "".join(
        tarjeta_factura(nombre, datos, tipo_bloque, carpeta_original, carpeta_cliente)
        for nombre, datos in facturas
    )
    resumen = f"{titulo} — {len(facturas)} factures, {len(revisar)} a revisar — {total_ok:.2f} €"
    return f"""
    <details>
      <summary>{esc(resumen)}</summary>
      {tarjetas if tarjetas else "<p>Sense factures.</p>"}
    </details>""", total_ok, cuota_ok


def seccion_trimestre(trimestre, datos_t, carpeta_cliente, origen_gastos, origen_ingressos):
    html_gastos, total_g, cuota_g = seccion_flujo("DESPESES", datos_t["gastos"], origen_gastos, carpeta_cliente)
    html_ingresos, total_i, cuota_i = seccion_flujo("INGRESSOS", datos_t["ingresos"], origen_ingressos, carpeta_cliente)
    n_total = len(datos_t["gastos"]) + len(datos_t["ingresos"])
    resumen = f"{trimestre} — {n_total} factures ({total_g:.2f} € despeses, {total_i:.2f} € ingressos)"

    resultat = cuota_i - cuota_g
    if resultat > 0:
        etiqueta = " (a ingressar)"
    elif resultat < 0:
        etiqueta = " (a compensar)"
    else:
        etiqueta = ""
    resultat_html = (
        f'<p class="resultat-iva"><strong>RESULTAT IVA (repercutit − suportat): {resultat:.2f} €{etiqueta}</strong><br>'
        f'<span class="nota">resultat de treball, no liquidació oficial</span></p>'
    )

    return f"""
    <details>
      <summary class="resumen-trimestre">{esc(resumen)}</summary>
      {resultat_html}
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
      <thead><tr><th>Trimestre</th><th>Flux</th><th>Agent (€)</th>
      <th>Càlcul manual (€)</th><th>Diferència</th></tr></thead>
      <tbody>{filas_html}</tbody>
    </table>"""


def panel_conciliacio_flux(nombre_flux, presentes, extraidas, ok, pendents):
    """Nomes lectura de disc -- cap re-execucio. 'Errors' es infereix
    com presents-ok-pendents: valid nomes si extraer_todas.py ja ha
    corregut a fons abans de generar aquest informe (aixi es com
    ejecutar.py ho invoca -- informe.py es sempre l'ultima maquina)."""
    errors = max(0, presentes - ok - pendents)
    quadra = presentes == ok + pendents + errors
    return f"""
    <div class="conciliacio-flux">
      <h4>{esc(nombre_flux)}</h4>
      <table class="conciliacio">
        <tr><td>Presents</td><td class="num">{presentes}</td></tr>
        <tr><td>Fitxes extretes</td><td class="num">{extraidas}</td></tr>
        <tr><td>OK</td><td class="num">{ok}</td></tr>
        <tr><td>Pendents</td><td class="num">{pendents}</td></tr>
        <tr><td>Errors</td><td class="num">{errors}</td></tr>
      </table>
      <p class="quadre">{presentes} presents = {ok} OK + {pendents} pendents + {errors} errors {"✓" if quadra else "⚠"}</p>
    </div>"""


def seccion_avisos(carpeta_cliente, gastos, ingresos, origen_ingressos):
    # a) Documents apartats
    carpeta_albarans = f"{carpeta_cliente}/apartados/albarans"
    nombres_albaran = []
    if os.path.isdir(carpeta_albarans):
        nombres_albaran = sorted(
            n for n in os.listdir(carpeta_albarans) if n.lower().endswith(EXTENSIONES_ORIGINAL)
        )
    if not nombres_albaran:
        html_apartats = "<p>Sense documents apartats.</p>"
    else:
        filas = "".join(
            f"""<tr><td><a href="{ruta_relativa_html(os.path.join(carpeta_albarans, n), carpeta_cliente)}"
                target="_blank">{esc(n)}</a></td>
                <td>Albarà: no es comptabilitza — la factura posterior n'agrupa els albarans.
                Verificar que aquesta factura ha arribat</td></tr>"""
            for n in nombres_albaran
        )
        html_apartats = f'<table><thead><tr><th>Fitxer</th><th>Justificació</th></tr></thead><tbody>{filas}</tbody></table>'

    # b) Pagines descartades com a soroll
    manifiestos = cargar_manifiestos(f"{carpeta_cliente}/rebudes/lotes_procesados")
    filas_ruido = [
        (nombre_lote, doc)
        for nombre_lote, documentos in manifiestos
        for doc in documentos
        if doc.get("tipo") == "ruido"
    ]
    if not manifiestos:
        html_ruido = "<p>Sense dades de lots anteriors.</p>"
    elif not filas_ruido:
        html_ruido = "<p>Sense pàgines de soroll.</p>"
    else:
        filas = ""
        for nombre_lote, doc in filas_ruido:
            ruta_lote = os.path.join(carpeta_cliente, "rebudes/lotes_procesados", nombre_lote)
            if os.path.exists(ruta_lote):
                lote_html = f'<a href="{ruta_relativa_html(ruta_lote, carpeta_cliente)}" target="_blank">{esc(nombre_lote)}</a>'
            else:
                lote_html = esc(nombre_lote)
            filas += f"""<tr><td>{lote_html}</td><td>p{doc['pagina_inicio']}-{doc['pagina_fin']}</td>
                <td>{esc(doc.get('emisor_pista'))}</td></tr>"""
        html_ruido = f'<table><thead><tr><th>Lot</th><th>Pàgines</th><th>Pista</th></tr></thead><tbody>{filas}</tbody></table>'

    # c) Avisos de consistencia
    filas_consistencia = [("DESPESA", a) for a in avisos_consistencia(gastos)] + [
        ("INGRÉS", a) for a in avisos_consistencia(ingresos)
    ]
    if not filas_consistencia:
        html_consistencia = "<p>Sense avisos de consistència.</p>"
    else:
        filas = ""
        for flujo, (proveedor, nifs, tipos, archivos) in filas_consistencia:
            filas += f"""<tr><td>{esc(flujo)}</td><td>{esc(proveedor)}</td><td>{esc(" / ".join(nifs))}</td>
                <td>{{{esc(", ".join(str(t) for t in tipos))}}}</td><td>{esc("; ".join(archivos))}</td></tr>"""
        html_consistencia = (
            '<table><thead><tr><th>Flux</th><th>Proveïdor</th><th>NIF</th>'
            f'<th>Tipus d\'IVA</th><th>Fitxers afectats</th></tr></thead><tbody>{filas}</tbody></table>'
        )

    return f"""
    <h2>Avisos</h2>
    <h3>Documents apartats</h3>
    {html_apartats}
    <h3>Pàgines descartades com a soroll</h3>
    {html_ruido}
    <h3>Avisos de consistència</h3>
    {html_consistencia}
    """


ESTILO = """
body { font-family: -apple-system, "Segoe UI", Roboto, sans-serif; max-width: 1100px;
       margin: 2rem auto; padding: 0 1rem; color: #1a1a1a; line-height: 1.5; }
h1 { margin-bottom: 0.2rem; }
.subtitulo { color: #555; margin-top: 0; }
h2 { border-bottom: 2px solid #D9E1F2; padding-bottom: 0.3rem; margin-top: 2.5rem; }

.conciliacio { display: flex; flex-wrap: wrap; gap: 1.5rem; margin: 1rem 0; }
.conciliacio-flux { background: #f5f5f5; border-radius: 8px; padding: 1rem 1.5rem; min-width: 220px; }
.conciliacio-flux table.conciliacio { width: 100%; margin: 0; }
.conciliacio-flux table.conciliacio td { padding: 0.15rem 0.3rem; border: none; }
.conciliacio-flux .quadre { font-size: 0.85rem; color: #555; margin-top: 0.5rem; }
.aparte { color: #555; font-size: 0.9rem; }

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

p.resultat-iva { background: #eef2f8; border-radius: 6px; padding: 0.6rem 1rem; margin: 0.6rem 0; }
p.resultat-iva .nota { font-size: 0.8rem; font-style: italic; color: #555; }

.tarjeta { display: flex; gap: 1.5rem; border-radius: 8px; padding: 1rem 1.5rem;
           margin: 0.8rem 0 0.8rem 1rem; border: 1px solid #ddd; }
.tarjeta.revisar { background: #FCE4D6; border-color: #e8b48f; }
.tarjeta.ok { background: #E2EFDA; border-color: #b7d7a8; }
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

table:not(.comparacion):not(.conciliacio) { border-collapse: collapse; width: 100%; margin: 0.6rem 0 1.5rem 0; font-size: 0.9rem; }
table:not(.comparacion):not(.conciliacio) th, table:not(.comparacion):not(.conciliacio) td {
  border: 1px solid #ddd; padding: 0.4rem 0.6rem; text-align: left; }
table:not(.comparacion):not(.conciliacio) th { background: #D9E1F2; }

@media (prefers-color-scheme: dark) {
  body { background: #1a1a1a; color: #eee; }
  summary, table.comparacion th, table:not(.comparacion) th { background: #2a3a52; }
  summary.resumen-trimestre { background: #35496b; }
  .tarjeta { border-color: #444; }
  .tarjeta.ok { background: #22331f; border-color: #3a5230; }
  .tarjeta.revisar { background: #3a2a1f; border-color: #5c4128; }
  .etiqueta-tipo, .etiqueta-estado { background: #333; }
  .archivo { color: #aaa; }
  .conciliacio-flux { background: #262626; }
  p.resultat-iva { background: #223; }
}

@media print {
  details > * { display: block !important; }
  summary::-webkit-details-marker { display: none; }
  .tarjeta { page-break-inside: avoid; }
}
"""


clientes_generados = []

for fila_cliente in leer_clientes():
    carpeta = fila_cliente["carpeta"]
    carpeta_cliente = f"clientes/{carpeta}"

    gastos = cargar_validadas(f"{carpeta_cliente}/rebudes/validadas")
    ingresos = cargar_validadas(f"{carpeta_cliente}/apartados/ingressos_validadas")

    if not gastos and not ingresos:
        continue

    origen_gastos = f"{carpeta_cliente}/rebudes"
    origen_ingressos = f"{carpeta_cliente}/{RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS.get(carpeta, 'apartados/ingressos')}"

    # Panel CONCILIACIO -- solo lectura de disco
    presentes_gastos = len(listar_archivos_rebudes(origen_gastos))
    extraidas_gastos = contar_json(f"{carpeta_cliente}/rebudes/extraidas")
    ok_gastos = sum(1 for _, d in gastos if d.get("estado") == "OK")
    pendents_gastos = len(gastos) - ok_gastos

    presentes_ingressos = len([
        f for f in os.listdir(origen_ingressos) if f.lower().endswith(EXTENSIONES_ORIGINAL)
    ]) if os.path.isdir(origen_ingressos) else 0
    extraidas_ingressos = contar_json(f"{carpeta_cliente}/apartados/ingressos_extraidas")
    ok_ingressos = sum(1 for _, d in ingresos if d.get("estado") == "OK")
    pendents_ingressos = len(ingresos) - ok_ingressos

    carpeta_albarans = f"{carpeta_cliente}/apartados/albarans"
    n_albarans = len([
        f for f in os.listdir(carpeta_albarans) if f.lower().endswith(EXTENSIONES_ORIGINAL)
    ]) if os.path.isdir(carpeta_albarans) else 0

    manifiestos_conciliacio = cargar_manifiestos(f"{carpeta_cliente}/rebudes/lotes_procesados")
    n_ruido = sum(
        1 for _, documentos in manifiestos_conciliacio for doc in documentos if doc.get("tipo") == "ruido"
    )

    panel_html = f"""
    <h2>Conciliació</h2>
    <div class="conciliacio">
      {panel_conciliacio_flux("Despeses", presentes_gastos, extraidas_gastos, ok_gastos, pendents_gastos)}
      {panel_conciliacio_flux("Ingressos", presentes_ingressos, extraidas_ingressos, ok_ingressos, pendents_ingressos)}
    </div>
    <p class="aparte">Informat a part (no entra en el quadre de dalt): {n_albarans} albarans apartats,
    {n_ruido} pàgines de soroll{" (sense dades de lots anteriors)" if not manifiestos_conciliacio else ""}.</p>
    """

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
            html_gastos, total_g, _ = seccion_flujo("DESPESES", datos_t["gastos"], origen_gastos, carpeta_cliente)
            html_ingresos, total_i, _ = seccion_flujo("INGRESSOS", datos_t["ingresos"], origen_ingressos, carpeta_cliente)
            secciones_html.append(f"""
            <details>
              <summary class="resumen-trimestre">SENSE DATA — {len(datos_t["gastos"]) + len(datos_t["ingresos"])} factures</summary>
              {html_gastos}
              {html_ingresos}
            </details>""")
            continue

        html_trimestre, total_g, total_i = seccion_trimestre(
            trimestre, datos_t, carpeta_cliente, origen_gastos, origen_ingressos
        )
        secciones_html.append(html_trimestre)
        filas_comparacion.append((trimestre, "DESPESES", total_g))
        filas_comparacion.append((trimestre, "INGRESSOS", total_i))

    html_avisos = seccion_avisos(carpeta_cliente, gastos, ingresos, origen_ingressos)

    html_final = f"""<!DOCTYPE html>
<html lang="ca">
<head>
<meta charset="utf-8">
<title>Informe {esc(fila_cliente["nombre"])}</title>
<style>{ESTILO}</style>
</head>
<body>
  <h1>{esc(fila_cliente["nombre"])}</h1>
  <p class="subtitulo">NIF {esc(fila_cliente["nif"])} · Generat el {GENERADO_EL}</p>

  {panel_html}

  <h2>Full de comparació</h2>
  {tabla_comparacion(filas_comparacion)}

  <h2>Factures per trimestre</h2>
  {"".join(secciones_html)}

  {html_avisos}
</body>
</html>"""

    ruta_html = f"{carpeta_cliente}/informe_2026.html"
    with open(ruta_html, "w") as f:
        f.write(html_final)

    print(f"{carpeta} / informe: {n_tarjetas} tarjetas")
    print(f"Escrito: {ruta_html}")

    clientes_generados.append({
        "nombre": fila_cliente["nombre"],
        "nif": fila_cliente["nif"],
        "carpeta": carpeta,
        "informe_existe": os.path.exists(ruta_html),
        "excel_existe": os.path.exists(f"{carpeta_cliente}/sumatorios_2026.xlsx"),
    })


# clientes/index.html -- portada con enlace al informe y al Excel de cada cliente
filas_index = ""
for c in clientes_generados:
    enlace_informe = (
        f'<a href="{c["carpeta"]}/informe_2026.html">Informe</a>' if c["informe_existe"] else "no disponible"
    )
    enlace_excel = (
        f'<a href="{c["carpeta"]}/sumatorios_2026.xlsx">Excel</a>' if c["excel_existe"] else "no disponible"
    )
    filas_index += f"""<tr><td>{esc(c["nombre"])}</td><td>{esc(c["nif"])}</td>
        <td>{enlace_informe}</td><td>{enlace_excel}</td></tr>"""

index_html = f"""<!DOCTYPE html>
<html lang="ca">
<head>
<meta charset="utf-8">
<title>Clients — Agent TRIMESTRE</title>
<style>{ESTILO}</style>
</head>
<body>
  <h1>Clients</h1>
  <p class="subtitulo">Generat el {GENERADO_EL}</p>
  <table>
    <thead><tr><th>Client</th><th>NIF</th><th>Informe</th><th>Excel</th></tr></thead>
    <tbody>{filas_index}</tbody>
  </table>
</body>
</html>"""

with open("clientes/index.html", "w") as f:
    f.write(index_html)
print("Escrito: clientes/index.html")
