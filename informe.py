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

Piso 9.1: RESULTAT DEL TRIMESTRE amb taula de quotes per tipus d'IVA
(no nomes la linia). Etiqueta ABONAMENT (total negatiu). Verificacions
pures (lletra de NIF, quadre de retencio) -- avisen amb ⚠, mai canvien
estat ni suma. Seccio Errors dins d'Avisos -- fitxers presents sense
extraccio, amb motiu generic (extraer_todas.py no el guarda).

Piso 9.2: decisions.csv (archivo,accion,nota,qui,data) -- 'aprovar' fa
que una targeta REVISAR compti com OK i mostri un distintiu; 'descartar'
la treu de la llista normal i la porta a una seccio Descartats propia
per trimestre. Sense decisions.csv, o buit, el comportament es identic
al d'abans d'aquest pis.

Piso 9.3: construir_enllac() unica per a tots els enllaços -- nomes
generen <a> si l'arxiu existeix de veritat en disc (si no, text pla,
mai un enllaç mort). encontrar_original tambe busca a "procesadas"
(abans exclosa). verificar_enlaces() relegeix cada informe ja escrit i
distingeix "trencat" (bug de codi) de "corrupte" (arxiu buit o truncat
-- problema de dades/sincronitzacio, no de codi). El motiu d'un ERROR
distingeix aquest cas del generic quan es detecta.

Piso 9.4: boto prominent "Obrir l'Excel (còpia de treball)" a la
capçalera de cada informe i a clientes/index.html, amb `download`.
Nomes presentacio -- cap dada ni suma canvia.

Piso 10.5: logo d'Olivella incrustat en base64 (mateix patro que ja
fa servir tarjeta_factura per als jpg/png) a la capçalera de cada
informe i de l'index. Nomes presentacio.

Piso 11B: distintiu "Corregit" (mateix patro additiu que ABONAMENT o
"aprovat manualment") -- si la fitxa te camps_corregits (escrit per
validar.py, Piso 11B), es mostra el badge i el detall antic->nou. Cap
suma ni logica canvia aqui, nomes presentacio del que ja ve calculat.
"""

import base64
import csv
import html
import json
import os
import re
from datetime import datetime
from urllib.parse import quote, unquote

TIPOS_IVA = [0, 4, 5, 10, 12, 21]
EXTENSIONES_ORIGINAL = (".pdf", ".jpg", ".jpeg", ".png")
EXTENSIONES_IMAGEN = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png"}

RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS = {"davinstal": "Emeses/davinstal"}
SUBCARPETAS_RESERVADAS = {"extraidas", "validadas", "procesadas", "lotes_escaneados", "lotes_procesados"}
# Piso 9.3: para buscar un ORIGINAL (no para detectar errores), "procesadas"
# no se excluye -- ahi es donde un PDF ya tratado puede haberse movido
# (regla 5, run idempotente), y sigue siendo su ubicacion legitima.
SUBCARPETAS_NO_ORIGINALES = {"extraidas", "validadas", "lotes_escaneados", "lotes_procesados"}

GENERADO_EL = datetime.now().strftime("%d/%m/%Y %H:%M")

# Piso 10.5: logo incrustado en base64 (mismo patro que ya fa servir
# tarjeta_factura per als jpg/png dels originals) -- aixi es veu igual
# es obri l'informe des d'on es obri, sense dependre de la distancia
# relativa fins a assets/.
LOGO_HTML = ""
if os.path.exists("assets/logo_olivella.png"):
    with open("assets/logo_olivella.png", "rb") as f:
        _logo_b64 = base64.standard_b64encode(f.read()).decode("utf-8")
    LOGO_HTML = f'<img src="data:image/png;base64,{_logo_b64}" alt="Olivella" class="logo-capcalera">'

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


TABLA_LETRA_DNI = "TRWAGMYFPDXBNJZSQVHLCKE"


def validar_nif(nif):
    """Igual que en sumar.py: verificacion pura de la letra de control
    de un NIF espanyol (DNI, NIE o CIF). True/False si reconoce el
    formato; None si no tiene forma reconocible (ej. NIF extranjero) --
    no se avisa en ese caso. No cambia ningun estado ni suma."""
    if nif is None:
        return None
    n = "".join(c for c in nif if c.isalnum()).upper()

    if len(n) == 9 and n[:8].isdigit() and n[8].isalpha():
        return TABLA_LETRA_DNI[int(n[:8]) % 23] == n[8]

    if len(n) == 9 and n[0] in "XYZ" and n[1:8].isdigit() and n[8].isalpha():
        prefijo = {"X": "0", "Y": "1", "Z": "2"}[n[0]]
        return TABLA_LETRA_DNI[int(prefijo + n[1:8]) % 23] == n[8]

    if len(n) == 9 and n[0] in "ABCDEFGHJKLMNPQRSUVW" and n[1:8].isdigit() and (n[8].isdigit() or n[8].isalpha()):
        digitos = n[1:8]
        suma_par = sum(int(d) for d in digitos[1::2])
        suma_impar = 0
        for d in digitos[0::2]:
            doble = int(d) * 2
            suma_impar += doble // 10 + doble % 10
        digito_control = (10 - (suma_par + suma_impar) % 10) % 10
        letra_control = "JABCDEFGHI"[digito_control]
        if n[0] in "ABEH":
            return n[8] == str(digito_control)
        if n[0] in "KPQS":
            return n[8] == letra_control
        return n[8] == str(digito_control) or n[8] == letra_control

    return None


def verificar_retencion(datos):
    """Igual que en sumar.py: True si no hay retencion o si cuadra
    (tolerancia 0,02). No cambia ningun estado ni suma."""
    retencion_pct = datos.get("retencion_pct") or 0
    retencion_cuota = datos.get("retencion_cuota") or 0
    if not retencion_pct and not retencion_cuota:
        return True
    suma_base = sum((l.get("base") or 0) for l in (datos.get("lineas_iva") or []))
    esperado = suma_base * retencion_pct / 100
    return abs(esperado - retencion_cuota) <= 0.02


MOTIVO_ERROR_GENERICO = (
    "No s'ha pogut generar la fitxa — l'arxiu és present però no hi ha extracció. "
    "Cal revisar l'escaneig o tornar-ho a intentar."
)

MOTIVO_ERROR_CORRUPTE = (
    "L'arxiu original és present però buit o corromput — probablement un "
    "problema de sincronització. Cal tornar-lo a sincronitzar o demanar-lo de nou."
)


def archivo_corrupto(ruta):
    """True si la ruta existe pero esta vacia, o es un PDF sense trailer
    %%EOF en l'ultim KB (truncat a mitges -- p. ex. per una sincronitzacio
    d'iCloud interrompuda). No es una validacio generica de PDF: nomes
    detecta el patro de truncament vist en dades reals (Piso 9.3)."""
    if not os.path.exists(ruta):
        return False
    tamano = os.path.getsize(ruta)
    if tamano == 0:
        return True
    if ruta.lower().endswith(".pdf"):
        with open(ruta, "rb") as f:
            f.seek(max(0, tamano - 1024))
            cola = f.read()
        if b"%%EOF" not in cola:
            return True
    return False


def motivo_error(ruta):
    """Motiu derivat -- no inventat -- per a un archiu ERROR: distingeix
    corrupcio comprovable en disc del generic 'no hi ha extraccio'."""
    return MOTIVO_ERROR_CORRUPTE if archivo_corrupto(ruta) else MOTIVO_ERROR_GENERICO


def detectar_errores(rutas_presentes, carpeta_extraidas):
    """Igual que en sumar.py: archivos presentes sin ficha extraida --
    unica senal disponible, porque extraer_todas.py no persiste el
    motivo de un fallo de extraccion."""
    extraidos = set()
    if os.path.isdir(carpeta_extraidas):
        extraidos = {
            os.path.splitext(f)[0] for f in os.listdir(carpeta_extraidas) if f.lower().endswith(".json")
        }
    return [r for r in rutas_presentes if os.path.splitext(os.path.basename(r))[0] not in extraidos]


def avisos_verificacion(facturas):
    """Igual que en sumar.py: NIF sin letra valida o retencion que no
    cuadra. No cambia ningun estado ni suma."""
    avisos = []
    for nombre, datos in facturas:
        for campo, etiqueta in [("nif_proveedor", "proveïdor"), ("nif_receptor", "receptor")]:
            valor = datos.get(campo)
            if validar_nif(valor) is False:
                avisos.append((nombre, datos, f"NIF no supera la validació de la lletra ({etiqueta}: {valor})"))
        if not verificar_retencion(datos):
            avisos.append((nombre, datos, "la retenció calculada no quadra amb la retenció indicada"))
    return avisos


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


def cargar_decisiones(carpeta_cliente):
    """Igual que en sumar.py: lee decisions.csv si existe. Sin
    archivo, o vacio, devuelve {} -- comportamiento identico al de
    antes de este piso."""
    ruta = f"{carpeta_cliente}/decisions.csv"
    decisiones = {}
    if not os.path.exists(ruta):
        return decisiones
    with open(ruta) as f:
        for fila in csv.DictReader(f):
            archivo = fila.get("archivo")
            if archivo:
                decisiones[archivo] = fila
    return decisiones


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
    subcarpetas, ej. davinstal). Piso 9.3: "procesadas" SI se busca aqui
    (a diferencia de listar_archivos_rebudes) -- un original ya movido
    ahi tras procesarse sigue siendo su ubicacion legitima."""
    base = os.path.splitext(nombre_json)[0]
    for ext in EXTENSIONES_ORIGINAL:
        ruta = os.path.join(carpeta_origen, base + ext)
        if os.path.exists(ruta):
            return ruta
    if os.path.isdir(carpeta_origen):
        for nombre_sub in sorted(os.listdir(carpeta_origen)):
            ruta_sub = os.path.join(carpeta_origen, nombre_sub)
            if not os.path.isdir(ruta_sub) or nombre_sub.lower() in SUBCARPETAS_NO_ORIGINALES:
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


def sumar_bloque(facturas, decisiones):
    """Igual que en sumar.py: total y cuota por tipo de IVA de las que
    cuentan, y aparte revisar y descartados. decisions.csv puede
    anular el estado: 'aprovar' hace que una REVISAR cuente como OK;
    'descartar' nunca cuenta y va aparte. Sin decisions.csv, o vacio,
    el comportamiento es identico al de antes de este piso."""
    sumas = {tipo: {"cuota": 0.0} for tipo in TIPOS_IVA}
    sumas["otros"] = {"cuota": 0.0}
    total_ok = 0.0
    revisar = []
    descartados = []
    for nombre, datos in facturas:
        decision = decisiones.get(nombre)

        if decision and decision.get("accion") == "descartar":
            descartados.append((nombre, datos, decision))
            continue

        cuenta = datos.get("estado") == "OK" or (decision and decision.get("accion") == "aprovar")
        if not cuenta:
            revisar.append((nombre, datos))
            continue

        total_ok += datos.get("total") or 0
        for linea in datos.get("lineas_iva") or []:
            tipo = linea.get("tipo_iva")
            clave = tipo if tipo in TIPOS_IVA else "otros"
            sumas[clave]["cuota"] += linea.get("cuota") or 0
    return total_ok, sumas, revisar, descartados


def ruta_relativa_html(ruta_original, carpeta_cliente):
    """Ruta relativa a la carpeta del cliente, con espacios/comas
    escapados para que el navegador la resuelva bien."""
    relativa = os.path.relpath(ruta_original, carpeta_cliente)
    return quote(relativa)


def esc(valor):
    """Escapa texto para HTML; None se convierte en cadena vacia."""
    return html.escape(str(valor)) if valor is not None else ""


def construir_enllac(ruta_original, carpeta_cliente, texto, clase=""):
    """Piso 9.3: unica funcio d'enllaç del informe -- la fan servir
    totes les targetes i taules (OK, PENDENT, ERROR, DESCARTAT). Nomes
    torna un <a> si l'arxiu existeix DE VERITAT en disc; si no, None
    (el crider ja sap com mostrar el nom d'arxiu sense enllaç mort)."""
    if not (ruta_original and os.path.exists(ruta_original)):
        return None
    href = ruta_relativa_html(ruta_original, carpeta_cliente)
    clase_attr = f' class="{clase}"' if clase else ""
    return f'<a{clase_attr} href="{href}" target="_blank">{esc(texto)}</a>'


def verificar_enlaces(ruta_html, carpeta_cliente):
    """Piso 9.3: relegeix el HTML ja escrit i comprova que cada href
    apunta a un arxiu real. 'Trencat' (no existeix) es un bug de codi
    i no hauria de passar mai -- 'corrupte' (existeix pero buit o
    truncat) es un problema de dades/sincronitzacio que cap canvi de
    codi pot arreglar, nomes informar."""
    with open(ruta_html) as f:
        contenido = f.read()
    trencats = []
    corruptes = []
    verificados = 0
    for href in re.findall(r'href="([^"]+)"', contenido):
        if href.startswith(("http://", "https://", "mailto:")):
            continue
        verificados += 1
        ruta = os.path.join(carpeta_cliente, unquote(href))
        if not os.path.exists(ruta):
            trencats.append(href)
        elif archivo_corrupto(ruta):
            corruptes.append(href)
    print(
        f"{os.path.basename(carpeta_cliente)} / informe: {verificados} enllaços verificats, "
        f"{len(trencats)} trencats, {len(corruptes)} corruptes"
    )
    for href in trencats:
        print(f"  TRENCAT: {href}")
    return trencats, corruptes


def tarjeta_factura(nombre, datos, tipo_bloque, carpeta_original, carpeta_cliente, decision=None):
    """Fitxa completa d'UNA factura, OK o REVISAR. Esquerra: tots els
    camps extrets. Dreta: imatge incrustada (jpg/png) o enllaç gran (pdf).
    decision (Piso 9.2): fila de decisions.csv si n'hi ha una per aquest
    archivo -- 'aprovar' mostra un distintiu i pinta la targeta com OK."""
    ruta_original = encontrar_original(carpeta_original, nombre)
    extension = os.path.splitext(ruta_original)[1].lower() if ruta_original else None

    if ruta_original and extension in EXTENSIONES_IMAGEN:
        with open(ruta_original, "rb") as f:
            b64 = base64.standard_b64encode(f.read()).decode("utf-8")
        lado_derecho = f'<img loading="lazy" src="data:{EXTENSIONES_IMAGEN[extension]};base64,{b64}" alt="original">'
    else:
        lado_derecho = construir_enllac(ruta_original, carpeta_cliente, "Obrir original ↗", clase="btn-abrir") or (
            '<p class="sin-original">Original no localitzat</p>'
        )

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

    accion_decision = decision.get("accion") if decision else None
    clase_estado = "ok" if (estado != "REVISAR" or accion_decision == "aprovar") else "revisar"

    es_abonament = (datos.get("total") or 0) < 0
    abonament_html = '<span class="etiqueta-abonament">ABONAMENT</span>' if es_abonament else ""

    tiene_avisos = (
        validar_nif(datos.get("nif_proveedor")) is False
        or validar_nif(datos.get("nif_receptor")) is False
        or not verificar_retencion(datos)
    )
    avisos_html = '<span class="etiqueta-verificacio">⚠ verificació</span>' if tiene_avisos else ""

    decisio_html = ""
    nota_decisio_html = ""
    if accion_decision == "aprovar":
        qui = esc(decision.get("qui"))
        data_decisio = esc(decision.get("data"))
        decisio_html = f'<span class="etiqueta-decisio aprovat">★ Aprovat manualment ({qui}, {data_decisio})</span>'
        if decision.get("nota"):
            nota_decisio_html = f"<p class=\"nota-decisio\"><strong>Nota:</strong> {esc(decision.get('nota'))}</p>"

    # Piso 11B: distintiu "corregit" -- correccions.csv es aplica en
    # memoria dins de validar.py (cirugia minima alli). Aqui nomes es
    # mostra el que ja ve escrit a la fitxa, cap suma ni logica canvia.
    camps_corregits = datos.get("camps_corregits") or []
    corregit_html = '<span class="etiqueta-decisio corregit">✎ Corregit</span>' if camps_corregits else ""
    corregit_detall_html = ""
    if camps_corregits:
        items = "".join(
            f"<li>{esc(c['camp'])}: {esc(c['antic'])} → {esc(c['nou'])} ({esc(c['qui'])}, {esc(c['data'])})</li>"
            for c in camps_corregits
        )
        corregit_detall_html = f"<p><strong>Correccions aplicades:</strong></p><ul>{items}</ul>"

    return f"""
    <div class="tarjeta {clase_estado}">
      <div class="tarjeta-izq">
        <span class="etiqueta-tipo etiqueta-{tipo_bloque.lower()}">{tipo_bloque}</span>
        <span class="etiqueta-estado">{esc(estado)}</span>
        {abonament_html}
        {avisos_html}
        {decisio_html}
        {corregit_html}
        <h3>{esc(datos.get("proveedor"))}</h3>
        <p>NIF: {esc(datos.get("nif_proveedor"))} · Factura: {esc(datos.get("num_factura"))} · Data: {esc(datos.get("fecha_factura"))}</p>
        <ul class="lineas-iva">{lineas_html}</ul>
        <p>Total: {esc(datos.get("total"))} €</p>
        {retencion_html}
        <p class="archivo">{esc(nombre)}</p>
        {motivos_html}
        {observaciones_html}
        {nota_decisio_html}
        {corregit_detall_html}
      </div>
      <div class="tarjeta-der">{lado_derecho}</div>
    </div>"""


TIPO_BLOQUE_SINGULAR = {"DESPESES": "DESPESA", "INGRESSOS": "INGRÉS"}


def seccion_flujo(titulo, facturas, carpeta_original, carpeta_cliente, decisiones):
    total_ok, sumas, revisar, descartados = sumar_bloque(facturas, decisiones)
    tipo_bloque = TIPO_BLOQUE_SINGULAR[titulo]
    nombres_descartados = {nombre for nombre, _, _ in descartados}
    tarjetas = "".join(
        tarjeta_factura(nombre, datos, tipo_bloque, carpeta_original, carpeta_cliente, decisiones.get(nombre))
        for nombre, datos in facturas
        if nombre not in nombres_descartados
    )
    n_mostradas = len(facturas) - len(descartados)
    resumen = f"{titulo} — {n_mostradas} factures, {len(revisar)} a revisar — {total_ok:.2f} €"
    descartados_con_origen = [
        (nombre, datos, tipo_bloque, carpeta_original, decision) for nombre, datos, decision in descartados
    ]
    return f"""
    <details>
      <summary>{esc(resumen)}</summary>
      {tarjetas if tarjetas else "<p>Sense factures.</p>"}
    </details>""", total_ok, sumas, descartados_con_origen


def seccion_descartats(descartados, carpeta_cliente):
    """Facturas amb decisions.csv accio=descartar -- mai compten en
    cap suma, es llisten a part amb la seva nota (mirall del bloc
    DESCARTATS de sumar.py)."""
    filas = ""
    for nombre, datos, tipo_bloque, carpeta_original, decision in descartados:
        ruta_original = encontrar_original(carpeta_original, nombre)
        enlace = construir_enllac(ruta_original, carpeta_cliente, f"{nombre} ↗") or esc(nombre)
        nota = decision.get("nota") or ""
        qui = decision.get("qui") or ""
        data_decisio = decision.get("data") or ""
        filas += f"""<tr><td>{enlace}</td><td>{esc(tipo_bloque)}</td>
            <td>{esc(nota)} ({esc(qui)}, {esc(data_decisio)})</td></tr>"""
    return f"""
    <details>
      <summary>Descartats — {len(descartados)}</summary>
      <table class="descartats">
        <thead><tr><th>Fitxer</th><th>Tipus</th><th>Nota</th></tr></thead>
        <tbody>{filas}</tbody>
      </table>
    </details>"""


def tabla_resultat_trimestre(sumas_g, sumas_i):
    """RESULTAT DEL TRIMESTRE, protagonista al inicio -- misma Sigma
    cuota por tipo de IVA que sumar_bloque ya calcula, no se recalcula
    nada nuevo. No es una liquidacion oficial, solo un resultat de
    treball."""
    tipos_a_escribir = list(TIPOS_IVA)
    if sumas_g["otros"]["cuota"] or sumas_i["otros"]["cuota"]:
        tipos_a_escribir.append("otros")

    filas = ""
    cuota_g_total = 0.0
    cuota_i_total = 0.0
    for tipo in tipos_a_escribir:
        cg = sumas_g[tipo]["cuota"]
        ci = sumas_i[tipo]["cuota"]
        etiqueta_tipo = "ALTRES" if tipo == "otros" else tipo
        filas += f'<tr><td>{esc(etiqueta_tipo)}</td><td class="num">{cg:.2f}</td><td class="num">{ci:.2f}</td></tr>'
        cuota_g_total += cg
        cuota_i_total += ci

    resultat = cuota_i_total - cuota_g_total
    if resultat > 0:
        etiqueta = " (a ingressar)"
    elif resultat < 0:
        etiqueta = " (a compensar)"
    else:
        etiqueta = ""

    return f"""
    <table class="resultat-trimestre">
      <thead><tr><th>Tipus IVA</th><th>Quota despeses</th><th>Quota ingressos</th></tr></thead>
      <tbody>{filas}</tbody>
    </table>
    <p class="resultat-iva"><strong>RESULTAT IVA (repercutit − suportat): {resultat:.2f} €{etiqueta}</strong><br>
    <span class="nota">IVA cobrat menys IVA pagat; positiu = a ingressar, negatiu = a compensar
    (resultat de treball, no liquidació oficial)</span></p>
    """


def seccion_trimestre(trimestre, datos_t, carpeta_cliente, origen_gastos, origen_ingressos, decisiones):
    html_gastos, total_g, sumas_g, descartats_g = seccion_flujo(
        "DESPESES", datos_t["gastos"], origen_gastos, carpeta_cliente, decisiones
    )
    html_ingresos, total_i, sumas_i, descartats_i = seccion_flujo(
        "INGRESSOS", datos_t["ingresos"], origen_ingressos, carpeta_cliente, decisiones
    )
    n_total = len(datos_t["gastos"]) + len(datos_t["ingresos"])
    resumen = f"{trimestre} — {n_total} factures ({total_g:.2f} € despeses, {total_i:.2f} € ingressos)"

    resultat_html = tabla_resultat_trimestre(sumas_g, sumas_i)
    descartats_totales = descartats_g + descartats_i
    descartats_html = seccion_descartats(descartats_totales, carpeta_cliente) if descartats_totales else ""

    return f"""
    <details>
      <summary class="resumen-trimestre">{esc(resumen)}</summary>
      {resultat_html}
      {html_gastos}
      {html_ingresos}
      {descartats_html}
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


def seccion_avisos(carpeta_cliente, gastos, ingresos, origen_gastos, origen_ingressos):
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
            f"""<tr><td>{construir_enllac(os.path.join(carpeta_albarans, n), carpeta_cliente, n) or esc(n)}</td>
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
            lote_html = construir_enllac(ruta_lote, carpeta_cliente, nombre_lote) or esc(nombre_lote)
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

    # d) Errors -- archivos presentes sin ficha extraida (no tienen fecha, no van en ningun trimestre)
    errores_gastos = detectar_errores(listar_archivos_rebudes(origen_gastos), f"{carpeta_cliente}/rebudes/extraidas")
    errores_ingresos = detectar_errores(
        [f for f in [os.path.join(origen_ingressos, n) for n in os.listdir(origen_ingressos)] if f.lower().endswith(EXTENSIONES_ORIGINAL)]
        if os.path.isdir(origen_ingressos) else [],
        f"{carpeta_cliente}/apartados/ingressos_extraidas",
    )
    filas_error = [("DESPESA", r) for r in errores_gastos] + [("INGRÉS", r) for r in errores_ingresos]
    if not filas_error:
        html_errores = "<p>Sense errors.</p>"
    else:
        filas = ""
        for flujo, ruta in filas_error:
            enlace = construir_enllac(ruta, carpeta_cliente, os.path.basename(ruta)) or esc(os.path.basename(ruta))
            filas += f"""<tr><td>[{esc(flujo)}] {enlace}</td>
                <td>{esc(motivo_error(ruta))}</td></tr>"""
        html_errores = f'<table class="errors"><thead><tr><th>Fitxer</th><th>Motiu</th></tr></thead><tbody>{filas}</tbody></table>'

    # e) Avisos de verificacio -- letra de NIF y cuadre de retencion, no cambian estado
    filas_verificacion = (
        [(origen_gastos, nombre, motivo) for nombre, _, motivo in avisos_verificacion(gastos)]
        + [(origen_ingressos, nombre, motivo) for nombre, _, motivo in avisos_verificacion(ingresos)]
    )
    if not filas_verificacion:
        html_verificacion = "<p>Sense avisos de verificació.</p>"
    else:
        filas = ""
        for carpeta_origen, nombre, motivo in filas_verificacion:
            ruta_original = encontrar_original(carpeta_origen, nombre)
            enlace = construir_enllac(ruta_original, carpeta_cliente, nombre) or esc(nombre)
            filas += f"<tr><td>{enlace}</td><td>{esc(motivo)}</td></tr>"
        html_verificacion = f'<table class="verificacio"><thead><tr><th>Fitxer</th><th>Motiu</th></tr></thead><tbody>{filas}</tbody></table>'

    return f"""
    <h2>Avisos</h2>
    <h3>Documents apartats</h3>
    {html_apartats}
    <h3>Pàgines descartades com a soroll</h3>
    {html_ruido}
    <h3>Avisos de consistència</h3>
    {html_consistencia}
    <h3>Errors</h3>
    {html_errores}
    <h3>Avisos de verificació</h3>
    {html_verificacion}
    """


ESTILO = """
body { font-family: -apple-system, "Segoe UI", Roboto, sans-serif; max-width: 1100px;
       margin: 2rem auto; padding: 0 1rem; color: #1a1a1a; line-height: 1.5; }
.capcalera { display: flex; align-items: center; gap: 0.8rem; }
.logo-capcalera { height: 40px; width: auto; }
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
.etiqueta-tipo, .etiqueta-estado, .etiqueta-abonament, .etiqueta-verificacio, .etiqueta-decisio {
  display: inline-block; padding: 0.15rem 0.6rem; border-radius: 4px;
  font-size: 0.8rem; font-weight: bold; background: #fff; margin-right: 0.3rem; }
.etiqueta-abonament { background: #FBE5D6; color: #833C00; }
.etiqueta-verificacio { background: #FFC7CE; color: #9C0006; }
.etiqueta-decisio.aprovat { background: #E2EFDA; color: #375623; }
.etiqueta-decisio.corregit { background: #D9E8F5; color: #0D3D6B; }
.lineas-iva { margin: 0.3rem 0; padding-left: 1.2rem; }
.archivo { font-family: monospace; font-size: 0.8rem; color: #555; }
.nota-decisio { font-style: italic; }
table.descartats tbody tr { background: #E7E6E6; }
.btn-abrir { display: inline-block; background: #0563C1; color: white; padding: 0.8rem 1.5rem;
             border-radius: 6px; text-decoration: none; font-weight: bold; }
.btn-excel { display: inline-block; background: #1D6F42; color: white; padding: 0.6rem 1.2rem;
             border-radius: 6px; text-decoration: none; font-weight: bold; margin: 0.5rem 0; }
.sin-original { color: #a33; font-style: italic; }

table:not(.comparacion):not(.conciliacio) { border-collapse: collapse; width: 100%; margin: 0.6rem 0 1.5rem 0; font-size: 0.9rem; }
table:not(.comparacion):not(.conciliacio) th, table:not(.comparacion):not(.conciliacio) td {
  border: 1px solid #ddd; padding: 0.4rem 0.6rem; text-align: left; }
table:not(.comparacion):not(.conciliacio) th { background: #D9E1F2; }
table.errors tbody tr { background: #FFC7CE; }

@media (prefers-color-scheme: dark) {
  body { background: #1a1a1a; color: #eee; }
  summary, table.comparacion th, table:not(.comparacion) th { background: #2a3a52; }
  summary.resumen-trimestre { background: #35496b; }
  .tarjeta { border-color: #444; }
  .tarjeta.ok { background: #22331f; border-color: #3a5230; }
  .tarjeta.revisar { background: #3a2a1f; border-color: #5c4128; }
  .etiqueta-tipo, .etiqueta-estado { background: #333; }
  .etiqueta-abonament { background: #4a3423; color: #f0c090; }
  .etiqueta-verificacio { background: #4a2323; color: #f0a0a0; }
  .etiqueta-decisio.aprovat { background: #22331f; color: #b7d7a8; }
  .etiqueta-decisio.corregit { background: #1a2e42; color: #a9cdf0; }
  table.errors tbody tr { background: #4a2323; }
  table.descartats tbody tr { background: #333; }
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

    decisiones = cargar_decisiones(carpeta_cliente)
    nombres_validos = {n for n, _ in gastos} | {n for n, _ in ingresos}
    for archivo in decisiones:
        if archivo not in nombres_validos:
            print(f"AVISO: decisions.csv de {carpeta} referencia un archivo que no existe entre las validadas: {archivo}")

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

    # Piso 9.4: boton prominente al Excel de trabajo del cliente, solo si existe
    ruta_excel_cliente = f"{carpeta_cliente}/sumatorios_2026.xlsx"
    boton_excel_html = (
        '<a class="btn-excel" href="sumatorios_2026.xlsx" download>Obrir l\'Excel (còpia de treball)</a>'
        if os.path.exists(ruta_excel_cliente) else ""
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
            html_gastos, total_g, _, descartats_g = seccion_flujo(
                "DESPESES", datos_t["gastos"], origen_gastos, carpeta_cliente, decisiones
            )
            html_ingresos, total_i, _, descartats_i = seccion_flujo(
                "INGRESSOS", datos_t["ingresos"], origen_ingressos, carpeta_cliente, decisiones
            )
            descartats_totales = descartats_g + descartats_i
            descartats_html = seccion_descartats(descartats_totales, carpeta_cliente) if descartats_totales else ""
            secciones_html.append(f"""
            <details>
              <summary class="resumen-trimestre">SENSE DATA — {len(datos_t["gastos"]) + len(datos_t["ingresos"])} factures</summary>
              {html_gastos}
              {html_ingresos}
              {descartats_html}
            </details>""")
            continue

        html_trimestre, total_g, total_i = seccion_trimestre(
            trimestre, datos_t, carpeta_cliente, origen_gastos, origen_ingressos, decisiones
        )
        secciones_html.append(html_trimestre)
        filas_comparacion.append((trimestre, "DESPESES", total_g))
        filas_comparacion.append((trimestre, "INGRESSOS", total_i))

    html_avisos = seccion_avisos(carpeta_cliente, gastos, ingresos, origen_gastos, origen_ingressos)

    html_final = f"""<!DOCTYPE html>
<html lang="ca">
<head>
<meta charset="utf-8">
<title>Informe {esc(fila_cliente["nombre"])}</title>
<style>{ESTILO}</style>
</head>
<body>
  <div class="capcalera">
    {LOGO_HTML}
    <h1>{esc(fila_cliente["nombre"])}</h1>
  </div>
  <p class="subtitulo">NIF {esc(fila_cliente["nif"])} · Generat el {GENERADO_EL}</p>
  {boton_excel_html}

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
    verificar_enlaces(ruta_html, carpeta_cliente)

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
        f'<a class="btn-excel" href="{c["carpeta"]}/sumatorios_2026.xlsx" download>Obrir l\'Excel (còpia de treball)</a>'
        if c["excel_existe"] else "no disponible"
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
  <div class="capcalera">
    {LOGO_HTML}
    <h1>Clients</h1>
  </div>
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
