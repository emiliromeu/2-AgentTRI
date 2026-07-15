"""Aplica la red de validacion de la seccion 3 del esquema a los JSON extraidos.

Piso 3: sin llamadas a la API, corre gratis. Lee extraidas/ (fixtures,
nunca se tocan) y escribe en validadas/ cada factura con su estado
(OK o REVISAR) y sus motivos.

Piso 4B: recorre todos los clientes de clientes.csv; el NIF de receptor
esperado ya no esta fijo, sale de la fila de cada cliente.

Piso 5: normaliza NIF antes de comparar, ignora duplicados (None, None),
y procesa tambien el flujo ingressos (liquidaciones de cooperativa)
ademas del flujo rebudes (facturas de compra).

Piso 6B: en ingressos el cliente puede ser receptor (liquidacion) o
emisor (factura emitida por el propio cliente) -- el giro. En rebudes
el cliente sigue siendo siempre el receptor, sin excepcion.

Piso 11B: unica cirugia del proyecto sobre este archivo -- ni una
linea de la red de validacion cambia. Antes de validar cada ficha,
si existen correcciones en clientes/<carpeta>/correccions.csv para
ese archivo (Piso 11B, escritas desde la app), se aplican EN MEMORIA
sobre una copia (el JSON de extraidas/ nunca se toca) y la ficha
corregida entra por la MISMA red sin atajos: corregir no aprova,
torna a passar l'examen. Si el resultado pasa, sale OK con
camps_corregits anotado (camp, antic->nou, qui, data); si no, sigue
REVISAR con los motivos que toquen. Sin correccions.csv, o vacio, el
comportamiento -- y el JSON de salida, byte a byte -- es identico al
de antes de este piso (camps_corregits solo se añade si hay algo que
corregir).

Piso 13K: bug de campo -- en ingressos donde el cliente es el emisor
(factura propia), la presentacion mostraba "proveedor" (el propio
cliente) en vez de su comprador. Nueva capa derivada -- la red de
validacion NO cambia -- contrapart_nom/contrapart_nif: la parte cuyo
NIF NO es el del cliente, calculada por NIF, no por posicion del
campo. Si las dos partes fueran el cliente, motivo nuevo "las dos
partes son el cliente" (REVISAR).

Piso 13J: normalizar_nif tambien quita el prefijo intracomunitario
"ES" (ej. "ES37266020V" == "37266020V") antes de comparar -- causaba
PENDENTS falsos de identidad cuando una factura trae el NIF en
formato UE. El xec de duplicados (nif_proveedor, num_factura) ya no
cuenta una factura con num_factura ausente -- antes bastaba que
coincidiera el proveedor para marcar "duplicada" encima del motivo
real de campo vacio.

Piso 13Q: bug real -- convertir_valor() llamaba float() a secas sobre
el texto de una correccion manual, sin try/except; si alguien tecleaba
un numero en formato español ("1.234,56") o basura ("N/A") en el
dialogo "Corregir camps" de la app, el siguiente Recalcular tumbaba
TODO el batch de validar.py (ValueError sin capturar, un solo script,
sin aislamiento por cliente). a_numero() nunca lanza: entiende formato
español, y si no puede convertir un valor no vacio, la ficha cae a
REVISAR con el motivo "importe numérico ilegible en...", nunca crashea
el lote. Ademas del camino de correcciones (unico donde se ha
reproducido el bug), normalizar_numeros() aplica el mismo blindaje de
forma defensiva a los campos ya cargados de extraidas/ -- pero SOLO si
ya son string (nunca toca un int/float existente, cero riesgo de
regresion en los datos reales, que siempre traen tipos nativos).
"""

import csv
import json
import os
import re
from pathlib import Path

# Piso 13B: ancla el cwd a la carpeta del propio script -- ver trocear.py.
RAIZ = Path(__file__).resolve().parent
os.chdir(RAIZ)

TOLERANCIA = 0.02

CAMPOS_OBLIGATORIOS = [
    "proveedor", "nif_proveedor", "num_factura", "fecha_factura",
    "receptor", "nif_receptor", "lineas_iva", "total",
    "retencion_pct", "retencion_cuota", "exenta",
]


def leer_clientes():
    with open("clientes/clientes.csv", encoding="utf-8") as f:
        return list(csv.DictReader(f))


PATRON_NIF_UE = re.compile(r"^ES([0-9A-Z]{9})$")


def normalizar_nif(nif):
    """Quita guiones/espacios y pasa a mayúsculas, para que un NIF con
    formato distinto (ej. "37266020-V" vs "37266020V") no dispare un
    motivo falso. Piso 13J: tambien quita el prefijo "ES" del formato
    intracomunitario (ej. "ES37266020V") si va delante de un NIF/CIF
    domestico de 9 caracteres -- un CIF domestico normal tiene 9
    caracteres exactos, nunca 11, asi que no hay riesgo de falso
    positivo (ej. "B66515529" se queda intacto)."""
    if nif is None:
        return None
    n = "".join(c for c in nif if c.isalnum()).upper()
    m = PATRON_NIF_UE.match(n)
    return m.group(1) if m else n


def cargar_correcciones(carpeta_cliente):
    """Piso 11B: lee correccions.csv (arxiu,camp,valor_antic,valor_nou,
    motiu,qui,data) si existe -- una fila por CAMPO corregido, varias
    filas pueden compartir el mismo arxiu. Sin archivo, o vacio,
    devuelve {} -- mismo criterio que decisions.csv (Piso 9.2/11A):
    nada inventado, la app lo rellena a partir de acciones reales."""
    ruta = f"{carpeta_cliente}/correccions.csv"
    correcciones = {}
    if not os.path.exists(ruta):
        return correcciones
    with open(ruta, encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            archivo = fila.get("arxiu")
            if archivo:
                correcciones.setdefault(archivo, []).append(fila)
    return correcciones


def cargar_decisiones(carpeta_cliente):
    """Piso 13T: mismo criterio que app.py/sumar.py/informe.py (duplicado
    a proposito, ninguna "maquina" importa otra) -- decisions.csv es
    llibre major (Piso 13M): la ULTIMA fila de cada arxiu es el seu
    estat efectiu, i un "revertir" l'elimina del diccionari. Aqui nomes
    interessa saber quins arxius estan "descartar" -- per treure'ls del
    xec de duplicats, mai per canviar estado/motivos (aixo segueix sent
    nomes de la xarxa de validacio)."""
    ruta = f"{carpeta_cliente}/decisions.csv"
    decisiones = {}
    if not os.path.exists(ruta):
        return decisiones
    with open(ruta, encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            archivo = fila.get("archivo")
            if not archivo:
                continue
            if fila.get("accion") == "revertir":
                decisiones.pop(archivo, None)
            else:
                decisiones[archivo] = fila
    return decisiones


CAMPOS_NUMERICOS_LINEA = {"tipo_iva", "base", "cuota"}
CAMPOS_NUMERICOS_TOP = {"total", "retencion_pct", "retencion_cuota"}


def a_numero(valor):
    """Piso 13Q: convierte texto en formato español ("1.234,56", "1234,56",
    con espacios o "€") a float -- NUNCA lanza excepcion. Si ya es un
    numero (int/float), lo devuelve tal cual. Si no se puede convertir
    (texto vacio o basura tipo "abc"/"N/A"), devuelve None -- quien
    llama decide si eso es motivo de REVISAR."""
    if valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace("€", "").replace(" ", "")
    if texto == "":
        return None
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None


def convertir_valor(campo, valor_texto):
    """Piso 11B: correccions.csv llega siempre como texto -- conversion
    minima al tipo que espera el esquema canonico. Piso 13Q: a_numero()
    en vez de float() a secas -- nunca crashea el lote entero por una
    correccion mal tecleada."""
    if campo in CAMPOS_NUMERICOS_LINEA or campo in CAMPOS_NUMERICOS_TOP:
        return a_numero(valor_texto)
    if campo == "exenta":
        return valor_texto.strip().lower() in ("true", "1", "si", "sí")
    return valor_texto


def normalizar_numeros(datos):
    """Piso 13Q: aplica a_numero() a cada campo numerico, pero SOLO si
    ya es un string -- un int/float existente (todo lo real, hoy) se
    deja intacto, cero riesgo de regresion. Defensa en el punto de
    carga por si extraer_todas.py dejara pasar algun dia un string
    (la API siempre ha devuelto tipos nativos hasta ahora). Devuelve
    (datos, illegibles) -- illegibles es una lista de (campo,
    valor_crudo) para los que no se ha podido convertir."""
    datos = dict(datos)
    illegibles = []

    for campo in CAMPOS_NUMERICOS_TOP:
        crudo = datos.get(campo)
        if isinstance(crudo, str):
            convertido = a_numero(crudo)
            if convertido is None and crudo.strip():
                illegibles.append((campo, crudo))
            datos[campo] = convertido

    if datos.get("lineas_iva"):
        lineas_norm = []
        for i, linea in enumerate(datos["lineas_iva"], start=1):
            linea = dict(linea)
            for campo in CAMPOS_NUMERICOS_LINEA:
                crudo = linea.get(campo)
                if isinstance(crudo, str):
                    convertido = a_numero(crudo)
                    if convertido is None and crudo.strip():
                        illegibles.append((f"lineas_iva[{i}].{campo}", crudo))
                    linea[campo] = convertido
            lineas_norm.append(linea)
        datos["lineas_iva"] = lineas_norm

    return datos, illegibles


CAMPOS_NUMERICOS = CAMPOS_NUMERICOS_LINEA | CAMPOS_NUMERICOS_TOP


def aplicar_correcciones(datos, correcciones_archivo):
    """Piso 11B: aplica en memoria las correcciones de UN archivo sobre
    una COPIA de datos -- el JSON de extraidas/ nunca se toca. "camp"
    admite un nombre simple ("total", "nif_receptor"...) o
    "lineas_iva[N].subcampo" para corregir una linea de IVA concreta.
    La ficha corregida sigue despues por la red de validacion sin
    tocar, tal cual -- esta funcion NO valida nada, solo sustituye
    valores.

    Piso 13Q: tambien devuelve correcciones_illegibles -- (camp,
    valor_cru) para cada correccion de un campo numerico cuyo texto
    (no vacio) a_numero() no ha podido convertir. Antes esto lanzaba
    ValueError sin capturar (float() a secas) y tumbaba el batch
    entero; ahora simplemente se anota para que el bucle principal lo
    convierta en un motivo REVISAR."""
    if not correcciones_archivo:
        return datos, [], []

    datos = dict(datos)
    if datos.get("lineas_iva") is not None:
        datos["lineas_iva"] = [dict(linea) for linea in datos["lineas_iva"]]

    camps_corregits = []
    correcciones_illegibles = []
    for correccion in correcciones_archivo:
        camp = correccion["camp"]
        coincide_linea = re.match(r"lineas_iva\[(\d+)\]\.(\w+)", camp)
        if coincide_linea:
            indice, subcampo = int(coincide_linea.group(1)), coincide_linea.group(2)
            antic = datos["lineas_iva"][indice].get(subcampo)
            nuevo = convertir_valor(subcampo, correccion["valor_nou"])
            if nuevo is None and subcampo in CAMPOS_NUMERICOS and correccion["valor_nou"].strip():
                correcciones_illegibles.append((camp, correccion["valor_nou"]))
            datos["lineas_iva"][indice][subcampo] = nuevo
        else:
            antic = datos.get(camp)
            nuevo = convertir_valor(camp, correccion["valor_nou"])
            if nuevo is None and camp in CAMPOS_NUMERICOS and correccion["valor_nou"].strip():
                correcciones_illegibles.append((camp, correccion["valor_nou"]))
            datos[camp] = nuevo
        camps_corregits.append({
            "camp": camp,
            "antic": antic,
            "nou": nuevo,
            "qui": correccion.get("qui"),
            "data": correccion.get("data"),
        })
    return datos, camps_corregits, correcciones_illegibles


# (etiqueta, carpeta origen, carpeta destino) -- mismo esquema y mismas reglas
# para rebudes (facturas de compra) e ingressos (liquidaciones de cooperativa)
FLUJOS = [
    ("rebudes", "rebudes/extraidas", "rebudes/validadas"),
    ("ingressos", "apartados/ingressos_extraidas", "apartados/ingressos_validadas"),
]

ok_total = {"rebudes": 0, "ingressos": 0}
revisar_total = {"rebudes": 0, "ingressos": 0}
ilegibles_total = {"rebudes": 0, "ingressos": 0}

# Piso 13Q: capturado una vez (no dentro del bucle) para poder comprobar,
# cuando una factura no coincide con NINGUN NIF del cliente actual, si su
# contrapart coincide con el NIF de OTRO cliente del registro (sugerencia
# de "sembla que sigui de X", nunca un movimiento automatico).
todos_clientes = leer_clientes()

for fila in todos_clientes:
    carpeta = fila["carpeta"]
    nif_receptor_esperado = fila["nif"]

    # Piso 11B: una unica correccions.csv por cliente, comparte flujo
    # rebudes/ingressos igual que decisions.csv (Piso 9.2/11A).
    correcciones = cargar_correcciones(f"clientes/{carpeta}")
    correcciones_aplicadas = {}
    # Piso 13T: nomes per treure les fitxes "descartar" del xec de
    # duplicats -- una descartada mai compta com a "altre" per a ningu.
    decisiones = cargar_decisiones(f"clientes/{carpeta}")

    for etiqueta, origen_rel, destino_rel in FLUJOS:
        carpeta_entrada = f"clientes/{carpeta}/{origen_rel}"
        carpeta_salida = f"clientes/{carpeta}/{destino_rel}"

        if not os.path.isdir(carpeta_entrada):
            continue

        os.makedirs(carpeta_salida, exist_ok=True)

        nombres_json = sorted(
            f for f in os.listdir(carpeta_entrada) if f.lower().endswith(".json")
        )

        print(f"\n== {carpeta} / {etiqueta} ==")

        # Primera pasada -- cargar cada JSON, detectar los ilegibles
        facturas = []
        ilegibles = 0
        for nombre in nombres_json:
            ruta = os.path.join(carpeta_entrada, nombre)
            try:
                with open(ruta, encoding="utf-8") as f:
                    datos = json.load(f)
                datos, camps_corregits, correcciones_illegibles = aplicar_correcciones(
                    datos, correcciones.get(nombre, [])
                )
                if camps_corregits:
                    correcciones_aplicadas[nombre] = camps_corregits
                datos, illegibles_raw = normalizar_numeros(datos)
                facturas.append((nombre, datos, correcciones_illegibles + illegibles_raw))
            except (json.JSONDecodeError, OSError) as e:
                print(f"AVISO: {nombre} ilegible: {e}")
                ilegibles += 1

        # Entre pasadas -- mapa de (nif_proveedor, num_factura) -> nombres de archivo,
        # para detectar duplicados dentro de este cliente/flujo antes de validar
        # ninguna factura. Piso 13J: num_factura ausente NUNCA cuenta -- ya
        # cae en REVISAR por "campo obligatorio vacío" (mas abajo), y si dos
        # facturas del mismo proveedor tienen las dos num_factura=None,
        # antes disparaban un "duplicada" falso y redundante encima del
        # motivo real. Piso 13T: una fitxa amb decisio efectiva "descartar"
        # tampoc compta -- fora del joc es fora del joc, no contamina cap
        # altra fitxa amb un "duplicada" contra un arxiu que ja no importa.
        claves = {}
        for nombre, datos, _ in facturas:
            if datos.get("num_factura") is None:
                continue
            decision_propia = decisiones.get(nombre)
            if decision_propia and decision_propia.get("accion") == "descartar":
                continue
            clave = (datos.get("nif_proveedor"), datos.get("num_factura"))
            claves.setdefault(clave, []).append(nombre)

        # Segunda pasada -- validar cada factura
        ok = 0
        revisar = 0
        for nombre, datos, illegibles_numeros in facturas:
            motivos = []

            campos_illegibles = {campo for campo, _ in illegibles_numeros}
            for campo, valor_cru in illegibles_numeros:
                motivos.append(f"importe numérico ilegible en {campo}: '{valor_cru}'")

            for campo in CAMPOS_OBLIGATORIOS:
                # Piso 13Q: si ya sabemos que estaba ilegible (no vacio, sino
                # basura que a_numero() no ha podido convertir), no hace
                # falta el motivo generico de "vacío" encima -- seria
                # confuso, el campo no estaba vacío.
                if datos.get(campo) is None and campo not in campos_illegibles:
                    motivos.append(f"campo obligatorio vacío: {campo}")

            lineas = datos.get("lineas_iva") or []
            for i, linea in enumerate(lineas, start=1):
                tipo = linea.get("tipo_iva")
                base = linea.get("base")
                cuota = linea.get("cuota")
                if tipo is None or base is None or cuota is None:
                    motivos.append(f"línea {i} de IVA con campo vacío")
                    continue
                esperado = base * tipo / 100
                if abs(esperado - cuota) > TOLERANCIA:
                    motivos.append(
                        f"línea {i}: {base} × {tipo}% = {esperado:.2f}, pero cuota indica {cuota}"
                    )

            # Piso 13R: coherencia exenta -- antes "exenta" no se contrastaba
            # contra las propias líneas de IVA, así que una ficha podía
            # quedar marcada exenta con líneas de IVA real por debajo,
            # silenciosamente incoherente. corregir sigue sin autoaprobarse
            # (la doctrina de aplicar_correcciones no cambia): esta ficha
            # simplemente vuelve a REVISAR con el motivo de abajo hasta que
            # las líneas queden a 0 o se desmarque exenta.
            if datos.get("exenta") and any((l.get("tipo_iva") or 0) > 0 or (l.get("cuota") or 0) > 0 for l in lineas):
                motivos.append(
                    "marcada como exenta pero tiene líneas con IVA -- corrige las líneas "
                    "(tipo 0, cuota 0) o desmarca exenta"
                )

            total = datos.get("total")
            if total is not None:
                suma = sum((l.get("base") or 0) + (l.get("cuota") or 0) for l in lineas)
                if abs(suma - total) > TOLERANCIA:
                    # Piso 13Y: alguns papers (lloguers, professionals,
                    # agràries) imprimeixen com a TOTAL el líquid --
                    # el brut ja amb la retenció restada. Abans de
                    # marcar motiu es prova aquesta segona identitat;
                    # si només quadra pel net, es normalitza el total
                    # al brut (el que la resta de la màquina espera
                    # sempre) i queda constància -- mai en silenci.
                    retencion_cuota = datos.get("retencion_cuota") or 0
                    if retencion_cuota and abs(suma - retencion_cuota - total) <= TOLERANCIA:
                        total_brut = round(suma, 2)
                        nota = f"total normalitzat: el paper duia el líquid {total} → {total_brut}"
                        datos["observaciones"] = (
                            f"{datos['observaciones']} | {nota}" if datos.get("observaciones") else nota
                        )
                        datos["total"] = total_brut
                        datos["total_normalitzat"] = True
                    else:
                        motivos.append(f"total no cuadra: bases+cuotas={suma:.2f}, total indica {total}")

            # En rebudes el cliente SIEMPRE es el receptor (factura de compra).
            # En ingressos puede ser receptor (liquidacion de cooperativa) o
            # emisor (factura emitida por el propio cliente) -- el giro.
            cliente_normalizado = normalizar_nif(nif_receptor_esperado)
            if etiqueta == "rebudes":
                nif_receptor = datos.get("nif_receptor")
                if nif_receptor is not None and normalizar_nif(nif_receptor) != cliente_normalizado:
                    motivos.append(
                        f"nif_receptor no coincide: esperado {nif_receptor_esperado}, encontrado {nif_receptor}"
                    )
            elif etiqueta == "ingressos":
                nif_proveedor_doc = datos.get("nif_proveedor")
                nif_receptor = datos.get("nif_receptor")
                es_receptor = nif_receptor is not None and normalizar_nif(nif_receptor) == cliente_normalizado
                es_emisor = nif_proveedor_doc is not None and normalizar_nif(nif_proveedor_doc) == cliente_normalizado
                if not es_receptor and not es_emisor:
                    motivos.append("el cliente no aparece ni como emisor ni como receptor")

            # Piso 13Q: suggeriment de client -- NOMES quan cap NIF ha
            # coincidit amb el client actual (el motiu d'identitat de dalt
            # ja s'ha disparat), es comprova si el NIF que sobra coincideix
            # amb un ALTRE client del registre. Mai mou res -- nomes marca
            # la ficha perque la interfície pugui suggerir-ho.
            suggerit_carpeta, suggerit_nom = None, None
            identitat_no_coincideix = any(
                m.startswith("nif_receptor no coincide")
                or m == "el cliente no aparece ni como emisor ni como receptor"
                for m in motivos
            )
            if identitat_no_coincideix:
                if etiqueta == "rebudes":
                    candidats_nif = [datos.get("nif_receptor")]
                else:
                    candidats_nif = [datos.get("nif_receptor"), datos.get("nif_proveedor")]
                for nif_candidat in candidats_nif:
                    nif_candidat_norm = normalizar_nif(nif_candidat)
                    if nif_candidat_norm is None:
                        continue
                    for otro in todos_clientes:
                        if otro["carpeta"] == carpeta:
                            continue
                        if normalizar_nif(otro["nif"]) == nif_candidat_norm:
                            suggerit_carpeta, suggerit_nom = otro["carpeta"], otro["nombre"]
                            break
                    if suggerit_carpeta:
                        break

            if datos.get("num_factura") is not None:
                clave = (datos.get("nif_proveedor"), datos.get("num_factura"))
                # Piso 13T: .get(clave, []) en comptes de claves[clave] --
                # una fitxa que sigui ELLA MATEIXA descartada pot no
                # apareixer al mapa (si era l'unica amb aquesta clau, o
                # totes les que la comparteixen tambe estan descartades).
                otros = [n for n in claves.get(clave, []) if n != nombre]
                if otros:
                    motivos.append(f"factura duplicada: mismo proveedor+num_factura que {', '.join(otros)}")

            # Piso 13S: la regla vieja ("retención con cuota > 0 en rebudes
            # -> REVISAR, el llibre no tiene columna") queda RETIRADA -- ara
            # sumar.py SÍ té columna (% Ret./Retenció) tant a DESPESES com a
            # INGRESSOS. La retención bien cuadrada ya no es una anomalía,
            # es un dato mas. Se sustituye por el mismo xec aritmético que
            # verificar_retencion() (app.py/sumar.py/informe.py) ya hacia
            # como AVISO -- ahora tambien vive aqui como motivo de veritat,
            # igual en los dos flujos (el xec no tiene direccion).
            retencion_pct = datos.get("retencion_pct") or 0
            retencion_cuota = datos.get("retencion_cuota") or 0
            if retencion_pct or retencion_cuota:
                base_total_retencion = sum((l.get("base") or 0) for l in lineas)
                esperado_retencion = base_total_retencion * retencion_pct / 100
                if abs(esperado_retencion - retencion_cuota) > TOLERANCIA:
                    motivos.append(
                        f"retención no cuadra: {base_total_retencion} × {retencion_pct}% = "
                        f"{esperado_retencion:.2f}, pero retencion_cuota indica {retencion_cuota}"
                    )

            # Piso 13K: contrapart determinista per NIF, no per posicio del
            # camp -- "proveedor" es siempre quien EMITE, pero en rebudes
            # el cliente es SIEMPRE receptor (contrapart = proveedor) mientras
            # que en ingressos puede ser receptor (liquidacion) O emisor
            # (factura propia, "el giro" de mas arriba) -- ahi contrapart
            # pasa a ser el receptor. Se calcula para las DOS flujos (en
            # rebudes da lo mismo que proveedor de siempre) para que la
            # presentacion (sumar.py/informe.py/app.py) no tenga que saber
            # de flujos, solo lea este campo ya resuelto.
            nif_prov_norm = normalizar_nif(datos.get("nif_proveedor"))
            nif_rec_norm = normalizar_nif(datos.get("nif_receptor"))
            prov_es_client = nif_prov_norm is not None and nif_prov_norm == cliente_normalizado
            rec_es_client = nif_rec_norm is not None and nif_rec_norm == cliente_normalizado
            if prov_es_client and rec_es_client:
                contrapart_nom, contrapart_nif = None, None
                motivos.append("las dos partes son el cliente")
            elif prov_es_client:
                contrapart_nom, contrapart_nif = datos.get("receptor"), datos.get("nif_receptor")
            else:
                # rec_es_client, o cap dels dos coincideix (ambigu) -- per
                # defecte la contrapart es el proveedor, el mateix criteri
                # que ja es feia servir sempre abans d'aquest pis.
                contrapart_nom, contrapart_nif = datos.get("proveedor"), datos.get("nif_proveedor")

            estado = "OK" if not motivos else "REVISAR"
            if estado == "OK":
                ok += 1
            else:
                revisar += 1

            salida = dict(datos)
            salida["estado"] = estado
            salida["motivos"] = motivos
            salida["contrapart_nom"] = contrapart_nom
            salida["contrapart_nif"] = contrapart_nif
            salida["suggerit_carpeta"] = suggerit_carpeta
            salida["suggerit_nom"] = suggerit_nom
            if nombre in correcciones_aplicadas:
                salida["camps_corregits"] = correcciones_aplicadas[nombre]

            ruta_salida = os.path.join(carpeta_salida, nombre)
            with open(ruta_salida, "w", encoding="utf-8") as f:
                json.dump(salida, f, indent=2, ensure_ascii=False)

            if motivos:
                print(f"REVISAR: {nombre} -- {'; '.join(motivos)}")
            else:
                print(f"OK: {nombre}")

        print(f"{carpeta} / {etiqueta}: {ok} OK, {revisar} REVISAR con motivos, {ilegibles} ilegibles")
        ok_total[etiqueta] += ok
        revisar_total[etiqueta] += revisar
        ilegibles_total[etiqueta] += ilegibles

print(f"\nResumen total rebudes: {ok_total['rebudes']} OK, {revisar_total['rebudes']} REVISAR con motivos, {ilegibles_total['rebudes']} ilegibles")
print(f"Resumen total ingressos: {ok_total['ingressos']} OK, {revisar_total['ingressos']} REVISAR con motivos, {ilegibles_total['ingressos']} ilegibles")
