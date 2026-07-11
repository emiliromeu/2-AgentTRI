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
"""

import csv
import json
import os
import re

TOLERANCIA = 0.02

CAMPOS_OBLIGATORIOS = [
    "proveedor", "nif_proveedor", "num_factura", "fecha_factura",
    "receptor", "nif_receptor", "lineas_iva", "total",
    "retencion_pct", "retencion_cuota", "exenta",
]


def leer_clientes():
    with open("clientes/clientes.csv") as f:
        return list(csv.DictReader(f))


def normalizar_nif(nif):
    """Quita guiones/espacios y pasa a mayúsculas, para que un NIF con
    formato distinto (ej. "37266020-V" vs "37266020V") no dispare un
    motivo falso."""
    if nif is None:
        return None
    return "".join(c for c in nif if c.isalnum()).upper()


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
    with open(ruta) as f:
        for fila in csv.DictReader(f):
            archivo = fila.get("arxiu")
            if archivo:
                correcciones.setdefault(archivo, []).append(fila)
    return correcciones


CAMPOS_NUMERICOS_LINEA = {"tipo_iva", "base", "cuota"}
CAMPOS_NUMERICOS_TOP = {"total", "retencion_pct", "retencion_cuota"}


def convertir_valor(campo, valor_texto):
    """Piso 11B: correccions.csv llega siempre como texto -- conversion
    minima al tipo que espera el esquema canonico."""
    if campo in CAMPOS_NUMERICOS_LINEA or campo in CAMPOS_NUMERICOS_TOP:
        return float(valor_texto)
    if campo == "exenta":
        return valor_texto.strip().lower() in ("true", "1", "si", "sí")
    return valor_texto


def aplicar_correcciones(datos, correcciones_archivo):
    """Piso 11B: aplica en memoria las correcciones de UN archivo sobre
    una COPIA de datos -- el JSON de extraidas/ nunca se toca. "camp"
    admite un nombre simple ("total", "nif_receptor"...) o
    "lineas_iva[N].subcampo" para corregir una linea de IVA concreta.
    La ficha corregida sigue despues por la red de validacion sin
    tocar, tal cual -- esta funcion NO valida nada, solo sustituye
    valores."""
    if not correcciones_archivo:
        return datos, []

    datos = dict(datos)
    if datos.get("lineas_iva") is not None:
        datos["lineas_iva"] = [dict(linea) for linea in datos["lineas_iva"]]

    camps_corregits = []
    for correccion in correcciones_archivo:
        camp = correccion["camp"]
        coincide_linea = re.match(r"lineas_iva\[(\d+)\]\.(\w+)", camp)
        if coincide_linea:
            indice, subcampo = int(coincide_linea.group(1)), coincide_linea.group(2)
            antic = datos["lineas_iva"][indice].get(subcampo)
            nuevo = convertir_valor(subcampo, correccion["valor_nou"])
            datos["lineas_iva"][indice][subcampo] = nuevo
        else:
            antic = datos.get(camp)
            nuevo = convertir_valor(camp, correccion["valor_nou"])
            datos[camp] = nuevo
        camps_corregits.append({
            "camp": camp,
            "antic": antic,
            "nou": nuevo,
            "qui": correccion.get("qui"),
            "data": correccion.get("data"),
        })
    return datos, camps_corregits


# (etiqueta, carpeta origen, carpeta destino) -- mismo esquema y mismas reglas
# para rebudes (facturas de compra) e ingressos (liquidaciones de cooperativa)
FLUJOS = [
    ("rebudes", "rebudes/extraidas", "rebudes/validadas"),
    ("ingressos", "apartados/ingressos_extraidas", "apartados/ingressos_validadas"),
]

ok_total = {"rebudes": 0, "ingressos": 0}
revisar_total = {"rebudes": 0, "ingressos": 0}
ilegibles_total = {"rebudes": 0, "ingressos": 0}

for fila in leer_clientes():
    carpeta = fila["carpeta"]
    nif_receptor_esperado = fila["nif"]

    # Piso 11B: una unica correccions.csv por cliente, comparte flujo
    # rebudes/ingressos igual que decisions.csv (Piso 9.2/11A).
    correcciones = cargar_correcciones(f"clientes/{carpeta}")
    correcciones_aplicadas = {}

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
                with open(ruta) as f:
                    datos = json.load(f)
                datos, camps_corregits = aplicar_correcciones(datos, correcciones.get(nombre, []))
                if camps_corregits:
                    correcciones_aplicadas[nombre] = camps_corregits
                facturas.append((nombre, datos))
            except (json.JSONDecodeError, OSError) as e:
                print(f"AVISO: {nombre} ilegible: {e}")
                ilegibles += 1

        # Entre pasadas -- mapa de (nif_proveedor, num_factura) -> nombres de archivo,
        # para detectar duplicados dentro de este cliente/flujo antes de validar
        # ninguna factura. (None, None) no cuenta -- dos facturas totalmente
        # vacías ya caen en REVISAR por sus propios campos, no son "duplicadas".
        claves = {}
        for nombre, datos in facturas:
            clave = (datos.get("nif_proveedor"), datos.get("num_factura"))
            if clave == (None, None):
                continue
            claves.setdefault(clave, []).append(nombre)

        # Segunda pasada -- validar cada factura
        ok = 0
        revisar = 0
        for nombre, datos in facturas:
            motivos = []

            for campo in CAMPOS_OBLIGATORIOS:
                if datos.get(campo) is None:
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

            total = datos.get("total")
            if total is not None:
                suma = sum((l.get("base") or 0) + (l.get("cuota") or 0) for l in lineas)
                if abs(suma - total) > TOLERANCIA:
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

            clave = (datos.get("nif_proveedor"), datos.get("num_factura"))
            otros = [n for n in claves[clave] if n != nombre] if clave != (None, None) else []
            if otros:
                motivos.append(f"factura duplicada: mismo proveedor+num_factura que {', '.join(otros)}")

            # La retención sin columna en el llibre solo es un problema en rebudes
            # (facturas de compra) -- en ingressos (liquidaciones de cooperativa)
            # la retención es normal y se suma, no se marca a revisar.
            retencion_cuota = datos.get("retencion_cuota")
            if etiqueta == "rebudes" and retencion_cuota is not None and retencion_cuota > 0:
                motivos.append("retención con cuota > 0: el llibre no tiene columna para representarla")

            estado = "OK" if not motivos else "REVISAR"
            if estado == "OK":
                ok += 1
            else:
                revisar += 1

            salida = dict(datos)
            salida["estado"] = estado
            salida["motivos"] = motivos
            if nombre in correcciones_aplicadas:
                salida["camps_corregits"] = correcciones_aplicadas[nombre]

            ruta_salida = os.path.join(carpeta_salida, nombre)
            with open(ruta_salida, "w") as f:
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
