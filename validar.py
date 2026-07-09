"""Aplica la red de validacion de la seccion 3 del esquema a los JSON extraidos.

Piso 3: sin llamadas a la API, corre gratis. Lee extraidas/ (fixtures,
nunca se tocan) y escribe en validadas/ cada factura con su estado
(OK o REVISAR) y sus motivos.
"""

import json
import os

carpeta_entrada = "clientes/penedes_languages/rebudes/extraidas"
carpeta_salida = "clientes/penedes_languages/rebudes/validadas"
os.makedirs(carpeta_salida, exist_ok=True)

NIF_RECEPTOR_ESPERADO = "B66515529"  # en el piso 4 vendra de la tabla de clientes
TOLERANCIA = 0.02

CAMPOS_OBLIGATORIOS = [
    "proveedor", "nif_proveedor", "num_factura", "fecha_factura",
    "receptor", "nif_receptor", "lineas_iva", "total",
    "retencion_pct", "retencion_cuota", "exenta",
]

nombres_json = sorted(
    f for f in os.listdir(carpeta_entrada) if f.lower().endswith(".json")
)

# Primera pasada -- cargar cada JSON, detectar los ilegibles
facturas = []
ilegibles = 0
for nombre in nombres_json:
    ruta = os.path.join(carpeta_entrada, nombre)
    try:
        with open(ruta) as f:
            datos = json.load(f)
        facturas.append((nombre, datos))
    except (json.JSONDecodeError, OSError) as e:
        print(f"AVISO: {nombre} ilegible: {e}")
        ilegibles += 1

# Entre pasadas -- mapa de (nif_proveedor, num_factura) -> nombres de archivo,
# para detectar duplicados en todo el lote antes de validar ninguna factura
claves = {}
for nombre, datos in facturas:
    clave = (datos.get("nif_proveedor"), datos.get("num_factura"))
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

    nif_receptor = datos.get("nif_receptor")
    if nif_receptor is not None and nif_receptor != NIF_RECEPTOR_ESPERADO:
        motivos.append(
            f"nif_receptor no coincide: esperado {NIF_RECEPTOR_ESPERADO}, encontrado {nif_receptor}"
        )

    clave = (datos.get("nif_proveedor"), datos.get("num_factura"))
    otros = [n for n in claves[clave] if n != nombre]
    if otros:
        motivos.append(f"factura duplicada: mismo proveedor+num_factura que {', '.join(otros)}")

    retencion_cuota = datos.get("retencion_cuota")
    if retencion_cuota is not None and retencion_cuota > 0:
        motivos.append("retención con cuota > 0: el llibre no tiene columna para representarla")

    estado = "OK" if not motivos else "REVISAR"
    if estado == "OK":
        ok += 1
    else:
        revisar += 1

    salida = dict(datos)
    salida["estado"] = estado
    salida["motivos"] = motivos

    ruta_salida = os.path.join(carpeta_salida, nombre)
    with open(ruta_salida, "w") as f:
        json.dump(salida, f, indent=2, ensure_ascii=False)

    if motivos:
        print(f"REVISAR: {nombre} -- {'; '.join(motivos)}")
    else:
        print(f"OK: {nombre}")

print(f"\nResumen: {ok} OK, {revisar} REVISAR con motivos, {ilegibles} ilegibles")
