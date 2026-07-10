"""Extrae todas las facturas de entrada/ y guarda un JSON por factura.

Piso 2: recorre la carpeta, salta las que ya tienen JSON (idempotente,
no gasta creditos de mas), y cada factura falla en su propio try/except
sin tirar el lote entero abajo. Acepta PDF y tambien imagenes sueltas
(jpg/png), enviando cada una con el tipo de bloque que le toca.

Piso 4B: recorre todos los clientes de clientes.csv, no solo Penedes.

Piso 5: procesa tambien el flujo ingressos (liquidaciones de cooperativa
en apartados/ingressos/) ademas del flujo rebudes (entrada/), con
resumenes separados por flujo.
"""

import base64
import csv
import json
import os

from anthropic import Anthropic
from dotenv import load_dotenv

# Bloque a -- cargar la clave desde .env, sin imprimirla jamas (regla 6 de CLAUDE.md)
load_dotenv()
api_key = os.environ.get("ANTHROPIC_API_KEY")
cliente = Anthropic(api_key=api_key)

EXTENSIONES_ACEPTADAS = (".pdf", ".jpg", ".jpeg", ".png")

# tipo de bloque y media_type que espera la API segun la extension del archivo
MEDIA_TYPE_POR_EXTENSION = {
    ".pdf": ("document", "application/pdf"),
    ".jpg": ("image", "image/jpeg"),
    ".jpeg": ("image", "image/jpeg"),
    ".png": ("image", "image/png"),
}


def leer_clientes():
    with open("clientes/clientes.csv") as f:
        return list(csv.DictReader(f))


def limpiar_json(texto):
    """Quita las vallas ```json ... ``` si Claude las añade, pese a
    que el prompt le pide no hacerlo."""
    texto = texto.strip()
    if texto.startswith("```"):
        lineas = texto.split("\n")[1:]
        if lineas and lineas[-1].strip() == "```":
            lineas = lineas[:-1]
        texto = "\n".join(lineas)
    return texto.strip()


prompt = """Extrae los datos de esta factura y devuelve UNICAMENTE este JSON,
sin texto antes ni despues, sin bloques de markdown:

{
  "proveedor": "...",
  "nif_proveedor": "...",
  "num_factura": "...",
  "fecha_factura": "YYYY-MM-DD",
  "receptor": "...",
  "nif_receptor": "...",
  "lineas_iva": [
    { "tipo_iva": 0, "base": 0.0, "cuota": 0.0 }
  ],
  "total": 0.0,
  "retencion_pct": 0,
  "retencion_cuota": 0.0,
  "exenta": false,
  "observaciones": ""
}

Reglas:
- lineas_iva es una lista: si la factura tiene dos tipos de IVA, incluye dos elementos.
- exenta = true si la factura indica que esta exenta de IVA.
- Si un campo no aparece en el PDF, ponlo a null. Nunca inventes un valor.
- Si la factura no menciona retención, retencion_pct y retencion_cuota son 0. null se reserva para campos que deberían verse y no se pueden leer.
- Si el documento es un abono o factura rectificativa, conserva bases, cuotas y total en NEGATIVO.
- observaciones es texto libre para anotar cualquier cosa rara en la factura.
"""

# (etiqueta, carpeta origen, carpeta destino) -- mismo esquema y mismas reglas
# para rebudes (facturas de compra) e ingressos (liquidaciones de cooperativa)
FLUJOS = [
    ("rebudes", "rebudes/entrada", "rebudes/extraidas"),
    ("ingressos", "apartados/ingressos", "apartados/ingressos_extraidas"),
]

# Bloque c -- contadores totales por flujo, sumados a lo largo de todos los clientes
extraidas_total = {"rebudes": 0, "ingressos": 0}
saltadas_total = {"rebudes": 0, "ingressos": 0}
con_error_total = {"rebudes": 0, "ingressos": 0}

# Bloque d -- bucle por cliente, luego por flujo, y dentro el bucle de siempre por archivo
for fila in leer_clientes():
    carpeta = fila["carpeta"]

    for etiqueta, origen_rel, destino_rel in FLUJOS:
        carpeta_entrada = f"clientes/{carpeta}/{origen_rel}"
        carpeta_salida = f"clientes/{carpeta}/{destino_rel}"

        if not os.path.isdir(carpeta_entrada):
            continue

        os.makedirs(carpeta_salida, exist_ok=True)

        nombres_archivo = sorted(
            f for f in os.listdir(carpeta_entrada) if f.lower().endswith(EXTENSIONES_ACEPTADAS)
        )

        print(f"\n== {carpeta} / {etiqueta} ==")
        extraidas = 0
        saltadas = 0
        con_error = 0

        for nombre_archivo in nombres_archivo:
            ruta_archivo = os.path.join(carpeta_entrada, nombre_archivo)
            nombre_json = os.path.splitext(nombre_archivo)[0] + ".json"
            ruta_json = os.path.join(carpeta_salida, nombre_json)

            if os.path.exists(ruta_json):
                print(f"saltada: {nombre_archivo}")
                saltadas += 1
                continue

            try:
                extension = os.path.splitext(nombre_archivo)[1].lower()
                tipo_bloque, media_type = MEDIA_TYPE_POR_EXTENSION[extension]

                with open(ruta_archivo, "rb") as f:
                    archivo_base64 = base64.standard_b64encode(f.read()).decode("utf-8")

                respuesta = cliente.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=1024,
                    messages=[
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": tipo_bloque,
                                    "source": {
                                        "type": "base64",
                                        "media_type": media_type,
                                        "data": archivo_base64,
                                    },
                                },
                                {"type": "text", "text": prompt},
                            ],
                        }
                    ],
                )

                texto = respuesta.content[0].text
                datos = json.loads(limpiar_json(texto))

                with open(ruta_json, "w") as f:
                    json.dump(datos, f, indent=2, ensure_ascii=False)

                print(f"extraída: {nombre_archivo}")
                extraidas += 1

            except Exception as e:
                print(f"AVISO: error en {nombre_archivo}: {e}")
                con_error += 1

        print(f"{carpeta} / {etiqueta}: {extraidas} extraídas, {saltadas} saltadas, {con_error} con error")
        extraidas_total[etiqueta] += extraidas
        saltadas_total[etiqueta] += saltadas
        con_error_total[etiqueta] += con_error

# Bloque e -- resumen final, separado por flujo
print(f"\nResumen total rebudes: {extraidas_total['rebudes']} extraídas, {saltadas_total['rebudes']} saltadas, {con_error_total['rebudes']} con error")
print(f"Resumen total ingressos: {extraidas_total['ingressos']} extraídas, {saltadas_total['ingressos']} saltadas, {con_error_total['ingressos']} con error")
