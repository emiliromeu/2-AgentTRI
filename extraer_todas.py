"""Extrae todas las facturas de entrada/ y guarda un JSON por factura.

Piso 2: recorre la carpeta, salta las que ya tienen JSON (idempotente,
no gasta creditos de mas), y cada factura falla en su propio try/except
sin tirar el lote entero abajo.
"""

import base64
import json
import os

from anthropic import Anthropic
from dotenv import load_dotenv

# Bloque a -- cargar la clave desde .env, sin imprimirla jamas (regla 6 de CLAUDE.md)
load_dotenv()
api_key = os.environ.get("ANTHROPIC_API_KEY")

carpeta_entrada = "clientes/penedes_languages/rebudes/entrada"
carpeta_salida = "clientes/penedes_languages/rebudes/extraidas"
os.makedirs(carpeta_salida, exist_ok=True)

# Bloque b -- listar solo los PDF, en orden estable
nombres_pdf = sorted(
    f for f in os.listdir(carpeta_entrada) if f.lower().endswith(".pdf")
)

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
- observaciones es texto libre para anotar cualquier cosa rara en la factura.
"""

cliente = Anthropic(api_key=api_key)

# Bloque c -- contadores para el resumen final
extraidas = 0
saltadas = 0
con_error = 0

# Bloque d -- el bucle: cada factura en su propio try/except
for nombre_pdf in nombres_pdf:
    ruta_pdf = os.path.join(carpeta_entrada, nombre_pdf)
    nombre_json = os.path.splitext(nombre_pdf)[0] + ".json"
    ruta_json = os.path.join(carpeta_salida, nombre_json)

    if os.path.exists(ruta_json):
        print(f"saltada: {nombre_pdf}")
        saltadas += 1
        continue

    try:
        with open(ruta_pdf, "rb") as f:
            pdf_base64 = base64.standard_b64encode(f.read()).decode("utf-8")

        respuesta = cliente.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1024,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "document",
                            "source": {
                                "type": "base64",
                                "media_type": "application/pdf",
                                "data": pdf_base64,
                            },
                        },
                        {"type": "text", "text": prompt},
                    ],
                }
            ],
        )

        texto = respuesta.content[0].text
        datos = json.loads(texto)

        with open(ruta_json, "w") as f:
            json.dump(datos, f, indent=2, ensure_ascii=False)

        print(f"extraída: {nombre_pdf}")
        extraidas += 1

    except Exception as e:
        print(f"AVISO: error en {nombre_pdf}: {e}")
        con_error += 1

# Bloque e -- resumen final
print(f"\nResumen: {extraidas} extraídas, {saltadas} saltadas, {con_error} con error")
