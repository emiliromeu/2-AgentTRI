"""Lee una factura en PDF y devuelve el JSON extraído por Claude.

Bala trazadora: procesa UNA sola factura, hardcodeada. Sin bucles,
sin validacion, sin Excel -- eso viene en otra sesion.
"""

import base64
import json
import os

from anthropic import Anthropic
from dotenv import load_dotenv

# Bloque a -- cargar la clave desde .env, sin imprimirla jamas (regla 6 de CLAUDE.md)
load_dotenv()
api_key = os.environ.get("ANTHROPIC_API_KEY")

# Bloque b -- ruta de la factura piloto, hardcodeada a proposito
ruta_pdf = "clientes/penedes_languages/rebudes/entrada/Romeu Olivella, Francesc Fra. 2619.pdf"

# Bloque c -- leer el PDF en bytes y codificarlo en base64 para mandarlo a la API
with open(ruta_pdf, "rb") as f:
    pdf_base64 = base64.standard_b64encode(f.read()).decode("utf-8")

# Bloque d -- llamar a la API pidiendo SOLO el JSON de la seccion 1 del esquema
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
- observaciones es texto libre para anotar cualquier cosa rara en la factura.
"""

cliente = Anthropic(api_key=api_key)

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

# Bloque e -- parsear el JSON; si falla, mostrar la respuesta cruda con un aviso
texto = respuesta.content[0].text

try:
    datos = json.loads(texto)
    print(json.dumps(datos, indent=2, ensure_ascii=False))
except json.JSONDecodeError:
    print("AVISO: la respuesta de Claude no es JSON valido. Respuesta cruda:")
    print(texto)
