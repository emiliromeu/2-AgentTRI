"""Trocea lotes escaneados multi-documento en facturas, abonos, albaranes
y liquidaciones sueltas.

Piso 4A: los clientes de papel mandan un PDF de 17-39 paginas con varios
documentos pegados. Este script le pide a Claude un manifiesto (que
paginas son que documento), lo valida en codigo, y corta cada trozo con
pypdf a su carpeta de destino -- para que extraer_todas.py y validar.py
puedan seguir trabajando sobre facturas individuales como ya hacen.

Piso 6A: cada lote procesado con exito guarda su manifiesto completo
en lotes_procesados/ -- registro permanente de que se decidio sobre
cada pagina (regla 9 de CLAUDE.md), que sumar.py lee para la hoja AVISOS.
"""

import base64
import csv
import io
import json
import os

from anthropic import Anthropic
from dotenv import load_dotenv
from pypdf import PdfReader, PdfWriter

load_dotenv()
api_key = os.environ.get("ANTHROPIC_API_KEY")
cliente = Anthropic(api_key=api_key)

LIMITE_BYTES = 18 * 1024 * 1024

TIPOS_VALIDOS = {"factura", "abono", "albara", "liquidacion_ingreso", "ruido"}

DESTINO_POR_TIPO = {
    "factura": "rebudes/entrada",
    "abono": "rebudes/entrada",
    "albara": "apartados/albarans",
    "liquidacion_ingreso": "apartados/ingressos",
}


def leer_clientes():
    with open("clientes/clientes.csv", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def limpiar_json(texto):
    """Quita las vallas ```json ... ``` si Claude las añade, pese a
    que el prompt le pide no hacerlo. Si encima hay prosa antes del
    JSON (Claude "pensando en voz alta" con un documento incompleto),
    recorta desde la primera { hasta la última }."""
    texto = texto.strip()
    if texto.startswith("```"):
        lineas = texto.split("\n")[1:]
        if lineas and lineas[-1].strip() == "```":
            lineas = lineas[:-1]
        texto = "\n".join(lineas)
    texto = texto.strip()
    if not texto.startswith("{") and not texto.startswith("["):
        inicio = texto.find("{")
        fin = texto.rfind("}")
        if inicio != -1 and fin != -1 and fin > inicio:
            texto = texto[inicio:fin + 1]
    return texto.strip()


def construir_prompt(nombre_cliente, nif_cliente):
    return f"""Este PDF es un lote escaneado de correspondencia recibida por
{nombre_cliente} (NIF {nif_cliente}). Contiene varios documentos distintos,
uno detras de otro. Devuelve UNICAMENTE este JSON, sin texto antes ni
despues, sin bloques de markdown:

{{
  "documentos": [
    {{"tipo": "factura", "pagina_inicio": 1, "pagina_fin": 2, "emisor_pista": "..."}}
  ]
}}

Tipos validos: factura, abono, albara, liquidacion_ingreso, ruido.

Reglas de clasificacion:
- Las paginas de este PDF se numeran empezando en 1.
- Cada pagina pertenece a EXACTAMENTE un documento: no dejes huecos ni la repitas en dos documentos.
- pagina_inicio y pagina_fin son inclusivos.
- ruido: condiciones generales, paginas legales, publicidad, o cualquier pagina que no sea un documento de gestion.
- liquidacion_ingreso: liquidaciones de una cooperativa donde {nombre_cliente} entrega producto (aparece como quien vende/entrega, no como comprador).
- abono: facturas rectificativas o notas de abono.
- emisor_pista: nombre o pista del proveedor/emisor que aparece en el documento. Si no se distingue, usa "desconocido".
"""


def pedir_manifiesto(pdf_bytes, prompt):
    pdf_base64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")
    respuesta = cliente.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
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
    return json.loads(limpiar_json(respuesta.content[0].text))["documentos"]


def escribir_sub_pdf(reader, pagina_inicio, pagina_fin, ruta_salida):
    writer = PdfWriter()
    for i in range(pagina_inicio - 1, pagina_fin):
        writer.add_page(reader.pages[i])
    with open(ruta_salida, "wb") as f:
        writer.write(f)


def nombre_seguro(texto):
    limpio = "".join(c if c.isalnum() or c in " -_" else "_" for c in texto)
    return limpio.strip()[:40] or "desconocido"


# --- bucle principal ---

lotes_procesados_ok = 0
lotes_saltados = 0
lotes_con_error = 0
contador_tipos = {t: 0 for t in TIPOS_VALIDOS}

for fila in leer_clientes():
    carpeta = fila["carpeta"]
    ruta_lotes = f"clientes/{carpeta}/rebudes/lotes_escaneados"
    if not os.path.isdir(ruta_lotes):
        continue

    ruta_procesados = f"clientes/{carpeta}/rebudes/lotes_procesados"
    os.makedirs(ruta_procesados, exist_ok=True)

    nombres_lote = sorted(
        f for f in os.listdir(ruta_lotes) if f.lower().endswith(".pdf")
    )

    for nombre_lote in nombres_lote:
        ruta_lote = os.path.join(ruta_lotes, nombre_lote)
        ruta_destino_procesado = os.path.join(ruta_procesados, nombre_lote)
        base_lote = os.path.splitext(nombre_lote)[0]

        if os.path.exists(ruta_destino_procesado):
            print(f"saltado: {nombre_lote}")
            lotes_saltados += 1
            continue

        try:
            reader = PdfReader(ruta_lote)
            total_paginas = len(reader.pages)
            prompt = construir_prompt(fila["nombre"], fila["nif"])
            peso = os.path.getsize(ruta_lote)

            if peso > LIMITE_BYTES:
                mitad = total_paginas // 2

                writer_a = PdfWriter()
                for i in range(0, mitad):
                    writer_a.add_page(reader.pages[i])
                buffer_a = io.BytesIO()
                writer_a.write(buffer_a)

                writer_b = PdfWriter()
                for i in range(mitad, total_paginas):
                    writer_b.add_page(reader.pages[i])
                buffer_b = io.BytesIO()
                writer_b.write(buffer_b)

                documentos_a = pedir_manifiesto(buffer_a.getvalue(), prompt)
                documentos_b = pedir_manifiesto(buffer_b.getvalue(), prompt)
                for doc in documentos_b:
                    doc["pagina_inicio"] += mitad
                    doc["pagina_fin"] += mitad
                documentos = documentos_a + documentos_b
            else:
                with open(ruta_lote, "rb") as f:
                    documentos = pedir_manifiesto(f.read(), prompt)

            # Validar el manifiesto: cada pagina debe pertenecer a EXACTAMENTE un documento
            dueno = [None] * (total_paginas + 1)  # indice 1..total_paginas
            errores = []
            for doc in documentos:
                if doc.get("tipo") not in TIPOS_VALIDOS:
                    errores.append(f"tipo desconocido '{doc.get('tipo')}' en páginas {doc.get('pagina_inicio')}-{doc.get('pagina_fin')}")
                    continue
                for p in range(doc["pagina_inicio"], doc["pagina_fin"] + 1):
                    if p < 1 or p > total_paginas:
                        errores.append(f"página {p} fuera de rango (el lote tiene {total_paginas})")
                        continue
                    if dueno[p] is not None:
                        errores.append(f"página {p} asignada dos veces")
                    dueno[p] = doc

            huecos = [p for p in range(1, total_paginas + 1) if dueno[p] is None]
            if huecos:
                errores.append(f"páginas sin asignar: {huecos}")

            if errores:
                print(f"AVISO: manifiesto inválido en {nombre_lote}: {'; '.join(errores)}")
                lotes_con_error += 1
                continue

            # Cortar cada documento clasificado a su carpeta de destino
            for doc in documentos:
                tipo = doc["tipo"]
                contador_tipos[tipo] += 1

                if tipo == "ruido":
                    print(f"ruido: {nombre_lote} p{doc['pagina_inicio']}-{doc['pagina_fin']}")
                    continue

                carpeta_destino = f"clientes/{carpeta}/{DESTINO_POR_TIPO[tipo]}"
                os.makedirs(carpeta_destino, exist_ok=True)

                emisor = nombre_seguro(doc.get("emisor_pista") or "desconocido")
                nombre_salida = f"{base_lote}_p{doc['pagina_inicio']:03d}-{doc['pagina_fin']:03d}_{emisor}.pdf"
                ruta_salida = os.path.join(carpeta_destino, nombre_salida)

                escribir_sub_pdf(reader, doc["pagina_inicio"], doc["pagina_fin"], ruta_salida)
                print(f"cortado ({tipo}): {nombre_salida}")

            # Guardar el manifiesto completo (regla 9: nada muere en silencio,
            # tampoco el enrutado -- sumar.py lo lee para la hoja AVISOS)
            ruta_manifiesto = os.path.join(ruta_procesados, base_lote + "_manifiesto.json")
            with open(ruta_manifiesto, "w", encoding="utf-8") as f:
                json.dump(documentos, f, indent=2, ensure_ascii=False)

            os.rename(ruta_lote, ruta_destino_procesado)
            print(f"procesado: {nombre_lote}")
            lotes_procesados_ok += 1

        except Exception as e:
            print(f"AVISO: error procesando {nombre_lote}: {e}")
            lotes_con_error += 1

print(f"\nResumen: {lotes_procesados_ok} lotes procesados, {lotes_saltados} saltados, {lotes_con_error} con error")
print(f"Documentos cortados por tipo: {contador_tipos}")
