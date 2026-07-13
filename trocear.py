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
from pathlib import Path

from anthropic import Anthropic
from dotenv import load_dotenv
from pypdf import PdfReader, PdfWriter

# Piso 13B: ancla el cwd a la carpeta del propio script -- todo el
# resto del archivo usa rutas relativas ("clientes/...") que solo son
# correctas si el proceso arranca ya posicionado aqui. Antes de esto
# dependia de quien lo lanzara (ejecutar.py, un .bat, una terminal a
# mano) dejara el cwd bien puesto; ahora se auto-corrige siempre.
RAIZ = Path(__file__).resolve().parent
os.chdir(RAIZ)

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

# Piso 13K: moll bessó per a lots de VENDES (abans nomes existia el de
# compres -- un lot de vendes hi acabava barrejat, bug confirmat en
# camp). Aqui el "correu rebut" no te sentit: son copies escanejades
# de documents que EL PROPI CLIENT ha emes, per aixo el vocabulari de
# tipus es mes curt (sense albara/liquidacio_ingreso, que nomes tenen
# sentit com a correspondencia rebuda) i el prompt es diferent.
TIPOS_VALIDOS_VENDES = {"factura", "abono", "ruido"}

# Mateixa convencio duplicada que a app.py/sumar.py/extraer_todas.py/
# informe.py (cap maquina es importable) -- si "Vendes" ho ignores,
# els arxius de Davinstal caurien on extraer_todas.py no els mira mai.
RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS = {"davinstal": "Emeses/davinstal"}


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


def construir_prompt_vendes(nombre_cliente, nif_cliente):
    """Piso 13K: mismo contrato de JSON que construir_prompt, pero el
    lote no es correspondencia recibida -- son copias escaneadas de
    facturas/abonos que {nombre_cliente} ha EMITIDO a sus propios
    clientes, asi que emisor_pista pasa a significar el destinatario
    (a quien se le vendio), no el proveedor."""
    return f"""Este PDF es un lote escaneado de copias de facturas/abonos
EMITIDOS (de venda) por {nombre_cliente} (NIF {nif_cliente}) a sus propios
clientes -- no es correspondencia recibida, son documentos que
{nombre_cliente} ha generado y guardado escaneados. Devuelve UNICAMENTE
este JSON, sin texto antes ni despues, sin bloques de markdown:

{{
  "documentos": [
    {{"tipo": "factura", "pagina_inicio": 1, "pagina_fin": 2, "emisor_pista": "..."}}
  ]
}}

Tipos validos: factura, abono, ruido.

Reglas de clasificacion:
- Las paginas de este PDF se numeran empezando en 1.
- Cada pagina pertenece a EXACTAMENTE un documento: no dejes huecos ni la repitas en dos documentos.
- pagina_inicio y pagina_fin son inclusivos.
- ruido: paginas en blanco, separadores, o cualquier pagina que no sea una factura/abono.
- abono: facturas rectificativas o notas de abono EMITIDAS por {nombre_cliente}.
- emisor_pista: nombre del client/comprador al que se le emitio el documento (el destinatario, NO {nombre_cliente}). Si no se distingue, usa "desconocido".
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


def procesar_flujo_lotes(carpeta, nombre_cliente, nif_cliente, ruta_lotes, ruta_procesados,
                          construir_prompt_fn, tipos_validos, destino_por_tipo):
    """Piso 13K: cos comu als dos molls (compres i vendes) -- abans
    nomes existia el de compres, escrit inline al bucle principal.
    Es parametritza amb el prompt, els tipus valids i el mapa de
    destins de cada flux perque la logica de trossejar/validar/desar
    (identica als dos) no es dupliqui. Retorna (ok, saltados, con_error,
    contador_tipos) d'aquest moll nomes -- el bucle principal ho suma."""
    ok = saltados = con_error = 0
    contador_tipos = {}

    if not os.path.isdir(ruta_lotes):
        return ok, saltados, con_error, contador_tipos
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
            saltados += 1
            continue

        try:
            reader = PdfReader(ruta_lote)
            total_paginas = len(reader.pages)
            prompt = construir_prompt_fn(nombre_cliente, nif_cliente)
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
                if doc.get("tipo") not in tipos_validos:
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
                con_error += 1
                continue

            # Cortar cada documento clasificado a su carpeta de destino
            for doc in documentos:
                tipo = doc["tipo"]
                contador_tipos[tipo] = contador_tipos.get(tipo, 0) + 1

                if tipo == "ruido":
                    print(f"ruido: {nombre_lote} p{doc['pagina_inicio']}-{doc['pagina_fin']}")
                    continue

                carpeta_destino = f"clientes/{carpeta}/{destino_por_tipo[tipo]}"
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
            ok += 1

        except Exception as e:
            print(f"AVISO: error procesando {nombre_lote}: {e}")
            con_error += 1

    return ok, saltados, con_error, contador_tipos


lotes_procesados_ok = 0
lotes_saltados = 0
lotes_con_error = 0
contador_tipos = {t: 0 for t in TIPOS_VALIDOS | TIPOS_VALIDOS_VENDES}

for fila in leer_clientes():
    carpeta = fila["carpeta"]
    origen_ingressos = RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS.get(carpeta, "apartados/ingressos")

    resultados_flujos = [
        procesar_flujo_lotes(
            carpeta, fila["nombre"], fila["nif"],
            f"clientes/{carpeta}/rebudes/lotes_escaneados",
            f"clientes/{carpeta}/rebudes/lotes_procesados",
            construir_prompt, TIPOS_VALIDOS, DESTINO_POR_TIPO,
        ),
        # Piso 13K: moll bessó de vendes -- factura/abono van a
        # l'origen d'ingressos d'aquest client (personalitzat per a
        # davinstal, com qualsevol Vendes solta).
        procesar_flujo_lotes(
            carpeta, fila["nombre"], fila["nif"],
            f"clientes/{carpeta}/apartados/lotes_vendes_escaneados",
            f"clientes/{carpeta}/apartados/lotes_vendes_procesados",
            construir_prompt_vendes, TIPOS_VALIDOS_VENDES,
            {"factura": origen_ingressos, "abono": origen_ingressos},
        ),
    ]
    for ok, saltados, con_error, tipos_flujo in resultados_flujos:
        lotes_procesados_ok += ok
        lotes_saltados += saltados
        lotes_con_error += con_error
        for tipo, n in tipos_flujo.items():
            contador_tipos[tipo] = contador_tipos.get(tipo, 0) + n

print(f"\nResumen: {lotes_procesados_ok} lotes procesados, {lotes_saltados} saltados, {lotes_con_error} con error")
print(f"Documentos cortados por tipo: {contador_tipos}")
