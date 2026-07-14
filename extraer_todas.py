"""Extrae todas las facturas de entrada/ y guarda un JSON por factura.

Piso 2: recorre la carpeta, salta las que ya tienen JSON (idempotente,
no gasta creditos de mas), y cada factura falla en su propio try/except
sin tirar el lote entero abajo. Acepta PDF y tambien imagenes sueltas
(jpg/png), enviando cada una con el tipo de bloque que le toca.

Piso 4B: recorre todos los clientes de clientes.csv, no solo Penedes.

Piso 5: procesa tambien el flujo ingressos (liquidaciones de cooperativa
en apartados/ingressos/) ademas del flujo rebudes (entrada/), con
resumenes separados por flujo.

Piso 6B: el origen de cada flujo puede personalizarse por cliente
(RUTAS_ORIGEN_PERSONALIZADAS), para clientes con carpetas ya
organizadas antes de este pipeline. El flujo rebudes tambien lee
subcarpetas por proveedor, no solo entrada/ (listar_archivos_rebudes).

Piso 13K: regla explicita en el prompt -- proveedor/receptor son
siempre emisor/destinatario, nunca se intercambian aunque el
documento sea una venta del propio cliente. Blindaje a futuro, no
se re-extrae nada existente (las fichas ya extraidas se resuelven
por NIF en validar.py).

Piso 13J: regla explicita para num_factura -- captar el numero
COMPLETO con su serie/prefijo, nunca solo el numero suelto (una
causa real de falsos duplicados era comparar "14" contra "14" cuando
en realidad eran "Serie A-14" y "Serie B-14"). Blindaje a futuro,
sin re-extraer nada.
"""

import base64
import csv
import json
import os
from pathlib import Path

from anthropic import Anthropic
from dotenv import load_dotenv

# Piso 13B: ancla el cwd a la carpeta del propio script -- ver trocear.py.
RAIZ = Path(__file__).resolve().parent
os.chdir(RAIZ)

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
    with open("clientes/clientes.csv", encoding="utf-8") as f:
        return list(csv.DictReader(f))


CAMPOS_NUMERICOS_LINEA = {"tipo_iva", "base", "cuota"}
CAMPOS_NUMERICOS_TOP = {"total", "retencion_pct", "retencion_cuota"}


def a_numero(valor):
    """Piso 13Q: misma logica que a_numero() de validar.py (duplicada
    a proposito -- ninguna "maquina" importa otra, igual que
    leer_clientes()) -- nunca lanza. Defensa en el punto de carga: la
    API siempre ha devuelto tipos nativos hasta ahora, pero si algun
    dia devolviera un numero como texto ("1.234,56"), se guarda ya
    convertido en vez de dejar pasar un string a extraidas/*.json."""
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


def normalizar_numeros(datos):
    """Piso 13Q: aplica a_numero() SOLO a los campos que ya son string
    -- un int/float ya valido se deja intacto. Nunca lanza, nunca hace
    fallar la extraccion de una factura por esto."""
    for campo in CAMPOS_NUMERICOS_TOP:
        if isinstance(datos.get(campo), str):
            datos[campo] = a_numero(datos[campo])
    for linea in datos.get("lineas_iva") or []:
        for campo in CAMPOS_NUMERICOS_LINEA:
            if isinstance(linea.get(campo), str):
                linea[campo] = a_numero(linea[campo])
    return datos


SUBCARPETAS_RESERVADAS = {"extraidas", "validadas", "procesadas", "lotes_escaneados", "lotes_procesados"}


def listar_archivos_rebudes(carpeta_rebudes):
    """Recoge PDF/imagenes de entrada/ y de cualquier subcarpeta hermana
    que no sea una reservada del pipeline -- algunos clientes (davinstal)
    organizan sus facturas de compra por proveedor (Rebudes/biosca/,
    Rebudes/SALTOKI/...) en vez de dejarlas sueltas en entrada/."""
    rutas = []
    if not os.path.isdir(carpeta_rebudes):
        return rutas
    for nombre in sorted(os.listdir(carpeta_rebudes)):
        ruta = os.path.join(carpeta_rebudes, nombre)
        if not os.path.isdir(ruta) or nombre.lower() in SUBCARPETAS_RESERVADAS:
            continue
        for nombre_archivo in sorted(os.listdir(ruta)):
            if nombre_archivo.lower().endswith(EXTENSIONES_ACEPTADAS):
                rutas.append(os.path.join(ruta, nombre_archivo))
    return rutas


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
- El campo total es SIEMPRE el total de la factura: suma de bases más cuotas de IVA. NUNCA el líquido a percibir tras retenciones (llamado total acompte, líquido, o a percibir). La retención va exclusivamente en retencion_pct y retencion_cuota — jamás como línea de IVA, jamás descontada del total ni de las bases.
- Si el documento es un abono o factura rectificativa, conserva bases, cuotas y total en NEGATIVO.
- observaciones es texto libre para anotar cualquier cosa rara en la factura.
- proveedor = quien EMITE la factura (aparece en la cabecera/logo como emisor). receptor = a quien va dirigida (el destinatario). Esto es SIEMPRE así, incluso si el documento es una venta hecha por el propio cliente de esta gestoria (él sería el proveedor/emisor, su comprador el receptor). Nunca los intercambies aunque el documento sea una venta.
- num_factura: el número COMPLETO de la factura, incluyendo cualquier serie o prefijo que aparezca (ej. si la factura indica "Serie A - 14" o "2026/A-14", el campo debe ser ese texto completo, no solo "14").
"""

# (etiqueta, carpeta origen, carpeta destino) -- mismo esquema y mismas reglas
# para rebudes (facturas de compra) e ingressos (liquidaciones de cooperativa).
# El origen de "rebudes" es la carpeta rebudes/ entera (no solo entrada/),
# para poder mirar tambien sus subcarpetas por proveedor si las tiene.
FLUJOS = [
    ("rebudes", "rebudes", "rebudes/extraidas"),
    ("ingressos", "apartados/ingressos", "apartados/ingressos_extraidas"),
]

# Algunos clientes ya tenian su carpeta de facturas organizada antes de
# este pipeline -- no se ha movido nada, el codigo apunta ahi en vez de
# al origen por defecto de FLUJOS.
RUTAS_ORIGEN_PERSONALIZADAS = {
    ("davinstal", "ingressos"): "Emeses/davinstal",
}

# Bloque c -- contadores totales por flujo, sumados a lo largo de todos los clientes
extraidas_total = {"rebudes": 0, "ingressos": 0}
saltadas_total = {"rebudes": 0, "ingressos": 0}
con_error_total = {"rebudes": 0, "ingressos": 0}

# Piso 13S: processar.stop -- comprovat ENTRE factures (mai a mitges d'una
# crida a l'API). Aquest es el pas costos del pipeline (API de pagament),
# per aixo es la maquina amb el xec mes fi -- ejecutar.py nomes comprova
# ENTRE maquines, no te visibilitat de "quina factura toca ara".
RAIZ_STOP = Path(__file__).resolve().parent
RUTA_STOP = RAIZ_STOP / "processar.stop"
aturat_per_stop = False

# Bloque d -- bucle por cliente, luego por flujo, y dentro el bucle de siempre por archivo
for fila in leer_clientes():
    if aturat_per_stop:
        break
    carpeta = fila["carpeta"]

    for etiqueta, origen_rel, destino_rel in FLUJOS:
        if aturat_per_stop:
            break
        origen = RUTAS_ORIGEN_PERSONALIZADAS.get((carpeta, etiqueta), origen_rel)
        carpeta_entrada = f"clientes/{carpeta}/{origen}"
        carpeta_salida = f"clientes/{carpeta}/{destino_rel}"

        if not os.path.isdir(carpeta_entrada):
            continue

        os.makedirs(carpeta_salida, exist_ok=True)

        if etiqueta == "rebudes":
            rutas_archivo = listar_archivos_rebudes(carpeta_entrada)
        else:
            rutas_archivo = [
                os.path.join(carpeta_entrada, f)
                for f in sorted(os.listdir(carpeta_entrada))
                if f.lower().endswith(EXTENSIONES_ACEPTADAS)
            ]

        print(f"\n== {carpeta} / {etiqueta} ==")
        extraidas = 0
        saltadas = 0
        con_error = 0

        for indice, ruta_archivo in enumerate(rutas_archivo):
            if RUTA_STOP.exists():
                pendientes = len(rutas_archivo) - indice
                print(
                    f"\nATURAT per petició de l'usuari a {carpeta}/{etiqueta} -- "
                    f"{extraidas} processades aquí, {pendientes} pendents aquí (pot haver-hi més "
                    "en altres clients/fluxos encara no visitats). Reprèn amb Processar quan vulguis."
                )
                aturat_per_stop = True
                break

            nombre_archivo = os.path.basename(ruta_archivo)
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
                datos = normalizar_numeros(datos)

                with open(ruta_json, "w", encoding="utf-8") as f:
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
