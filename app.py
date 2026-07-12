"""App Streamlit -- closca sobre la fabrica. Piso 10.1.

Nomes coloca arxius on ja els busca extraer_todas.py/trocear.py, i
llanca ejecutar.py com a subproces -- exactament com ho faria Emili a
ma. No calcula ni valida res: la logica de negoci viu nomes a les
cinc maquines (trocear, extraer_todas, validar, sumar, informe), que
aquesta app no toca ni importa (totes executen el seu bucle principal
en carregar-se -- convencio des del piso 2).

Piso 10.2: verificada clic a clic en navegador real (Playwright, la
extensio de Chrome no connectava). Tres bugs reals arreglats: (1) la
carpeta de "Nou client" ara es deriva en viu i no s'edita a ma --
calia treure el camp del st.form, que no reexecuta en cada tecla;
(2) "Obrir informe"/"Obrir Excel" no feien res -- Chromium bloqueja
la navegacio file:// des d'un origen http://localhost (confirmat amb
la consola real), aixi que ara s'obren amb `open` (macOS) des del
costat Python, en el seu lloc real; (3) la pantalla final nomes
oferia un client -- ara ensenya tots (ejecutar.py els processa tots
de cop). De pas: totes les rutes s'ancoren a la carpeta del propi
app.py (RAIZ_PROYECTO), no al directori des d'on es llanci streamlit.

Piso 10.3: "Arxivar client" -- filosofia A3, mai destruir. Nomes
shutil.move (mai shutil.rmtree) a arxivats/<carpeta>_<data_hora>/,
amb recompte real llegit de disc i confirmacio escrivint el nom
exacte. Treu la fila de clientes.csv i deixa rastre a
arxivats/registre_arxivat.csv -- recuperable a ma en qualsevol
moment.

Piso 10.5: identitat visual Olivella. Colors del tema (config.toml)
extrets amb PIL del logo real, no a ull. st.logo natiu + un unic
bloc CSS_A3 nomes per allo que config.toml no cobreix (l'stack de
tipografia -apple-system exacte i la subtil ombra de les targetes).
Cap logica canvia -- nomes presentacio.

Piso 11A: vista "Revisió" -- la primera pantalla que ESCRIU
decisions.csv (abans nomes ho llegien sumar.py/informe.py, Piso 9.2).
Aprovar/Descartar sobre CADA fitxa (OK o REVISAR -- simetria: es pot
descartar un OK amb nota, sumar_bloque ja ho suporta des del Piso 9.2
sense cap canvi). Errors accionables: "Retirar" mou (mai esborra) a
clientes/<carpeta>/errors_retirats/, fora de rebudes/apartados perque
extraer_todas.py no el torni a veure. Boto RECALCULAR encadena nomes
sumar.py + informe.py (sense API). La fabrica no es toca.

Piso 11A-fix: reproduit en directe el bug que Emili patia -- un
st.text_input solt nomes aplica el seu valor amb blur/Enter, mai
tecla a tecla, aixi que Aprovar/Descartar/Retirar (tots amb
disabled=not qui) es quedaven grisos encara que ja hagues escrit el
seu nom. "Qui revisa?" ara es una porta bloquejant amb st.form (que
si aplica el valor en polsar el seu boto, sense blur) -- no es veu
cap targeta fins que es confirma. Cada targeta te la nota+botons
dins del seu propi st.form pel mateix motiu. Cap boto es desactiva
ja mai: si falta la nota per descartar, el clic es processa sempre i
respon amb un st.error visible -- fallada sorollosa, mai muda.

Piso 11B: "Corregir camps" -- unica capa d'escriptura que toca la
xarxa de validacio, i ho fa nomes a traves de validar.py (cirugia
minima, documentada alli). Aqui nomes s'escriu correccions.csv (una
fila per camp que de veritat canvia) -- mai s'edita cap JSON
d'extraidas/. RECALCULAR ara encadena validar.py abans de sumar.py/
informe.py perque la correccio torni a passar tot l'examen.

Piso 12B: les targetes de Revisió passen a portar els mateixos colors
que informe.py (verd .ok / taronja .revisar / vermell .error) via
st.container(key=...) + CSS_A3 -- cap canvi de logica, nomes la key
del container i tres regles CSS mes. Vista previa del PDF original:
primera pagina renderitzada com a imatge amb PyMuPDF (fitz), mantenint
"Obrir original" per accedir-hi sencer; si no es pot renderitzar, cau
al comportament d'abans (nomes el boto) -- mai un error sense explicar.

Piso 11C: diagnosticat (navegador real contra davinstal, 240 targetes)
que un sol clic a Aprovar costava ~33s -- un st.rerun() sencer
reexecutava tota la pantalla, re-renderitzant les 240 vistes previes
de PDF sense cap cache (fins i tot les de dins d'un expander tancat,
que Streamlit igualment munta) i saltant el scroll a dalt. Arreglat de
tres maneres alhora: (1) tarjeta_revisio/tarjeta_error passen a
@st.fragment -- un clic reexecuta NOMES aquesta targeta
(st.rerun(scope="fragment")), mai la resta; cada targeta relegeix
decisions.csv ella mateixa per saber si ja esta decidida, ja que un
rerun de fragment no torna a executar el bucle exterior que la crida.
(2) previsualizar_pdf cacheada per (ruta, mtime) amb st.cache_data --
mai mes es recalcula si l'arxiu no ha canviat. (3) paginacio de 15 en
15 a Pendents/Fitxes OK/Errors, que acota el cost de qualsevol futur
rerun complet. Selecció múltiple amb intenció: checkbox per targeta
(fora del fragment, perque el comptador de seleccionades es vegi
sempre al dia) + barra "N seleccionades" que NOMES apareix si hi ha
selecció -- a Pendents amb nota compartida opcional, a Errors amb nota
obligatoria i un avis fix sobre arxius corromputs per iCloud. Els
comptadors de dalt (N pendents...) ja no es poden refrescar sols quan
es decideix una targeta -- la pantalla ho avisa sempre de forma fixa
amb un boto "Actualitzar comptadors", mai fingint un numero en viu.
"""

import csv
import io
import json
import os
import shutil
import subprocess
import sys
from datetime import datetime

import fitz
import streamlit as st
from PIL import Image
from pillow_heif import register_heif_opener
from pypdf import PdfReader

register_heif_opener()

# Piso 10.2: ancorat a la carpeta d'aquest arxiu, no al directori des
# d'on es llanci `streamlit run` -- abans "clientes/..." nomes
# funcionava si s'arrencava amb el cwd a l'arrel del projecte.
RAIZ_PROYECTO = os.path.dirname(os.path.abspath(__file__))


def ruta_proyecto(*partes):
    return os.path.join(RAIZ_PROYECTO, *partes)


RUTA_CLIENTES_CSV = ruta_proyecto("clientes", "clientes.csv")
CAMPOS_CLIENTES_CSV = ["nif", "nombre", "carpeta"]
EXTENSIONES_PERMITIDAS = ["pdf", "jpg", "jpeg", "png", "heic", "heif"]
EXTENSIONES_HEIC = (".heic", ".heif")
# Piso 13D: valor intern -> text mostrat al radio "Afegir factures".
DESTINOS_AFEGIR = {
    "Compres": "Compres",
    "Vendes": "Vendes",
    "Lot": "Lot d'escàner (diverses factures en un sol PDF)",
}

# Igual que en extraer_todas.py/sumar.py/informe.py -- convencion duplicada
# a proposito (ninguna maquina es importable). Si "Vendes" ignorase esto,
# los archivos de Davinstal caerian en una carpeta que extraer_todas.py
# no mira nunca.
RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS = {"davinstal": "Emeses/davinstal"}

TABLA_LETRA_DNI = "TRWAGMYFPDXBNJZSQVHLCKE"


def validar_nif(nif):
    """Copia identica a sumar.py/informe.py (Piso 9.1) -- no importable
    de alli (top-level code se ejecutaria al cargarlas)."""
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


def normalizar_nif(nif):
    return "".join(c for c in (nif or "") if c.isalnum()).upper()


def slug(texto):
    equivalencias = str.maketrans("àáèéíòóúüçñ", "aaeeioouucn")
    limpio = "".join(c if c.isalnum() else "_" for c in texto.lower().translate(equivalencias))
    while "__" in limpio:
        limpio = limpio.replace("__", "_")
    return limpio.strip("_")


def leer_clientes():
    if not os.path.exists(RUTA_CLIENTES_CSV):
        return []
    with open(RUTA_CLIENTES_CSV, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def escribir_clientes(filas):
    """Reescriu clientes.csv sencer -- l'usen anadir_cliente (afegir
    una fila) i arxivar_cliente (Piso 10.3, treure'n una).

    Piso 13B: open(..., "w") mai crea el directori pare -- si
    clientes/ no existis (PC nou sense el USB copiat), aixo era
    exactament el punt real on petava el bug ("Errno 2"). La guardia
    de dades absents (mes avall) ja hauria de bloquejar abans
    d'arribar aqui, pero aquest makedirs tanca el forat del tot."""
    os.makedirs(os.path.dirname(RUTA_CLIENTES_CSV), exist_ok=True)
    with open(RUTA_CLIENTES_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CAMPOS_CLIENTES_CSV)
        writer.writeheader()
        writer.writerows(filas)


def registrar_instalacio(cami):
    """Piso 13C: rastre escrit de quina bifurcació es va triar quan
    faltava clientes/ -- decisió humana, mai en silenci (regla 4).
    Es crida sempre DESPRES que clientes/ ja existeixi (als dos
    camins), aixi que l'open("a") mai falla per directori absent."""
    ruta = ruta_proyecto("clientes", "instalacio.log")
    with open(ruta, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {cami}\n")


def anadir_cliente(nif, nombre, carpeta):
    filas = leer_clientes()
    filas.append({"nif": nif, "nombre": nombre, "carpeta": carpeta})
    escribir_clientes(filas)
    os.makedirs(ruta_proyecto("clientes", carpeta, "rebudes", "entrada"), exist_ok=True)
    # Piso 13D: lotes_escaneados/ (entrada de trocear.py) i
    # lotes_procesados/ (on deixa el PDF sencer + manifest un cop
    # trocejat) -- trocear.py ja les tolera absents, pero "Nou client"
    # deixa l'arbre complet des del primer moment.
    os.makedirs(ruta_proyecto("clientes", carpeta, "rebudes", "lotes_escaneados"), exist_ok=True)
    os.makedirs(ruta_proyecto("clientes", carpeta, "rebudes", "lotes_procesados"), exist_ok=True)
    os.makedirs(ruta_proyecto("clientes", carpeta, "apartados", "ingressos"), exist_ok=True)


def contar_archivos_cliente(carpeta):
    """Piso 10.3: recompte real de disc per al panel de confirmacio --
    nomes comptar per extensio, cap interpretacio de contingut (igual
    d'"embolcall" que guardar_archivo())."""
    base = ruta_proyecto("clientes", carpeta)
    originals = 0
    fitxes = 0
    for _, _, archivos in os.walk(base):
        for nombre_archivo in archivos:
            extension = os.path.splitext(nombre_archivo)[1].lower()
            if extension in EXTENSIONES_HEIC or extension in (".pdf", ".jpg", ".jpeg", ".png"):
                originals += 1
            elif extension == ".json":
                fitxes += 1
    informes = sum(
        os.path.exists(os.path.join(base, nombre_archivo))
        for nombre_archivo in ("informe_2026.html", "sumatorios_2026.xlsx")
    )
    return {"originals": originals, "fitxes": fitxes, "informes": informes}


def asegurar_gitignore_arxivats():
    """Piso 10.3: 'arxivats/' no s'ha de versionar (son dades fiscals
    reals, mateix criteri que clientes/). S'afegeix nomes si encara
    no hi es -- sense duplicar ni trencar el fitxer existent."""
    ruta_gitignore = ruta_proyecto(".gitignore")
    lineas = []
    if os.path.exists(ruta_gitignore):
        with open(ruta_gitignore, encoding="utf-8") as f:
            lineas = f.read().splitlines()
    if "arxivats/" not in lineas:
        with open(ruta_gitignore, "a", encoding="utf-8") as f:
            f.write("arxivats/\n")


def arxivar_cliente(fila):
    """Piso 10.3: filosofia A3 -- mai shutil.rmtree, nomes shutil.move.
    Mou clientes/<carpeta>/ sencera a arxivats/<carpeta>_<data_hora>/,
    treu la fila de clientes.csv, i deixa rastre a
    arxivats/registre_arxivat.csv. Res s'esborra mai."""
    carpeta = fila["carpeta"]
    origen = ruta_proyecto("clientes", carpeta)
    marca = datetime.now().strftime("%Y%m%d_%H%M%S")
    carpeta_arxivats = ruta_proyecto("arxivats")
    destino = os.path.join(carpeta_arxivats, f"{carpeta}_{marca}")

    if os.path.exists(destino):
        raise RuntimeError(f"La carpeta destí '{destino}' ja existeix -- aturat, no s'ha mogut res.")

    os.makedirs(carpeta_arxivats, exist_ok=True)
    asegurar_gitignore_arxivats()

    n_archivos = sum(len(archivos) for _, _, archivos in os.walk(origen))
    shutil.move(origen, destino)

    filas_restantes = [f for f in leer_clientes() if f["carpeta"] != carpeta]
    escribir_clientes(filas_restantes)

    ruta_registro = os.path.join(carpeta_arxivats, "registre_arxivat.csv")
    escribir_cabecera = not os.path.exists(ruta_registro)
    with open(ruta_registro, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if escribir_cabecera:
            writer.writerow(["fecha", "cliente", "nif", "carpeta_origen", "carpeta_destino", "n_archivos"])
        writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            fila["nombre"], fila["nif"], origen, destino, n_archivos,
        ])

    return destino, n_archivos


def ruta_destino_factures(carpeta, destino):
    if destino == "Compres":
        return ruta_proyecto("clientes", carpeta, "rebudes", "entrada")
    if destino == "Lot":
        return ruta_proyecto("clientes", carpeta, "rebudes", "lotes_escaneados")
    origen_ingressos = RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS.get(carpeta, "apartados/ingressos")
    return os.path.join(ruta_proyecto("clientes", carpeta), *origen_ingressos.split("/"))


def guardar_archivo(archivo_subido, carpeta_destino):
    """Guarda un archivo subido; .heic/.heif se convierten a .jpg con
    pillow-heif. Si el nombre ya existe, anade un sufijo numerico en
    vez de sobrescribir (regla 9: nada en silencio)."""
    os.makedirs(carpeta_destino, exist_ok=True)
    nombre_base, extension = os.path.splitext(archivo_subido.name)
    es_heic = extension.lower() in EXTENSIONES_HEIC
    extension_final = ".jpg" if es_heic else extension

    nombre_final = nombre_base + extension_final
    ruta_final = os.path.join(carpeta_destino, nombre_final)
    contador = 2
    while os.path.exists(ruta_final):
        nombre_final = f"{nombre_base}_{contador}{extension_final}"
        ruta_final = os.path.join(carpeta_destino, nombre_final)
        contador += 1

    if es_heic:
        imagen = Image.open(io.BytesIO(archivo_subido.getvalue())).convert("RGB")
        imagen.save(ruta_final, "JPEG")
    else:
        with open(ruta_final, "wb") as f:
            f.write(archivo_subido.getvalue())

    return nombre_final


def guardar_y_reportar(archivos, carpeta_destino):
    """Piso 13D: factoritzat perque "Afegir factures" el crida des de
    tres llocs (desar normal, i els dos camins de la guardia de lots
    despistats) sense repetir el mateix bloc tres cops."""
    nombres_finales = [guardar_archivo(a, carpeta_destino) for a in archivos]
    st.success(f"S'han desat {len(nombres_finales)} arxius a `{carpeta_destino}`:")
    for nombre in nombres_finales:
        st.write(f"- {nombre}")


@st.cache_data(show_spinner=False)
def previsualizar_pdf(ruta_pdf, mtime):
    """Piso 12B: primera pagina del PDF renderitzada com a imatge en
    memoria (mai es desa a disc). Si el PDF no es pot obrir o
    renderitzar (mateix esperit que archivo_corrupto de sumar.py/
    informe.py), retorna None -- qui la crida cau llavors al boto
    "Obrir original" sol, mai un error sense explicar.

    Piso 11C: cacheada per (ruta, mtime) -- abans es renderitzava de
    zero a CADA rerun, per CADA targeta (el cost dominant que feia
    triguar ~33s un rerun complet amb 240 targetes). mtime a la clau
    de cache perque si mai es resuja l'original la cache s'invalida
    sola, no cal netejar-la a ma."""
    try:
        with fitz.open(ruta_pdf) as doc:
            return doc[0].get_pixmap(dpi=100).tobytes("png")
    except Exception:
        return None


def boton_obrir(etiqueta, ruta_absoluta, key):
    """Piso 10.2: st.link_button con file:// no funciona -- Chromium
    bloquea la navegacion a file:// desde un origen http://localhost
    (confirmado con la consola real: "Not allowed to load local
    resource"). En vez de eso, se abre del lado Python, directamente
    en su sitio -- para que los enlaces relativos del informe a sus
    originales sigan vivos.

    Piso 13: multiplataforma -- Windows no tiene el comando `open` de
    macOS. os.startfile() (Windows) no devuelve returncode y puede
    lanzar OSError (ej. sin asociacion de archivo) -- capturado
    explicitamente, nunca falla en silencio (regla 4/10)."""
    existe = os.path.exists(ruta_absoluta)
    if st.button(etiqueta, disabled=not existe, key=key):
        try:
            if sys.platform == "win32":
                os.startfile(ruta_absoluta)
            elif sys.platform == "darwin":
                resultado = subprocess.run(["open", ruta_absoluta])
                if resultado.returncode != 0:
                    st.error(f"No s'ha pogut obrir {ruta_absoluta} (codi {resultado.returncode}).")
            else:
                st.error(f"Plataforma no suportada per obrir arxius: {sys.platform}.")
        except OSError as e:
            st.error(f"No s'ha pogut obrir {ruta_absoluta}: {e}")


def tarjeta_cliente(fila, prefijo, lineas_log=None):
    """Piso 10.2: componente reutilizado en la vista Clients y en el
    resumen final de Processar (antes cada uno tenia su propio bloque
    ad-hoc, y el de Processar solo mostraba un cliente)."""
    carpeta = fila["carpeta"]
    ruta_informe = ruta_proyecto("clientes", carpeta, "informe_2026.html")
    ruta_excel = ruta_proyecto("clientes", carpeta, "sumatorios_2026.xlsx")
    with st.container(border=True):
        col1, col2, col3 = st.columns([3, 1, 1])
        with col1:
            st.markdown(f"**{fila['nombre']}**  \nNIF {fila['nif']} · carpeta `{carpeta}`")
            if os.path.exists(ruta_informe):
                mtime = datetime.fromtimestamp(os.path.getmtime(ruta_informe))
                st.caption(f"Últim run: {mtime.strftime('%d/%m/%Y %H:%M')}")
            else:
                st.caption("Encara no s'ha processat.")
        with col2:
            boton_obrir("Obrir informe", ruta_informe, key=f"{prefijo}_informe_{carpeta}")
        with col3:
            boton_obrir("Obrir Excel", ruta_excel, key=f"{prefijo}_excel_{carpeta}")
        if lineas_log is not None:
            with st.expander("Registre d'aquest client"):
                st.code("\n".join(lineas_log) if lineas_log else "Sense línies per a aquest client.")

        # Piso 10.3: opcio discreta -- el boto disparador d'un popover
        # ja es secundari/gris per defecte, no cal cap "boto vermell".
        with st.popover("Arxivar client"):
            conteo = contar_archivos_cliente(carpeta)
            st.markdown(f"**{fila['nombre']}** · NIF {fila['nif']} · carpeta `{carpeta}`")
            st.write(
                f"{conteo['originals']} documents originals, {conteo['fitxes']} fitxes, "
                f"{conteo['informes']} informes."
            )
            st.caption(
                "Res s'esborra -- es mou sencer a `arxivats/`, recuperable manualment."
            )
            confirmacion = st.text_input(
                f"Escriu \"{fila['nombre']}\" per confirmar", key=f"{prefijo}_arxivar_nom_{carpeta}"
            )
            if st.button(
                "Arxivar definitivament",
                key=f"{prefijo}_arxivar_boto_{carpeta}",
                disabled=confirmacion != fila["nombre"],
            ):
                try:
                    destino, n_archivos = arxivar_cliente(fila)
                except RuntimeError as e:
                    st.error(str(e))
                else:
                    st.success(f"Client arxivat a `{destino}`. Recuperable manualment.")
                    st.rerun()


# =======================================================================
# Piso 11A: vista "Revisio" -- duplicado de solo-lectura de informe.py
# (ninguna maquina es importable, misma convencion desde el piso 2) mas
# las primeras funciones que ESCRIBEN decisions.csv (hasta ahora solo lo
# leian sumar.py/informe.py).
# =======================================================================

EXTENSIONES_ORIGINAL = (".pdf", ".jpg", ".jpeg", ".png")
EXTENSIONES_IMAGEN = (".jpg", ".jpeg", ".png")
SUBCARPETAS_RESERVADAS = {"extraidas", "validadas", "procesadas", "lotes_escaneados", "lotes_procesados"}
SUBCARPETAS_NO_ORIGINALES = {"extraidas", "validadas", "lotes_escaneados", "lotes_procesados"}
CAMPOS_DECISIONS_CSV = ["archivo", "accion", "nota", "qui", "data"]
CAMPOS_CORRECCIONS_CSV = ["arxiu", "camp", "valor_antic", "valor_nou", "motiu", "qui", "data"]
CAMPOS_CORREGIBLES_TOP = [
    "proveedor", "nif_proveedor", "num_factura", "fecha_factura",
    "receptor", "nif_receptor", "total", "retencion_pct",
    "retencion_cuota", "exenta", "observaciones",
]

# Igual que en sumar.py/informe.py: traduccion de los motivos de
# validar.py por sustitucion de palabras fijas (validar.py no se toca).
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


def verificar_retencion(datos):
    """Igual que en informe.py/sumar.py: True si no hay retencion o si
    cuadra (tolerancia 0,02). No cambia ningun estado ni suma."""
    retencion_pct = datos.get("retencion_pct") or 0
    retencion_cuota = datos.get("retencion_cuota") or 0
    if not retencion_pct and not retencion_cuota:
        return True
    suma_base = sum((l.get("base") or 0) for l in (datos.get("lineas_iva") or []))
    esperado = suma_base * retencion_pct / 100
    return abs(esperado - retencion_cuota) <= 0.02


def avisos_verificacion_ficha(datos):
    """Igual criterio que avisos_verificacion() de informe.py, pero para
    UNA ficha (aqui interesa mostrarlo en su propia tarjeta, no listarlo
    aparte para todo un cliente)."""
    avisos = []
    for campo, etiqueta in [("nif_proveedor", "proveïdor"), ("nif_receptor", "receptor")]:
        valor = datos.get(campo)
        if validar_nif(valor) is False:
            avisos.append(f"NIF no supera la validació de la lletra ({etiqueta}: {valor})")
    if not verificar_retencion(datos):
        avisos.append("la retenció calculada no quadra amb la retenció indicada")
    return avisos


MOTIVO_ERROR_GENERICO = (
    "No s'ha pogut generar la fitxa — l'arxiu és present però no hi ha extracció. "
    "Cal revisar l'escaneig o tornar-ho a intentar."
)
MOTIVO_ERROR_CORRUPTE = (
    "L'arxiu original és present però buit o corromput — probablement un "
    "problema de sincronització. Cal tornar-lo a sincronitzar o demanar-lo de nou."
)


def archivo_corrupto(ruta):
    """Igual que en informe.py (Piso 9.3): True si la ruta existe pero
    esta vacia, o es un PDF sense trailer %%EOF en l'ultim KB."""
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
    return MOTIVO_ERROR_CORRUPTE if archivo_corrupto(ruta) else MOTIVO_ERROR_GENERICO


def cargar_validadas(carpeta):
    facturas = []
    if not os.path.isdir(carpeta):
        return facturas
    for nombre in sorted(os.listdir(carpeta)):
        if not nombre.lower().endswith(".json"):
            continue
        with open(os.path.join(carpeta, nombre), encoding="utf-8") as f:
            facturas.append((nombre, json.load(f)))
    return facturas


def cargar_decisiones(carpeta_cliente):
    """Igual que en sumar.py/informe.py: lee decisions.csv si existe.
    Sin archivo, o vacio, devuelve {}."""
    ruta = os.path.join(carpeta_cliente, "decisions.csv")
    decisiones = {}
    if not os.path.exists(ruta):
        return decisiones
    with open(ruta, encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            archivo = fila.get("archivo")
            if archivo:
                decisiones[archivo] = fila
    return decisiones


def encontrar_original(carpeta_origen, nombre_json):
    """Igual que en informe.py/sumar.py: busca directamente, y si no, en
    subcarpetas hermanas no reservadas. "procesadas" SI se busca (un
    original ya movido ahi tras procesarse sigue siendo su ubicacion
    legitima)."""
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
    """Igual que en informe.py/extraer_todas.py: entrada/ y cualquier
    subcarpeta hermana no reservada (proveedores por subcarpeta)."""
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


def detectar_errores(rutas_presentes, carpeta_extraidas):
    """Igual que en informe.py/sumar.py: archivos presentes sin ficha
    extraida -- unica senal disponible."""
    extraidos = set()
    if os.path.isdir(carpeta_extraidas):
        extraidos = {
            os.path.splitext(f)[0] for f in os.listdir(carpeta_extraidas) if f.lower().endswith(".json")
        }
    return [r for r in rutas_presentes if os.path.splitext(os.path.basename(r))[0] not in extraidos]


def escribir_decision(carpeta_cliente, archivo, accion, nota, qui, data=None):
    """Piso 11A: primera funcion que ESCRIBE decisions.csv -- hasta
    ahora sumar.py/informe.py solo lo leian (Piso 9.2, rellenado a mano
    por Emili). Si ya habia una decision para este archivo, la
    sustituye -- no se acumulan filas contradictorias del mismo archivo.

    Piso 11C: 'data' opcional -- una accion en lote captura UNA marca
    de tiempo antes del bucle y la pasa a cada llamada, para que las
    N filas del mismo lote compartan exactamente la misma data (no
    unos milisegundos distintos cada una)."""
    ruta = os.path.join(carpeta_cliente, "decisions.csv")
    filas = []
    if os.path.exists(ruta):
        with open(ruta, encoding="utf-8") as f:
            filas = list(csv.DictReader(f))
    filas = [f for f in filas if f.get("archivo") != archivo]
    filas.append({
        "archivo": archivo,
        "accion": accion,
        "nota": nota,
        "qui": qui,
        "data": data or datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })
    with open(ruta, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CAMPOS_DECISIONS_CSV)
        writer.writeheader()
        writer.writerows(filas)


def escribir_correccion(carpeta_cliente, archivo, cambios, motiu, qui):
    """Piso 11B: anexa una fila por cada campo que de verdad cambio a
    correccions.csv. validar.py (cargar_correcciones/aplicar_correcciones)
    las aplica en memoria antes de volver a validar -- nunca toca
    extraidas/, y la ficha corregida vuelve a pasar por TODA la red sin
    atajos (RECALCULAR ahora incluye validar.py primero)."""
    ruta = os.path.join(carpeta_cliente, "correccions.csv")
    escribir_cabecera = not os.path.exists(ruta)
    data_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(ruta, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if escribir_cabecera:
            writer.writerow(CAMPOS_CORRECCIONS_CSV)
        for camp, valor_antic, valor_nou in cambios:
            writer.writerow([archivo, camp, valor_antic, valor_nou, motiu, qui, data_actual])


def retirar_error(carpeta_cliente, ruta_archivo, motivo, qui, nota=None):
    """Piso 11A: filosofia A3 igual que arxivar_cliente (Piso 10.3) --
    nunca shutil.rmtree, solo shutil.move. Vive fuera de rebudes/ y
    apartados/ a proposito: asi extraer_todas.py (que si escanea esas
    carpetas) nunca vuelve a ver el archivo retirado.

    Piso 11C: 'nota' opcional -- la retirada en lot demana una nota
    compartida obligatoria (el motiu tecnic ja ve derivat sol de
    motivo_error, aixo es el que escriu la persona). Columna extra a
    registre.csv, mai llegida per cap altra maquina (comprovat)."""
    carpeta_retirats = os.path.join(carpeta_cliente, "errors_retirats")
    os.makedirs(carpeta_retirats, exist_ok=True)
    nombre_base, extension = os.path.splitext(os.path.basename(ruta_archivo))
    nombre_final = nombre_base + extension
    destino = os.path.join(carpeta_retirats, nombre_final)
    contador = 2
    while os.path.exists(destino):
        nombre_final = f"{nombre_base}_{contador}{extension}"
        destino = os.path.join(carpeta_retirats, nombre_final)
        contador += 1

    shutil.move(ruta_archivo, destino)

    ruta_registro = os.path.join(carpeta_retirats, "registre.csv")
    escribir_cabecera = not os.path.exists(ruta_registro)
    with open(ruta_registro, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if escribir_cabecera:
            writer.writerow(["arxiu", "motiu", "qui", "data", "nota"])
        writer.writerow([
            nombre_final, motivo, qui, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), nota or "",
        ])
    return destino


def contar_decisiones_sin_recalcular(carpeta_cliente, decisiones):
    """Compara la fecha de cada decision (formato propio, parseable --
    la escribe escribir_decision) contra el mtime de informe_2026.html.
    Sin informe todavia, cuentan todas. Fechas no parseables (alguien
    las escribio a mano en otro formato) no cuentan -- no rompen nada."""
    ruta_informe = os.path.join(carpeta_cliente, "informe_2026.html")
    if not os.path.exists(ruta_informe):
        return len(decisiones)
    mtime_informe = os.path.getmtime(ruta_informe)
    contador = 0
    for fila in decisiones.values():
        try:
            marca = datetime.strptime(fila.get("data", ""), "%Y-%m-%d %H:%M:%S").timestamp()
        except ValueError:
            continue
        if marca > mtime_informe:
            contador += 1
    return contador


def estado_revision_cliente(carpeta):
    """Recoge todo lo que necesita la vista Revisio para un cliente:
    fichas (rebudes+ingressos), decisiones, y errores -- todo leido de
    disco en cada rerun, nada cacheado en session_state."""
    carpeta_cliente = ruta_proyecto("clientes", carpeta)
    origen_gastos = os.path.join(carpeta_cliente, "rebudes")
    origen_ingressos_rel = RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS.get(carpeta, "apartados/ingressos")
    origen_ingressos = os.path.join(carpeta_cliente, *origen_ingressos_rel.split("/"))

    gastos = cargar_validadas(os.path.join(carpeta_cliente, "rebudes", "validadas"))
    ingresos = cargar_validadas(os.path.join(carpeta_cliente, "apartados", "ingressos_validadas"))
    todas = (
        [(n, d, "rebudes", origen_gastos) for n, d in gastos]
        + [(n, d, "ingressos", origen_ingressos) for n, d in ingresos]
    )

    decisiones = cargar_decisiones(carpeta_cliente)

    errores_gastos = detectar_errores(
        listar_archivos_rebudes(origen_gastos), os.path.join(carpeta_cliente, "rebudes", "extraidas")
    )
    archivos_ingressos_presentes = (
        [os.path.join(origen_ingressos, n) for n in os.listdir(origen_ingressos)
         if n.lower().endswith(EXTENSIONES_ORIGINAL)]
        if os.path.isdir(origen_ingressos) else []
    )
    errores_ingresos = detectar_errores(
        archivos_ingressos_presentes, os.path.join(carpeta_cliente, "apartados", "ingressos_extraidas")
    )

    pendents = [
        (n, d, flujo, origen) for n, d, flujo, origen in todas
        if d.get("estado") == "REVISAR" and n not in decisiones
    ]
    oks = [(n, d, flujo, origen) for n, d, flujo, origen in todas if d.get("estado") == "OK"]

    return {
        "carpeta_cliente": carpeta_cliente,
        "todas": todas,
        "pendents": pendents,
        "oks": oks,
        "decisiones": decisiones,
        "errores": [("rebudes", r) for r in errores_gastos] + [("ingressos", r) for r in errores_ingresos],
    }


def paginar(lista, key_pagina, por_pagina=15):
    """Piso 11C: acota a como mucho `por_pagina` elementos por pantalla
    -- necesario para clientes grandes (davinstal: 240 fitxes), tanto
    para no saturar la vista com per acotar el cost de qualsevol futur
    rerun complet (nomes es renderitzen les targetes de LA pagina
    actual, mai totes). La pagina viu a session_state -- sobreviu a
    reruns, es reseteja sola si el total de pagines encongeix per sota
    de la pagina on estaves (ex. despres d'un lot)."""
    total = len(lista)
    total_paginas = max(1, (total + por_pagina - 1) // por_pagina)
    pagina = st.session_state.get(key_pagina, 0)
    pagina = min(max(pagina, 0), total_paginas - 1)
    st.session_state[key_pagina] = pagina

    if total_paginas > 1:
        col_ant, col_info, col_seg = st.columns([1, 3, 1])
        with col_ant:
            if st.button("← Anterior", key=f"{key_pagina}_ant", disabled=pagina == 0):
                st.session_state[key_pagina] = pagina - 1
                st.rerun()
        with col_info:
            st.caption(f"Pàgina {pagina + 1} de {total_paginas} ({total} en total)")
        with col_seg:
            if st.button("Següent →", key=f"{key_pagina}_seg", disabled=pagina >= total_paginas - 1):
                st.session_state[key_pagina] = pagina + 1
                st.rerun()

    inicio = pagina * por_pagina
    return lista[inicio:inicio + por_pagina]


@st.fragment
def tarjeta_revisio(nombre, datos, origen, carpeta_cliente, qui, prefijo):
    """Piso 11A: tarjeta de UNA ficha (PENDENT u OK) con las mismas dos
    acciones (Aprovar/Descartar) -- simetria del punto 3: se puede
    descartar un OK con nota igual que se aprova un REVISAR.

    Piso 11A-fix: nota+botones dentro de un st.form -- un form aplica
    TODO lo tecleado en el momento de pulsar cualquiera de sus botones,
    sin necesitar Tab/blur (eso era lo que dejaba Aprovar/Descartar
    "sin responder" cuando Emili escribia y clicaba seguido). La nota
    ahora sirve para las dos acciones: obligatoria para Descartar,
    opcional para Aprovar (antes no habia ningun sitio para comentar
    al aprobar). qui ya llega siempre relleno -- la puerta de entrada
    a Revisio lo garantiza -- asi que los botones nunca se deshabilitan;
    si falta la nota al descartar, el clic SIEMPRE se procesa y se
    responde con un st.error visible, nunca en silencio.

    Piso 11C: @st.fragment -- un clic (Aprovar/Descartar/Corregir) fa
    rerun NOMES d'aquesta funcio (st.rerun(scope="fragment")), mai de
    la resta de la pantalla: ni el scroll salta ni les altres targetes
    es re-renderitzen (abans, ~33s per rerun amb un client gran, tot
    per re-renderitzar 240 vistes previes de PDF sense cap cache).
    Com un rerun de fragment NO torna a executar el bucle exterior que
    la crida, aquesta targeta no pot desapareixer sola de "Pendents"
    en decidir-se -- per aixo relegeix decisions.csv ELLA MATEIXA al
    principi (un sol CSV petit, barat) i mostra una confirmacio en
    comptes del formulari si ja hi ha decisio, en lloc de confiar en
    la 'decisio' que li va passar el bucle exterior (que pot estar
    desactualitzada tot i que la ficha ja s'hagi decidit fa un
    moment). Els comptadors de dalt (N pendents...) NO poden
    refrescar-se sols quan es decideix una targeta (un rerun de
    fragment no reexecuta el bloc que els calcula) -- per aixo la
    pantalla ho avisa sempre de forma fixa amb un boto "Actualitzar
    comptadors", mai fingint un numero en viu que no ho es."""
    decision = cargar_decisiones(carpeta_cliente).get(nombre)

    ruta_original = encontrar_original(origen, nombre)
    extension = os.path.splitext(ruta_original)[1].lower() if ruta_original else None

    # Piso 12B: mateixos colors que informe.py (.tarjeta.ok/.tarjeta.revisar)
    # -- la key del container es l'unica cosa que canvia, el CSS_A3 de sota
    # la reconeix pel prefix. prefijo aqui nomes val "pendent" (REVISAR
    # sense decisio) o "ok" (OK sense decisio) -- els ja decidits no criden
    # aquesta funcio (van a l'expander "Ja decidits" en text pla).
    clase_color = "revisar" if prefijo == "pendent" else "ok"
    with st.container(border=True, key=f"tarjeta_{clase_color}_{prefijo}_{nombre}"):
        col_izq, col_der = st.columns([2, 1])
        with col_izq:
            st.markdown(f"**{datos.get('proveedor')}**")
            st.caption(
                f"NIF {datos.get('nif_proveedor')} · Factura {datos.get('num_factura')} · "
                f"{datos.get('fecha_factura')} · estat {datos.get('estado')}"
            )
            for linea in datos.get("lineas_iva") or []:
                st.write(f"Base {linea.get('base')} € × {linea.get('tipo_iva')}% = {linea.get('cuota')} €")
            st.write(f"Total: {datos.get('total')} €")
            if datos.get("retencion_cuota"):
                st.write(f"Retenció: {datos.get('retencion_pct')}% = {datos.get('retencion_cuota')} €")
            motivos = [traducir_motivo(m) for m in (datos.get("motivos") or [])]
            if motivos:
                st.warning("\n".join(f"- {m}" for m in motivos))
            for aviso in avisos_verificacion_ficha(datos):
                st.caption(f"⚠ {aviso}")
            if datos.get("observaciones"):
                st.caption(f"Observacions: {datos['observaciones']}")
            camps_corregits = datos.get("camps_corregits") or []
            if camps_corregits:
                detall = "; ".join(
                    f"{c['camp']}: {c['antic']} → {c['nou']} ({c['qui']}, {c['data']})"
                    for c in camps_corregits
                )
                st.success(f"Corregit — {detall}")
            st.caption(nombre)
        with col_der:
            if ruta_original and extension in EXTENSIONES_IMAGEN:
                st.image(ruta_original)
            else:
                if ruta_original and extension == ".pdf":
                    imagen_pdf = previsualizar_pdf(ruta_original, os.path.getmtime(ruta_original))
                    if imagen_pdf:
                        st.image(imagen_pdf)
                boton_obrir("Obrir original", ruta_original or "", key=f"{prefijo}_original_{nombre}")

        if decision:
            etiqueta_accio = "Aprovada" if decision.get("accion") == "aprovar" else "Descartada"
            nota_mostrada = f" — _{decision.get('nota')}_" if decision.get("nota") else ""
            st.success(f"✓ {etiqueta_accio} per {decision.get('qui')} el {decision.get('data')}{nota_mostrada}")
        else:
            with st.form(key=f"form_{prefijo}_{nombre}", border=False):
                nota = st.text_input(
                    "Nota (obligatòria per descartar; opcional per aprovar)",
                    key=f"{prefijo}_nota_{nombre}",
                )
                col_a, col_b = st.columns([1, 2])
                with col_a:
                    click_aprovar = st.form_submit_button("Aprovar", type="primary")
                with col_b:
                    click_descartar = st.form_submit_button("Descartar")

            if click_aprovar:
                escribir_decision(carpeta_cliente, nombre, "aprovar", nota, qui)
                st.rerun(scope="fragment")
            if click_descartar:
                if not nota:
                    st.error("Cal escriure una nota per descartar.")
                else:
                    escribir_decision(carpeta_cliente, nombre, "descartar", nota, qui)
                    st.rerun(scope="fragment")

        # Piso 11B: "Corregir camps" -- capa de correccio, mai edita
        # extraidas/. Camps precarregats amb el valor ACTUAL; nomes es
        # desa el que de veritat canvia. Dins d'un st.form pel mateix
        # motiu que Aprovar/Descartar (Piso 11A-fix): sense form, un
        # camp editat sense Tab no es aplicaria en clicar Guardar.
        with st.popover("Corregir camps"):
            st.caption(
                "La fitxa corregida torna a passar tota la validació -- "
                "corregir no aprova. Cal RECALCULAR després de desar."
            )
            with st.form(key=f"form_correccio_{prefijo}_{nombre}", border=False):
                valors_nous = {}
                for camp in CAMPOS_CORREGIBLES_TOP:
                    valor_actual = datos.get(camp)
                    valors_nous[camp] = st.text_input(
                        camp,
                        value="" if valor_actual is None else str(valor_actual),
                        key=f"{prefijo}_correccio_{camp}_{nombre}",
                    )
                lineas_nuevas = []
                for i, linea in enumerate(datos.get("lineas_iva") or []):
                    st.caption(f"Línia {i}")
                    col_t, col_b, col_c = st.columns(3)
                    with col_t:
                        t = st.text_input(
                            f"tipo_iva[{i}]", value=str(linea.get("tipo_iva")),
                            key=f"{prefijo}_correccio_tipo_{i}_{nombre}",
                        )
                    with col_b:
                        b = st.text_input(
                            f"base[{i}]", value=str(linea.get("base")),
                            key=f"{prefijo}_correccio_base_{i}_{nombre}",
                        )
                    with col_c:
                        c = st.text_input(
                            f"cuota[{i}]", value=str(linea.get("cuota")),
                            key=f"{prefijo}_correccio_cuota_{i}_{nombre}",
                        )
                    lineas_nuevas.append((i, t, b, c))
                motiu_correccio = st.text_input("Motiu (obligatori)", key=f"{prefijo}_correccio_motiu_{nombre}")
                click_guardar = st.form_submit_button("Guardar correccions")

            if click_guardar:
                cambios = []
                for camp in CAMPOS_CORREGIBLES_TOP:
                    valor_antic = datos.get(camp)
                    valor_antic_str = "" if valor_antic is None else str(valor_antic)
                    if valors_nous[camp] != valor_antic_str:
                        cambios.append((camp, valor_antic_str, valors_nous[camp]))
                for i, t, b, c in lineas_nuevas:
                    linea_actual = (datos.get("lineas_iva") or [])[i]
                    if t != str(linea_actual.get("tipo_iva")):
                        cambios.append((f"lineas_iva[{i}].tipo_iva", str(linea_actual.get("tipo_iva")), t))
                    if b != str(linea_actual.get("base")):
                        cambios.append((f"lineas_iva[{i}].base", str(linea_actual.get("base")), b))
                    if c != str(linea_actual.get("cuota")):
                        cambios.append((f"lineas_iva[{i}].cuota", str(linea_actual.get("cuota")), c))

                if not cambios:
                    st.info("Cap canvi per desar.")
                elif not motiu_correccio:
                    st.error("Cal escriure un motiu per desar correccions.")
                else:
                    escribir_correccion(carpeta_cliente, nombre, cambios, motiu_correccio, qui)
                    st.success(f"{len(cambios)} camp(s) corregit(s). Cal RECALCULAR perquè es reavaluïn.")
                    st.rerun(scope="fragment")


@st.fragment
def tarjeta_error(flujo, ruta, carpeta_cliente, qui, prefijo):
    """Piso 11A: tarjeta por ERROR (archivo presente sin ficha) --
    motivo derivado, enlace, instruccion, y accion RETIRAR (mai
    esborrar, mou a errors_retirats/). Piso 11A-fix: qui ya llega
    siempre relleno (puerta de entrada), no hace falta disabled=.

    Piso 11C: @st.fragment igual que tarjeta_revisio -- Retirar nomes
    fa rerun d'aquesta targeta. Com el bucle exterior no torna a
    executar-se, comprova ELLA MATEIXA si l'arxiu encara existeix
    (pot haver-lo retirat un lot mentre aquesta targeta seguia
    muntada) abans de mostrar el boto."""
    nombre = os.path.basename(ruta)
    with st.container(border=True, key=f"tarjeta_error_{prefijo}_{nombre}"):
        st.markdown(f"**[{flujo.upper()}] {nombre}**")
        if not os.path.exists(ruta):
            st.success("✓ Arxiu retirat.")
        else:
            st.warning(motivo_error(ruta))
            boton_obrir("Obrir arxiu", ruta, key=f"{prefijo}_error_original_{nombre}")
            st.caption("Torna a pujar l'arxiu bo des d'\"Afegir factures\".")
            if st.button("Retirar arxiu il·legible", key=f"{prefijo}_retirar_{nombre}"):
                destino = retirar_error(carpeta_cliente, ruta, motivo_error(ruta), qui)
                st.success(f"Arxiu retirat a `{destino}`.")
                st.rerun(scope="fragment")


# Piso 10.5: unic bloc CSS -- nomes per allo que config.toml no cobreix
# (l'stack de tipografia exacte demanat i la subtil ombra de les
# targetes). Colors/radi/etc ja viuen a .streamlit/config.toml.
CSS_A3 = """
<style>
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
}
[data-testid="stVerticalBlockBorderWrapper"] {
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.06);
    border-radius: 12px;
}
[data-testid="stVerticalBlock"] > [data-testid="stVerticalBlockBorderWrapper"] {
    margin-bottom: 0.5rem;
}
h1, h2, h3 {
    letter-spacing: -0.01em;
}
/* Piso 12B: mateixos colors que informe.py (.tarjeta.ok/.tarjeta.revisar)
   a les targetes de Revisió -- en aquesta versió de Streamlit la vora,
   el radi i el padding del container(border=True) viuen directament al
   propi stVerticalBlock que porta la classe st-key-<key> (no hi ha cap
   embolcall stVerticalBlockBorderWrapper com en versions anteriors). */
[class*="st-key-tarjeta_revisar_"] {
    background: #FCE4D6 !important;
    border-color: #e8b48f !important;
}
[class*="st-key-tarjeta_ok_"] {
    background: #E2EFDA !important;
    border-color: #b7d7a8 !important;
}
[class*="st-key-tarjeta_error_"] {
    background: #FFC7CE !important;
    border-color: #e2929c !important;
}
</style>
"""

# Piso 13B: assets/ absent no pot tirar avall l'app (regla "jamas
# caida" -- degradar, mai petar). Sense aquests dos arxius, l'app
# arrenca igual, nomes sense icona/logo.
_icono = ruta_proyecto("assets", "olivella.ico")
st.set_page_config(
    page_title="Agent TRIMESTRE — Gestoria Olivella",
    page_icon=_icono if os.path.exists(_icono) else "📊",
    layout="wide",
)
_logo = ruta_proyecto("assets", "logo_olivella.png")
if os.path.exists(_logo):
    st.logo(_logo)
st.markdown(CSS_A3, unsafe_allow_html=True)

# Piso 13B: guàrdia de dades absents -- bloqueja TOTES les vistes amb
# un únic check (mateix patró que la porta "Qui revisa" de Revisió)
# en comptes d'un traceback cru.
# Piso 13C: bifurcació, no mur -- "clientes/" absent pot ser un PC nou
# amb dades per copiar del USB, o un PC que comença net de veritat.
# Cap dels dos botons porta type="primary": mateix pes visual, cap
# camí es "el recomanat". Els dos deixen rastre a instalacio.log --
# mai s'auto-crea en silenci sense que algú ho hagi triat.
if not os.path.isdir(ruta_proyecto("clientes")):
    st.title("Agent TRIMESTRE")
    st.warning("Aquest PC encara no té la carpeta de dades (`clientes/`). Tria una opció:")

    col_usb, col_zero = st.columns(2)
    with col_usb:
        st.subheader("Tinc dades prèvies (USB)")
        st.write(
            "Copia la carpeta `clientes/` des del USB dins de:\n\n"
            f"`{ruta_proyecto('clientes')}`"
        )
        if st.button("Torna-ho a comprovar", key="guardia_comprovar"):
            if os.path.isdir(ruta_proyecto("clientes")):
                registrar_instalacio("USB (dades prèvies trobades)")
            st.rerun()
    with col_zero:
        st.subheader("Començar de zero en aquest PC")
        st.write(
            "Crea una carpeta `clientes/` buida en aquest PC. "
            "Podràs donar d'alta clients nous des de la pestanya "
            "\"Clients\" tot seguit -- \"Nou client\" ja crea la "
            "carpeta i subcarpetes de cadascun, com sempre."
        )
        if st.button("Començar de zero en aquest PC", key="guardia_zero"):
            escribir_clientes([])
            registrar_instalacio("Començar de zero en aquest PC")
            st.rerun()
    st.stop()

if "log_proces" not in st.session_state:
    st.session_state["log_proces"] = None

st.title("Agent TRIMESTRE")
vista = st.sidebar.radio("Navegació", ["Clients", "Afegir factures", "Processar", "Revisió"])

# ----------------------------------------------------------------------
if vista == "Clients":
    st.header("Clients")
    clientes = leer_clientes()

    if not clientes:
        st.info("Encara no hi ha cap client donat d'alta.")
    else:
        for fila in clientes:
            tarjeta_cliente(fila, "clients")

    with st.expander("Nou client"):
        # Piso 10.2: FORA de st.form -- els widgets dins d'un form no
        # reexecuten l'script en cada tecla, i la carpeta derivada
        # nomes pot ser una vista previa en viu si aquest bloc si ho fa.
        # Netejar session_state d'un key JA instanciat aquest run peta
        # (StreamlitAPIException) -- cal fer-ho ABANS d'instanciar el
        # widget, al rerun seguent, no en el mateix click.
        if st.session_state.get("nou_client_creat"):
            st.session_state["nou_nombre"] = ""
            st.session_state["nou_nif"] = ""
            st.session_state["nou_client_creat"] = False

        nombre = st.text_input("Nom del client", key="nou_nombre").strip()
        nif = normalizar_nif(st.text_input("NIF", key="nou_nif"))
        carpeta_derivada = slug(nombre) if nombre else ""
        if carpeta_derivada:
            st.caption(f"Es crearà la carpeta: `{carpeta_derivada}`")
        elif nombre:
            st.caption("El nom no genera cap identificador vàlid -- afegeix lletres o números.")
        continuar_igualment = st.checkbox(
            "El NIF no quadra la lletra de control -- continuar igualment", key="nou_continuar"
        )

        if st.button("Crear client", key="nou_crear", type="primary"):
            carpetas_existentes = {f["carpeta"] for f in leer_clientes()}
            if not nombre or not nif:
                st.error("Falten camps: nom i NIF són obligatoris.")
            elif not carpeta_derivada:
                st.error("El nom no genera cap identificador vàlid -- afegeix lletres o números.")
            elif carpeta_derivada in carpetas_existentes or os.path.isdir(ruta_proyecto("clientes", carpeta_derivada)):
                st.error(f"Ja existeix un client amb la carpeta '{carpeta_derivada}'. Ajusta el nom per diferenciar-lo.")
            elif validar_nif(nif) is False and not continuar_igualment:
                st.warning(
                    "El NIF no supera la validació de la lletra de control. "
                    "Marca 'continuar igualment' si n'estàs segur i torna a enviar."
                )
            else:
                anadir_cliente(nif, nombre, carpeta_derivada)
                st.session_state["nou_client_creat"] = True
                st.success(f"Client '{nombre}' creat (carpeta `{carpeta_derivada}`).")
                st.rerun()

# ----------------------------------------------------------------------
elif vista == "Afegir factures":
    st.header("Afegir factures")
    clientes = leer_clientes()

    if not clientes:
        st.info("Primer cal donar d'alta un client a la pestanya 'Clients'.")
    else:
        opciones = {f"{f['nombre']} ({f['carpeta']})": f["carpeta"] for f in clientes}
        eleccion = st.selectbox("Client", list(opciones.keys()))
        carpeta = opciones[eleccion]

        # Piso 13D: tercer desti -- "Lot d'escaner". format_func desacobla
        # el valor intern ("Lot") del text llarg mostrat, per no haver de
        # comparar contra el string sencer a ruta_destino_factures.
        destino = st.radio(
            "Destí", list(DESTINOS_AFEGIR.keys()),
            format_func=lambda k: DESTINOS_AFEGIR[k], horizontal=True,
        )
        st.caption(
            "Factura solta: un únic document per arxiu. Lot d'escàner: un "
            "sol PDF escanejat amb diverses factures/albarans a dins -- es "
            "trocejarà automàticament en processar."
        )

        tipos_permitidos = ["pdf"] if destino == "Lot" else EXTENSIONES_PERMITIDAS
        archivos = st.file_uploader(
            "Arrossega els arxius aquí",
            type=tipos_permitidos,
            accept_multiple_files=True,
        )

        # Piso 13D: guardia de lots despistats -- nomes te sentit si
        # l'usuari NO ha triat ja "Lot d'escaner" expressament. Mira
        # cada PDF en memoria (mai es desa res a disc encara).
        sospitos = None
        if destino in ("Compres", "Vendes"):
            for a in archivos or []:
                if os.path.splitext(a.name)[1].lower() == ".pdf":
                    try:
                        n_pag = len(PdfReader(io.BytesIO(a.getvalue())).pages)
                    except Exception:
                        continue  # PDF no llegible -- no es el moment de bloquejar
                    if n_pag > 3:
                        sospitos = (a.name, n_pag)
                        break

        if sospitos:
            nombre_sospitos, n_pag = sospitos
            st.warning(f"Aquest PDF ({nombre_sospitos}) té {n_pag} pàgines — és un lot d'escàner?")
            col_si, col_no = st.columns(2)
            with col_si:
                confirmar_lot = st.button("Sí, és un lot", key="lot_confirmar")
            with col_no:
                confirmar_solta = st.button("No, és una factura llarga", key="lot_descartar")
            if confirmar_lot:
                guardar_y_reportar(archivos, ruta_destino_factures(carpeta, "Lot"))
            elif confirmar_solta:
                guardar_y_reportar(archivos, ruta_destino_factures(carpeta, destino))
        else:
            if st.button("Desar arxius", disabled=not archivos, type="primary"):
                guardar_y_reportar(archivos, ruta_destino_factures(carpeta, destino))

# ----------------------------------------------------------------------
elif vista == "Processar":
    st.header("Processar")
    clientes = leer_clientes()

    if not clientes:
        st.info("Encara no hi ha cap client donat d'alta -- res per processar.")
    else:
        if st.button("Processar", type="primary"):
            placeholder = st.empty()
            buffer = ""
            proceso = subprocess.Popen(
                [sys.executable, "ejecutar.py"],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                cwd=RAIZ_PROYECTO,
            )
            for linea in proceso.stdout:
                buffer += linea
                placeholder.code(buffer)
            proceso.wait()
            st.session_state["log_proces"] = buffer
            if proceso.returncode == 0:
                st.success("Procés acabat.")
            else:
                st.error(f"El procés ha acabat amb error (codi {proceso.returncode}).")

        if st.session_state["log_proces"]:
            st.subheader("Resum per client")
            st.caption(
                "S'han processat tots els clients (els que ja estaven al dia "
                "s'han saltat per idempotència)."
            )
            for fila in clientes:
                linias_cliente = [
                    linia for linia in st.session_state["log_proces"].splitlines()
                    if linia.startswith(f"{fila['carpeta']} ")
                ]
                tarjeta_cliente(fila, "final", lineas_log=linias_cliente)
            boton_obrir("Obrir portada", ruta_proyecto("clientes", "index.html"), key="final_portada")

# ----------------------------------------------------------------------
elif vista == "Revisió":
    st.header("Revisió")
    clientes = leer_clientes()

    if not clientes:
        st.info("Encara no hi ha cap client donat d'alta.")
    elif not st.session_state.get("qui_revisa_confirmat"):
        # Piso 11A-fix: puerta BLOQUEANTE -- dentro de un st.form el valor
        # tecleado se aplica al pulsar el boton, sin necesitar Tab/blur
        # (a diferencia de un st.text_input suelto, que es lo que hacia
        # que Aprovar/Descartar/Retirar parecieran "no responder").
        st.info("Cal indicar qui revisa abans de veure cap fitxa.")
        with st.form("form_qui_revisa"):
            nom_qui = st.text_input("Qui revisa? (una vegada per sessió)", key="qui_revisa_input")
            entrar = st.form_submit_button("Entrar a revisar", type="primary")
        if entrar:
            if not nom_qui.strip():
                st.error("Cal escriure un nom abans de continuar.")
            else:
                st.session_state["qui_revisa_confirmat"] = nom_qui.strip()
                st.rerun()
        st.stop()
    else:
        qui = st.session_state["qui_revisa_confirmat"]
        col_qui, col_canviar = st.columns([4, 1])
        with col_qui:
            st.caption(f"Revisant com: **{qui}**")
        with col_canviar:
            if st.button("Canviar qui revisa", key="qui_canviar"):
                st.session_state["qui_revisa_confirmat"] = None
                st.rerun()

        opciones = {f"{f['nombre']} ({f['carpeta']})": f for f in clientes}
        eleccion = st.selectbox("Client", list(opciones.keys()), key="revisio_client")
        fila_cliente = opciones[eleccion]
        carpeta = fila_cliente["carpeta"]

        estado = estado_revision_cliente(carpeta)
        n_pendents = len(estado["pendents"])
        n_errores = len(estado["errores"])
        n_decidits = len(estado["decisiones"])
        n_sin_recalcular = contar_decisiones_sin_recalcular(estado["carpeta_cliente"], estado["decisiones"])

        st.subheader(
            f"{n_pendents} pendents per decidir · {n_errores} errors per resoldre · "
            f"{n_decidits} decidits · {n_sin_recalcular} decisions noves sense recalcular"
        )

        # Piso 11C: cada targeta es un @st.fragment -- Aprovar/Descartar/
        # Retirar nomes reexecuten la targeta, mai aquest bloc exterior,
        # aixi que aquests numeros NO es poden refrescar sols quan es
        # decideix una targeta (series fals ensenyar un comptador "en viu"
        # que en realitat no pot actualitzar-se sense un rerun complet).
        # Explicit i sempre visible en comptes de silenciós (regla 10):
        # es diu clarament quan es van calcular i es dona un boto a ma.
        col_avis, col_boto = st.columns([4, 1])
        with col_avis:
            st.caption(
                "Aquests números són de l'última vegada que es va carregar o refrescar la "
                "pantalla -- decidir una targeta individual no els actualitza sols."
            )
        with col_boto:
            if st.button("Actualitzar comptadors", key="revisio_actualitzar_comptadors"):
                st.rerun()

        if st.button(
            "RECALCULAR", key="revisio_recalcular",
            type="primary" if n_sin_recalcular > 0 else "secondary",
        ):
            placeholder = st.empty()
            buffer = ""
            # Piso 11B: validar.py entra en la cadena -- las correccions.csv
            # solo se aplican en memoria dentro de validar.py, asi que hace
            # falta volver a correrlo para que una correccio "torni a passar
            # l'examen". Sigue sin llamadas a la API (validar.py es pura
            # logica Python), igual de gratis que sumar.py/informe.py.
            for maquina in ["validar.py", "sumar.py", "informe.py"]:
                proceso = subprocess.Popen(
                    [sys.executable, maquina],
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, bufsize=1, cwd=RAIZ_PROYECTO,
                )
                for linea in proceso.stdout:
                    buffer += linea
                    placeholder.code(buffer)
                proceso.wait()
            st.success("Recalculat.")
            st.rerun()

        col1, col2 = st.columns(2)
        with col1:
            boton_obrir("Obrir informe", os.path.join(estado["carpeta_cliente"], "informe_2026.html"), key="revisio_informe")
        with col2:
            boton_obrir("Obrir Excel", os.path.join(estado["carpeta_cliente"], "sumatorios_2026.xlsx"), key="revisio_excel")

        # Piso 11C: checkbox de lot FORA del fragment de cada targeta --
        # marcar/desmarcar es un rerun normal (ja barat gracies a la
        # paginacio + cache de PDF), aixi el comptador de seleccionades
        # es veu sempre al dia sense dependre de fragments. La barra
        # d'accio en lot NOMES es pinta si hi ha alguna seleccionada --
        # "no existeix aprovar-ho-tot sense selecció" literalment: no hi
        # ha res a clicar, no un boto desactivat.
        st.markdown(f"### Pendents sense decisió ({n_pendents})")
        if not estado["pendents"]:
            st.caption("Cap pendent sense decidir.")
        else:
            pagina_pendents = paginar(estado["pendents"], f"revisio_pag_pendents_{carpeta}")
            for nombre, datos, flujo, origen in pagina_pendents:
                st.checkbox("Seleccionar per al lot", key=f"revisio_sel_pendent_{carpeta}_{nombre}")
                tarjeta_revisio(nombre, datos, origen, estado["carpeta_cliente"], qui, "pendent")

            seleccionats_pendents = [
                nombre for nombre, _, _, _ in pagina_pendents
                if st.session_state.get(f"revisio_sel_pendent_{carpeta}_{nombre}")
            ]
            if seleccionats_pendents:
                st.info(f"{len(seleccionats_pendents)} seleccionades")
                nota_lot_pendents = st.text_input(
                    "Nota compartida (opcional)", key=f"revisio_lot_nota_pendents_{carpeta}"
                )
                if st.button("APROVAR SELECCIONADES", type="primary", key=f"revisio_lot_aprovar_{carpeta}"):
                    data_lote = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    for nombre in seleccionats_pendents:
                        escribir_decision(
                            estado["carpeta_cliente"], nombre, "aprovar", nota_lot_pendents, qui, data=data_lote,
                        )
                    st.rerun()

        oks_no_decididas = [
            (n, d, flujo, origen) for n, d, flujo, origen in estado["oks"]
            if n not in estado["decisiones"]
        ]
        # Piso 11C: key= fixa -- abans l'etiqueta portava el recompte
        # dins del text i Streamlit en derivava la identitat del propi
        # text, aixi que CADA Aprovar/Descartar (que canvia el recompte)
        # feia que semblés un expander NOU i el tancava sol encara que
        # l'usuari l'hagués obert. Amb key= fixa sobreviu a qualsevol
        # rerun complet futur. Sense lot -- les OK ja sumen.
        with st.expander(
            f"Fitxes OK — revisar-les també ({len(oks_no_decididas)})",
            key=f"revisio_expander_oks_{carpeta}",
        ):
            oks_ordenadas = sorted(
                oks_no_decididas, key=lambda t: 0 if avisos_verificacion_ficha(t[1]) else 1
            )
            if not oks_ordenadas:
                st.caption("Sense fitxes OK.")
            else:
                pagina_oks = paginar(oks_ordenadas, f"revisio_pag_oks_{carpeta}")
                for nombre, datos, flujo, origen in pagina_oks:
                    tarjeta_revisio(nombre, datos, origen, estado["carpeta_cliente"], qui, "ok")

        st.markdown(f"### Errors ({n_errores})")
        if not estado["errores"]:
            st.caption("Cap error per resoldre.")
        else:
            st.caption(
                "Si l'arxiu és bo però va quedar corromput (iCloud), torna'l a pujar "
                "abans de retirar res."
            )
            pagina_errores = paginar(estado["errores"], f"revisio_pag_errors_{carpeta}")
            for flujo, ruta in pagina_errores:
                st.checkbox("Seleccionar per al lot", key=f"revisio_sel_error_{carpeta}_{ruta}")
                tarjeta_error(flujo, ruta, estado["carpeta_cliente"], qui, "error")

            seleccionats_errores = [
                (flujo, ruta) for flujo, ruta in pagina_errores
                if st.session_state.get(f"revisio_sel_error_{carpeta}_{ruta}")
            ]
            if seleccionats_errores:
                st.info(f"{len(seleccionats_errores)} seleccionats")
                nota_lot_errores = st.text_input(
                    "Nota compartida (obligatòria)", key=f"revisio_lot_nota_errors_{carpeta}"
                )
                if st.button("RETIRAR SELECCIONATS", key=f"revisio_lot_retirar_{carpeta}"):
                    if not nota_lot_errores:
                        st.error("Cal escriure una nota abans de retirar en lot.")
                    else:
                        for flujo, ruta in seleccionats_errores:
                            retirar_error(
                                estado["carpeta_cliente"], ruta, motivo_error(ruta), qui, nota=nota_lot_errores,
                            )
                        st.rerun()

        with st.expander(f"Ja decidits ({n_decidits})", key=f"revisio_expander_decidits_{carpeta}"):
            if not estado["decisiones"]:
                st.caption("Cap decisió encara.")
            for archivo, decision in estado["decisiones"].items():
                nota = f" — _{decision.get('nota')}_" if decision.get("nota") else ""
                st.write(
                    f"**{archivo}** — {decision.get('accion')} per {decision.get('qui')} "
                    f"el {decision.get('data')}{nota}"
                )
