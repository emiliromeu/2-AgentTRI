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
import hashlib
import html
import io
import json
import os
import re
import shutil
import subprocess
import sys
import unicodedata
from datetime import datetime

import fitz
import streamlit as st
from PIL import Image
from pillow_heif import register_heif_opener
from pypdf import PdfReader

# Piso 13L: migrar_lot.py es l'unica excepcio a "ninguna maquina es
# importable" -- no te bucle de nivell superior (nomes definicions),
# pensat expressament per ser importat des d'aqui.
import migrar_lot

register_heif_opener()

# Piso 10.2: ancorat a la carpeta d'aquest arxiu, no al directori des
# d'on es llanci `streamlit run` -- abans "clientes/..." nomes
# funcionava si s'arrencava amb el cwd a l'arrel del projecte.
RAIZ_PROYECTO = os.path.dirname(os.path.abspath(__file__))


def ruta_proyecto(*partes):
    return os.path.join(RAIZ_PROYECTO, *partes)


# Piso 13S: candau de Processar -- mai dues "fàbriques" corrent alhora
# sobre els mateixos arxius. Mateix criteri que ejecutar.py (duplicat a
# proposit, cap "maquina" n'importa una altra); aquest fitxer, a més,
# es qui LLANÇA ejecutar.py, així que ha de saber comprovar-ho abans de
# fer-ho i mentre corre en segon pla.
RUTA_LOCK_PROCESSAR = ruta_proyecto("processar.lock")
RUTA_STOP_PROCESSAR = ruta_proyecto("processar.stop")
RUTA_LOG_PROCESSAR = ruta_proyecto("proces_log.txt")


def proceso_vivo(pid):
    """Piso 13U: REGRESSIÓ GREU trobada i arreglada -- a Windows,
    signal.CTRL_C_EVENT val 0, i os.kill(pid, 0) NO és un simple xec
    d'existència com a Unix: crida GenerateConsoleCtrlEvent, que envia
    l'esdeveniment a TOT el grup de consola que comparteix `pid`. Com
    ejecutar.py es llança amb un Popen normal (mateix grup de consola
    que el Streamlit pare), la primera comprovació del candau (el
    banner, a CADA càrrega de pàgina) enviava un Ctrl+C real que matava
    Streamlit sencer -- reproduït i diagnosticat via git diff +
    documentació de Python/Windows, confirmat que aquest projecte corre
    en Windows (app.py usa os.startfile(), sumar.py gestiona rutes
    "C:/..."). A Windows, OpenProcess/CloseHandle (ctypes, stdlib, cap
    dependència nova) NOMÉS consulten -- mai envien cap senyal."""
    if os.name == "nt":
        import ctypes
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        handle = ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if handle:
            ctypes.windll.kernel32.CloseHandle(handle)
            return True
        return False
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except OSError:
        return True
    return True


def candau_processar_viu():
    """Retorna la info (dict) del candau si un Processar és viu de
    veritat (PID encara existeix), None si no n'hi ha o està mort
    (orfe -- es deixa tal qual, ejecutar.py ja el neteja en arrencar).

    Piso 13U: try/except ample a tota la funció -- el candau el
    gestiona NOMÉS ejecutar.py; aquí NOMÉS es llegeix, i un candau
    corrupte, absent, o qualsevol sorpresa de plataforma mai pot tombar
    el script principal (aquest era exactament el mecanisme de la
    regressió greu d'aquest pis)."""
    try:
        if not os.path.exists(RUTA_LOCK_PROCESSAR):
            return None
        with open(RUTA_LOCK_PROCESSAR, encoding="utf-8") as f:
            info = json.load(f)
        pid = info.get("pid")
        if pid is None or not proceso_vivo(pid):
            return None
        return info
    except Exception:
        return None


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

# Piso 13J: quan es tria "Lot" directament (sense passar per la guardia
# de Compres/Vendes), cal preguntar quin flux es -- tradueix la tria
# cap al mateix vocabulari intern que ja fa servir ruta_destino_factures.
DESTI_LOT_DIRECTE = {"Compres": "Compres", "Vendes (emeses)": "Vendes"}

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


PATRON_NIF_UE = re.compile(r"^ES([0-9A-Z]{9})$")


def normalizar_nif(nif):
    """Piso 13J: mateix criteri que validar.py -- treu el prefix
    intracomunitari "ES" davant d'un NIF/CIF domestic de 9 caracters."""
    n = "".join(c for c in (nif or "") if c.isalnum()).upper()
    m = PATRON_NIF_UE.match(n)
    return m.group(1) if m else n


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
    # Piso 13K: moll bessó dels dos d'amunt, pero per a lots de VENDES --
    # abans nomes existia un moll de lots (implicitament de compres), i
    # un lot de vendes hi acabava barrejat (bug de camp).
    os.makedirs(ruta_proyecto("clientes", carpeta, "apartados", "lotes_vendes_escaneados"), exist_ok=True)
    os.makedirs(ruta_proyecto("clientes", carpeta, "apartados", "lotes_vendes_procesados"), exist_ok=True)


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


def regenerar_index_clientes():
    """Piso 13V: reescriu clientes/index.html NOMES amb els clients
    actuals de leer_clientes() -- es crida just despres d'arxivar un
    client perque la portada mai ensenyi un enllaç trencat fins al
    proper Processar/RECALCULAR. A proposit NO crida informe.py: aixo
    tocaria l'informe_2026.html de TOTS els clients i n'esborraria el
    semafor "sense recalcular" sense que sumar.py s'hagi tornat a
    executar de veritat. Estil propi i minim (no el d'informe.py,
    pensat per a tarjetes/conciliacio que aquesta portada no fa
    servir) -- cap re-execucio de cap maquina, nomes os.path.exists()."""
    filas_taula = ""
    for f in leer_clientes():
        carpeta = f["carpeta"]
        carpeta_cliente = ruta_proyecto("clientes", carpeta)
        ruta_informe = os.path.join(carpeta_cliente, "informe_2026.html")
        ruta_excel = os.path.join(carpeta_cliente, "sumatorios_2026.xlsx")
        enlace_informe = (
            f'<a href="{carpeta}/informe_2026.html">Informe</a>'
            if os.path.exists(ruta_informe) else "no disponible"
        )
        enlace_excel = (
            f'<a href="{carpeta}/sumatorios_2026.xlsx" download>Obrir l\'Excel</a>'
            if os.path.exists(ruta_excel) else "no disponible"
        )
        filas_taula += (
            f"<tr><td>{html.escape(f['nombre'])}</td><td>{html.escape(f['nif'])}</td>"
            f"<td>{enlace_informe}</td><td>{enlace_excel}</td></tr>"
        )
    index_html = f"""<!DOCTYPE html>
<html lang="ca">
<head>
<meta charset="utf-8">
<title>Clients — Agent TRIMESTRE</title>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 2rem; }}
table {{ border-collapse: collapse; width: 100%; }}
th, td {{ border: 1px solid #ccc; padding: 0.4rem 0.8rem; text-align: left; }}
th {{ background: #D9E1F2; }}
</style>
</head>
<body>
  <h1>Clients</h1>
  <p>Actualitzat el {datetime.now().strftime("%d/%m/%Y %H:%M")}</p>
  <table>
    <thead><tr><th>Client</th><th>NIF</th><th>Informe</th><th>Excel</th></tr></thead>
    <tbody>{filas_taula}</tbody>
  </table>
</body>
</html>"""
    with open(ruta_proyecto("clientes", "index.html"), "w", encoding="utf-8") as f:
        f.write(index_html)


# =======================================================================
# Piso 13N: DESTRUIR -- l'unic verb del projecte que esborra de veritat
# (os.remove/shutil.rmtree). Fins ara tot movia (arxivar_cliente,
# retirar_error) o nomes excloia de sumes (decisions.csv). Nomes
# arriba aqui el que ja esta retirat/descartat o un client ja
# arxivat -- mai un document viu al flux -- i sempre amb certificat
# escrit ABANS d'esborrar res.
# =======================================================================

CAMPOS_REGISTRE_DESTRUCCIONS = ["data", "qui", "client", "detall", "motiu"]


def asegurar_gitignore_destruccions():
    """Mateix patro que asegurar_gitignore_arxivats -- el certificat
    de l'arrel (clients arxivats sencers) porta dades fiscals reals
    als noms/motius, mai a git."""
    ruta_gitignore = ruta_proyecto(".gitignore")
    lineas = []
    if os.path.exists(ruta_gitignore):
        with open(ruta_gitignore, encoding="utf-8") as f:
            lineas = f.read().splitlines()
    if "registre_destruccions.csv" not in lineas:
        with open(ruta_gitignore, "a", encoding="utf-8") as f:
            f.write("registre_destruccions.csv\n")


def escribir_registre_destruccio(ruta_registro, client, detall, motiu, qui):
    """El certificat -- l'unica empremta que sobreviu a la destruccio.
    S'escriu SEMPRE abans d'esborrar cap fitxer (si aixo falla, no
    s'ha destruit encara res). Llibre major: nomes afegeix."""
    asegurar_gitignore_destruccions()
    escribir_cabecera = not os.path.exists(ruta_registro)
    os.makedirs(os.path.dirname(ruta_registro), exist_ok=True)
    with open(ruta_registro, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CAMPOS_REGISTRE_DESTRUCCIONS)
        if escribir_cabecera:
            writer.writeheader()
        writer.writerow({
            "data": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "qui": qui,
            "client": client,
            "detall": detall,
            "motiu": motiu,
        })


def candidatos_destruccio(carpeta):
    """Piso 13N: les DUES UNIQUES fonts destruibles -- mai un document
    viu al flux. retirats: errors_retirats/ (sense JSON, l'original
    tot sol). descartats: fitxes amb decisio EFECTIVA "descartar"
    (cargar_decisiones ja aplica el llibre major del Piso 13M -- si
    s'ha revertit el descartar, l'arxiu ja no surt aqui)."""
    carpeta_cliente = ruta_proyecto("clientes", carpeta)

    retirats = []
    carpeta_retirats = os.path.join(carpeta_cliente, "errors_retirats")
    if os.path.isdir(carpeta_retirats):
        for nombre in sorted(os.listdir(carpeta_retirats)):
            if nombre == "registre.csv":
                continue
            ruta = os.path.join(carpeta_retirats, nombre)
            if os.path.isfile(ruta):
                retirats.append({
                    "tipo": "retirat", "nombre": nombre,
                    "rutas": [ruta], "num_factura": None,
                })

    descartats = []
    decisiones = cargar_decisiones(carpeta_cliente)
    origen_ingressos_rel = RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS.get(carpeta, "apartados/ingressos")
    flujos = [
        ("rebudes", os.path.join(carpeta_cliente, "rebudes", "validadas"),
         os.path.join(carpeta_cliente, "rebudes", "extraidas"),
         os.path.join(carpeta_cliente, "rebudes")),
        ("ingressos", os.path.join(carpeta_cliente, "apartados", "ingressos_validadas"),
         os.path.join(carpeta_cliente, "apartados", "ingressos_extraidas"),
         os.path.join(carpeta_cliente, *origen_ingressos_rel.split("/"))),
    ]
    for flujo, carpeta_validadas, carpeta_extraidas, carpeta_origen in flujos:
        for nombre, datos in cargar_validadas(carpeta_validadas):
            decision = decisiones.get(nombre)
            if estat_efectiu(decision) != "descartar":
                continue
            rutas = []
            ruta_original = encontrar_original(carpeta_origen, nombre)
            if ruta_original:
                rutas.append(ruta_original)
            ruta_extraida = os.path.join(carpeta_extraidas, nombre)
            if os.path.exists(ruta_extraida):
                rutas.append(ruta_extraida)
            ruta_validada = os.path.join(carpeta_validadas, nombre)
            if os.path.exists(ruta_validada):
                rutas.append(ruta_validada)
            descartats.append({
                "tipo": "descartat", "nombre": nombre, "flujo": flujo,
                "rutas": rutas, "num_factura": datos.get("num_factura"),
            })

    return retirats, descartats


def destruir_documentos(carpeta, items, motiu, qui):
    """Piso 13N: re-verifica CADA item contra candidatos_destruccio
    fresc just abans d'actuar -- bloqueig contra condicio de cursa
    (algu podria haver revertit un descartar en un altre tab mentre
    aquesta pantalla estava oberta). Els que ja no son destruibles es
    salten i es reporten -- mai en silenci (regla 4). Escriu el
    certificat ABANS d'esborrar, despres os.remove de veritat, i deixa
    nota permanent a decisions.csv per als descartats (llibre major
    del Piso 13M -- guanya un tercer valor d'accio, mai reescriu)."""
    carpeta_cliente = ruta_proyecto("clientes", carpeta)
    retirats_valids, descartats_valids = candidatos_destruccio(carpeta)
    valids_per_nom = {r["nombre"]: r for r in retirats_valids}
    valids_per_nom.update({d["nombre"]: d for d in descartats_valids})

    a_destruir = []
    omesos = []
    for item in items:
        actual = valids_per_nom.get(item["nombre"])
        if actual is None or actual["tipo"] != item["tipo"]:
            omesos.append((item["nombre"], "ja no és destruïble (potser algú ha revertit la decisió)"))
            continue
        a_destruir.append(actual)

    if not a_destruir:
        return [], omesos

    detall = "; ".join(
        f"{it['nombre']} ({it['num_factura']})" if it.get("num_factura") else it["nombre"]
        for it in a_destruir
    )
    ruta_registro = os.path.join(carpeta_cliente, "registre_destruccions.csv")
    escribir_registre_destruccio(ruta_registro, carpeta, detall, motiu, qui)

    destruits = []
    data_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for it in a_destruir:
        for ruta in it["rutas"]:
            if os.path.exists(ruta):
                os.remove(ruta)
        if it["tipo"] == "descartat":
            escribir_decision(
                carpeta_cliente, it["nombre"], "destruir",
                f"Document destruït el {data_actual}: {motiu}", qui,
            )
        destruits.append(it["nombre"])

    return destruits, omesos


def destruir_client_arxivat(nombre_carpeta_arxivada, motiu, qui):
    """Piso 13N: nomes clients ja arxivats (arxivats/<carpeta>_<data>/)
    -- un client actiu primer s'arxiva (dues portes, mai una). El
    certificat viu a l'arrel (ja no hi haura clientes/<carpeta>/ per
    escriure-hi res un cop fet el rmtree)."""
    carpeta_arxivats = ruta_proyecto("arxivats")
    ruta_origen = os.path.join(carpeta_arxivats, nombre_carpeta_arxivada)
    if not os.path.isdir(ruta_origen):
        raise RuntimeError(f"'{ruta_origen}' no existeix -- aturat, no s'ha esborrat res.")

    n_archivos = sum(len(archivos) for _, _, archivos in os.walk(ruta_origen))
    detall = f"Client arxivat sencer ({nombre_carpeta_arxivada}): {n_archivos} arxius"
    ruta_registro = ruta_proyecto("registre_destruccions.csv")
    escribir_registre_destruccio(ruta_registro, nombre_carpeta_arxivada, detall, motiu, qui)

    shutil.rmtree(ruta_origen)


def auto_recalcular_sumar_informe():
    """Piso 13V: crida sumar.py + informe.py (gratuïts, sense API) just
    despres de destruir documents d'un client actiu -- perque l'informe
    no quedi desquadrat (comptava com a "presents" arxius que ja no hi
    son) fins al proper Processar/RECALCULAR manual. Mateix candau que
    RECALCULAR (app.py, vista Revisió): si hi ha un Processar en marxa,
    s'avisa i se salta -- la destrucció ja s'ha fet, nomes queda
    pendent recalcular a ma des de Revisió."""
    candau = candau_processar_viu()
    if candau:
        st.warning(
            f"Hi ha un Processar en marxa (iniciat per {candau.get('qui')} a les "
            f"{candau.get('data_inici')}) -- sumar+informe NO s'han pogut actualitzar sols. "
            "Recalcula des de Revisió quan acabi."
        )
        return
    with st.spinner("Actualitzant sumar+informe..."):
        for maquina in ["sumar.py", "informe.py"]:
            proceso = subprocess.run(
                [sys.executable, maquina], cwd=RAIZ_PROYECTO, capture_output=True, text=True,
            )
            if proceso.returncode != 0:
                st.error(
                    f"{maquina} ha fallat en actualitzar-se (codi {proceso.returncode}). "
                    "Recalcula des de Revisió."
                )
                return
    st.success("sumar+informe actualitzats.")


def ruta_destino_factures(carpeta, destino, es_lot=False):
    """Piso 13K: es_lot es un eix INDEPENDENT del flux (Compres/Vendes)
    -- abans la guardia de lots despistats ignorava quin flux havia
    triat l'usuari i sempre cridava amb el valor literal "Lot", que
    aterrava al moll de Compres encara que vinguessis de Vendes (bug
    confirmat en camp). Ara cada flux te el seu propi moll de lots."""
    if destino == "Compres":
        if es_lot:
            return ruta_proyecto("clientes", carpeta, "rebudes", "lotes_escaneados")
        return ruta_proyecto("clientes", carpeta, "rebudes", "entrada")
    if destino == "Vendes":
        if es_lot:
            return ruta_proyecto("clientes", carpeta, "apartados", "lotes_vendes_escaneados")
        origen_ingressos = RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS.get(carpeta, "apartados/ingressos")
        return os.path.join(ruta_proyecto("clientes", carpeta), *origen_ingressos.split("/"))
    # destino == "Lot": triat directament al radio (sense passar per la
    # guardia de Compres/Vendes) -- comportament historic, sempre
    # compres, perque aqui no hi ha manera de saber quin flux volia.
    return ruta_proyecto("clientes", carpeta, "rebudes", "lotes_escaneados")


def hash_sha256_bytes(contenido):
    return hashlib.sha256(contenido).hexdigest()


def hash_sha256_archivo(ruta):
    with open(ruta, "rb") as f:
        return hash_sha256_bytes(f.read())


def cargar_indice_hashos(carpeta_cliente, carpetas_a_indexar):
    """Piso 13T: hashos.csv (hash,nombre) -- auto-reparador: es carrega
    el que ja hi ha i, per cada arxiu de les carpetes indicades que
    encara no hi surti, es hasheja i s'hi afegeix (al csv i a l'índex
    en memòria). "Es manté sol" -- no cal cap migració a part, els
    arxius que ja existien abans d'aquest pis es van indexant sols la
    primera vegada que es comproven."""
    ruta_csv = os.path.join(carpeta_cliente, "hashos.csv")
    indice = {}
    nombres_indexados = set()
    if os.path.exists(ruta_csv):
        with open(ruta_csv, encoding="utf-8") as f:
            for fila in csv.DictReader(f):
                indice[fila["hash"]] = fila["nombre"]
                nombres_indexados.add(fila["nombre"])

    filas_nuevas = []
    for carpeta in carpetas_a_indexar:
        if not carpeta or not os.path.isdir(carpeta):
            continue
        for nombre in sorted(os.listdir(carpeta)):
            ruta_archivo = os.path.join(carpeta, nombre)
            if not os.path.isfile(ruta_archivo) or nombre in nombres_indexados:
                continue
            hash_archivo = hash_sha256_archivo(ruta_archivo)
            indice[hash_archivo] = nombre
            filas_nuevas.append((hash_archivo, nombre))
            nombres_indexados.add(nombre)

    if filas_nuevas:
        os.makedirs(carpeta_cliente, exist_ok=True)
        escribir_cabecera = not os.path.exists(ruta_csv)
        with open(ruta_csv, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if escribir_cabecera:
                writer.writerow(["hash", "nombre"])
            writer.writerows(filas_nuevas)

    return indice


def registrar_hash(carpeta_cliente, hash_archivo, nombre):
    """Afegeix una fila nova a hashos.csv -- crida just després de desar
    un arxiu nou (cargar_indice_hashos ja l'hauria trobat sol la propera
    vegada, però registrar-lo ara evita haver-lo de re-hashejar)."""
    ruta_csv = os.path.join(carpeta_cliente, "hashos.csv")
    os.makedirs(carpeta_cliente, exist_ok=True)
    escribir_cabecera = not os.path.exists(ruta_csv)
    with open(ruta_csv, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if escribir_cabecera:
            writer.writerow(["hash", "nombre"])
        writer.writerow([hash_archivo, nombre])


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


def guardar_y_reportar(archivos, carpeta_destino, carpeta_cliente, carpetas_hash_extra=()):
    """Piso 13D: factoritzat perque "Afegir factures" el crida des de
    tres llocs (desar normal, i els dos camins de la guardia de lots
    despistats) sense repetir el mateix bloc tres cops.

    Piso 13R: bug real reproduit en directe -- st.file_uploader no es
    buida sol despres de desar, i l'avis de "es un lot?" es recalcula
    igual a cada rerun mentre el fitxer hi segueixi, aixi que tornava
    a estar visible i clicable DESPRES d'un desat ja fet amb exit.
    Cada clic addicional (doble clic, confusio, qualsevol rerun futur)
    tornava a cridar guardar_archivo, que MAI sobreescriu -- creava
    una copia _2/_3... del mateix contingut. Registre per file_id
    (identificador estable de Streamlit per a AQUESTA pujada concreta,
    Piso 13R) a session_state: cada archivo es desa EXACTAMENT una
    vegada per sessio, independentment de quants cops es re-renderitzi
    o es re-clicki el mateix avis (regla 5: idempotent).

    Piso 13T: idempotencia per CONTINGUT (mai per nom) -- un escaner
    reintrodueix el mateix paper amb un nom nou; el hash SHA-256 el
    reconeix igual. carpetas_hash_extra permet incloure un sibling ja
    conegut ("rebudes/procesadas") a la comprovacio, a mes de la propia
    carpeta_destino."""
    ja_desats = st.session_state.setdefault("afegir_ja_desats", set())
    indice_hash = cargar_indice_hashos(carpeta_cliente, [carpeta_destino, *carpetas_hash_extra])
    nombres_finales = []
    ja_existents = []
    for a in archivos:
        if a.file_id in ja_desats:
            continue
        hash_arxiu = hash_sha256_bytes(a.getvalue())
        if hash_arxiu in indice_hash:
            ja_existents.append((a.name, indice_hash[hash_arxiu]))
            ja_desats.add(a.file_id)
            continue
        nombre_final = guardar_archivo(a, carpeta_destino)
        indice_hash[hash_arxiu] = nombre_final
        registrar_hash(carpeta_cliente, hash_arxiu, nombre_final)
        nombres_finales.append(nombre_final)
        ja_desats.add(a.file_id)

    if nombres_finales:
        st.success(f"S'han desat {len(nombres_finales)} arxius a `{carpeta_destino}`:")
        for nombre in nombres_finales:
            st.write(f"- {nombre}")
    for nombre_pujat, nombre_existent in ja_existents:
        st.warning(f"Aquest arxiu ja existeix com a `{nombre_existent}` -- no s'ha tornat a pujar ({nombre_pujat}).")
    if not nombres_finales and not ja_existents:
        st.info("Aquests arxius ja s'havien desat abans en aquesta sessió -- no s'ha tornat a escriure res.")


def entrada_manual_factura(fila_afegir, carpeta):
    """Piso 13X: tercer mode d'"Afegir factures" -- una fitxa tecleja a
    ma (rebut sense escanejar, nota de despesa), sense passar mai per
    extraer_todas.py. Es desa un JSON directament al MATEIX canal
    "extraidas" que la LLM ja fa servir -- validar.py no es toca gens:
    el seu bucle nomes llista *.json a extraidas/, aritmètica, checksum
    de NIF, duplicats i identitat s'apliquen sols en el proper
    Recalcular, exactament igual que a qualsevol altra fitxa.

    Piso 13X: mateix fantasma "keys estables" que Corregir camps --
    despres de desar, un senyal petit assigna un valor buit a cada key
    "manual_*" ABANS que es tornin a instanciar al proper run, perque la
    propera fitxa tecleja comenci en blanc (mai amb l'anterior parada).
    Assignar un valor buit (mai `pop()`) -- confirmat en directe que
    `pop()` no basta: el frontend del widget li reenvia el seu propi
    valor cachejat abans que el script corri, i el "buit" es perd;
    nomes una assignacio explicita (mateix patró que `entrada_cerca`)
    guanya a aquesta rehidratacio."""
    if st.session_state.pop("manual_reset_senyal", False):
        n_linies_previes = st.session_state.get("manual_n_linies", 1)
        for i in range(n_linies_previes):
            st.session_state[f"manual_tipo_{i}"] = ""
            st.session_state[f"manual_base_{i}"] = ""
            st.session_state[f"manual_cuota_{i}"] = ""
        st.session_state["manual_n_linies"] = 1
        st.session_state["manual_qui"] = ""
        st.session_state["manual_contrapart_nom"] = ""
        st.session_state["manual_contrapart_nif"] = ""
        st.session_state["manual_num_factura"] = ""
        st.session_state["manual_exenta"] = False
        st.session_state["manual_porta_retencio"] = False
        st.session_state["manual_retencio_pct"] = ""
        st.session_state["manual_retencio_cuota"] = ""
        st.session_state["manual_total"] = ""

    st.caption(
        "Escriu les dades directament -- sense escanejar res. Passa per la "
        "MATEIXA xarxa de validació que qualsevol altra fitxa (aritmètica, "
        "NIF, duplicats, identitat) en el proper Recalcular."
    )

    qui_manual = st.text_input("Qui introdueix aquesta factura?", key="manual_qui")

    flux_manual = st.radio("Flux", ["Compres", "Vendes"], horizontal=True, key="manual_flux")
    if flux_manual == "Compres":
        st.caption(f"Receptor: **{fila_afegir['nombre']}** (NIF {fila_afegir['nif']}) -- s'omple sol.")
    else:
        st.caption(f"Emissor: **{fila_afegir['nombre']}** (NIF {fila_afegir['nif']}) -- s'omple sol.")

    col_nom, col_nif = st.columns(2)
    with col_nom:
        contrapart_nom_manual = st.text_input("Contrapart (nom)", key="manual_contrapart_nom")
    with col_nif:
        contrapart_nif_manual = st.text_input("Contrapart (NIF)", key="manual_contrapart_nif")

    col_num, col_data = st.columns(2)
    with col_num:
        num_factura_manual = st.text_input("Número de factura", key="manual_num_factura")
    with col_data:
        data_factura_manual = st.date_input("Data de factura", key="manual_data_factura")

    exempta_manual = st.checkbox("Exempta", key="manual_exenta")
    if exempta_manual:
        st.caption("⚠️ una factura exempta no porta IVA — revisa les línies")

    # Piso 13X: llista dinàmica de línies d'IVA -- comptador propi a
    # session_state (afegir/treure una línia), mateix mode de la resta
    # del formulari (fora de st.form, per a les pistes en viu).
    if "manual_n_linies" not in st.session_state:
        st.session_state["manual_n_linies"] = 1

    lineas_manuals = []
    for i in range(st.session_state["manual_n_linies"]):
        st.caption(f"Línia {i + 1}")
        col_t, col_b, col_c = st.columns(3)
        key_tipo, key_base, key_cuota = f"manual_tipo_{i}", f"manual_base_{i}", f"manual_cuota_{i}"
        if exempta_manual:
            # Assignat ABANS de crear el text_input d'aquesta key en
            # aquesta mateixa execució -- mateixa cascada que Corregir camps.
            st.session_state[key_tipo] = "0"
            st.session_state[key_cuota] = "0"
        with col_t:
            t = st.text_input(f"Tipus IVA [{i + 1}]", key=key_tipo)
        with col_b:
            b = st.text_input(f"Base [{i + 1}]", key=key_base)
        with col_c:
            c = st.text_input(f"Quota [{i + 1}]", key=key_cuota)
        base_num, tipo_num = a_numero(b), a_numero(t)
        if base_num is not None and tipo_num is not None:
            st.caption(f"{base_num} × {tipo_num}% = {base_num * tipo_num / 100:.2f} — coincideix amb el paper?")
        lineas_manuals.append((t, b, c))

    col_add, col_treure = st.columns(2)
    with col_add:
        if st.button("+ Afegir línia", key="manual_afegir_linia"):
            st.session_state["manual_n_linies"] += 1
            st.rerun()
    with col_treure:
        if st.button("- Treure línia", key="manual_treure_linia", disabled=st.session_state["manual_n_linies"] <= 1):
            st.session_state["manual_n_linies"] -= 1
            st.rerun()

    porta_retencio = st.checkbox("Aquesta factura porta retenció?", key="manual_porta_retencio")
    if porta_retencio:
        col_pct, col_cuota_ret = st.columns(2)
        with col_pct:
            retencio_pct_manual = st.text_input("% Retenció", key="manual_retencio_pct")
        with col_cuota_ret:
            retencio_cuota_manual = st.text_input("Quota retenció", key="manual_retencio_cuota")
    else:
        retencio_pct_manual, retencio_cuota_manual = "0", "0.0"

    suma_calculada = sum((a_numero(b) or 0) + (a_numero(c) or 0) for _, b, c in lineas_manuals)
    st.caption(f"Suma de línies (bases+quotes): {suma_calculada:.2f} — coincideix amb el total del paper?")
    total_manual = st.text_input("Total", key="manual_total")

    if st.button("Desar fitxa", type="primary", key="manual_desar"):
        errors_manual = []
        if not qui_manual:
            errors_manual.append("Cal escriure qui introdueix aquesta factura.")
        if not contrapart_nom_manual:
            errors_manual.append("Cal escriure el nom de la contrapart.")
        if not contrapart_nif_manual:
            errors_manual.append("Cal escriure el NIF de la contrapart.")
        if not num_factura_manual:
            errors_manual.append("Cal escriure el número de factura.")
        if a_numero(total_manual) is None:
            errors_manual.append("El total no és un import vàlid.")

        if errors_manual:
            for error in errors_manual:
                st.error(error)
        else:
            carpeta_cliente_manual = ruta_proyecto("clientes", carpeta)
            if flux_manual == "Compres":
                proveedor_final, nif_proveedor_final = contrapart_nom_manual, contrapart_nif_manual
                receptor_final, nif_receptor_final = fila_afegir["nombre"], fila_afegir["nif"]
                carpeta_extraidas_manual = os.path.join(carpeta_cliente_manual, "rebudes", "extraidas")
            else:
                proveedor_final, nif_proveedor_final = fila_afegir["nombre"], fila_afegir["nif"]
                receptor_final, nif_receptor_final = contrapart_nom_manual, contrapart_nif_manual
                carpeta_extraidas_manual = os.path.join(carpeta_cliente_manual, "apartados", "ingressos_extraidas")

            fitxa_manual = {
                "proveedor": proveedor_final,
                "nif_proveedor": nif_proveedor_final,
                "num_factura": num_factura_manual,
                "fecha_factura": data_factura_manual.isoformat(),
                "receptor": receptor_final,
                "nif_receptor": nif_receptor_final,
                "lineas_iva": [
                    {"tipo_iva": a_numero(t), "base": a_numero(b), "cuota": a_numero(c)}
                    for t, b, c in lineas_manuals
                ],
                "total": a_numero(total_manual),
                "retencion_pct": a_numero(retencio_pct_manual) or 0,
                "retencion_cuota": a_numero(retencio_cuota_manual) or 0.0,
                "exenta": exempta_manual,
                "observaciones": None,
                "origen": "manual",
                "qui": qui_manual,
                "data_entrada": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

            # Piso 13T: mateix patró de collisió que retirar_error -- mai
            # sobreescriu, bump numèric si dues entrades manuals cauen
            # dins el mateix segon.
            os.makedirs(carpeta_extraidas_manual, exist_ok=True)
            nombre_base_manual = f"manual_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            nombre_final_manual = f"{nombre_base_manual}.json"
            contador_manual = 2
            while os.path.exists(os.path.join(carpeta_extraidas_manual, nombre_final_manual)):
                nombre_final_manual = f"{nombre_base_manual}_{contador_manual}.json"
                contador_manual += 1
            with open(os.path.join(carpeta_extraidas_manual, nombre_final_manual), "w", encoding="utf-8") as f:
                json.dump(fitxa_manual, f, ensure_ascii=False, indent=2)

            # Piso 14: entrada manual es una mutacio com qualsevol altra
            # -- ha d'encendre el semafor "sense recalcular" igual que
            # decisions/moviments/retirs/correccions. escribir_entrada_manual
            # deixa constancia a un ledger propi (mateix format que
            # errors_retirats/registre.csv) perque
            # contar_manuals_sense_recalcular pugui detectar-ho.
            escribir_entrada_manual(carpeta_cliente_manual, nombre_final_manual, qui_manual)

            st.success(
                f"Fitxa manual desada: `{nombre_final_manual}`. "
                "Recalcula per veure-la als resultats."
            )
            st.session_state["manual_reset_senyal"] = True
            st.rerun()


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
    explicitamente, nunca falla en silencio (regla 4/10).

    Piso 13E: abans el boto se deshabilitava en silenci si l'arxiu no
    existia -- exactament el que la regla 10 prohibeix. Ara sempre es
    clicable (mateix patro que Aprovar/Descartar): la comprovacio es
    fa dins del clic, i si falta mostra la RUTA COMPLETA que ha
    intentat, perque qualsevol fallada s'autodiagnostiqui sola."""
    if st.button(etiqueta, key=key):
        if not os.path.exists(ruta_absoluta):
            st.error(f"No s'ha trobat l'arxiu:\n\n`{ruta_absoluta}`")
        else:
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
                    # Piso 13V: la portada mai ha de mostrar aquest
                    # client mort ni un enllaç trencat, ni un segon.
                    regenerar_index_clientes()
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


def a_numero(valor):
    """Piso 13X: mateix algorisme tolerant que validar.py (duplicat a
    proposit, cap "maquina" n'importa una altra) -- converteix text en
    format espanyol ("1.234,56", "1234,56") a float, MAI llança
    excepcio. Buit o il·legible -> None."""
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
    ("el cliente no aparece ni como emisor ni como receptor",
     "el client no apareix ni com a emissor ni com a receptor"),
    ("las dos partes son el cliente", "les dues parts són el client"),
    ("importe numérico ilegible en", "import numèric il·legible a"),
    ("marcada como exenta pero tiene líneas con IVA -- corrige las líneas (tipo 0, cuota 0) o desmarca exenta",
     "marcada com a exempta però té línies amb IVA -- corregeix les línies (tipus 0, quota 0) o desmarca exempta"),
    ("retención no cuadra:", "la retenció no quadra:"),
    ("pero retencion_cuota indica", "però la retenció indica"),
    ("IVA incluido sin desglosar (tipo impreso:", "IVA inclòs sense desglosar (tipus imprès:"),
    ("ninguno)", "cap)"),
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
    Sin archivo, o vacio, devuelve {}.

    Piso 13M: decisions.csv ya es llibre major (cada decision es una
    fila nova, mai una sobreescriptura) -- l'estat EFECTIU d'un
    archiu es la seva ULTIMA fila. Si aquesta ultima fila es
    "revertir", NO hi ha decisio efectiva: es treu del diccionari en
    comptes de deixar-hi la fila "revertir" (aixi "if decision:" i
    "decision.get('accion') == 'aprovar'" seguixen funcionant igual a
    sumar.py/informe.py/aqui mateix, sense haver de tocar cap altre
    lloc que ja les fa servir)."""
    ruta = os.path.join(carpeta_cliente, "decisions.csv")
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


def estat_efectiu(decision):
    """Piso 13Y: igual que en sumar.py/informe.py -- 'aprovar',
    'descartar', o None a partir d'una decisio ja obtinguda de
    cargar_decisiones()."""
    return decision.get("accion") if decision else None


def historial_decisiones(carpeta_cliente, archivo):
    """Piso 13M: totes les files d'aquest archiu a decisions.csv, en
    ordre cronologic (inclou aprovar/descartar/revertir) -- per a
    l'auditoria completa (badge de "desfet N vegades", historial
    compacte a l'Excel/informe)."""
    ruta = os.path.join(carpeta_cliente, "decisions.csv")
    if not os.path.exists(ruta):
        return []
    with open(ruta, encoding="utf-8") as f:
        return [fila for fila in csv.DictReader(f) if fila.get("archivo") == archivo]


TRADUCCION_ACCION = {"aprovar": "aprovada", "descartar": "descartada", "revertir": "revertit"}


def resumen_historial_decisiones(historial):
    """Piso 13M: text compacte tipus "descartada per X el D — revertit
    per Y el D2" -- None si no hi ha cap "revertir" (sense soroll per
    a les fitxes que mai s'han tocat, regla 10)."""
    if not any(fila.get("accion") == "revertir" for fila in historial):
        return None
    return " — ".join(
        f"{TRADUCCION_ACCION.get(fila.get('accion'), fila.get('accion'))} per {fila.get('qui')} el {fila.get('data')}"
        for fila in historial
    )


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
    """Piso 11A: primera funcion que ESCRIBE decisions.csv.

    Piso 13M: decisions.csv es un llibre major -- NOMES CREIX. Abans
    (Piso 11A) aquesta funcio rellegia tot el fitxer i sobreescrivia
    la fila anterior d'aquest archivo; ara sempre AFEGEIX una fila
    nova, mai reescriu ni esborra les que ja hi son. accion admet ara
    "revertir" a mes d'"aprovar"/"descartar" -- desfer una decisio es
    una fila mes, no una edicio. L'estat efectiu d'un archiu es la
    seva ULTIMA fila (cargar_decisiones ho aplica).

    Piso 11C: 'data' opcional -- una accion en lote captura UNA marca
    de tiempo antes del bucle y la pasa a cada llamada, para que las
    N filas del mismo lote compartan exactamente la misma data (no
    unos milisegundos distintos cada una)."""
    ruta = os.path.join(carpeta_cliente, "decisions.csv")
    escribir_cabecera = not os.path.exists(ruta)
    with open(ruta, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CAMPOS_DECISIONS_CSV)
        if escribir_cabecera:
            writer.writeheader()
        writer.writerow({
            "archivo": archivo,
            "accion": accion,
            "nota": nota,
            "qui": qui,
            "data": data or datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })


def revertir_decision(carpeta_cliente, archivo, motiu, qui):
    """Piso 13M: desfa la decisio EFECTIVA d'aquest archiu -- escriu
    una fila "revertir" nova (mai esborra res). Si ja no hi ha cap
    decisio efectiva (algu ja l'ha desfet, o mai n'hi va haver cap),
    NO escriu res i retorna None -- el crider mostra l'error visible
    (regla 10, i la trampa explicita del Piso 13M: revertir dues
    vegades seguides ha de quedar bloquejat amb missatge clar, mai en
    silenci). Si nota queda buida, es genera una referencia llegible
    a la decisio que es desfa (evita haver d'afegir columnes noves a
    un csv que ja existeix en clients reals)."""
    decision_actual = cargar_decisiones(carpeta_cliente).get(archivo)
    if decision_actual is None:
        return None
    nota_final = motiu or (
        f"Reverteix '{decision_actual.get('accion')}' de {decision_actual.get('qui')} "
        f"({decision_actual.get('data')})"
    )
    escribir_decision(carpeta_cliente, archivo, "revertir", nota_final, qui)
    return decision_actual


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


def historial_correccions(carpeta_cliente, nombre):
    """Piso 13X: mateix esperit que historial_decisiones (13M) però per a
    correccions.csv -- agrupa les files CONSECUTIVES d'aquest arxiu que
    comparteixen (qui, data) en un "grup" (escribir_correccion escriu
    totes les files d'un mateix clic amb la mateixa marca de temps).
    Retorna una llista de grups en ordre cronològic, cada un
    {"qui", "data", "canvis": [fila_csv, ...]} -- "Desfer últim canvi"
    nomes necessita el darrer."""
    ruta = os.path.join(carpeta_cliente, "correccions.csv")
    if not os.path.exists(ruta):
        return []
    grups = []
    with open(ruta, encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            if fila.get("arxiu") != nombre:
                continue
            clau = (fila.get("qui"), fila.get("data"))
            if grups and (grups[-1]["qui"], grups[-1]["data"]) == clau:
                grups[-1]["canvis"].append(fila)
            else:
                grups.append({"qui": fila.get("qui"), "data": fila.get("data"), "canvis": [fila]})
    return grups


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


def contar_decisiones_sin_recalcular(carpeta_cliente):
    """Compara la fecha de cada decision (formato propio, parseable --
    la escribe escribir_decision) contra el mtime de informe_2026.html.
    Sin informe todavia, cuentan todas. Fechas no parseables (alguien
    las escribio a mano en otro formato) no cuentan -- no rompen nada.

    Piso 13V: abans rebia el diccionari ja col·lapsat de
    cargar_decisiones() (que POPA qualsevol archiu la última fila del
    qual sigui "revertir") -- un Desfer era estructuralment invisible
    per aquest comptador. Ara llegeix decisions.csv EN CRU i compta
    QUALSEVOL fila (aprovar/descartar/revertir), mateix patró que
    contar_moviments_sense_recalcular i contar_retirs_sense_recalcular."""
    ruta_decisions = os.path.join(carpeta_cliente, "decisions.csv")
    if not os.path.exists(ruta_decisions):
        return 0
    ruta_informe = os.path.join(carpeta_cliente, "informe_2026.html")
    mtime_informe = os.path.getmtime(ruta_informe) if os.path.exists(ruta_informe) else 0
    contador = 0
    with open(ruta_decisions, encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            try:
                marca = datetime.strptime(fila.get("data", ""), "%Y-%m-%d %H:%M:%S").timestamp()
            except ValueError:
                continue
            if marca > mtime_informe:
                contador += 1
    return contador


def contar_retirs_sense_recalcular(carpeta_cliente):
    """Piso 13V: mateix patró exacte que contar_moviments_sense_recalcular
    però per a errors_retirats/registre.csv (escrit per retirar_error) --
    abans Retirar no tenia cap senyal de "sense recalcular"."""
    ruta_registre = os.path.join(carpeta_cliente, "errors_retirats", "registre.csv")
    if not os.path.exists(ruta_registre):
        return 0
    ruta_informe = os.path.join(carpeta_cliente, "informe_2026.html")
    mtime_informe = os.path.getmtime(ruta_informe) if os.path.exists(ruta_informe) else 0
    contador = 0
    with open(ruta_registre, encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            try:
                marca = datetime.strptime(fila.get("data", ""), "%Y-%m-%d %H:%M:%S").timestamp()
            except ValueError:
                continue
            if marca > mtime_informe:
                contador += 1
    return contador


def escribir_entrada_manual(carpeta_cliente, archivo, qui):
    """Piso 14: ledger propi per a entrada_manual_factura -- mateix
    format exacte (arxiu, qui, data) que errors_retirats/registre.csv,
    perque contar_manuals_sense_recalcular pugui detectar-ho amb el
    mateix patró que ja fan servir decisions/moviments/retirs. Abans
    d'aquest piso una fitxa manual no encenia cap senyal de "sense
    recalcular" -- una mutació real i invisible."""
    ruta = os.path.join(carpeta_cliente, "entrades_manuals.csv")
    escribir_cabecera = not os.path.exists(ruta)
    with open(ruta, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if escribir_cabecera:
            writer.writerow(["arxiu", "qui", "data"])
        writer.writerow([archivo, qui, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])


def contar_manuals_sense_recalcular(carpeta_cliente):
    """Piso 14: mateix patró exacte que contar_retirs_sense_recalcular
    però per a entrades_manuals.csv (escrit per escribir_entrada_manual)."""
    ruta_manuals = os.path.join(carpeta_cliente, "entrades_manuals.csv")
    if not os.path.exists(ruta_manuals):
        return 0
    ruta_informe = os.path.join(carpeta_cliente, "informe_2026.html")
    mtime_informe = os.path.getmtime(ruta_informe) if os.path.exists(ruta_informe) else 0
    contador = 0
    with open(ruta_manuals, encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            try:
                marca = datetime.strptime(fila.get("data", ""), "%Y-%m-%d %H:%M:%S").timestamp()
            except ValueError:
                continue
            if marca > mtime_informe:
                contador += 1
    return contador


def tiene_correccion_pendiente(carpeta_cliente, nombre):
    """Piso 13H: mateix patro que contar_decisiones_sin_recalcular pero
    per a correccions.csv -- hi ha alguna fila d'aquesta fitxa escrita
    DESPRES de l'ultim informe_2026.html generat. Apareix nomes en
    guardar una correccio, desapareix sol quan RECALCULAR el
    regenera (regla 10: mai una senyal que menteixi)."""
    ruta_correccions = os.path.join(carpeta_cliente, "correccions.csv")
    if not os.path.exists(ruta_correccions):
        return False
    ruta_informe = os.path.join(carpeta_cliente, "informe_2026.html")
    mtime_informe = os.path.getmtime(ruta_informe) if os.path.exists(ruta_informe) else 0
    with open(ruta_correccions, encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            if fila.get("arxiu") != nombre:
                continue
            try:
                marca = datetime.strptime(fila.get("data", ""), "%Y-%m-%d %H:%M:%S").timestamp()
            except ValueError:
                continue
            if marca > mtime_informe:
                return True
    return False


def contar_moviments_sense_recalcular(carpeta_cliente):
    """Piso 13L: mateix patró exacte que contar_decisiones_sin_recalcular
    però per a moviments_flux.csv (escrit per migrar_lot.moure_de_flux)."""
    ruta_moviments = os.path.join(carpeta_cliente, "moviments_flux.csv")
    if not os.path.exists(ruta_moviments):
        return 0
    ruta_informe = os.path.join(carpeta_cliente, "informe_2026.html")
    mtime_informe = os.path.getmtime(ruta_informe) if os.path.exists(ruta_informe) else 0
    contador = 0
    with open(ruta_moviments, encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            try:
                marca = datetime.strptime(fila.get("data", ""), "%Y-%m-%d %H:%M:%S").timestamp()
            except ValueError:
                continue
            if marca > mtime_informe:
                contador += 1
    return contador


def obtener_moviment(carpeta_cliente, nombre_base):
    """Piso 13L: mateix patró que cargar_decisiones -- l'últim moviment
    d'aquest document a moviments_flux.csv, si n'hi ha cap. Es fa
    servir per mostrar "Mogut" a la targeta en comptes del formulari
    (mateix motiu que decision/camps_corregits: un rerun de fragment
    no fa desaparèixer la targeta de la llista exterior)."""
    ruta_moviments = os.path.join(carpeta_cliente, "moviments_flux.csv")
    if not os.path.exists(ruta_moviments):
        return None
    ultimo = None
    with open(ruta_moviments, encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            if fila.get("arxiu") == nombre_base:
                ultimo = fila
    return ultimo


def formatar_moviment_costat(valor):
    """Piso 13Q: moviments_flux.csv guarda "carpeta:flux" quan el
    moviment ha estat entre CLIENTS diferents (migrar_lot.moure_a_client)
    i nomes el nom pla del flux quan ha estat dins del mateix client
    (moure_de_flux, sense canvis -- cap moviment antic es reescriu).
    Aquesta funció tradueix el primer format a "Nom (flux)" per mostrar-ho;
    el segon es retorna tal qual."""
    valor = valor or ""
    if ":" not in valor:
        return valor
    carpeta_val, flux_val = valor.split(":", 1)
    nom = next((f["nombre"] for f in leer_clientes() if f["carpeta"] == carpeta_val), carpeta_val)
    return f"{nom} ({flux_val})"


def sembla_factura_de_lautre_flux(datos, flujo, nif_client):
    """Piso 13L: senyal (mai automàtic, regla 10 en positiu) -- només
    ben definit per a rebudes: si el motiu d'identitat ha saltat I el
    client resulta ser qui EMET (no qui rep, com tocaria en una
    compra), sembla realment una factura de VENDES mal enviada pel
    canal de compres. No hi ha senyal simètric fiable per a ingressos
    (ser emissor hi és un estat vàlid, el "giro" del Piso 6B, no un
    error)."""
    if flujo != "rebudes":
        return False
    if not any("nif_receptor no coincide" in m for m in (datos.get("motivos") or [])):
        return False
    return normalizar_nif(datos.get("nif_proveedor")) == normalizar_nif(nif_client)


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


# Piso 13V: etiqueta visible (nom + color) per a cada flux intern --
# "rebudes"/"ingressos" son els noms de carpeta de tota la vida, mai
# ensenyats a l'usuari; "Compres"/"Vendes" son el que ja es veu a la
# resta de la interfície (ex. ruta_destino_factures).
FLUX_ETIQUETA = {"rebudes": "🔵 Compres", "ingressos": "🟢 Vendes"}
FLUX_INTERN = {"Compres": "rebudes", "Vendes": "ingressos"}


def selector_flux(key_base):
    """Piso 13V: control segmentat "Tots | Compres | Vendes" per a cada
    secció paginada de Revisió -- conviu amb entrada_cerca (els dos
    filtres s'apliquen junts sobre el conjunt COMPLET, abans de
    paginar()). key_base ha de portar la carpeta del client (mateix
    motiu que entrada_cerca: en canviar de client, la selecció de
    l'anterior no s'ha d'arrossegar). Retorna "rebudes"/"ingressos" o
    None (= Tots, cap filtre)."""
    seleccio = st.segmented_control(
        "Flux", ["Tots", "Compres", "Vendes"], default="Tots", required=True, key=f"{key_base}_flux",
    )
    return FLUX_INTERN.get(seleccio)


def normalizar_cerca(texto):
    """Piso 13U: minúscules + sense accents (NFD treu els diacrítics) +
    coma->punt (perquè "1234,56" i "1234.56" es trobin igual) -- aplicat
    als DOS costats (consulta i text buscable) de la cerca de Revisió."""
    texto = str(texto or "").replace(",", ".")
    sense_accents = "".join(
        c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn"
    )
    return sense_accents.lower()


def texto_buscable_ficha(nombre, datos):
    """Piso 13U: tots els camps VISIBLES a la targeta (contrapart, NIF,
    núm. factura, imports, nom d'arxiu, motius ja TRADUÏTS -- el que es
    veu, no el cru intern) concatenats i normalitzats per a la cerca."""
    partes = [
        datos.get("contrapart_nom"),
        datos.get("contrapart_nif"),
        datos.get("num_factura"),
        datos.get("total"),
        nombre,
    ]
    for linea in datos.get("lineas_iva") or []:
        partes.append(linea.get("base"))
        partes.append(linea.get("cuota"))
    partes.extend(traducir_motivo(m) for m in (datos.get("motivos") or []))
    return normalizar_cerca(" ".join(str(p) for p in partes if p is not None))


def entrada_cerca(etiqueta, key_base):
    """Piso 13U: camp de cerca reutilitzable per a cada secció paginada
    de Revisió. FORA de cap st.form (rerun natiu a cada tecla, sense
    esperar Tab/Enter) i amb key= FIXA (sobreviu als reruns).

    Netejar-lo MAI toca st.session_state[key_input] directament -- el
    mateix fantasma "cannot be modified after the widget... is
    instantiated" que ha causat la regressió greu d'aquest pis. El botó
    "✕ Netejar" marca un senyal petit consumit ABANS que el widget
    existeixi al proper run (mateix patró que "vista_radio")."""
    key_input = f"{key_base}_input"
    key_senyal_netejar = f"{key_base}_netejar_senyal"
    if st.session_state.pop(key_senyal_netejar, False):
        st.session_state[key_input] = ""

    col_cerca, col_netejar = st.columns([5, 1])
    with col_cerca:
        query = st.text_input(
            etiqueta, key=key_input, placeholder="contrapart, NIF, núm. factura, import...",
        )
    with col_netejar:
        st.markdown("<div style='height: 1.7rem'></div>", unsafe_allow_html=True)
        if query and st.button("✕ Netejar", key=f"{key_input}_boto_netejar"):
            st.session_state[key_senyal_netejar] = True
            st.rerun()
    return query


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
def tarjeta_revisio(nombre, datos, origen, carpeta_cliente, qui, prefijo, flujo, nif_client):
    """Piso 11A: tarjeta de UNA ficha (PENDENT u OK) con las mismas dos
    acciones (Aprovar/Descartar) -- simetria del punto 3: se puede
    descartar un OK con nota igual que se aprova un REVISAR.

    Piso 13L: flujo ("rebudes"/"ingressos") y nif_client permiten
    ofrecer "Moure a l'altre flux" (solo si la fitxa no tiene decisió
    encara) y, cuando aplica, el suggeriment visible de
    sembla_factura_de_lautre_flux.

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
    nombre_base = os.path.splitext(nombre)[0]
    moviment = obtener_moviment(carpeta_cliente, nombre_base)

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
            # Piso 13K: contrapart (validar.py), no "proveedor" a seques --
            # a ingressos on el client es l'emissor (factura propia),
            # "proveedor" era el propi client, no el seu comprador.
            st.markdown(f"**{datos.get('contrapart_nom')}**")
            st.caption(
                f"{FLUX_ETIQUETA[flujo]} · NIF {datos.get('contrapart_nif')} · "
                f"Factura {datos.get('num_factura')} · "
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
            # Piso 14: guàrdia de multituds -- suggeriment nomes textual,
            # mai automatitzat (cap re-extracció ni moviment sol). El
            # camí real ja existeix: tornar a pujar l'arxiu com a Lot
            # d'escàner des d'Afegir factures.
            if any("possible arxiu multi-factura" in m for m in motivos):
                st.caption(
                    "💡 Torna a pujar-lo com a Lot d'escàner des d'Afegir factures "
                    "perquè es trocegi automàticament."
                )

            # Piso 13Z: "IVA inclòs" sense desglosar -- la màquina mai
            # inventa el desglose sola (extraer_todas.py ja ho deixa
            # lineas_iva buida), però un cop la PERSONA confirma el tipus
            # (imprès al paper, o triat aquí), la inversa matemàtica quadra
            # sola per a QUALSEVOL tipus per construcció -- l'àncora és el
            # paper, mai l'aritmètica. Un clic escriu la línia com a
            # correcció firmada (mateix camí que "Corregir camps").
            iva_inclos_sense_desglosar = (
                bool(datos.get("iva_inclos_detectat")) and not (datos.get("lineas_iva") or [])
            )
            if iva_inclos_sense_desglosar and not tiene_correccion_pendiente(carpeta_cliente, nombre):
                total_factura = datos.get("total") or 0
                tipus_impres = datos.get("tipus_impres")
                tipus_a_mostrar = [tipus_impres] if tipus_impres is not None else [21, 10, 4]
                st.caption(
                    "Desglose invers (base = total / (1 + tipus%), quota = total - base) -- "
                    "tria el tipus que digui el paper:"
                )
                cols_desglose = st.columns(len(tipus_a_mostrar))
                for col_desglose, tipo in zip(cols_desglose, tipus_a_mostrar):
                    base_calc = round(total_factura / (1 + tipo / 100), 2)
                    quota_calc = round(total_factura - base_calc, 2)
                    with col_desglose:
                        st.caption(f"{tipo}%: base {base_calc} €, quota {quota_calc} €")
                        if st.button(
                            "Aplicar desglose" if tipus_impres is not None else f"Aplicar {tipo}%",
                            key=f"{prefijo}_desglose_{tipo}_{nombre}",
                        ):
                            cambios = [
                                ("lineas_iva[0].tipo_iva", "", str(tipo)),
                                ("lineas_iva[0].base", "", str(base_calc)),
                                ("lineas_iva[0].cuota", "", str(quota_calc)),
                            ]
                            escribir_correccion(
                                carpeta_cliente, nombre, cambios,
                                f"desglose invers, IVA inclòs {tipo}%", qui,
                            )
                            st.toast(
                                f"Desglose aplicat: {tipo}% -> base {base_calc} €, quota {quota_calc} €",
                                icon="✅",
                            )
                            st.rerun()
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
            if tiene_correccion_pendiente(carpeta_cliente, nombre):
                st.info("🕓 Corregit (pendent de recalcular)")
            # Piso 13X: historial visible SEMPRE (no nomes despres de
            # recalcular, com camps_corregits de dalt) -- llegeix
            # correccions.csv en cru, mateix patró que tiene_correccion_pendiente.
            grups_correccio_card = historial_correccions(carpeta_cliente, nombre)
            if grups_correccio_card:
                with st.expander(
                    f"Historial de correccions ({len(grups_correccio_card)})",
                    key=f"hist_correccio_{prefijo}_{nombre}",
                ):
                    for grup in grups_correccio_card:
                        resum_grup = "; ".join(
                            f"{c['camp']}: {c['valor_antic']} → {c['valor_nou']}" for c in grup["canvis"]
                        )
                        st.caption(f"{grup['data']} per {grup['qui']}: {resum_grup}")
            st.caption(nombre)
        with col_der:
            if ruta_original and extension in EXTENSIONES_IMAGEN:
                st.image(ruta_original)
            elif not ruta_original and datos.get("origen") == "manual":
                # Piso 14: mateix fallback que informe.py::tarjeta_factura
                # ja fa des de Piso 13X -- una fitxa manual MAI té
                # original, mostrar el botó "Obrir original" nomes portava
                # a "No s'ha trobat l'arxiu: ``" (ruta buida), confús i
                # sense explicar per què (regla 10).
                st.info(f"📝 ENTRADA MANUAL\n\n{datos.get('qui')}, {datos.get('data_entrada')}")
            else:
                if ruta_original and extension == ".pdf":
                    imagen_pdf = previsualizar_pdf(ruta_original, os.path.getmtime(ruta_original))
                    if imagen_pdf:
                        st.image(imagen_pdf)
                boton_obrir("Obrir original", ruta_original or "", key=f"{prefijo}_original_{nombre}")

        # Piso 13M: el bloc de moviment es INDEPENDENT del de decisio --
        # una fitxa pot alhora haver estat moguda (historic) i estar
        # pendent d'una decisio nova al seu flux actual (abans era un
        # if/elif/else excloent que amagava el formulari si hi havia
        # moviment). Es llegeix en cada render, mai es confia en
        # l'estat que va passar el bucle exterior.
        if moviment:
            motiu_mostrat = f" — _{moviment.get('motiu')}_" if moviment.get("motiu") else ""
            de_mostrat = formatar_moviment_costat(moviment.get("de"))
            a_mostrat = formatar_moviment_costat(moviment.get("a"))
            st.info(
                f"↔️ Mogut de {de_mostrat} a {a_mostrat} per {moviment.get('qui')} "
                f"el {moviment.get('data')}{motiu_mostrat} · Recalcula per refer els veredictes."
            )
            with st.form(key=f"form_tornar_{prefijo}_{nombre}", border=False):
                motiu_tornar = st.text_input("Motiu (opcional)", key=f"{prefijo}_motiu_tornar_{nombre}")
                click_tornar = st.form_submit_button(f"Tornar a {de_mostrat}")
            if click_tornar:
                # Piso 13Q: "a" amb ":" vol dir que el moviment va ser
                # ENTRE clients (moure_a_client) -- la tornada ha de fer
                # el mateix camí a l'inrevés. Sense ":" es el format pla
                # d'un moviment de flux (moure_de_flux), tal com sempre.
                a_valor = moviment.get("a")
                if ":" in a_valor:
                    carpeta_actual_mov, flux_actual_mov = a_valor.split(":", 1)
                    de_valor = moviment.get("de")
                    if ":" in de_valor:
                        carpeta_origen_mov, flux_origen_mov = de_valor.split(":", 1)
                    else:
                        carpeta_origen_mov, flux_origen_mov = os.path.basename(carpeta_cliente), de_valor
                    migrar_lot.moure_a_client(
                        carpeta_actual_mov, carpeta_origen_mov, [(nombre_base, flux_actual_mov)],
                        motiu_tornar or "Moviment individual des de Revisió (tornada)", qui,
                        flujo_desti=flux_origen_mov,
                    )
                else:
                    migrar_lot.moure_de_flux(
                        os.path.basename(carpeta_cliente), [(nombre_base, a_valor)],
                        motiu_tornar or "Moviment individual des de Revisió (tornada)", qui,
                    )
                st.toast(f"Tornat a {de_mostrat}. Recalcula per refer els veredictes.", icon="↩️")
                st.rerun(scope="fragment")

        historial = historial_decisiones(carpeta_cliente, nombre)
        n_revertits = sum(1 for f in historial if f.get("accion") == "revertir")
        if n_revertits:
            st.caption(f"↩️ Decisió desfeta ({n_revertits} {'vegada' if n_revertits == 1 else 'vegades'})")

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

        # Piso 13R: endreça -- Aprovar/Descartar es queden sempre visibles
        # (l'acció principal), la resta d'eines (Corregir camps, Moure a
        # l'altre flux, Moure a un altre client, Piso 13N/13Q) es reculen
        # dins d'un "⋮ Més accions" per no saturar la targeta. Cap flux
        # intern canvia -- mateixes crides, mateixos paràmetres, només
        # canvia on viuen els widgets. Els avisos de suggeriment es queden
        # SEMPRE visibles fora del popover (avisar no és una acció que
        # calgui amagar, i sense decisió encara té sentit oferir-los).
        flujo_contrari = "Vendes" if flujo == "rebudes" else "Compres"
        clients_altres = [f for f in leer_clientes() if f["carpeta"] != os.path.basename(carpeta_cliente)]
        if not decision:
            if prefijo == "pendent" and sembla_factura_de_lautre_flux(datos, flujo, nif_client):
                st.warning(f"Sembla una factura de {flujo_contrari.upper()} — vols moure-la?")
            if datos.get("suggerit_carpeta"):
                st.warning(f"Sembla de **{datos.get('suggerit_nom')}** — vols moure-ho?")

        # Piso 13S: clau del popover "⋮ Més accions" de sota -- st.popover
        # accepta key= des d'aquesta versió de Streamlit (no ho feia al
        # 13H). Posar aquesta clau a False just abans de CADA rerun que
        # ja existeix dins seu (moure de flux, moure a client, guardar/
        # cancelar la correcció) el tanca de veritat -- confirmat amb un
        # mini-repro aïllat abans de tocar aquest codi.
        popover_key = f"popover_{prefijo}_{nombre}"

        def _tancar_popover():
            st.session_state[popover_key] = False

        # Piso 11B: "Corregir camps" -- capa de correccio, mai edita
        # extraidas/. Camps precarregats amb el valor ACTUAL; nomes es
        # desa el que de veritat canvia.
        #
        # Piso 13H: abans era un st.popover -- reproduit en directe
        # (captura + correccions.csv real) que en guardar la correccio
        # SI s'escrivia be, pero el popover es quedava obert sense cap
        # senyal visible (regla 10: violacio). Causa: Streamlit no te
        # cap API per tancar un popover programaticament -- el seu
        # estat obert/tancat viu al client, independent del rerun de
        # fragment. st.dialog SI es tanca sol en fer st.rerun() des de
        # dins seu (aixi ho documenta Streamlit per a aquest patro
        # exacte de "formulari modal"), aixi que es canvia pel modal
        # en comptes de lluitar contra el popover.
        #
        # Piso 13S: el mateix fantasma reaparegut un nivell amunt -- ARA
        # es aquest dialog qui viu DINS del popover "⋮ Més accions", i
        # obrir-lo no tancava el popover que el conté (reproduit en
        # directe, captura feta). on_dismiss tanca el popover si es
        # cancel·la sense guardar; el camí de "Guardar" ho fa explicitament
        # just abans del seu propi st.rerun() de sempre.
        # Piso 13X: totes les keys que aquest diàleg fa servir -- calen per
        # "oblidar" qualsevol edició abandonada (vegeu _reiniciar_correccio).
        n_linies_actual = len(datos.get("lineas_iva") or [])

        def _keys_correccio():
            keys = [f"{prefijo}_correccio_{camp}_{nombre}" for camp in CAMPOS_CORREGIBLES_TOP if camp != "exenta"]
            keys.append(f"{prefijo}_correccio_exenta_{nombre}")
            for i in range(n_linies_actual):
                keys += [
                    f"{prefijo}_correccio_tipo_{i}_{nombre}",
                    f"{prefijo}_correccio_base_{i}_{nombre}",
                    f"{prefijo}_correccio_cuota_{i}_{nombre}",
                ]
            keys.append(f"{prefijo}_correccio_motiu_{nombre}")
            return keys

        def _reiniciar_correccio():
            """Piso 13X: FORENSE confirmat en directe -- les keys d'aquest
            diàleg son estables entre obertures. Esborrar un camp i tancar
            SENSE guardar mai escriu res a correccions.csv (confirmat), però
            el valor esborrat es queda parat a session_state; en reobrir el
            MATEIX diàleg (sense recalcular pel mig), value= s'ignora
            (Streamlit no reinicialitza una key ja existent) i es veu
            l'edició abandonada com si fos la real -- un cop real de
            correccions.csv (client penedes_languages, 2026-07-14 23:09)
            en dona proba directa: una segona "Guardar" 9s després d'una
            primera va re-escriure "0 -> 0" perquè el camp encara mostrava
            el "0" de l'intent anterior, no el 21 real de la fitxa. Es
            crida ABANS d'obrir el diàleg (botó "Corregir camps") i en
            Cancel·lar -- mai després d'instanciar cap widget d'aquestes
            keys en aquesta mateixa execució."""
            for key in _keys_correccio():
                st.session_state.pop(key, None)

        @st.dialog("Corregir camps", on_dismiss=_tancar_popover)
        def dialog_corregir_camps():
            st.caption(
                "La fitxa corregida torna a passar tota la validació -- "
                "corregir no aprova. Cal RECALCULAR després de desar."
            )
            # Piso 13X: exenta és l'ÚNIC camp booleà de l'esquema -- casella,
            # mai text "True"/"False". Viu FORA del form perquè la cascada
            # de línies (tipus=0, quota=0) sigui EN VIU en marcar-la (un
            # checkbox dins d'un st.form no reacciona fins que se sotmet).
            key_exenta = f"{prefijo}_correccio_exenta_{nombre}"
            exempta_marcada = st.checkbox(
                "exenta", value=bool(datos.get("exenta")), key=key_exenta,
            )
            if exempta_marcada:
                st.caption("⚠️ una factura exempta no porta IVA — revisa les línies")
                for i in range(n_linies_actual):
                    # Assignat ABANS de crear el text_input d'aquesta key en
                    # aquesta mateixa execució -- mai després (regla establerta
                    # al projecte).
                    st.session_state[f"{prefijo}_correccio_tipo_{i}_{nombre}"] = "0"
                    st.session_state[f"{prefijo}_correccio_cuota_{i}_{nombre}"] = "0"

            with st.form(key=f"form_correccio_{prefijo}_{nombre}", border=False):
                valors_nous = {}
                for camp in CAMPOS_CORREGIBLES_TOP:
                    if camp == "exenta":
                        continue
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
                col_guardar, col_cancelar = st.columns(2)
                with col_guardar:
                    click_guardar = st.form_submit_button("Guardar canvis", type="primary")
                with col_cancelar:
                    click_cancelar = st.form_submit_button("Cancel·lar")

            if click_cancelar:
                # Piso 13X: "restaura els valors originals" -- com les keys
                # es netegen, en reobrir el diàleg value= torna a manar.
                _reiniciar_correccio()
                _tancar_popover()
                st.rerun()

            if click_guardar:
                cambios = []
                valor_exenta_str = "True" if exempta_marcada else "False"
                valor_exenta_antic_str = "True" if datos.get("exenta") else "False"
                if valor_exenta_str != valor_exenta_antic_str:
                    cambios.append(("exenta", valor_exenta_antic_str, valor_exenta_str))
                for camp in CAMPOS_CORREGIBLES_TOP:
                    if camp == "exenta":
                        continue
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
                    # Piso 13X: el toast resumeix el QUÈ -- mai un genèric
                    # "guardat" que obliga a rellegir tota la fitxa per saber
                    # què ha canviat de veritat.
                    resum_canvis = ", ".join(f"{c}: {a}→{n}" for c, a, n in cambios)
                    st.toast(f"{len(cambios)} camps: {resum_canvis}", icon="✅")
                    # Piso 13R: avís inline (mai bloquejant -- la doctrina no
                    # canvia, validar.py fa l'examen de veritat en el proper
                    # Recalcular) si la correcció deixa la fitxa amb
                    # exenta=true i alguna línia amb IVA. st.toast en comptes
                    # de st.warning perquè el st.rerun() de sota esborraria
                    # qualsevol missatge normal abans que es pogués llegir
                    # (mateixa lliçó que Piso 13P).
                    def _es_positiu(texto):
                        try:
                            return float(texto) > 0
                        except (TypeError, ValueError):
                            return False

                    te_iva = any(_es_positiu(t) or _es_positiu(c) for _, t, b, c in lineas_nuevas)
                    if exempta_marcada and te_iva:
                        st.toast(
                            "Marcada com a exempta amb línies d'IVA -- tornarà a REVISAR fins "
                            "que corregeixis les línies (tipus 0, quota 0) o desmarquis exempta.",
                            icon="⚠️",
                        )
                    # Piso 13H: comprobado en directe -- st.rerun(scope="fragment")
                    # NO tanca el dialog quan es crida des de dins (el "obert"
                    # d'un dialog viu fora del propi re-render del fragment).
                    # st.rerun() a seques si el tanca -- nomes aquesta accio
                    # paga aquest cost (es rara, no cada clic).
                    _tancar_popover()
                    st.rerun()

        with st.popover("⋮ Més accions", key=popover_key):
            if not decision and prefijo == "pendent":
                st.markdown("**Moure a l'altre flux**")
                with st.form(key=f"form_moure_{prefijo}_{nombre}", border=False):
                    motiu_moure = st.text_input("Motiu (opcional)", key=f"{prefijo}_motiu_moure_{nombre}")
                    click_moure = st.form_submit_button(f"Moure a {flujo_contrari}")
                if click_moure:
                    migrar_lot.moure_de_flux(
                        os.path.basename(carpeta_cliente), [(nombre_base, flujo)],
                        motiu_moure or "Moviment individual des de Revisió", qui,
                    )
                    st.toast(f"Mogut a {flujo_contrari}. Recalcula per refer els veredictes.", icon="↔️")
                    _tancar_popover()
                    st.rerun(scope="fragment")

            if not decision and clients_altres:
                st.markdown("**Moure a un altre client**")
                opcions_client_desti = {f"{f['nombre']} ({f['carpeta']})": f["carpeta"] for f in clients_altres}
                etiquetes_client = list(opcions_client_desti.keys())
                index_defecte = next(
                    (i for i, k in enumerate(etiquetes_client) if opcions_client_desti[k] == datos.get("suggerit_carpeta")),
                    0,
                )
                with st.form(key=f"form_moure_client_{prefijo}_{nombre}", border=False):
                    eleccio_desti = st.selectbox(
                        "Moure a un altre client", etiquetes_client, index=index_defecte,
                        key=f"{prefijo}_client_desti_{nombre}",
                    )
                    flux_desti_mostrat = st.radio(
                        "Flux al client destí", ["Compres", "Vendes"],
                        index=0 if flujo == "rebudes" else 1,
                        horizontal=True, key=f"{prefijo}_flux_desti_{nombre}",
                    )
                    motiu_client = st.text_input("Motiu (opcional)", key=f"{prefijo}_motiu_client_{nombre}")
                    click_moure_client = st.form_submit_button("Moure a aquest client")
                if click_moure_client:
                    carpeta_desti = opcions_client_desti[eleccio_desti]
                    flux_desti_intern = "rebudes" if flux_desti_mostrat == "Compres" else "ingressos"
                    migrar_lot.moure_a_client(
                        os.path.basename(carpeta_cliente), carpeta_desti, [(nombre_base, flujo)],
                        motiu_client or "Moviment individual des de Revisió (altre client)", qui,
                        flujo_desti=flux_desti_intern,
                    )
                    st.toast(f"Mogut a {eleccio_desti}. Recalcula per refer els veredictes.", icon="↔️")
                    _tancar_popover()
                    st.rerun(scope="fragment")

            st.markdown("**Corregir camps**")
            if st.button("Corregir camps", key=f"{prefijo}_obrir_correccio_{nombre}"):
                # Piso 13X: mai obrir el diàleg amb una edició abandonada
                # d'una obertura anterior parada a session_state (vegeu
                # _reiniciar_correccio).
                _reiniciar_correccio()
                dialog_corregir_camps()

            grups_correccio = historial_correccions(carpeta_cliente, nombre)
            if grups_correccio:
                st.markdown("**Desfer última correcció**")
                ultim_grup = grups_correccio[-1]
                resum_ultim = "; ".join(
                    f"{c['camp']}: {c['valor_antic']} → {c['valor_nou']}" for c in ultim_grup["canvis"]
                )
                st.caption(f"Última: {resum_ultim} ({ultim_grup['qui']}, {ultim_grup['data']})")
                if st.button("Desfer últim canvi", key=f"{prefijo}_desfer_correccio_{nombre}"):
                    # Piso 13X: simetria amb el 13M (decisions.csv) -- el
                    # llibre només creix, mai s'edita ni s'esborra cap fila.
                    # El grup INVERS és una fila nova, no una marxa enrere.
                    cambios_inversos = [
                        (c["camp"], c["valor_nou"], c["valor_antic"]) for c in ultim_grup["canvis"]
                    ]
                    motiu_original = ultim_grup["canvis"][0].get("motiu", "")
                    escribir_correccion(
                        carpeta_cliente, nombre, cambios_inversos,
                        f"Desfet: {motiu_original}", qui,
                    )
                    st.toast("Últim canvi desfet — recalcula per refer els veredictes.", icon="↩️")
                    _tancar_popover()
                    st.rerun(scope="fragment")


@st.fragment
def bloque_decidit(archivo, carpeta_cliente, qui):
    """Piso 13M: "Ja decidits" passa de text pla a poder desfer la
    decisió -- mateix patró que tarjeta_revisio (fragment, es llegeix
    tot fresc en cada render, mai es confia en l'estat que va passar
    el bucle exterior). Si la decisió ja no existeix (algú ja l'ha
    desfet, o l'ha tornat a decidir), es mostra la confirmació i prou
    -- la fitxa reapareixerà a Pendents/OK en el proper rerun complet
    (aquest fragment no reexecuta el bucle exterior)."""
    decision = cargar_decisiones(carpeta_cliente).get(archivo)
    historial = historial_decisiones(carpeta_cliente, archivo)
    n_revertits = sum(1 for f in historial if f.get("accion") == "revertir")

    if decision is None:
        st.write(
            f"**{archivo}** — decisió desfeta ({n_revertits} "
            f"{'vegada' if n_revertits == 1 else 'vegades'}). Recalcula per veure-la al seu lloc."
        )
        return

    nota = f" — _{decision.get('nota')}_" if decision.get("nota") else ""
    st.write(
        f"**{archivo}** — {decision.get('accion')} per {decision.get('qui')} "
        f"el {decision.get('data')}{nota}"
    )
    if n_revertits:
        st.caption(f"↩️ Decisió desfeta ({n_revertits} {'vegada' if n_revertits == 1 else 'vegades'}) anteriorment")

    # Piso 13R: mateix "⋮ Més accions" que tarjeta_revisio -- una sola
    # acció aquí (Desfer), però mateix criteri d'endreça consistent a
    # tota la vista Revisió. No es converteix aquesta fila en targeta
    # amb vora -- només l'acció es reculla darrere el popover.
    #
    # Piso 13S: key= al popover + tancar-lo abans del st.rerun() que ja
    # feia -- mateix arranjament que tarjeta_revisio (el fantasma del
    # popover que no es tanca).
    popover_key = f"popover_desfer_{archivo}"
    with st.popover("⋮ Més accions", key=popover_key):
        with st.form(key=f"form_desfer_{archivo}", border=False):
            motiu_desfer = st.text_input("Motiu (opcional)", key=f"desfer_motiu_{archivo}")
            click_desfer = st.form_submit_button("Desfer aquesta decisió")
        if click_desfer:
            # Piso 13M: guàrdia explícita -- revertir_decision ja recomprova
            # si encara hi ha decisió efectiva just abans d'escriure (evita
            # una condició de cursa si es clica dues vegades seguides).
            revertida = revertir_decision(carpeta_cliente, archivo, motiu_desfer, qui)
            if revertida is None:
                st.error("No hi ha res a revertir.")
            else:
                st.toast(f"Decisió desfeta — {archivo} torna a pendents.", icon="↩️")
                st.session_state[popover_key] = False
                st.rerun()


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
        st.markdown(f"**{FLUX_ETIQUETA[flujo]} {nombre}**")
        if not os.path.exists(ruta):
            st.success("✓ Arxiu retirat.")
        else:
            st.warning(motivo_error(ruta))
            boton_obrir("Obrir arxiu", ruta, key=f"{prefijo}_error_original_{nombre}")
            st.caption("Torna a pujar l'arxiu bo des d'\"Afegir factures\".")
            col_retirar, col_reprocessar = st.columns(2)
            with col_retirar:
                if st.button("Retirar arxiu il·legible", key=f"{prefijo}_retirar_{nombre}"):
                    destino = retirar_error(carpeta_cliente, ruta, motivo_error(ruta), qui)
                    st.success(f"Arxiu retirat a `{destino}`.")
                    st.rerun(scope="fragment")
            with col_reprocessar:
                # Piso 13T: reprocessar NOMÉS aquest arxiu -- una sola
                # crida a extraer_todas.py (mode --archivo), mai tot el
                # lot. Mateix candau que Processar/RECALCULAR (Piso
                # 13S) -- toca els mateixos arxius.
                if st.button("Tornar a processar aquest arxiu", key=f"{prefijo}_reprocessar_{nombre}"):
                    candau_reprocessar = candau_processar_viu()
                    if candau_reprocessar:
                        st.error(
                            f"Hi ha un Processar en marxa (iniciat per {candau_reprocessar.get('qui')} a les "
                            f"{candau_reprocessar.get('data_inici')}) -- espera que acabi abans de reprocessar."
                        )
                    else:
                        nombre_base = os.path.splitext(nombre)[0]
                        if flujo == "rebudes":
                            carpeta_extraidas = os.path.join(carpeta_cliente, "rebudes", "extraidas")
                            carpeta_validadas = os.path.join(carpeta_cliente, "rebudes", "validadas")
                        else:
                            carpeta_extraidas = os.path.join(carpeta_cliente, "apartados", "ingressos_extraidas")
                            carpeta_validadas = os.path.join(carpeta_cliente, "apartados", "ingressos_validadas")
                        ruta_json_extraidas = os.path.join(carpeta_extraidas, nombre_base + ".json")
                        ruta_json_validadas = os.path.join(carpeta_validadas, nombre_base + ".json")
                        # Neteja l'estat parcial d'AQUEST arxiu concret,
                        # mai de la resta -- mai un residu vell que faci
                        # "saltada" sense voler en el reintent.
                        for ruta_neteja in (ruta_json_extraidas, ruta_json_validadas):
                            if os.path.exists(ruta_neteja):
                                os.remove(ruta_neteja)

                        with st.spinner("Reprocessant..."):
                            proceso_extraccio = subprocess.run(
                                [sys.executable, "extraer_todas.py", "--archivo", ruta, "--json", ruta_json_extraidas],
                                cwd=RAIZ_PROYECTO, capture_output=True, text=True,
                            )
                            subprocess.run(
                                [sys.executable, "validar.py"], cwd=RAIZ_PROYECTO, capture_output=True, text=True,
                            )

                        if proceso_extraccio.returncode == 0:
                            st.toast(f"{nombre} reprocessat -- Recalcula per veure-ho a Revisió.", icon="✅")
                        else:
                            st.toast(f"{nombre} ha tornat a fallar en reprocessar-lo.", icon="⚠️")
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
/* Piso 13N: el boto final de DESTRUIR -- vermell perque es l'unic
   verb del projecte que esborra de veritat, i mai type="primary"
   (no ha de ser el boto per defecte de la pantalla, regla del pis).
   Mai disabled= (regla 10): el clic sempre es processa i valida amb
   st.error visible, per aixo nomes cal l'estil actiu. */
[class*="st-key-destruir_boto"] button {
    background-color: #b91c1c !important;
    border-color: #b91c1c !important;
    color: white !important;
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

if "log_recalcular" not in st.session_state:
    st.session_state["log_recalcular"] = None

st.title("Agent TRIMESTRE")

# Piso 13S: Streamlit no deixa tocar st.session_state["vista_radio"]
# UN COP el widget amb aquesta key ja s'ha instanciat en aquest mateix
# run (StreamlitAPIException, reproduit en directe) -- per aixo "Veure
# progrés" (banner de sota) nomes marca un senyal petit ABANS de crear
# el radio; aquest bloc el consumeix i fixa el valor de veritat abans
# que el widget existeixi.
if st.session_state.pop("_saltar_a_processar", False):
    st.session_state["vista_radio"] = "Processar"

# Piso 13S: key= perquè "Veure progrés" pugui saltar aquí mateix des de
# qualsevol altra vista sense rellançar res -- reconnectar, no rellançar.
vista = st.sidebar.radio(
    "Navegació", ["Clients", "Afegir factures", "Processar", "Revisió", "Manteniment"], key="vista_radio",
)

# Piso 13S: semàfor -- banner persistent a TOTES les vistes mentre hi ha
# un Processar viu (candau real, PID comprovat). Navegar fora no atura
# res -- el procés ja és independent del run que el va llançar -- per
# aixo el banner ho diu explícitament en lloc de deixar-ho ambigu.
_candau_global = candau_processar_viu()
if _candau_global:
    st.warning(
        f"⚙ Processant [{_candau_global.get('abast')}]... iniciat a les "
        f"{_candau_global.get('data_inici')} per {_candau_global.get('qui')}. "
        "El procés continua en segon pla encara que naveguis fora."
    )
    col_veure, col_parar_global = st.columns(2)
    with col_veure:
        if st.button("Veure progrés", key="banner_veure_progres"):
            st.session_state["_saltar_a_processar"] = True
            st.rerun()
    with col_parar_global:
        if st.button("PARAR", key="banner_parar"):
            with open(RUTA_STOP_PROCESSAR, "w", encoding="utf-8"):
                pass
            st.toast(
                "S'ha demanat aturar el procés -- acabarà la factura/lot que té entre mans.",
                icon="🛑",
            )
            st.rerun()

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
        opciones = {f"{f['nombre']} ({f['carpeta']})": f for f in clientes}
        eleccion = st.selectbox("Client", list(opciones.keys()))
        fila_afegir = opciones[eleccion]
        carpeta = fila_afegir["carpeta"]

        # Piso 13X: mode nou -- pujar un arxiu escanejat (de sempre) o
        # teclejar una fitxa a ma (mai passa per extraer_todas.py). Son
        # dos eixos diferents (mètode d'entrada, no destí), per això un
        # selector propi a dalt en comptes d'un quart valor barrejat al
        # radio "Destí" de sota.
        mode_afegir = st.radio(
            "Com vols afegir factures?", ["Pujar arxius", "Entrada manual"], horizontal=True,
        )

        if mode_afegir == "Pujar arxius":
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

            # Piso 13J: "Lot d'escaner" triat DIRECTAMENT (sense passar per la
            # guardia de mes avall) abans assumia sempre Compres en silenci --
            # bug real de camp (~70 factures emeses processades com a compres).
            # Bloquejant (regla 10): sense triar-ho, no es pot ni pujar el PDF.
            tipus_lot = None
            if destino == "Lot":
                tipus_lot = st.radio(
                    "Aquest lot és de...", list(DESTI_LOT_DIRECTE.keys()),
                    index=None, horizontal=True, key="tipus_lot_directe",
                )
                if tipus_lot is None:
                    st.info("Selecciona si el lot és de compres o de vendes abans de pujar el PDF.")

            # Piso 13Q: mentre "Lot" encara no te tipus triat, el desti real
            # no es coneix -- ni banner ni confirmacio tenen sentit encara
            # (el guardia de dalt ja bloqueja el pas).
            if destino != "Lot" or tipus_lot is not None:
                # Piso 13J: si destino == "Lot", el flux real es el triat a
                # tipus_lot (traduit amb DESTI_LOT_DIRECTE), no el literal "Lot".
                destino_efectiu = DESTI_LOT_DIRECTE.get(tipus_lot) if destino == "Lot" else destino
                context_afegir = (carpeta, destino_efectiu)
                banner = (
                    f"📥 Pujant a: **{fila_afegir['nombre']} ({carpeta})** → "
                    f"**{destino_efectiu.upper()}**"
                )
                if destino_efectiu == "Vendes":
                    st.success(banner)
                else:
                    st.info(banner)

                # Piso 13Q: PREVENIR -- confirmacio NOMES quan el context canvia
                # (primera pujada de la sessio, o client/flux diferent de
                # l'anterior). Un cop confirmat, pujar mes arxius al MATEIX
                # context no torna a preguntar -- mai una confirmacio rutinaria.
                if st.session_state.get("afegir_context_confirmat") != context_afegir:
                    st.warning(
                        "Has canviat de client o de destí -- confirma abans de pujar arxius."
                    )
                    if st.button("Confirmar destinació", type="primary", key="afegir_confirmar_context"):
                        st.session_state["afegir_context_confirmat"] = context_afegir
                        st.rerun()
                else:
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

                    # Piso 13T: hashos.csv viu a l'arrel del client; "rebudes/
                    # procesadas" es l'unic sibling "ja processat" conegut avui
                    # (Piso 13Q/migrar_lot.py), i nomes te sentit quan la
                    # destinacio es Compres solta (mai un lot sencer, que mai
                    # coincidira per hash amb una factura solta ja processada).
                    carpeta_arrel_client = ruta_proyecto("clientes", carpeta)
                    carpetes_procesadas = [os.path.join(carpeta_arrel_client, "rebudes", "procesadas")]

                    if sospitos:
                        nombre_sospitos, n_pag = sospitos
                        st.warning(f"Aquest PDF ({nombre_sospitos}) té {n_pag} pàgines — és un lot d'escàner?")
                        col_si, col_no = st.columns(2)
                        with col_si:
                            confirmar_lot = st.button("Sí, és un lot", key="lot_confirmar")
                        with col_no:
                            confirmar_solta = st.button("No, és una factura llarga", key="lot_descartar")
                        if confirmar_lot:
                            # Piso 13K: abans "Lot" literal, ignorant el flux triat --
                            # ara respecta destino (Compres/Vendes) i nomes marca
                            # es_lot=True per triar el moll corresponent.
                            guardar_y_reportar(
                                archivos, ruta_destino_factures(carpeta, destino, es_lot=True), carpeta_arrel_client,
                            )
                        elif confirmar_solta:
                            guardar_y_reportar(
                                archivos, ruta_destino_factures(carpeta, destino), carpeta_arrel_client,
                                carpetas_hash_extra=carpetes_procesadas if destino == "Compres" else (),
                            )
                    else:
                        if st.button("Desar arxius", disabled=not archivos, type="primary"):
                            guardar_y_reportar(
                                archivos, ruta_destino_factures(carpeta, destino_efectiu, es_lot=(destino == "Lot")),
                                carpeta_arrel_client,
                                carpetas_hash_extra=carpetes_procesadas if destino_efectiu == "Compres" and destino != "Lot" else (),
                            )
        else:
            entrada_manual_factura(fila_afegir, carpeta)

# ----------------------------------------------------------------------
elif vista == "Processar":
    st.header("Processar")
    clientes = leer_clientes()

    if not clientes:
        st.info("Encara no hi ha cap client donat d'alta -- res per processar.")
    elif not st.session_state.get("qui_processa_confirmat"):
        st.info("Cal indicar qui processa abans de llançar-ho.")
        with st.form("form_qui_processa"):
            nom_qui_processa = st.text_input("Qui processa?", key="qui_processa_input")
            entrar_processa = st.form_submit_button("Entrar", type="primary")
        if entrar_processa:
            if not nom_qui_processa.strip():
                st.error("Cal escriure un nom abans de continuar.")
            else:
                st.session_state["qui_processa_confirmat"] = nom_qui_processa.strip()
                st.rerun()
        st.stop()
    else:
        qui_processa = st.session_state["qui_processa_confirmat"]
        col_qui_p, col_canviar_p = st.columns([4, 1])
        with col_qui_p:
            st.caption(f"Processant com: **{qui_processa}**")
        with col_canviar_p:
            if st.button("Canviar qui processa", key="qui_processa_canviar"):
                st.session_state["qui_processa_confirmat"] = None
                st.rerun()

        # Piso 13S: Processar ja no bloqueja la sessió sencera -- llança
        # ejecutar.py DESLLIGAT (stdout a un fitxer, mai una pipe llegida
        # en un bucle síncron) i torna de seguida. El candau
        # (processar.lock) és la font de veritat de si encara corre;
        # "Veure progrés"/aquesta mateixa vista només RECONNECTEN
        # llegint proces_log.txt fresc del disc, mai rellancen res.
        candau = candau_processar_viu()
        if candau:
            st.warning(
                f"⚙ Ja hi ha un processament en marxa (iniciat per {candau.get('qui')} "
                f"a les {candau.get('data_inici')}). Espera que acabi o para'l."
            )
            log_actual = ""
            if os.path.exists(RUTA_LOG_PROCESSAR):
                with open(RUTA_LOG_PROCESSAR, encoding="utf-8") as f:
                    log_actual = f.read()
            st.code(log_actual or "(encara sense sortida)")
            col_actualitza, col_parar = st.columns(2)
            with col_actualitza:
                if st.button("Actualitza el registre", key="processar_actualitza"):
                    st.rerun()
            with col_parar:
                if st.button("PARAR", type="primary", key="processar_parar"):
                    with open(RUTA_STOP_PROCESSAR, "w", encoding="utf-8"):
                        pass
                    st.toast(
                        "S'ha demanat aturar el procés -- acabarà la factura/lot que té entre mans.",
                        icon="🛑",
                    )
                    st.rerun()
        else:
            if st.button("Processar", type="primary"):
                # Piso 13S: neteja d'un stop antic (idempotent -- si no
                # n'hi havia, no fa res) perquè no bloquegi aquest run nou.
                if os.path.exists(RUTA_STOP_PROCESSAR):
                    os.remove(RUTA_STOP_PROCESSAR)

                # Piso 13U: es restaura el patró de sempre (Piso 10.x,
                # pre-13S) per a QUI llança el procés -- bloquejant, amb
                # el log en directe a la mateixa pantalla (el "munyeco"
                # que dona confiança que està corrent de veritat). El
                # 13S ho havia canviat per un subprocés desslligat sense
                # cap senyal de vida directa -- es manté NOMÉS com a
                # extra per a qui NO ha llançat el procés (candau/banner
                # de dalt): cada línia també es escriu a proces_log.txt
                # (tee) perquè aquestes altres sessions puguin seguir-lo.
                placeholder = st.empty()
                buffer = ""
                with open(RUTA_LOG_PROCESSAR, "w", encoding="utf-8") as log_file:
                    proceso = subprocess.Popen(
                        [sys.executable, "ejecutar.py", qui_processa],
                        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                        text=True, bufsize=1, cwd=RAIZ_PROYECTO,
                    )
                    for linea in proceso.stdout:
                        buffer += linea
                        placeholder.code(buffer)
                        log_file.write(linea)
                        log_file.flush()
                    proceso.wait()

                if proceso.returncode == 0:
                    st.success("Procés acabat.")
                else:
                    st.error(f"El procés ha acabat amb error (codi {proceso.returncode}).")

            if os.path.exists(RUTA_LOG_PROCESSAR):
                with open(RUTA_LOG_PROCESSAR, encoding="utf-8") as f:
                    log_final = f.read()
                if log_final:
                    if "ATURAT per petició" in log_final:
                        st.warning("L'últim Processar es va aturar per petició de l'usuari abans d'acabar.")
                    elif "AVISO:" in log_final and "Pipeline completo" not in log_final:
                        st.error("L'últim Processar va acabar amb un error. Revisa el registre de sota.")
                    elif "Pipeline completo" in log_final:
                        st.success("Últim Processar acabat.")

                    with st.expander("Registre de l'últim Processar", expanded=False):
                        st.code(log_final)

                    st.subheader("Resum per client")
                    st.caption(
                        "S'han processat tots els clients (els que ja estaven al dia "
                        "s'han saltat per idempotència)."
                    )
                    for fila in clientes:
                        linias_cliente = [
                            linia for linia in log_final.splitlines()
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
        # Piso 13V: desglossat per flux -- calculat sempre sobre el
        # conjunt SENCER (mai el filtrat pel selector_flux de sota).
        n_pendents_compres = sum(1 for _, _, f, _ in estado["pendents"] if f == "rebudes")
        n_pendents_vendes = n_pendents - n_pendents_compres
        n_sin_recalcular = contar_decisiones_sin_recalcular(estado["carpeta_cliente"])
        n_moguts_sense_recalcular = contar_moviments_sense_recalcular(estado["carpeta_cliente"])
        n_retirs_sense_recalcular = contar_retirs_sense_recalcular(estado["carpeta_cliente"])
        n_manuals_sense_recalcular = contar_manuals_sense_recalcular(estado["carpeta_cliente"])

        st.subheader(
            f"{n_pendents} pendents per decidir ({n_pendents_compres} compres · {n_pendents_vendes} vendes) · "
            f"{n_errores} errors per resoldre · "
            f"{n_decidits} decidits · {n_sin_recalcular} decisions noves sense recalcular · "
            f"{n_moguts_sense_recalcular} moguts sense recalcular · "
            f"{n_retirs_sense_recalcular} retirs sense recalcular · "
            f"{n_manuals_sense_recalcular} entrades manuals sense recalcular"
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
            type="primary" if (
                n_sin_recalcular > 0 or n_moguts_sense_recalcular > 0 or n_retirs_sense_recalcular > 0
                or n_manuals_sense_recalcular > 0
            ) else "secondary",
        ):
            # Piso 13S: RECALCULAR crida la MATEIXA cadena (validar->sumar->
            # informe) sobre els MATEIXOS arxius que Processar -- mai a la
            # vegada, per no arriscar una escriptura concurrent real.
            candau_recalcular = candau_processar_viu()
            if candau_recalcular:
                st.error(
                    f"Hi ha un Processar en marxa (iniciat per {candau_recalcular.get('qui')} a les "
                    f"{candau_recalcular.get('data_inici')}) -- espera que acabi abans de Recalcular, "
                    "per no escriure als mateixos arxius alhora."
                )
            else:
                placeholder = st.empty()
                buffer = ""
                # Piso 11B: validar.py entra en la cadena -- las correccions.csv
                # solo se aplican en memoria dentro de validar.py, asi que hace
                # falta volver a correrlo para que una correccio "torni a passar
                # l'examen". Sigue sin llamadas a la API (validar.py es pura
                # logica Python), igual de gratis que sumar.py/informe.py.
                #
                # Piso 13P: bug de confiança trobat en directe -- si l'Excel
                # (o l'informe) d'un client estava obert, sumar.py/informe.py
                # ja el saltaven be (Piso 13G/13P) i seguien amb la resta, pero
                # aquest bloc mai comprovava proceso.returncode ni guardava el
                # buffer -- sempre acabava en "Recalculat." + st.rerun()
                # immediat, que esborrava la pantalla ABANS que ningu pogues
                # llegir l'AVISO. Mateix patró que "Processar" (mes amunt): es
                # distingeix un crash de veritat (returncode != 0 -- atura la
                # cadena, com ejecutar.py) d'un avis de fitxer bloquejat
                # (returncode 0 pero "AVISO:" al log -- el lot ja ha continuat
                # sol, regla 4, pero mai es disfressa d'exit). Cap dels dos
                # casos fa rerun immediat: el missatge i el log s'han de poder
                # llegir.
                aturat_per_error = False
                maquina_fallida = None
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
                    if proceso.returncode != 0:
                        aturat_per_error = True
                        maquina_fallida = maquina
                        break

                st.session_state["log_recalcular"] = buffer

                if aturat_per_error:
                    st.error(
                        f"{maquina_fallida} ha acabat amb un error (codi {proceso.returncode}) -- "
                        f"aturant la cadena aquí. Revisa el registre de sota."
                    )
                elif "AVISO:" in buffer:
                    st.warning(
                        "Recalculat, PERÒ algun client no s'ha pogut actualitzar (Excel o informe "
                        "obert?). Revisa el registre de sota abans de confiar en les xifres."
                    )
                else:
                    st.success("Recalculat.")
                    st.rerun()

        if st.session_state["log_recalcular"]:
            with st.expander("Registre de l'últim Recalcular", expanded=False):
                st.code(st.session_state["log_recalcular"])

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
        flux_pendents = selector_flux(f"revisio_flux_pendents_{carpeta}")
        pendents_en_flux = (
            [(n, d, f, o) for n, d, f, o in estado["pendents"] if f == flux_pendents]
            if flux_pendents else estado["pendents"]
        )
        cerca_pendents = entrada_cerca("🔍 Cercar a Pendents", f"revisio_cerca_pendents_{carpeta}")
        if cerca_pendents:
            q_pendents = normalizar_cerca(cerca_pendents)
            pendents_mostrats = [
                (n, d, f, o) for n, d, f, o in pendents_en_flux
                if q_pendents in texto_buscable_ficha(n, d)
            ]
            st.caption(f"{len(pendents_mostrats)} resultats de {len(pendents_en_flux)}")
        else:
            pendents_mostrats = pendents_en_flux

        if not pendents_mostrats:
            st.caption("Cap pendent sense decidir." if not cerca_pendents else "Cap resultat per aquesta cerca.")
        else:
            pagina_pendents = paginar(pendents_mostrats, f"revisio_pag_pendents_{carpeta}")
            for nombre, datos, flujo, origen in pagina_pendents:
                st.checkbox("Seleccionar per al lot", key=f"revisio_sel_pendent_{carpeta}_{nombre}")
                tarjeta_revisio(nombre, datos, origen, estado["carpeta_cliente"], qui, "pendent", flujo, fila_cliente["nif"])

            seleccionats_pendents_info = [
                (nombre, flujo) for nombre, _, flujo, _ in pagina_pendents
                if st.session_state.get(f"revisio_sel_pendent_{carpeta}_{nombre}")
            ]
            seleccionats_pendents = [nombre for nombre, _ in seleccionats_pendents_info]
            if seleccionats_pendents:
                st.info(f"{len(seleccionats_pendents)} seleccionades")
                nota_lot_pendents = st.text_input(
                    "Nota compartida (opcional)", key=f"revisio_lot_nota_pendents_{carpeta}"
                )
                col_aprovar_lot, col_moure_lot = st.columns(2)
                with col_aprovar_lot:
                    if st.button("APROVAR SELECCIONADES", type="primary", key=f"revisio_lot_aprovar_{carpeta}"):
                        data_lote = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        for nombre in seleccionats_pendents:
                            escribir_decision(
                                estado["carpeta_cliente"], nombre, "aprovar", nota_lot_pendents, qui, data=data_lote,
                            )
                        st.rerun()
                with col_moure_lot:
                    # Piso 13L: porta de lot -- cada document viatja al SEU
                    # propi flux contrari (una seleccio pot barrejar
                    # rebudes+ingressos alhora).
                    if st.button("MOURE SELECCIONADES A L'ALTRE FLUX", key=f"revisio_lot_moure_{carpeta}"):
                        documentos = [
                            (os.path.splitext(nombre)[0], flujo) for nombre, flujo in seleccionats_pendents_info
                        ]
                        migrar_lot.moure_de_flux(
                            carpeta, documentos,
                            nota_lot_pendents or "Moviment en lot des de Revisió", qui,
                        )
                        st.toast(
                            f"{len(documentos)} mogudes a l'altre flux. Recalcula per refer els veredictes.",
                            icon="↔️",
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
            flux_oks = selector_flux(f"revisio_flux_oks_{carpeta}")
            oks_en_flux = (
                [(n, d, f, o) for n, d, f, o in oks_ordenadas if f == flux_oks] if flux_oks else oks_ordenadas
            )
            cerca_oks = entrada_cerca("🔍 Cercar a Fitxes OK", f"revisio_cerca_oks_{carpeta}")
            if cerca_oks:
                q_oks = normalizar_cerca(cerca_oks)
                oks_mostrades = [
                    (n, d, f, o) for n, d, f, o in oks_en_flux if q_oks in texto_buscable_ficha(n, d)
                ]
                st.caption(f"{len(oks_mostrades)} resultats de {len(oks_en_flux)}")
            else:
                oks_mostrades = oks_en_flux

            if not oks_mostrades:
                st.caption("Sense fitxes OK." if not cerca_oks else "Cap resultat per aquesta cerca.")
            else:
                pagina_oks = paginar(oks_mostrades, f"revisio_pag_oks_{carpeta}")
                for nombre, datos, flujo, origen in pagina_oks:
                    tarjeta_revisio(nombre, datos, origen, estado["carpeta_cliente"], qui, "ok", flujo, fila_cliente["nif"])

        st.markdown(f"### Errors ({n_errores})")
        flux_errors = selector_flux(f"revisio_flux_errors_{carpeta}")
        errores_en_flux = (
            [(f, r) for f, r in estado["errores"] if f == flux_errors] if flux_errors else estado["errores"]
        )
        cerca_errors = entrada_cerca("🔍 Cercar a Errors", f"revisio_cerca_errors_{carpeta}")
        if cerca_errors:
            q_errors = normalizar_cerca(cerca_errors)
            errores_mostrats = [
                (f, r) for f, r in errores_en_flux
                if q_errors in normalizar_cerca(f"{os.path.basename(r)} {f}")
            ]
            st.caption(f"{len(errores_mostrats)} resultats de {len(errores_en_flux)}")
        else:
            errores_mostrats = errores_en_flux

        if not errores_mostrats:
            st.caption("Cap error per resoldre." if not cerca_errors else "Cap resultat per aquesta cerca.")
        else:
            st.caption(
                "Si l'arxiu és bo però va quedar corromput (iCloud), torna'l a pujar "
                "abans de retirar res."
            )
            pagina_errores = paginar(errores_mostrats, f"revisio_pag_errors_{carpeta}")
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

        # Piso 13Q: NETEJA DE DUPLICATS -- validar.py (Piso 13J) ja calcula
        # el motiu "factura duplicada: mismo proveedor+num_factura que ..."
        # per cada pendent; aquesta secció NOMÉS agrupa el que ja existeix
        # (cap detecció nova) i ofereix descartar les còpies amb un clic en
        # comptes d'una a una des de la targeta.
        grups_duplicats = {}
        for nombre_dup, datos_dup, flujo_dup, origen_dup in estado["pendents"]:
            if not any((m or "").startswith("factura duplicada") for m in (datos_dup.get("motivos") or [])):
                continue
            clau = (datos_dup.get("nif_proveedor"), datos_dup.get("num_factura"))
            grups_duplicats.setdefault(clau, []).append((nombre_dup, datos_dup, flujo_dup, origen_dup))

        # Piso 13T: l'ordre per data de cada grup es calcula UN SOL COP --
        # el reutilitzen tant el botó gros com cada grup individual (abans
        # es recalculava dins del bucle, ara nomes cal per mostrar-lo).
        grups_ordenats = {}
        for clau, membres in grups_duplicats.items():
            membres_amb_data = []
            for nombre_dup, datos_dup, flujo_dup, origen_dup in membres:
                ruta_orig_dup = encontrar_original(origen_dup, nombre_dup)
                mtime_dup = os.path.getmtime(ruta_orig_dup) if ruta_orig_dup else 0
                membres_amb_data.append((nombre_dup, flujo_dup, mtime_dup))
            membres_amb_data.sort(key=lambda t: t[2])
            grups_ordenats[clau] = membres_amb_data

        st.markdown(f"### Duplicats ({len(grups_duplicats)} grups)")
        flux_duplicats = selector_flux(f"revisio_flux_duplicats_{carpeta}")
        if flux_duplicats:
            # Piso 13V: mateix criteri que la cerca de sota -- un grup
            # sobreviu si ALGUN membre coincideix amb el flux triat.
            grups_ordenats = {
                clau: membres for clau, membres in grups_ordenats.items()
                if any(f == flux_duplicats for _, f, _ in membres)
            }
        n_grups_en_flux = len(grups_ordenats)
        cerca_duplicats = entrada_cerca("🔍 Cercar a Duplicats", f"revisio_cerca_duplicats_{carpeta}")
        if cerca_duplicats:
            q_duplicats = normalizar_cerca(cerca_duplicats)
            # Piso 13U: es filtra per GRUP sencer -- es manté si QUALSEVOL
            # membre coincideix (resoldre un duplicat es fa en bloc).
            grups_ordenats = {
                clau: membres for clau, membres in grups_ordenats.items()
                if any(q_duplicats in texto_buscable_ficha(n, d) for n, d, f, o in grups_duplicats[clau])
            }
            st.caption(f"{len(grups_ordenats)} grups de {n_grups_en_flux}")

        if not grups_ordenats:
            st.caption("Cap duplicat detectat." if not cerca_duplicats else "Cap resultat per aquesta cerca.")
        else:
            total_copies = sum(len(m) - 1 for m in grups_ordenats.values())
            # Piso 13T: el botó gros aplica la MATEIXA preselecció que ja fa
            # cada grup per defecte (totes menys la més antiga) a TOTS els
            # grups d'un sol clic -- els comptes al mateix botó ja son la
            # confirmació d'una línia, no cal cap diàleg a part.
            if st.button(
                f"DESCARTAR TOTES LES CÒPIES ({len(grups_ordenats)} grups, {total_copies} còpies)",
                key=f"dup_descartar_tot_{carpeta}", type="primary",
            ):
                for (nif_dup, num_dup), membres_amb_data in grups_ordenats.items():
                    original_citat = membres_amb_data[0][0]
                    for nombre_dup, _, _ in membres_amb_data[1:]:
                        escribir_decision(
                            estado["carpeta_cliente"], nombre_dup, "descartar",
                            f"duplicat de {original_citat}", qui,
                        )
                st.rerun()

            for (nif_dup, num_dup), membres_amb_data in grups_ordenats.items():
                with st.container(border=True, key=f"dup_grup_{carpeta}_{nif_dup}_{num_dup}"):
                    st.markdown(f"**Factura {num_dup}** · NIF {nif_dup} · {len(membres_amb_data)} còpies")
                    seleccio = {}
                    for i, (nombre_dup, flujo_dup, mtime_dup) in enumerate(membres_amb_data):
                        data_llegible = (
                            datetime.fromtimestamp(mtime_dup).strftime("%d/%m/%Y %H:%M") if mtime_dup else "—"
                        )
                        etiqueta_rol = "original (més antiga)" if i == 0 else "còpia"
                        seleccio[nombre_dup] = st.checkbox(
                            f"{nombre_dup} — {data_llegible} — {FLUX_ETIQUETA[flujo_dup]} — {etiqueta_rol}",
                            value=(i != 0),
                            key=f"dup_sel_{carpeta}_{nombre_dup}",
                        )
                    marcades = [n for n, sel in seleccio.items() if sel]
                    if st.button(
                        f"DESCARTAR LES CÒPIES SELECCIONADES ({len(marcades)})",
                        key=f"dup_descartar_{carpeta}_{nif_dup}_{num_dup}",
                    ):
                        no_marcades = [n for n in seleccio if n not in marcades]
                        if not marcades:
                            st.error("Selecciona almenys una còpia abans de descartar.")
                        elif not no_marcades:
                            st.error("Cal deixar almenys un document viu -- no es pot descartar tot el grup.")
                        else:
                            original_citat = no_marcades[0]
                            for n in marcades:
                                escribir_decision(
                                    estado["carpeta_cliente"], n, "descartar",
                                    f"duplicat de {original_citat}", qui,
                                )
                            st.rerun()

        with st.expander(f"Ja decidits ({n_decidits})", key=f"revisio_expander_decidits_{carpeta}"):
            if not estado["decisiones"]:
                st.caption("Cap decisió encara.")
            for archivo in estado["decisiones"]:
                bloque_decidit(archivo, estado["carpeta_cliente"], qui)

# ----------------------------------------------------------------------
elif vista == "Manteniment":
    st.header("Manteniment")
    st.caption(
        "L'únic lloc del projecte que esborra de veritat -- excloure i retirar "
        "són el dia a dia; destruir és excepcional, lent i queda certificat. "
        "Mai res que estigui viu al flux: primer descarta o retira, després destrueix."
    )

    if not st.session_state.get("qui_manteniment_confirmat"):
        st.info("Cal indicar qui fa aquest manteniment abans de veure les eines.")
        with st.form("form_qui_manteniment"):
            nom_qui_manteniment = st.text_input("Qui fa el manteniment?", key="qui_manteniment_input")
            entrar_manteniment = st.form_submit_button("Entrar", type="primary")
        if entrar_manteniment:
            if not nom_qui_manteniment.strip():
                st.error("Cal escriure un nom abans de continuar.")
            else:
                st.session_state["qui_manteniment_confirmat"] = nom_qui_manteniment.strip()
                st.rerun()
        st.stop()
    else:
        qui_m = st.session_state["qui_manteniment_confirmat"]
        col_qui_m, col_canviar_m = st.columns([4, 1])
        with col_qui_m:
            st.caption(f"Fent manteniment com: **{qui_m}**")
        with col_canviar_m:
            if st.button("Canviar qui fa manteniment", key="qui_manteniment_canviar"):
                st.session_state["qui_manteniment_confirmat"] = None
                st.rerun()

        tab_docs, tab_client, tab_massiva = st.tabs(
            ["Documents d'un client", "Client arxivat", "Destrucció massiva"]
        )

        with tab_docs:
            clientes_m = leer_clientes()
            if not clientes_m:
                st.info("Encara no hi ha cap client donat d'alta.")
            else:
                opciones_m = {f"{f['nombre']} ({f['carpeta']})": f["carpeta"] for f in clientes_m}
                eleccion_m = st.selectbox("Client", list(opciones_m.keys()), key="manteniment_client")
                carpeta_m = opciones_m[eleccion_m]

                retirats, descartats = candidatos_destruccio(carpeta_m)

                st.subheader(f"Retirats ({len(retirats)})")
                if not retirats:
                    st.caption("Cap document retirat.")
                for r in retirats:
                    st.checkbox(r["nombre"], key=f"destruir_sel_{carpeta_m}_{r['nombre']}")

                st.subheader(f"Descartats ({len(descartats)})")
                if not descartats:
                    st.caption("Cap document descartat.")
                for d in descartats:
                    etiqueta = d["nombre"]
                    if d.get("num_factura"):
                        etiqueta += f" (factura {d['num_factura']})"
                    st.checkbox(etiqueta, key=f"destruir_sel_{carpeta_m}_{d['nombre']}")

                seleccionats = [
                    it for it in (retirats + descartats)
                    if st.session_state.get(f"destruir_sel_{carpeta_m}_{it['nombre']}")
                ]

                if seleccionats:
                    bytes_totals = sum(
                        os.path.getsize(ruta)
                        for it in seleccionats for ruta in it["rutas"] if os.path.exists(ruta)
                    )
                    st.warning(
                        f"Es destruiran {len(seleccionats)} documents "
                        f"({bytes_totals / (1024 * 1024):.2f} MB):\n"
                        + "\n".join(f"- {it['nombre']}" for it in seleccionats)
                    )
                    motiu_destruir = st.text_input("Motiu (obligatori)", key=f"destruir_motiu_{carpeta_m}")
                    confirmacio_destruir = st.text_input(
                        'Escriu "DESTRUIR" per confirmar', key=f"destruir_confirmacio_{carpeta_m}"
                    )
                    # Piso 13N: MAI disabled= lligat a un text_input solt
                    # (regla 10 -- "nunca depender de que un text_input
                    # esté confirmado con Tab/Enter"). El clic SEMPRE es
                    # processa, mateix patró que Aprovar/Descartar: si
                    # falta el motiu o la paraula no és exacta, error
                    # visible i no es destrueix res.
                    if st.button("DESTRUIR DEFINITIVAMENT", key=f"destruir_boto_{carpeta_m}"):
                        if not motiu_destruir:
                            st.error("Cal escriure un motiu per destruir.")
                        elif confirmacio_destruir != "DESTRUIR":
                            st.error('Cal escriure exactament "DESTRUIR" per confirmar.')
                        else:
                            destruits, omesos = destruir_documentos(carpeta_m, seleccionats, motiu_destruir, qui_m)
                            if destruits:
                                st.success(
                                    f"{len(destruits)} documents destruïts. "
                                    f"Certificat escrit a clientes/{carpeta_m}/registre_destruccions.csv."
                                )
                                # Piso 13V: l'informe comptava aquests
                                # arxius com a "presents" -- s'actualitza
                                # sol, gratuit i en segons.
                                auto_recalcular_sumar_informe()
                            if omesos:
                                st.error(
                                    "Alguns documents ja no eren destruïbles i s'han saltat: "
                                    + "; ".join(f"{n} ({m})" for n, m in omesos)
                                )
                            st.rerun()

        with tab_client:
            carpeta_arxivats = ruta_proyecto("arxivats")
            carpetas_arxivades = sorted(
                n for n in os.listdir(carpeta_arxivats) if os.path.isdir(os.path.join(carpeta_arxivats, n))
            ) if os.path.isdir(carpeta_arxivats) else []

            if not carpetas_arxivades:
                st.info("Cap client arxivat.")
            else:
                eleccion_arx = st.selectbox("Client arxivat", carpetas_arxivades, key="manteniment_arxivat")
                ruta_arx = os.path.join(carpeta_arxivats, eleccion_arx)
                n_archivos_arx = sum(len(archivos) for _, _, archivos in os.walk(ruta_arx))
                bytes_arx = sum(
                    os.path.getsize(os.path.join(dirpath, f))
                    for dirpath, _, archivos in os.walk(ruta_arx) for f in archivos
                )
                st.warning(
                    f"Es destruirà TOT el client arxivat '{eleccion_arx}': "
                    f"{n_archivos_arx} arxius ({bytes_arx / (1024 * 1024):.2f} MB)."
                )
                motiu_arx = st.text_input("Motiu (obligatori)", key="destruir_motiu_arxivat")
                confirmacio_arx = st.text_input('Escriu "DESTRUIR" per confirmar', key="destruir_confirmacio_arxivat")
                # Piso 13N: mateix criteri que a l'altra pestanya -- mai
                # disabled= lligat a un text_input solt (regla 10).
                if st.button("DESTRUIR DEFINITIVAMENT", key="destruir_boto_arxivat"):
                    if not motiu_arx:
                        st.error("Cal escriure un motiu per destruir.")
                    elif confirmacio_arx != "DESTRUIR":
                        st.error('Cal escriure exactament "DESTRUIR" per confirmar.')
                    else:
                        try:
                            destruir_client_arxivat(eleccion_arx, motiu_arx, qui_m)
                        except RuntimeError as e:
                            st.error(str(e))
                        else:
                            st.success(f"Client arxivat '{eleccion_arx}' destruït. Certificat a l'arrel del projecte.")
                            st.rerun()

        with tab_massiva:
            st.caption(
                "Destrueix TOT l'arxivat i tot el retirat/descartat destruïble de TOTS els "
                "clients d'un sol cop -- res viu es toca (nomes el que candidatos_destruccio "
                "ja considera retirat/descartat, i clients ja arxivats)."
            )

            carpeta_arxivats_m = ruta_proyecto("arxivats")
            carpetes_arxivades_totes = sorted(
                n for n in os.listdir(carpeta_arxivats_m) if os.path.isdir(os.path.join(carpeta_arxivats_m, n))
            ) if os.path.isdir(carpeta_arxivats_m) else []

            resum_arxivats = []
            bytes_arxivats_total = 0
            for nom_carpeta in carpetes_arxivades_totes:
                ruta_arx_m = os.path.join(carpeta_arxivats_m, nom_carpeta)
                n_arx = sum(len(archivos) for _, _, archivos in os.walk(ruta_arx_m))
                bytes_arx_m = sum(
                    os.path.getsize(os.path.join(dirpath, f))
                    for dirpath, _, archivos in os.walk(ruta_arx_m) for f in archivos
                )
                resum_arxivats.append((nom_carpeta, n_arx, bytes_arx_m))
                bytes_arxivats_total += bytes_arx_m

            resum_documents = []
            bytes_documents_total = 0
            n_documents_total = 0
            for f in leer_clientes():
                retirats_f, descartats_f = candidatos_destruccio(f["carpeta"])
                items_f = retirats_f + descartats_f
                if not items_f:
                    continue
                bytes_f = sum(
                    os.path.getsize(ruta) for it in items_f for ruta in it["rutas"] if os.path.exists(ruta)
                )
                resum_documents.append((f["carpeta"], items_f, bytes_f))
                bytes_documents_total += bytes_f
                n_documents_total += len(items_f)

            bytes_total = bytes_arxivats_total + bytes_documents_total

            if not carpetes_arxivades_totes and n_documents_total == 0:
                st.info("Res per destruir -- cap client arxivat ni cap document retirat/descartat.")
            else:
                st.warning(
                    f"Es destruiran {len(carpetes_arxivades_totes)} clients arxivats i "
                    f"{n_documents_total} documents retirats/descartats de {len(resum_documents)} clients "
                    f"({bytes_total / (1024 * 1024):.2f} MB en total)."
                )
                with st.expander("Detall"):
                    for nom, n, b in resum_arxivats:
                        st.write(f"- Client arxivat **{nom}**: {n} arxius ({b / (1024 * 1024):.2f} MB)")
                    for carpeta_d, items_f, b in resum_documents:
                        st.write(f"- **{carpeta_d}**: {len(items_f)} documents retirats/descartats ({b / (1024 * 1024):.2f} MB)")

                motiu_massiu = st.text_input("Motiu (obligatori)", key="destruir_motiu_massiu")
                confirmacio_massiva = st.text_input(
                    'Escriu "DESTRUIR TOT" per confirmar', key="destruir_confirmacio_massiva"
                )
                # Piso 13V: mateix criteri de sempre -- mai disabled=
                # lligat a un text_input solt (regla 10). "DESTRUIR TOT"
                # (no "DESTRUIR") a proposit: confirmar aquest abast tan
                # ampli no pot compartir la mateixa paraula que un sol
                # document.
                if st.button("DESTRUIR TOT L'ARXIVAT I RETIRAT", key="destruir_boto_massiu"):
                    if not motiu_massiu:
                        st.error("Cal escriure un motiu per destruir.")
                    elif confirmacio_massiva != "DESTRUIR TOT":
                        st.error('Cal escriure exactament "DESTRUIR TOT" per confirmar.')
                    else:
                        n_clients_arxivats_destruits = 0
                        for nom_carpeta, _, _ in resum_arxivats:
                            try:
                                destruir_client_arxivat(nom_carpeta, motiu_massiu, qui_m)
                                n_clients_arxivats_destruits += 1
                            except RuntimeError as e:
                                st.error(str(e))

                        n_docs_destruits_total = 0
                        for carpeta_d, items_f, _ in resum_documents:
                            destruits_d, omesos_d = destruir_documentos(carpeta_d, items_f, motiu_massiu, qui_m)
                            n_docs_destruits_total += len(destruits_d)
                            if omesos_d:
                                st.error(
                                    f"{carpeta_d}: alguns documents ja no eren destruïbles i s'han saltat: "
                                    + "; ".join(f"{n} ({m})" for n, m in omesos_d)
                                )

                        st.success(
                            f"{n_clients_arxivats_destruits} clients arxivats i {n_docs_destruits_total} "
                            "documents destruïts. Certificats escrits."
                        )
                        if n_docs_destruits_total > 0:
                            # Piso 13V: mateix motiu que a "Documents d'un
                            # client" -- l'informe pot haver quedat
                            # desquadrat pels arxius que ja no hi son.
                            auto_recalcular_sumar_informe()
                        st.rerun()
