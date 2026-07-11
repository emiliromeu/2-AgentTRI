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
"""

import csv
import io
import os
import shutil
import subprocess
import sys
from datetime import datetime

import streamlit as st
from PIL import Image
from pillow_heif import register_heif_opener

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
    with open(RUTA_CLIENTES_CSV) as f:
        return list(csv.DictReader(f))


def escribir_clientes(filas):
    """Reescriu clientes.csv sencer -- l'usen anadir_cliente (afegir
    una fila) i arxivar_cliente (Piso 10.3, treure'n una)."""
    with open(RUTA_CLIENTES_CSV, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CAMPOS_CLIENTES_CSV)
        writer.writeheader()
        writer.writerows(filas)


def anadir_cliente(nif, nombre, carpeta):
    filas = leer_clientes()
    filas.append({"nif": nif, "nombre": nombre, "carpeta": carpeta})
    escribir_clientes(filas)
    os.makedirs(ruta_proyecto("clientes", carpeta, "rebudes", "entrada"), exist_ok=True)
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
        with open(ruta_gitignore) as f:
            lineas = f.read().splitlines()
    if "arxivats/" not in lineas:
        with open(ruta_gitignore, "a") as f:
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
    with open(ruta_registro, "a", newline="") as f:
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


def boton_obrir(etiqueta, ruta_absoluta, key):
    """Piso 10.2: st.link_button con file:// no funciona -- Chromium
    bloquea la navegacion a file:// desde un origen http://localhost
    (confirmado con la consola real: "Not allowed to load local
    resource"). En vez de eso, se abre del lado Python con `open`
    (macOS), directamente en su sitio -- para que los enlaces
    relativos del informe a sus originales sigan vivos."""
    existe = os.path.exists(ruta_absoluta)
    if st.button(etiqueta, disabled=not existe, key=key):
        resultado = subprocess.run(["open", ruta_absoluta])
        if resultado.returncode != 0:
            st.error(f"No s'ha pogut obrir {ruta_absoluta} (codi {resultado.returncode}).")


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
        with st.popover("🗄️ Arxivar client"):
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


st.set_page_config(page_title="Agent TRIMESTRE", layout="wide")

if "log_proces" not in st.session_state:
    st.session_state["log_proces"] = None

st.title("Agent TRIMESTRE")
vista = st.sidebar.radio("Navegació", ["Clients", "Afegir factures", "Processar"])

# ----------------------------------------------------------------------
if vista == "Clients":
    st.header("Clients")
    clientes = leer_clientes()

    if not clientes:
        st.info("Encara no hi ha cap client donat d'alta.")
    else:
        for fila in clientes:
            tarjeta_cliente(fila, "clients")

    with st.expander("➕ Nou client"):
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

        if st.button("Crear client", key="nou_crear"):
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
        destino = st.radio("Destí", ["Compres", "Vendes"], horizontal=True)
        archivos = st.file_uploader(
            "Arrossega els arxius aquí",
            type=EXTENSIONES_PERMITIDAS,
            accept_multiple_files=True,
        )

        if st.button("Desar arxius", disabled=not archivos):
            carpeta_destino = ruta_destino_factures(carpeta, destino)
            nombres_finales = [guardar_archivo(a, carpeta_destino) for a in archivos]
            st.success(f"S'han desat {len(nombres_finales)} arxius a `{carpeta_destino}`:")
            for nombre in nombres_finales:
                st.write(f"- {nombre}")

# ----------------------------------------------------------------------
elif vista == "Processar":
    st.header("Processar")
    clientes = leer_clientes()

    if not clientes:
        st.info("Encara no hi ha cap client donat d'alta -- res per processar.")
    else:
        if st.button("Processar"):
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
