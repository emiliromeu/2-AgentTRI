"""App Streamlit -- closca sobre la fabrica. Piso 10.1.

Nomes coloca arxius on ja els busca extraer_todas.py/trocear.py, i
llanca ejecutar.py com a subproces -- exactament com ho faria Emili a
ma. No calcula ni valida res: la logica de negoci viu nomes a les
cinc maquines (trocear, extraer_todas, validar, sumar, informe), que
aquesta app no toca ni importa (totes executen el seu bucle principal
en carregar-se -- convencio des del piso 2).
"""

import csv
import io
import os
import subprocess
import sys
from datetime import datetime

import streamlit as st
from PIL import Image
from pillow_heif import register_heif_opener

register_heif_opener()

RUTA_CLIENTES_CSV = "clientes/clientes.csv"
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


def anadir_cliente(nif, nombre, carpeta):
    filas = leer_clientes()
    filas.append({"nif": nif, "nombre": nombre, "carpeta": carpeta})
    with open(RUTA_CLIENTES_CSV, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CAMPOS_CLIENTES_CSV)
        writer.writeheader()
        writer.writerows(filas)
    os.makedirs(f"clientes/{carpeta}/rebudes/entrada", exist_ok=True)
    os.makedirs(f"clientes/{carpeta}/apartados/ingressos", exist_ok=True)


def ruta_destino_factures(carpeta, destino):
    if destino == "Compres":
        return f"clientes/{carpeta}/rebudes/entrada"
    origen_ingressos = RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS.get(carpeta, "apartados/ingressos")
    return f"clientes/{carpeta}/{origen_ingressos}"


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


def boton_obrir(etiqueta, ruta_relativa, key):
    if os.path.exists(ruta_relativa):
        st.link_button(etiqueta, f"file://{os.path.abspath(ruta_relativa)}", key=key)
    else:
        st.button(etiqueta, disabled=True, key=key)


st.set_page_config(page_title="Agent TRIMESTRE", layout="wide")

if "cliente_actiu" not in st.session_state:
    st.session_state["cliente_actiu"] = None
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
            carpeta = fila["carpeta"]
            ruta_informe = f"clientes/{carpeta}/informe_2026.html"
            ruta_excel = f"clientes/{carpeta}/sumatorios_2026.xlsx"
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
                    boton_obrir("Obrir informe", ruta_informe, key=f"informe_{carpeta}")
                with col3:
                    boton_obrir("Obrir Excel", ruta_excel, key=f"excel_{carpeta}")

    with st.expander("➕ Nou client"):
        with st.form("form_nou_client", clear_on_submit=True):
            nombre = st.text_input("Nom del client")
            nif_input = st.text_input("NIF")
            carpeta_sugerida = slug(nombre) if nombre else ""
            carpeta = st.text_input("Carpeta (identificador intern)", value=carpeta_sugerida)
            continuar_igualment = st.checkbox("El NIF no quadra la lletra de control -- continuar igualment")
            enviar = st.form_submit_button("Crear client")

        if enviar:
            nif = normalizar_nif(nif_input)
            carpetas_existentes = {f["carpeta"] for f in leer_clientes()}
            if not nombre or not nif or not carpeta:
                st.error("Falten camps: nom, NIF i carpeta són obligatoris.")
            elif carpeta in carpetas_existentes:
                st.error(f"Ja existeix un client amb la carpeta '{carpeta}'.")
            elif validar_nif(nif) is False and not continuar_igualment:
                st.warning(
                    "El NIF no supera la validació de la lletra de control. "
                    "Marca 'continuar igualment' si n'estàs segur i torna a enviar."
                )
            else:
                anadir_cliente(nif, nombre, carpeta)
                st.session_state["cliente_actiu"] = carpeta
                st.success(f"Client '{nombre}' creat (carpeta `{carpeta}`).")
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
            st.session_state["cliente_actiu"] = carpeta
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

        if st.session_state["log_proces"] and st.session_state["cliente_actiu"]:
            carpeta = st.session_state["cliente_actiu"]
            fila_cliente = next((f for f in clientes if f["carpeta"] == carpeta), None)
            if fila_cliente:
                st.subheader(f"Resum de {fila_cliente['nombre']}")
                linias_cliente = [
                    linia for linia in st.session_state["log_proces"].splitlines()
                    if linia.startswith(f"{carpeta} ")
                ]
                st.code("\n".join(linias_cliente) if linias_cliente else "Sense línies per a aquest client.")

                ruta_informe = f"clientes/{carpeta}/informe_2026.html"
                ruta_excel = f"clientes/{carpeta}/sumatorios_2026.xlsx"
                col1, col2 = st.columns(2)
                with col1:
                    boton_obrir("Obrir informe", ruta_informe, key="final_informe")
                with col2:
                    boton_obrir("Obrir Excel", ruta_excel, key="final_excel")
