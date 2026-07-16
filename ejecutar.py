"""Corre las cinco maquinas del pipeline en orden: trocear, extraer,
validar, sumar, informe.

Piso 5B: cada maquina es un proceso aparte (subprocess.run). Si una
revienta con una excepcion no controlada, el traceback se ve en la
consola igual (no se captura la salida) y la cadena para ahi -- las
demas maquinas no llegan a correr. Los AVISOS normales de cada una
(factura con error, REVISAR, lote saltado...) no cuentan como fallo:
esas maquinas siguen terminando con exito y la cadena sigue.

Piso 7: incorpora informe.py como quinta maquina.

Piso 13S: candau (processar.lock) -- mai dues "fàbriques" corrent
alhora sobre els mateixos arxius. Guarda PID/qui/hora, s'esborra
sempre per try/finally (mai un candau orfe per un crash); si en troba
un de mort (PID inexistent), l'ignora i avisa. processar.stop es
comprova ENTRE màquines -- si hi és, para la cadena neta (no és un
error) amb un resum parcial; només ejecutar.py l'esborra (mai les
màquines individuals, que no en són propietàries).
"""

import glob
import json
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

# Piso 13B: ancla el cwd a la carpeta del propio script -- ver trocear.py.
# Ademas fija cwd=RAIZ explicito en cada subproceso (pedido explicitamente):
# defensa en profundidad, correcto aunque este chdir se quite algun dia o
# ejecutar.py se llame de otro modo.
RAIZ = Path(__file__).resolve().parent
os.chdir(RAIZ)

# Piso 14E: l'auditoria de coherència tanca el pipeline -- --informatiu
# fa que mai retorni != 0 (la cadena no mor per una divergència, regla
# 4), però les divergències surten com a "AVISO: coherència ..." i el
# detector d'AVISO d'app.py les fa visibles.
MAQUINAS = ["trocear.py", "extraer_todas.py", "validar.py", "sumar.py", "informe.py", "auditar_coherencia.py --informatiu"]

RUTA_LOCK = RAIZ / "processar.lock"
RUTA_STOP = RAIZ / "processar.stop"


def banner(texto):
    print(f"\n{'=' * 60}\n{texto}\n{'=' * 60}", flush=True)


def proceso_vivo(pid):
    """Piso 13U: REGRESSIÓ GREU trobada i arreglada -- a Windows,
    signal.CTRL_C_EVENT val 0, i os.kill(pid, 0) NO és un simple xec
    d'existència com a Unix: crida GenerateConsoleCtrlEvent, que envia
    l'esdeveniment a TOT el grup de consola que comparteix `pid` --
    incloent el Streamlit pare, que comparteix grup de consola amb
    aquest procés (Popen normal, sense CREATE_NEW_PROCESS_GROUP).
    A Windows, OpenProcess/CloseHandle (ctypes, stdlib, cap dependència
    nova) NOMÉS consulten -- mai envien cap senyal. A Unix,
    ProcessLookupError vol dir que no hi ha ningú amb aquest PID -- un
    candau orfe d'un crash anterior que mai va arribar al finally.
    Qualsevol altre resultat (incloent PermissionError -- existeix pero
    no tenim permisos per senyalar-lo) es tracta com a viu."""
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


def candau_viu():
    """Retorna la info (dict) del candau si un altre Processar és viu
    de veritat, None si no n'hi ha o està mort (orfe)."""
    if not RUTA_LOCK.exists():
        return None
    try:
        info = json.loads(RUTA_LOCK.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
    pid = info.get("pid")
    if pid is None or not proceso_vivo(pid):
        return None
    return info


qui = sys.argv[1] if len(sys.argv) > 1 else "terminal"

candau_previ = candau_viu()
if candau_previ:
    print(
        f"AVISO: ja hi ha un processament en marxa (PID {candau_previ['pid']}, "
        f"iniciat per {candau_previ.get('qui')} a les {candau_previ.get('data_inici')}). Aturant-me."
    )
    sys.exit(1)
if RUTA_LOCK.exists():
    print("AVISO: candau orfe trobat (el procés que el va crear ja no existeix) -- s'ignora i se sobreescriu.")

RUTA_LOCK.write_text(
    json.dumps({
        "pid": os.getpid(),
        "abast": "tots els clients",
        "qui": qui,
        "data_inici": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }),
    encoding="utf-8",
)

try:
    aturat_per_stop = False
    for maquina in MAQUINAS:
        banner(f"Ejecutando {maquina}")
        resultado = subprocess.run([sys.executable, *maquina.split()], cwd=RAIZ)
        if resultado.returncode != 0:
            print(f"\nAVISO: {maquina} terminó con un error (código {resultado.returncode}). Parando la cadena aquí.")
            sys.exit(1)
        if RUTA_STOP.exists():
            print(f"\nATURAT per petició de l'usuari (després de {maquina}) -- reprèn amb Processar quan vulguis.")
            aturat_per_stop = True
            break

    if not aturat_per_stop:
        banner("Pipeline completo")
        print("Máquinas ejecutadas:", ", ".join(MAQUINAS))

        rutas = sorted(glob.glob("clientes/*/sumatorios_2026.xlsx"))
        print("\nSumatorios generados:")
        for ruta in rutas:
            print(f"  {ruta}")

        rutas_informe = sorted(glob.glob("clientes/*/informe_2026.html"))
        print("\nInformes generados:")
        for ruta in rutas_informe:
            print(f"  {ruta}")
finally:
    RUTA_LOCK.unlink(missing_ok=True)
    RUTA_STOP.unlink(missing_ok=True)
