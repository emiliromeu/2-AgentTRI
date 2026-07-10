"""Corre las cinco maquinas del pipeline en orden: trocear, extraer,
validar, sumar, informe.

Piso 5B: cada maquina es un proceso aparte (subprocess.run). Si una
revienta con una excepcion no controlada, el traceback se ve en la
consola igual (no se captura la salida) y la cadena para ahi -- las
demas maquinas no llegan a correr. Los AVISOS normales de cada una
(factura con error, REVISAR, lote saltado...) no cuentan como fallo:
esas maquinas siguen terminando con exito y la cadena sigue.

Piso 7: incorpora informe.py como quinta maquina.
"""

import glob
import subprocess
import sys

MAQUINAS = ["trocear.py", "extraer_todas.py", "validar.py", "sumar.py", "informe.py"]


def banner(texto):
    print(f"\n{'=' * 60}\n{texto}\n{'=' * 60}", flush=True)


for maquina in MAQUINAS:
    banner(f"Ejecutando {maquina}")
    resultado = subprocess.run([sys.executable, maquina])
    if resultado.returncode != 0:
        print(f"\nAVISO: {maquina} terminó con un error (código {resultado.returncode}). Parando la cadena aquí.")
        sys.exit(1)

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
