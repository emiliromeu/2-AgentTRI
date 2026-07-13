"""Repara un lot d'escaner que va entrar pel canal de compres quan en
realitat eren factures EMESES (vendes) -- cas real de camp (Piso 13J):
"Lot d'escaner" triat directament al radio "Desti" (sense passar per
la guardia de Compres/Vendes) assumia sempre compres en silenci.

Nomes MOU arxius -- mai n'esborra cap del mon, tret del cau derivat
de rebudes/validadas/ (validar.py sempre el regenera sencer a cada
run, no es font de veritat: es torna a escriure sol a
ingressos_validadas/, ja amb les regles d'ingressos).

Us:
    python migrar_lot.py <carpeta_client> <prefix_del_lot>

<prefix_del_lot> es el nom del PDF original del lot SENSE extensio
(el mateix "base_lote" que trocear.py fa servir per anomenar cada
tall: "{prefix}_p001-002_Emisor.pdf"). Es fan servir tots els arxius
de rebudes/entrada/ i rebudes/procesadas/ que comencin per
"{prefix}_p" -- aquesta segona perque el personal ja pot haver-los
arxivat a ma (encontrar_original tambe els busca alla, Piso 9.3).

Un cop mogut tot, encadena validar.py -> sumar.py -> informe.py
(gratis -- les fitxes ja estan extretes, no calen trocear.py ni
extraer_todas.py).
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path

# Piso 13B: ancla el cwd a la carpeta del propio script -- ver trocear.py.
RAIZ = Path(__file__).resolve().parent
os.chdir(RAIZ)

EXTENSIONES_ORIGINAL = (".pdf", ".jpg", ".jpeg", ".png")

# Misma convencion duplicada que en app.py/sumar.py/extraer_todas.py/
# trocear.py -- ninguna maquina es importable.
RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS = {"davinstal": "Emeses/davinstal"}


def mover(origen, destino):
    """Mueve un archivo si existe; False si no habia nada que mover.
    Nunca sobreescribe -- si el destino ya existe, para con un error
    claro (mismo criterio que arxivar_cliente en app.py, Piso 10.3)."""
    if not os.path.exists(origen):
        return False
    if os.path.exists(destino):
        raise RuntimeError(f"El destí '{destino}' ja existeix -- aturat, no s'ha mogut res més.")
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    shutil.move(origen, destino)
    return True


if len(sys.argv) != 3:
    print("Ús: python migrar_lot.py <carpeta_client> <prefix_del_lot>")
    sys.exit(1)

carpeta, prefijo = sys.argv[1], sys.argv[2]
carpeta_cliente = f"clientes/{carpeta}"
if not os.path.isdir(carpeta_cliente):
    print(f"AVISO: no existeix {carpeta_cliente}")
    sys.exit(1)

destino_ingressos = RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS.get(carpeta, "apartados/ingressos")

prefijo_pagina = f"{prefijo}_p"
nombres_base = set()
for sub in ("rebudes/entrada", "rebudes/procesadas"):
    ruta_sub = os.path.join(carpeta_cliente, sub)
    if not os.path.isdir(ruta_sub):
        continue
    for nombre in os.listdir(ruta_sub):
        if nombre.startswith(prefijo_pagina):
            nombres_base.add(os.path.splitext(nombre)[0])

if not nombres_base:
    print(
        f"AVISO: cap arxiu a rebudes/entrada ni rebudes/procesadas de {carpeta} "
        f"comença per '{prefijo_pagina}'. Res mogut."
    )
    sys.exit(1)

print(f"{len(nombres_base)} documents trobats amb prefix '{prefijo}':")

originals_moguts = 0
fitxes_mogudes = 0
validades_esborrades = 0

for nombre_base in sorted(nombres_base):
    movido_original = False
    for sub in ("rebudes/entrada", "rebudes/procesadas"):
        for ext in EXTENSIONES_ORIGINAL:
            origen = os.path.join(carpeta_cliente, sub, nombre_base + ext)
            if os.path.exists(origen):
                destino = os.path.join(carpeta_cliente, destino_ingressos, nombre_base + ext)
                mover(origen, destino)
                originals_moguts += 1
                movido_original = True

    origen_json = os.path.join(carpeta_cliente, "rebudes/extraidas", nombre_base + ".json")
    destino_json = os.path.join(carpeta_cliente, "apartados/ingressos_extraidas", nombre_base + ".json")
    fitxa_movida = mover(origen_json, destino_json)
    if fitxa_movida:
        fitxes_mogudes += 1

    ruta_validada_vieja = os.path.join(carpeta_cliente, "rebudes/validadas", nombre_base + ".json")
    validada_esborrada = os.path.exists(ruta_validada_vieja)
    if validada_esborrada:
        os.remove(ruta_validada_vieja)
        validades_esborrades += 1

    print(
        f"  {nombre_base}: "
        f"{'original movido' if movido_original else 'original NO trobat'}, "
        f"{'fitxa moguda' if fitxa_movida else 'fitxa NO trobada'}, "
        f"{'validacio antiga esborrada' if validada_esborrada else 'sense validacio antiga'}"
    )

print(
    f"\nResum: {originals_moguts} originals moguts, {fitxes_mogudes} fitxes d'extraidas "
    f"movides, {validades_esborrades} validacions antigues esborrades (es regeneraran)."
)

print("\nExecutant validar.py -> sumar.py -> informe.py (gratis)...")
for maquina in ("validar.py", "sumar.py", "informe.py"):
    resultado = subprocess.run([sys.executable, maquina], cwd=RAIZ)
    if resultado.returncode != 0:
        print(f"AVISO: {maquina} va acabar amb error (codi {resultado.returncode}). Aturant la cadena aquí.")
        sys.exit(1)

print("\nMigració completa.")
