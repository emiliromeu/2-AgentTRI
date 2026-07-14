"""Mou documents entre el flux de compres (rebudes) i el de vendes
(ingressos) -- perque un document mai queda enganxat a un flux erroni
sense sortida (Piso 13J: un lot sencer mal processat; Piso 13L: UNA
factura, des de l'app).

A diferencia de les cinc "maquines" (trocear/extraer_todas/validar/
sumar/informe -- "ninguna maquina es importable", s'executen senceres
en carregar-se), AQUEST fitxer SI esta pensat per ser importat: les
funcions de mes avall no tenen cap efecte en importar-les (cap bucle
de nivell superior), nomes es defineixen. app.py les crida directament
(import migrar_lot) per al cas d'una sola factura o d'un lot
seleccionat a Revisio; l'us de terminal (un lot sencer, sense passar
per l'app) viu tot dins de `if __name__ == "__main__":`, mes avall.

Nomes MOU arxius -- mai n'esborra cap del mon, tret del cau derivat
de validadas/ del flux ORIGEN (validar.py sempre el regenera sencer a
cada run, no es font de veritat: es torna a escriure sol al flux nou,
ja amb les regles que toquen).

Us de terminal:
    python migrar_lot.py <carpeta_client> <prefix_del_lot>

<prefix_del_lot> es el nom del PDF original del lot SENSE extensio
(el mateix "base_lote" que trocear.py fa servir per anomenar cada
tall: "{prefix}_p001-002_Emisor.pdf"). Es busca a TOTS DOS fluxos
(rebudes i ingressos, Piso 13L) per detectar sol la direccio.
"""

import csv
import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parent

EXTENSIONES_ORIGINAL = (".pdf", ".jpg", ".jpeg", ".png")

# Misma convencion duplicada que en app.py/sumar.py/extraer_todas.py/
# trocear.py -- ninguna maquina es importable (aquest fitxer es
# l'excepcio explicada al docstring, pero la constant es la mateixa).
RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS = {"davinstal": "Emeses/davinstal"}

CAMPOS_MOVIMENTS_CSV = ["data", "arxiu", "de", "a", "motiu", "qui"]


def ruta_cliente(carpeta, *partes):
    """Piso 13L: ruta absoluta ancorada a RAIZ (la carpeta d'aquest
    script), independent del cwd del proces que crida -- necessari
    perque app.py crida moure_de_flux() sense haver fet cap os.chdir
    (fa servir les seves propies rutes absolutes, RAIZ_PROYECTO)."""
    return os.path.join(RAIZ, "clientes", carpeta, *partes)


def config_flujo(carpeta, flujo):
    """Retorna (carpeta_original, carpeta_extraidas, carpeta_validadas)
    -- relatives a ruta_cliente(carpeta) -- del flux "rebudes" o
    "ingressos" d'aquest client."""
    if flujo == "rebudes":
        return ("rebudes/entrada", "rebudes/extraidas", "rebudes/validadas")
    origen_ingressos = RUTAS_ORIGEN_INGRESSOS_PERSONALIZADAS.get(carpeta, "apartados/ingressos")
    return (origen_ingressos, "apartados/ingressos_extraidas", "apartados/ingressos_validadas")


def mover(origen, destino):
    """Mou un arxiu si existeix; False si no hi havia res a moure.
    Mai sobreescriu -- si el desti ja existeix, para amb un error
    clar (mateix criteri que arxivar_cliente a app.py, Piso 10.3)."""
    if not os.path.exists(origen):
        return False
    if os.path.exists(destino):
        raise RuntimeError(f"El destí '{destino}' ja existeix -- aturat, no s'ha mogut res més.")
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    shutil.move(origen, destino)
    return True


def escribir_moviment_flux(carpeta, nombre_base, de, a, motiu, qui, data_actual):
    ruta = ruta_cliente(carpeta, "moviments_flux.csv")
    escribir_cabecera = not os.path.exists(ruta)
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    with open(ruta, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if escribir_cabecera:
            writer.writerow(CAMPOS_MOVIMENTS_CSV)
        writer.writerow([data_actual, nombre_base, de, a, motiu, qui])


def moure_de_flux(carpeta, documentos, motiu, qui):
    """documentos: llista de (nombre_base, flujo_actual) -- flujo_actual
    es "rebudes" o "ingressos". Mou cada document cap al flux
    CONTRARI: l'original (mirant tambe rebudes/procesadas si ve de
    rebudes, Piso 9.3) i la fitxa d'extraidas; esborra el veredicte
    vell de validadas/ (validar.py el refara amb les regles del flux
    nou). Escriu una fila a moviments_flux.csv per cada document que
    de veritat s'hagi mogut -- si es crida dues vegades sobre un
    document ja mogut, la segona no fa res ni deixa una fila fantasma
    (idempotent, regla 5). Retorna [(nombre_base, es_va_moure), ...]."""
    data_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    resultados = []
    for nombre_base, flujo_actual in documentos:
        flujo_nuevo = "ingressos" if flujo_actual == "rebudes" else "rebudes"
        origen_orig, origen_extr, origen_val = config_flujo(carpeta, flujo_actual)
        destino_orig, destino_extr, _ = config_flujo(carpeta, flujo_nuevo)

        candidatos_origen = [ruta_cliente(carpeta, origen_orig)]
        if flujo_actual == "rebudes":
            candidatos_origen.append(ruta_cliente(carpeta, "rebudes/procesadas"))

        movido_original = False
        for carpeta_candidata in candidatos_origen:
            for ext in EXTENSIONES_ORIGINAL:
                origen = os.path.join(carpeta_candidata, nombre_base + ext)
                if os.path.exists(origen):
                    destino = os.path.join(ruta_cliente(carpeta, destino_orig), nombre_base + ext)
                    mover(origen, destino)
                    movido_original = True

        origen_json = ruta_cliente(carpeta, origen_extr, nombre_base + ".json")
        destino_json = ruta_cliente(carpeta, destino_extr, nombre_base + ".json")
        fitxa_movida = mover(origen_json, destino_json)

        ruta_validada_vieja = ruta_cliente(carpeta, origen_val, nombre_base + ".json")
        if os.path.exists(ruta_validada_vieja):
            os.remove(ruta_validada_vieja)

        es_va_moure = movido_original or fitxa_movida
        if es_va_moure:
            escribir_moviment_flux(carpeta, nombre_base, flujo_actual, flujo_nuevo, motiu, qui, data_actual)
        resultados.append((nombre_base, es_va_moure))

    return resultados


def moure_a_client(carpeta_origen, carpeta_desti, documentos, motiu, qui, flujo_desti=None):
    """Piso 13Q: com moure_de_flux, pero cap a un client DIFERENT --
    l'error de destinació és pujar la factura al client equivocat, no
    nomes al flux equivocat. documentos: llista de (nombre_base,
    flujo_origen). Si flujo_desti es None, es queda al MATEIX flux al
    client nou (nomes canvia el client); si es dona, tambe canvia de
    flux en el mateix pas.

    No es toca moure_de_flux -- les tres crides existents (app.py)
    depenen del seu comportament actual, format de moviments_flux.csv
    inclos. Per distingir un moviment ENTRE clients d'un moviment de
    flux sense afegir cap columna al csv (que ja existeix a
    producció): aqui "de"/"a" es guarden com "carpeta:flux" (ex.
    "davinstal:rebudes") en comptes del nom pla del flux -- app.py ho
    reconeix mirant si hi ha ":" i ho tradueix per mostrar-ho."""
    data_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    resultados = []
    for nombre_base, flujo_origen in documentos:
        flujo_final = flujo_desti or flujo_origen
        origen_orig, origen_extr, origen_val = config_flujo(carpeta_origen, flujo_origen)
        destino_orig, destino_extr, _ = config_flujo(carpeta_desti, flujo_final)

        candidatos_origen = [ruta_cliente(carpeta_origen, origen_orig)]
        if flujo_origen == "rebudes":
            candidatos_origen.append(ruta_cliente(carpeta_origen, "rebudes/procesadas"))

        movido_original = False
        for carpeta_candidata in candidatos_origen:
            for ext in EXTENSIONES_ORIGINAL:
                origen = os.path.join(carpeta_candidata, nombre_base + ext)
                if os.path.exists(origen):
                    destino = os.path.join(ruta_cliente(carpeta_desti, destino_orig), nombre_base + ext)
                    mover(origen, destino)
                    movido_original = True

        origen_json = ruta_cliente(carpeta_origen, origen_extr, nombre_base + ".json")
        destino_json = ruta_cliente(carpeta_desti, destino_extr, nombre_base + ".json")
        fitxa_movida = mover(origen_json, destino_json)

        ruta_validada_vieja = ruta_cliente(carpeta_origen, origen_val, nombre_base + ".json")
        if os.path.exists(ruta_validada_vieja):
            os.remove(ruta_validada_vieja)

        es_va_moure = movido_original or fitxa_movida
        if es_va_moure:
            # Piso 13V: abans NOMES s'escrivia al desti -- l'origen mai
            # quedava marcat com "sense recalcular" tot i haver perdut
            # un document. Mateixa fila (mateix "de"/"a") als DOS csv.
            escribir_moviment_flux(
                carpeta_origen, nombre_base,
                f"{carpeta_origen}:{flujo_origen}", f"{carpeta_desti}:{flujo_final}",
                motiu, qui, data_actual,
            )
            escribir_moviment_flux(
                carpeta_desti, nombre_base,
                f"{carpeta_origen}:{flujo_origen}", f"{carpeta_desti}:{flujo_final}",
                motiu, qui, data_actual,
            )
        resultados.append((nombre_base, es_va_moure))

    return resultados


if __name__ == "__main__":
    # Piso 13B: ancla el cwd a la carpeta del propio script -- ver
    # trocear.py. Nomes aqui (mode terminal): les funcions d'amunt ja
    # fan servir rutes absolutes, no en depenen.
    os.chdir(RAIZ)

    if len(sys.argv) != 3:
        print("Ús: python migrar_lot.py <carpeta_client> <prefix_del_lot>")
        sys.exit(1)

    carpeta, prefijo = sys.argv[1], sys.argv[2]
    if not os.path.isdir(ruta_cliente(carpeta)):
        print(f"AVISO: no existeix clientes/{carpeta}")
        sys.exit(1)

    prefijo_pagina = f"{prefijo}_p"
    # Piso 13L: es busca als DOS fluxos (abans nomes rebudes) per
    # detectar sola la direccio -- un lot mal processat pot haver
    # entrat per qualsevol dels dos canals.
    documentos = set()
    for flujo in ("rebudes", "ingressos"):
        carpeta_original, _, _ = config_flujo(carpeta, flujo)
        candidatos = [ruta_cliente(carpeta, carpeta_original)]
        if flujo == "rebudes":
            candidatos.append(ruta_cliente(carpeta, "rebudes/procesadas"))
        for ruta_sub in candidatos:
            if not os.path.isdir(ruta_sub):
                continue
            for nombre in os.listdir(ruta_sub):
                if nombre.startswith(prefijo_pagina):
                    documentos.add((os.path.splitext(nombre)[0], flujo))

    documentos = sorted(documentos)
    if not documentos:
        print(
            f"AVISO: cap arxiu a cap dels dos fluxos de {carpeta} comença per "
            f"'{prefijo_pagina}'. Res mogut."
        )
        sys.exit(1)

    print(f"{len(documentos)} documents trobats amb prefix '{prefijo}':")
    resultados = moure_de_flux(
        carpeta, documentos, motiu=f"migrar_lot.py: lot {prefijo}", qui="terminal",
    )
    for nombre_base, es_va_moure in resultados:
        print(f"  {nombre_base}: {'mogut' if es_va_moure else 'NO trobat (res a moure)'}")

    print(f"\nResum: {sum(1 for _, ok in resultados if ok)} de {len(resultados)} documents moguts.")

    print("\nExecutant validar.py -> sumar.py -> informe.py (gratis)...")
    for maquina in ("validar.py", "sumar.py", "informe.py"):
        resultado = subprocess.run([sys.executable, maquina], cwd=RAIZ)
        if resultado.returncode != 0:
            print(f"AVISO: {maquina} va acabar amb error (codi {resultado.returncode}). Aturant la cadena aquí.")
            sys.exit(1)

    print("\nMigració completa.")
