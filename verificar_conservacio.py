"""Piso 13Y (B5): bateria de conservació total -- PERMANENT.

Recalcula EN INDEPENDÈNCIA (llegint només rebudes/validadas i
apartados/ingressos_validadas + decisions.csv -- MAI important
sumar.py/informe.py/app.py, un oracle que crida el codi que vol
auditar no serveix de res) els cubs per tipus d'IVA, Exempta, Total,
Σ retencions i RESULTAT que sumar.py escriu a l'Excel, per a CADA
trimestre, i els compara CEL·LA A CEL·LA (tolerància 0,01 €) contra el
fitxer ja generat. També comprova, fila a fila del DETALL, que Líquid
= Total − Retenció -- una consistència pura dins del mateix Excel.

Les correccions de correccions.csv NO calen llegir-les a part: quan
una factura arriba a validadas/*.json ja les porta aplicades (validar.py
les aplica abans d'escriure-hi) -- per això aquest script només
necessita validadas/*.json + decisions.csv per tenir la mateixa
informació que sumar.py fa servir per sumar.

Busca les etiquetes conegudes ("21", "10", "EXEMPTA", "ALTRES", "TOTAL",
"Σ RETENCIONS SUPORTADES", "TOTAL RETENCIONS", "RESULTAT IVA...") en
comptes d'assumir files fixes -- el layout de sumar.py pot canviar de
fila d'una execució a l'altra sense que aquest script hagi de canviar.

Regla CLAUDE.md (Piso 13Y): cap piso futur toca sumar.py sense passar
aquest script contra els 4 clients reals abans de commitejar.

Ús:
    python3 verificar_conservacio.py <carpeta_client> [<carpeta2> ...]
    python3 verificar_conservacio.py            # tots els clients
"""
import csv
import json
import os
import sys

import openpyxl

TIPOS_IVA = [0, 4, 5, 10, 12, 21]
TOLERANCIA = 0.01


def leer_clientes():
    ruta = "clientes/clientes.csv"
    if not os.path.exists(ruta):
        return []
    with open(ruta, encoding="utf-8") as f:
        return list(csv.DictReader(f))


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
    """Còpia INDEPENDENT de la mateixa gramàtica (Piso 13M) que
    app.py/sumar.py/informe.py -- última fila per arxiu, "revertir"
    treu l'arxiu del diccionari."""
    ruta = f"{carpeta_cliente}/decisions.csv"
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
    return decision.get("accion") if decision else None


def trimestre_de(fecha):
    if not fecha:
        return None
    try:
        mes = int(fecha[5:7])
    except (ValueError, IndexError):
        return None
    return f"{(mes - 1) // 3 + 1}T"


def cuadre_esperado(facturas, decisiones):
    """Base/quota per tipus (incl. "otros"/"exenta"), total i retenció
    de les factures que COMPTEN -- estado==OK o decisió aprovar, mai
    descartar. Mateixa regla que sumar_bloque, calculada aquí en
    independència total.

    Piso 14B: exenta és un camp DE LÍNIA (validar.py ja el resol i
    deriva el de factura) -- una factura mixta reparteix cada línia
    al seu cub, igual que sumar_bloque."""
    sumas = {t: {"base": 0.0, "cuota": 0.0} for t in TIPOS_IVA}
    sumas["otros"] = {"base": 0.0, "cuota": 0.0}
    sumas["exenta"] = {"base": 0.0, "cuota": 0.0}
    total = 0.0
    retencion = 0.0
    for nombre, datos in facturas:
        decision = decisiones.get(nombre)
        if estat_efectiu(decision) == "descartar":
            continue
        cuenta = datos.get("estado") == "OK" or estat_efectiu(decision) == "aprovar"
        if not cuenta:
            continue
        for linea in datos.get("lineas_iva") or []:
            tipo = linea.get("tipo_iva")
            clave = "exenta" if linea.get("exenta") else (tipo if tipo in TIPOS_IVA else "otros")
            sumas[clave]["base"] += linea.get("base") or 0
            sumas[clave]["cuota"] += linea.get("cuota") or 0
        total += datos.get("total") or 0
        retencion += datos.get("retencion_cuota") or 0
    return sumas, round(total, 2), round(retencion, 2)


SECCIONS_QUE_TANQUEN = {
    "RESULTAT DEL TRIMESTRE", "DESPESES", "INGRESSOS",
    "DETALL DESPESES", "DETALL INGRESSOS",
    "PENDENT DE REVISIÓ", "PENDENTS — NO SUMEN", "DESCARTATS",
}


def tipo_clave_de(etiqueta):
    if isinstance(etiqueta, (int, float)) and not isinstance(etiqueta, bool):
        return int(etiqueta)
    if etiqueta == "ALTRES":
        return "otros"
    if etiqueta == "EXEMPTA":
        return "exenta"
    return None


def leer_hoja_trimestre(ws):
    """Escaneja el full buscant les etiquetes conegudes -- mai assumeix
    files fixes. Retorna els cubs/totals de RESULTAT/DESPESES/INGRESSOS
    i, per cada fila de DETALL amb un Total, (total, retenció, líquid)
    per verificar la consistència Líquid = Total - Retenció."""
    valores = {
        "resultat_g": {}, "resultat_i": {}, "resultat_iva": None,
        "despeses_base": {}, "despeses_cuota": {}, "despeses_total": None, "despeses_retencio": None,
        "ingressos_base": {}, "ingressos_cuota": {}, "ingressos_total": None, "ingressos_retencio": None,
        "filas_liquid": [],
    }
    seccion = None
    mapa_cols_detall = None
    for fila in ws.iter_rows():
        etiqueta = fila[0].value

        if etiqueta in SECCIONS_QUE_TANQUEN:
            seccion = {
                "RESULTAT DEL TRIMESTRE": "resultat", "DESPESES": "despeses", "INGRESSOS": "ingressos",
            }.get(etiqueta)
            mapa_cols_detall = None
            continue
        if etiqueta is None:
            continue

        # DETALL: la fila de capçalera ("Data", "Núm. factura"...) marca
        # per etiqueta on és cada columna -- es rellegeix a cada bloc
        # (FACTURES QUE SUMEN i PENDENTS -- NO SUMEN en tenen una cada un)
        # perquè cap posició de columna es dona per fixa.
        if etiqueta == "Data":
            mapa_cols_detall = {cel.value: idx for idx, cel in enumerate(fila, start=1) if cel.value}
            continue
        if mapa_cols_detall is not None:
            col_total = mapa_cols_detall.get("Total")
            col_liquid = mapa_cols_detall.get("Líquid (€)")
            col_retencio = mapa_cols_detall.get("Retenció (€)")
            if col_total and col_liquid:
                valor_total = fila[col_total - 1].value
                if valor_total is not None:
                    valor_liquid = fila[col_liquid - 1].value
                    valor_retencio = (fila[col_retencio - 1].value if col_retencio else None) or 0
                    valores["filas_liquid"].append((
                        fila[1].value, valor_total, valor_retencio, valor_liquid,
                    ))

        tipo_clave = tipo_clave_de(etiqueta)
        if seccion == "resultat" and tipo_clave is not None:
            valores["resultat_g"][tipo_clave] = fila[1].value or 0
            valores["resultat_i"][tipo_clave] = fila[2].value or 0
        elif seccion == "resultat" and etiqueta == "RESULTAT IVA (repercutit - suportat)":
            valores["resultat_iva"] = fila[2].value or 0
        elif seccion == "despeses" and tipo_clave is not None:
            valores["despeses_base"][tipo_clave] = fila[1].value or 0
            valores["despeses_cuota"][tipo_clave] = fila[2].value or 0
        elif seccion == "despeses" and etiqueta == "TOTAL":
            valores["despeses_total"] = fila[3].value or 0
        elif seccion == "despeses" and etiqueta in ("Σ RETENCIONS SUPORTADES", "TOTAL RETENCIONS"):
            valores["despeses_retencio"] = fila[3].value or 0
        elif seccion == "ingressos" and tipo_clave is not None:
            valores["ingressos_base"][tipo_clave] = fila[1].value or 0
            valores["ingressos_cuota"][tipo_clave] = fila[2].value or 0
        elif seccion == "ingressos" and etiqueta == "TOTAL":
            valores["ingressos_total"] = fila[3].value or 0
        elif seccion == "ingressos" and etiqueta in ("Σ RETENCIONS SUPORTADES", "TOTAL RETENCIONS"):
            valores["ingressos_retencio"] = fila[3].value or 0

    return valores


def comparar(etiqueta, esperado, real, detalls=""):
    if abs((real or 0) - (esperado or 0)) <= TOLERANCIA:
        return True
    print(f"  ✗ {etiqueta}: esperat {esperado}, real {real}{detalls}")
    return False


def verificar_cliente(carpeta):
    carpeta_cliente = f"clientes/{carpeta}"
    ruta_xlsx = f"{carpeta_cliente}/sumatorios_2026.xlsx"
    if not os.path.exists(ruta_xlsx):
        print(f"{carpeta}: sense Excel generat, es salta.")
        return True

    gastos = cargar_validadas(f"{carpeta_cliente}/rebudes/validadas")
    ingresos = cargar_validadas(f"{carpeta_cliente}/apartados/ingressos_validadas")
    decisiones = cargar_decisiones(carpeta_cliente)

    trimestres = {}
    for nombre, datos in gastos:
        t = trimestre_de(datos.get("fecha_factura"))
        if t:
            trimestres.setdefault(t, {"gastos": [], "ingresos": []})["gastos"].append((nombre, datos))
    for nombre, datos in ingresos:
        t = trimestre_de(datos.get("fecha_factura"))
        if t:
            trimestres.setdefault(t, {"gastos": [], "ingresos": []})["ingresos"].append((nombre, datos))

    wb = openpyxl.load_workbook(ruta_xlsx, data_only=True)
    tot_ok = True

    for trimestre, datos_t in sorted(trimestres.items()):
        if trimestre not in wb.sheetnames:
            print(f"{carpeta} / {trimestre}: ✗ no hi ha full a l'Excel per a aquest trimestre")
            tot_ok = False
            continue

        esperado_g, total_g, retencio_g = cuadre_esperado(datos_t["gastos"], decisiones)
        esperado_i, total_i, retencio_i = cuadre_esperado(datos_t["ingresos"], decisiones)
        resultat_iva_esperat = round(
            sum(esperado_i[t]["cuota"] for t in esperado_i) - sum(esperado_g[t]["cuota"] for t in esperado_g), 2
        )

        real = leer_hoja_trimestre(wb[trimestre])
        ok = True

        for tipo in real["despeses_base"]:
            ok &= comparar(f"{trimestre} DESPESES base {tipo}", esperado_g.get(tipo, {}).get("base"), real["despeses_base"][tipo])
        for tipo in real["despeses_cuota"]:
            ok &= comparar(f"{trimestre} DESPESES quota {tipo}", esperado_g.get(tipo, {}).get("cuota"), real["despeses_cuota"][tipo])
        ok &= comparar(f"{trimestre} DESPESES total", total_g, real["despeses_total"])
        ok &= comparar(f"{trimestre} DESPESES retenció", retencio_g, real["despeses_retencio"])

        for tipo in real["ingressos_base"]:
            ok &= comparar(f"{trimestre} INGRESSOS base {tipo}", esperado_i.get(tipo, {}).get("base"), real["ingressos_base"][tipo])
        for tipo in real["ingressos_cuota"]:
            ok &= comparar(f"{trimestre} INGRESSOS quota {tipo}", esperado_i.get(tipo, {}).get("cuota"), real["ingressos_cuota"][tipo])
        ok &= comparar(f"{trimestre} INGRESSOS total", total_i, real["ingressos_total"])
        ok &= comparar(f"{trimestre} INGRESSOS retenció", retencio_i, real["ingressos_retencio"])

        ok &= comparar(f"{trimestre} RESULTAT IVA", resultat_iva_esperat, real["resultat_iva"])

        for num_factura, total, retencio, liquid in real["filas_liquid"]:
            esperat_liquid = round((total or 0) - (retencio or 0), 2)
            ok &= comparar(f"{trimestre} Líquid de factura {num_factura}", esperat_liquid, liquid)

        if ok:
            print(f"{carpeta} / {trimestre}: ✓ ({len(real['filas_liquid'])} files de detall comprovades)")
        tot_ok &= ok

    return tot_ok


def main():
    carpetas = sys.argv[1:]
    if not carpetas:
        carpetas = [f["carpeta"] for f in leer_clientes()]

    tot_ok = True
    for carpeta in carpetas:
        tot_ok &= verificar_cliente(carpeta)

    print()
    print("CONSERVACIÓ TOTAL: ✓ TOT QUADRA" if tot_ok else "CONSERVACIÓ TOTAL: ✗ HI HA DIVERGÈNCIES -- veure detall a dalt")
    sys.exit(0 if tot_ok else 1)


if __name__ == "__main__":
    main()
