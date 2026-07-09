# Esquema canónico — Agente TRIMESTRE (v1)

Contrato entre las tres piezas: el LLM extrae, el código deriva y valida, el humano decide.
Pilotado sobre 14 facturas rebudes reales de Penedes Languages SL (proveedores: Gestoria Olivella y Tim Institute).

## 1. Lo que extrae el LLM — una llamada por factura, sale este JSON

```json
{
  "proveedor": "Tim Institute S.L.",
  "nif_proveedor": "B65099673",
  "num_factura": "F260224",
  "fecha_factura": "2026-05-07",
  "receptor": "Penedes Languages S.L.",
  "nif_receptor": "B66515529",
  "lineas_iva": [
    { "tipo_iva": 4, "base": 272.89, "cuota": 10.91 }
  ],
  "total": 283.80,
  "retencion_pct": 0,
  "retencion_cuota": 0.0,
  "exenta": false,
  "observaciones": ""
}
```

Reglas de extracción:

- `lineas_iva` es una lista: una factura con dos tipos de IVA produce dos elementos, y el escritor genera **una fila del llibre por elemento**.
- `exenta = true` cuando la factura lo diga (ej. F260227, formación). Se registra con tipo 0 — el llibre lo recoge en la columna BASE 0.
- Campo que no aparece en el PDF → `null`. Jamás inventado. `observaciones` es texto libre para que el LLM anote rarezas.
- No se extrae nada que el llibre no consuma (el vencimiento queda fuera en v1: cada campo extra es superficie de error).

## 2. Lo que deriva el código — el LLM jamás decide esto

- `data_apunt`: fecha del procesado.
- `ordre`: correlativo del registro, asignado secuencialmente por hoja.
- Hoja destino (1t–4t): según `fecha_factura`. Las 14 del piloto caen en 1t, 2t y 3t.
- `categoria` (columna G): lookup en `memoria.csv` por `nif_proveedor`. Sin fila → estado PROVEEDOR NUEVO.
- Reparto a las columnas de gasto (M–AB) y a bases/cuotas por tipo: calculado en Python y escrito como **valores**, no fórmulas — el llibre original arrastra #REF! rotos y esa fragilidad no se hereda.

## 3. La red de validación — en código, y ruidosa

- Aritmética por línea: `|base × tipo/100 − cuota| ≤ 0,02`. La tolerancia no es teórica: la factura real F260224 da 272,89 × 4% = 10,9156 y el PDF dice 10,91 — un redondeo estricto la marcaría en falso.
- Total: `|Σ bases + Σ cuotas − total| ≤ 0,02`.
- Receptor: `nif_receptor` debe coincidir con el CIF del cliente de la carpeta. Las 14 del piloto: B66515529 ✓.
- Duplicado: `(nif_proveedor, num_factura)` único en todo el llibre.
- Retención: `retencion_pct > 0` con cuota real → REVISAR (el llibre no tiene columna para representarla).
- Cualquier `null` en campo obligatorio → REVISAR.

## 4. Estados

- **OK** — todo pasa → fila en la hoja de importación.
- **REVISAR** — algo no cuadra → fila en hoja aparte, con el motivo escrito al lado.
- **PROVEEDOR NUEVO** — NIF sin categoría en memoria → el humano asigna columna una vez, se añade la fila a `memoria.csv`, y no se vuelve a preguntar.
- **ERROR** — PDF ilegible o corrupto → se lista al final del run; nunca rompe el lote.

## 5. memoria.csv — semilla del piloto (rellenar CON el departamento)

```
nif_proveedor, proveedor,                              categoria
46630850Y,     Gestoria Olivella (F. Romeu),           ???
B65099673,     Tim Institute SL,                       ???
```

Pregunta de diseño abierta — decidir ANTES de codificar el lookup: ¿la categoría depende solo de QUIÉN factura, o también de QUÉ factura? Tim Institute emite royalties (¿servicios → VARIS?), libros (¿COMPRES?) y formación en el mismo trimestre. Si un mismo NIF cae en columnas distintas según el concepto, la memoria necesita una segunda clave (palabra clave del concepto) o un default por proveedor con excepciones.

## 6. Definición de HECHO de la bala trazadora

El run procesa las 14 de Penedes desde una carpeta, escribe un llibre con las filas repartidas en 1t/2t/3t con valores calculados y estados asignados, y la persona del departamento lo mira y dice: "esto es lo que yo tecleo". Métrica punto cero: X/14 perfectas sin tocar — ese número es la línea base de la métrica de profesionalidad del agente.
