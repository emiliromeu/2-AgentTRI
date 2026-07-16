# Agente TRIMESTRE — Gestoría Olivella

Pipeline que lee facturas rebudes en PDF, extrae campos con la API de
Anthropic, valida en código determinista y escribe el llibre Excel que
se importa a Geyce. El contrato de diseño completo vive en
docs/esquema_canonico_v1.md — leerlo antes de tocar código.

## Entorno
- Python 3.12 en .venv → `source .venv/bin/activate` antes de python/pip
- Librerías: anthropic, openpyxl, python-dotenv
- Clave API solo en .env, cargada con load_dotenv() + os.environ.get()

## Reglas innegociables
1. Emili debe entender TODO el código: cambios pequeños y explicados,
   nunca bloques grandes sin explicación. Concepto nuevo = explicarlo.
2. El LLM solo EXTRAE (extraer_factura(pdf) -> JSON). Clasificar,
   validar y escribir es código determinista.
3. El detalle va siempre en VALORES calculados en Python; los totales
   pueden ser fórmulas =SUMA() escritas por código sobre rangos que el
   código controla.
4. Fallos ruidosos: cada factura en su try/except con estado ERROR;
   el lote nunca muere y nada falla en silencio.
5. Run idempotente: ejecutarlo dos veces jamás duplica filas; los PDF
   procesados se mueven a procesadas/.
6. NUNCA imprimir, loguear ni leer el contenido de .env ni la API key.
7. clientes/ contiene datos fiscales reales: jamás a git, jamás en
   ejemplos, jamás en logs.
8. Nada está "hecho" sin probarlo contra las facturas reales de
   clientes/penedes_languages/rebudes/entrada.
9. Nada muere en silencio aplica también al enrutado — todo documento
   apartado o descartado debe ser visible en el Excel con su
   justificación.
10. La interfaz también es ruidosa: jamás botones deshabilitados sin
    explicación visible; los requisitos previos (firma, notas) se
    piden de forma bloqueante antes de mostrar acciones, o el clic
    produce un error visible. Nunca depender de que un text_input de
    Streamlit esté confirmado con Tab/Enter.
11. Ningún piso toca sumar.py sin pasar antes
    `python3 verificar_conservacio.py` contra los 4 clientes reales
    (0 divergencias) — bateria independiente (Piso 13Y) que recalcula
    los cubos por tipo de IVA, Total, Líquid y RESULTAT desde
    validadas/*.json + decisions.csv sin importar sumar.py, y los
    compara célula a célula contra el Excel generado.
12. Ningún piso toca MUTACIONES (mover de cliente/flujo, retirar,
    reactivar, corregir, decidir, deshacer, entrada manual) sin pasar
    `python3 auditar_coherencia.py` contra los 4 clientes reales
    (0 divergencias) — auditor independiente (Piso 14E) que verifica
    las CUATRO capas de cada documento: física (el archivo donde su
    estado manda), ficha (activa ↔ original, o manual legítima),
    libros (hashos/decisions/moviments consistentes) y vistas (lo
    vivo en el Excel, lo muerto fuera).
