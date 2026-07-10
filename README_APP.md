# Agent TRIMESTRE — app

## Arrancar

```
source .venv/bin/activate
streamlit run app.py
```

Obre una pestanya al navegador (normalment `http://localhost:8501`).

## Què fa i què no fa

L'app és només una closca: col·loca els arxius que puges a la carpeta
que ja mira `extraer_todas.py`/`trocear.py`, i llança `ejecutar.py` com
a subprocés — exactament el mateix camí que si l'executessis a mà.
No calcula ni valida res: la lògica de negoci viu només a les cinc
màquines (`trocear.py`, `extraer_todas.py`, `validar.py`, `sumar.py`,
`informe.py`).

## `python ejecutar.py` directe segueix disponible

L'app no substitueix res — el botó "Processar" fa exactament
`python ejecutar.py` per sota. Si es prefereix la terminal, segueix
funcionant igual que sempre:

```
source .venv/bin/activate
python ejecutar.py
```
