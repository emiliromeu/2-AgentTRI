# Agent TRIMESTRE — app

## Arrancar

```
source .venv/bin/activate
streamlit run app.py
```

Obre una pestanya al navegador (normalment `http://localhost:8501`).

## Arrencar amb doble clic (Mac)

Doble clic a `Iniciar_Agent.command` (a l'arrel del projecte) o a
l'àlies "Agent TRIMESTRE" de l'Escriptori — obre una finestra de
Terminal i arrenca l'app sola, sense escriure res.

**La primera vegada**, macOS pot bloquejar-ho per venir "d'un
desenvolupador no identificat". Si passa: clic dret (o Control+clic)
sobre l'arxiu → "Obrir" → confirmar al diàleg. A partir d'aquí, ja
s'obre amb doble clic normal com qualsevol altra aplicació.

Si no troba l'entorn virtual (`.venv`), el llançador ho diu clarament
en comptes de fallar en silenci — cal crear-lo primer (secció
"Arrancar" més amunt, amb `pip install -r requirements.txt`).

## Desplegament a Windows (pendent)

Ja existeix `Iniciar_Agent_Windows.bat` a l'arrel, com a peça del kit
de desplegament futur -- **encara no s'ha provat en un PC Windows
real** (no n'hi ha cap disponible ara mateix). Quan hi hagi una
màquina real, abans de fer-lo servir cal:

- Instal·lar Python al PC.
- Crear l'entorn virtual (`python -m venv .venv`) i instal·lar
  `requirements.txt` (`pip install -r requirements.txt`).
- Crear un `.env` amb una clau API pròpia d'aquella màquina — mai
  copiar la clau del Mac.
- Adaptar `boton_obrir()` d'`app.py`: fa servir la comanda `open`,
  exclusiva de macOS. A Windows caldria `os.startfile()` o `start`.
  S'adapta i es prova quan hi hagi el PC real, no abans.

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
