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

## Instal·lació a Windows

Piso 13: l'app ja és multiplataforma de codi (`boton_obrir()` detecta
`sys.platform` i fa servir `os.startfile()` a Windows / `open` a
macOS; tots els arxius de text que llegeix o escriu el pipeline
porten `encoding="utf-8"` explícit perquè els accents i la ç no es
corrompin amb el codi de pàgina per defecte de Windows). El que
**encara no s'ha fet és una prova en un PC Windows real** — vegeu
`CHECKLIST_PILOT_WINDOWS.md` per a la prova pilot pas a pas abans de
desplegar-ho als 6 PCs de la gestoria.

Passos per instal·lar en un PC Windows nou:

1. **Instal·lar Python** des de https://python.org si el PC no en té
   — durant la instal·lació, marca la casella **"Add Python to
   PATH"**.
2. **Copiar la carpeta sencera** del projecte a aquest PC (per
   exemple des d'un llapis USB o OneDrive).
3. **Doble clic a `instalar_windows.bat`** — crea l'entorn virtual
   `.venv`, instal·la `requirements.txt`, i si no existeix `.env`
   en crea una plantilla amb `ANTHROPIC_API_KEY=`. Qualsevol error
   (Python no trobat, instal·lació de dependències fallida...) es
   mostra clar a la pantalla i la finestra NO es tanca sola — sempre
   hi ha un `pause` al final, tant si acaba bé com si acaba amb
   error.
4. **Editar `.env`** amb el Bloc de notes i enganxar la clau API
   **d'aquest PC concret** — mai copiar la clau d'un altre
   ordinador (regla 6/7 de `CLAUDE.md`: la clau i les dades de
   clients mai es comparteixen entre màquines).
5. **Doble clic a `Iniciar_Agent_Windows.bat`** per arrencar l'app —
   si `.venv` no existeix encara (pas 3 no fet), ho diu clarament en
   comptes de fallar en silenci.

**Dades locals per PC**: cada ordinador té la seva pròpia carpeta
`clientes/` — no hi ha res compartit entre PCs. Els 6 PCs de la
gestoria funcionen com a instàncies independents, cadascuna amb els
seus propis clients, decisions i informes.

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
