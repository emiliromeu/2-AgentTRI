# Checklist — Prova pilot en Windows real (Piso 13)

Aquesta app no ha corregut mai en un PC Windows real. Tot el que hi
ha a `app.py`/`sumar.py`/`informe.py`/`validar.py`/`trocear.py`/
`extraer_todas.py` s'ha revisat i provat a fons en Mac (regressió
completa, totals al cèntim, 0 enllaços trencats), i s'ha auditat
línia a línia perquè els accents no es corrompin i els botons
d'obrir arxius funcionin a Windows -- però això és feina de codi, no
una prova real. **Aquesta checklist és la prova real**, pensada
perquè l'executis tu al primer PC pilot abans de fer-ho als altres 5.

Marca cada casella a mesura que la completis. Si alguna falla, para
aquí i digues-m'ho amb el missatge d'error exacte (captura de
pantalla si pot ser) -- no continuïs amb els altres PCs fins que
estigui arreglat.

## 1. Instal·lació

- [ ] Copia la carpeta sencera del projecte al PC pilot.
- [ ] Doble clic a `instalar_windows.bat`.
- [ ] Si no tenies Python instal·lat, l'script t'ho diu clarament
      (no es tanca en silenci) -- instal·la'l des de python.org
      marcant "Add Python to PATH" i torna a executar
      `instalar_windows.bat`.
- [ ] L'script acaba amb "Instal·lació completada correctament."
      sense cap ERROR pel mig.
- [ ] S'ha creat un arxiu `.env` nou (si no n'hi havia) amb una
      línia `ANTHROPIC_API_KEY=`.
- [ ] Obre `.env` amb el Bloc de notes i enganxa la clau API
      **d'aquest PC** (mai la clau d'un altre ordinador) després del
      signe `=`. Desa i tanca.

## 2. Arrencada

- [ ] Doble clic a `Iniciar_Agent_Windows.bat`.
- [ ] S'obre una finestra de consola i, al cap de pocs segons, el
      navegador amb l'app ("Agent TRIMESTRE").
- [ ] Si no existís `.venv` (per exemple perquè et vas saltar el pas
      1), l'script ho hauria de dir clarament ("Executa primer
      instalar_windows.bat") en comptes de fallar sense explicació --
      pots provar-ho expressament renombrant `.venv` un moment per
      confirmar-ho, i tornar-lo a renombrar després.

## 3. Client de prova (amb accents, a propòsit)

- [ ] Pestanya "Clients" → "Nou client" → nom amb accent, per exemple
      **"Comprovació SL"**, NIF qualsevol vàlid.
- [ ] El nom apareix bé a la llista de clients (sense símbols
      estranys en comptes de la ó/ç).

## 4. Pujar factures

- [ ] Pestanya "Afegir factures" → selecciona el client de prova.
- [ ] Puja un PDF real d'una factura.
- [ ] Puja una foto en format **HEIC** (típica d'iPhone) -- confirma
      que a la carpeta `clientes/<carpeta>/rebudes/entrada/` apareix
      convertida a `.jpg`, no com `.heic`.

## 5. Processar (crida real a l'API)

- [ ] Pestanya "Processar" → botó "Processar".
- [ ] Acaba amb "Procés acabat." sense error.
- [ ] Al resum per client, el client de prova mostra les factures
      pujades amb el seu estat (OK o REVISAR).

## 6. Els botons que abans no funcionaven a Windows (el bug arreglat)

- [ ] A la targeta del client de prova (pestanya "Clients" o al
      resum de "Processar"), prem **"Obrir informe"** -- s'ha
      d'obrir l'`informe_2026.html` al navegador per defecte.
- [ ] Prem **"Obrir Excel"** -- s'ha d'obrir `sumatorios_2026.xlsx`
      a Excel (o al programa que tinguis associat als `.xlsx`).
- [ ] Als dos arxius oberts, els accents i la ç es veuen bé (no
      símbols estranys tipus `Ã§` o `Ã³`).

## 7. Revisió

- [ ] Pestanya "Revisió" → escriu un nom a "Qui revisa?" → "Entrar a
      revisar".
- [ ] Selecciona el client de prova.
- [ ] Aprova una fitxa pendent (amb o sense nota).
- [ ] La targeta mostra "✓ Aprovada per..." sense recarregar tota la
      pantalla ni perdre el lloc on eres.

## 8. Neteja

- [ ] Pestanya "Clients" → targeta del client de prova → "Arxivar
      client" → escriu el nom exacte per confirmar → "Arxivar
      definitivament".
- [ ] El client de prova desapareix de la llista (queda recuperable
      a `arxivats/`, mai esborrat).

## Si alguna cosa falla

Anota exactament:
1. En quin pas exacte de la llista ha fallat.
2. El missatge d'error tal com surt a la pantalla (captura si pot
   ser).
3. Si la finestra de la consola s'ha tancat sola o ha quedat oberta
   amb l'error visible.

I envia-m'ho abans de provar-ho als altres 5 PCs.
