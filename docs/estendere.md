# Modificare l'Omni Report senza romperlo

Questa guida è per te che vuoi **cambiare** qualcosa: aggiungere una colonna, una
sezione nuova, una regola di malpractice, un report intero. Il manuale d'uso è in
`MANUALE.md`; qui si parla di cosa toccare e in quale ordine.

## Il principio, prima di tutto

Il progetto è costruito su una scommessa precisa:

> **Meglio un report che manca di un report sbagliato.**

Un file che non esce lo vedi subito. Un numero sbagliato viene archiviato, e sei
mesi dopo qualcuno prende una decisione su quel numero. Tutto il resto — il
contratto delle colonne, i controlli di coerenza, i test — esiste per servire
quella frase.

Da cui la regola operativa: **ogni modifica deve poter fallire in modo
rumoroso**. Se una tua modifica, andando male, produce un numero plausibile
invece di un errore, quella modifica non è finita.

---

## Le tre zone, in ordine di rischio

Prima di cambiare qualcosa, guarda in quale zona stai entrando.

### Zona verde — configurazione

`config\settings.yml`, `config\contratti.yml`, i fogli `Email Agenti` e
`Helper Malpractice!D:E` del template.

Sono **fatti**, non logica. Cambiali liberamente. Il preflight ti dice subito se
hai scritto una sciocchezza.

### Zona gialla — il contratto e il template

`config\columns.yml`, le formule del workbook, i fogli nuovi.

Qui si sbaglia in silenzio. La regola: **una modifica per volta, e rilancia
`preflight` dopo ognuna.** Non tre modifiche insieme, perché se il risultato
cambia non sai quale ha cambiato cosa.

### Zona rossa — il motore

Il VBA (`CreaMalpractice`), le formule di `Helper Turni` (15.128), il codice in
`src\fasterreports\core\`.

Non si entra senza aver prima **congelato il risultato di ieri come specifica**.
Vedi "Il test golden" più sotto.

---

## Le regole che non si negoziano

Sono cinque, e ognuna viene da un guasto vero.

### 1. Le colonne si cercano per nome, mai per posizione

Il difetto originale del processo manuale: `SF_DATABASE!$BB` nelle formule,
`Cells(r, "DY")` nel VBA. Se l'export inserisce una colonna, tutto scala e i
numeri diventano falsi senza che niente lo segnali.

Se scrivi codice nuovo che legge una fonte, passa dal contratto. Se scrivi una
formula nuova nel workbook, aggiungi la colonna al contratto con il suo
`consumers`.

### 2. Ogni colonna del contratto dichiara chi la legge

```yaml
- canonical: Agent Email
  target_col: B
  role: input
  consumers: ["Report Agenti", "Analisi Status", "VBA:AddATRules"]
```

`consumers` non è documentazione: c'è un test che confronta il contratto con le
formule reali del workbook. Se aggiungi una formula che legge una colonna non
dichiarata, il test lo segnala. Se togli l'ultima formula che leggeva una
colonna, sai che quella colonna non serve più.

### 3. Non usare openpyxl per scrivere

Distrugge le formule ad array dinamico e va in out-of-memory sui file grossi.
Per scrivere si usa **xlwings con Excel vero**; per leggere offline c'è
`core\xlsxsource.py`, che legge l'XML dentro lo zip.

### 4. Non ricalcolare con LibreOffice

Non valuta `_xlfn.XLOOKUP`. Il ricalcolo deve passare da Excel desktop, con
`CalculateFullRebuild()` — non un `calculate()` semplice, perché `INDIRECT`
rende le formule volatili.

### 5. Il VBA non si modifica riscrivendo il file

`xl\vbaProject.bin` contiene il sorgente **e il p-code compilato**. Cambiando
solo il testo si ottiene un file che *sembra* patchato e continua a eseguire il
codice vecchio. Per modificare il VBA: l'editor di Excel, oppure
`omni-report patch-template`, che passa dal modello a oggetti e fa ricompilare
VBA.

E in VBA **le dichiarazioni di modulo stanno tutte prima della prima
procedura**. Metterle dopo produce un modulo che si salva senza un lamento e
viene rifiutato al momento dell'esecuzione. `omni-report check` ora lo verifica
(voce `VBA compilabile`).

---

## Ricette

### Aggiungere una colonna da una fonte che già leggiamo

Esempio: serve `Case Language` da `SF_DATABASE`.

1. **Trova il nome esatto** nell'export. Deve essere quello, carattere per
   carattere.
2. **Scegli la colonna di destinazione** nel foglio `SF_DATABASE` del template.
   Deve essere libera, e non può essere già assegnata a un altro campo (il
   caricamento del contratto lo verifica e blocca).
3. **Aggiungi il campo** in `config\columns.yml`:
   ```yaml
   - canonical: Case Language
     target_col: EN
     role: input
     dtype: str
     consumers: []          # nessuno ancora: la formula la scrivi al passo 5
     notes: "aggiunta 2026-08-04 per il report lingue"
   ```
4. **`omni-report preflight --week NN`.** Il campo deve comparire con
   `via: esatto`. Se compare `alias` o `normalizzato`, il nome che hai scritto
   non è quello vero: correggilo, non lasciarlo così.
5. **Scrivi la formula** che la usa, e aggiungi il consumatore in `consumers`.
6. **`python -m pytest`** — il test di auto-verifica del contratto deve restare
   verde.

Se ti serve un valore che nell'export ha un nome instabile, usa `aliases`; se
esiste un'altra colonna che normalizzando collassa sullo stesso nome, usa
`match: exact` — meglio fallire che agganciare la colonna sbagliata.

### Aggiungere un foglio nuovo che legge i dati esistenti

Questo è il caso **più facile e più sicuro**, e quasi sempre è quello che serve.

1. Aggiungi il foglio nel template, con le sue formule.
2. Le formule leggono i fogli `AT_DATASET`, `SF_DATABASE`, `Turni`… che la
   pipeline riempie già.
3. Aggiorna `consumers` nel contratto per le colonne che il foglio nuovo legge.
4. `omni-report build --week NN`. Il tuo foglio si popola da sé al ricalcolo.

Non serve toccare Python. La pipeline non sa cosa sia `Report Agenti` e non
deve saperlo: scrive i dati, il workbook calcola.

**Attenzione a una cosa sola:** se il foglio nuovo usa formule ad array
dinamico, verifica che ci sia spazio sotto — le celle occupate producono
`#SPILL!`.

### Aggiungere una fonte dati nuova

1. **`config\columns.yml`**: un dataset nuovo, con `sheet`, `header_row`,
   `data_start_col` e i campi.
2. **`reader:`** dice *cosa è* il file, non come è scritto:
   - `csv` / `table` → una tabella (una riga per record). Il formato lo decide
     l'estensione: `.csv` passa da `csvsource`, `.xlsx` da `tablesource`.
   - `wfm_roster` / `wfm_backoffice` → una matrice larga da riportare in forma
     lunga.
   Se la tua fonte è una tabella, **non serve codice nuovo**.
3. **`config\settings.yml`**: lo schema del nome del file in `input_files`.
4. **Il foglio di destinazione** nel template.
5. `preflight`.

Se la fonte ha una forma diversa da queste due (per esempio un JSON), serve un
lettore nuovo in `core\`: deve restituire la stessa coppia `(headers, rows)` di
`read_csv`, e da lì tutto il resto della pipeline funziona senza modifiche.

### Aggiungere una regola di malpractice

Zona rossa: si tocca il VBA. Ordine obbligatorio.

1. **Congela il risultato attuale.** Metti da parte l'ultimo
   `Omni_Report_W<NN>.xlsm` prodotto: è la tua specifica.
2. Scrivi la regola nel modulo `CreaMalpractice`, seguendo il modello di
   `AddSFRules` / `AddATRules`: costruisci l'etichetta **una volta sola** dai
   parametri di `Helper Malpractice`, non ripetere la stringa.
3. La soglia va in `Helper Malpractice!B*`, **non** scritta nel codice: così si
   cambia senza toccare il VBA.
4. Rigenera e confronta:
   ```bash
   omni-report build --week NN
   python tools/golden_report.py "output/Omni_Report_W<NN>.xlsm" "il file congelato.xlsm"
   ```
   Le differenze devono essere **solo** le righe della regola nuova. Se ne vedi
   altre, hai cambiato qualcosa che non volevi cambiare.
5. `omni-report check` — le voci `MsgBox silenziati`, `errore propagato` e
   `VBA compilabile` devono restare `OK`.

### Aggiungere un controllo di coerenza

È il tipo di modifica con il miglior rapporto valore/rischio: non cambia nessun
numero, aggiunge una difesa.

1. Una funzione in `core\coherence.py`, sul modello delle esistenti.
2. Chiamala da `check_sources()`, **dentro il ramo che verifica di avere i dati**:
   un controllo che non ha i dati per essere fatto va saltato, non inventato.
3. Scegli il livello con onestà:
   - **`BLOCCA`** se il report sarebbe *sbagliato*;
   - **`SEGNALA`** se sarebbe *incompleto ma corretto*.

   Sbagliare in un verso lascia passare numeri falsi; sbagliare nell'altro
   produce rumore, e il rumore fa ignorare i controlli. Sono entrambi guasti.
4. Nel messaggio devono starci tre cose: **cosa** non torna, **dove** l'hai
   visto (cella, agente, file), **come** si sistema. Se manca il "come", chi
   legge non sa che fare.
5. Un test per il caso che scatta e uno per il caso che non deve scattare.
   Il secondo conta quanto il primo.

### Cambiare una soglia

`Helper Malpractice!B2:B8` nel template. Cambia i numeri: fallo di proposito e
scrivi da qualche parte quando l'hai fatto, altrimenti il confronto fra
settimane diventa incomprensibile.

---

## Il test golden: come si tocca il motore

Ogni volta che modifichi qualcosa che *calcola*, la procedura è la stessa:

1. **Il risultato di ieri è la specifica.** Metti da parte il workbook prodotto
   prima della modifica.
2. Fai **una** modifica.
3. Rigenera e confronta con `tools/golden_report.py`.
4. **Ogni differenza va spiegata prima di essere accettata.** "Sarà il
   ricalcolo" non è una spiegazione: se non sai perché una cella è cambiata, non
   sai cos'hai cambiato.

Ci sono già due golden test automatici in `tests\`:
`test_golden_wfm.py` ricostruisce `Turni` e `Slot Only Cases` dalle sorgenti e li
confronta cella per cella col workbook fatto a mano; il confronto dei numeri
finali si lancia a mano con `tools/golden_report.py`.

### Una storia che vale la regola

Per settimane un golden test ha certificato una differenza che non esisteva.
Il nostro lettore `.xlsx` non riconosceva le celle vuote autochiudenti
(`<c r="C1" s="120"/>`) e si prendeva il valore della cella successiva
attribuendolo a quella vuota. Non produceva errori: produceva **numeri nella
colonna sbagliata**. Tre affermazioni della documentazione erano false, e sono
state scoperte solo perché un valore era impossibile (`466` come frazione di
giorno).

Morale: **uno strumento di misura va misurato.** Se scrivi un tool di verifica,
scrivi anche i test del tool — `tests\test_xlsxsource.py` esiste solo per
questo.

---

## Prima di committare

```bash
python -m pytest                     # tutti verdi
omni-report check                    # ambiente e template
omni-report preflight --week NN      # sui dati veri di una settimana
```

E, se hai toccato qualcosa che calcola, il confronto golden.

Se un test diventa rosso, **prima capisci perché**. A volte il test ha ragione e
la modifica è sbagliata; a volte la modifica ha ragione e il test descriveva un
comportamento vecchio. In quel caso aggiorna il test **spiegando nel docstring
cosa è cambiato e perché**, come è stato fatto per il caso `NO BOT` qui sopra. Un
test cancellato per far passare la suite è una difesa in meno e nessuno se ne
accorgerà.

Nel messaggio di commit: **cosa hai cambiato e quale problema risolve**. Se il
problema era misurabile, mettici la misura — "3 agenti da 0 a 24/8/40 ore" dice
molto più di "fix skill".

---

## Le cose che vale la pena fare, in ordine

Se cerchi dove mettere le mani, questo è l'ordine per valore:

1. **Un controllo di coerenza nuovo.** Rischio quasi nullo, e ogni controllo
   trova cose che nessuno guardava.
2. **Un foglio di report nuovo** che legge i dati esistenti. Nessun codice, e la
   pipeline non se ne accorge.
3. **`config\contratti.yml`** riempito: l'unica colonna del report che oggi
   resta vuota.
4. **I due report WOW** (AHT Trend, CaseType Deepdive). Il piano è in
   `docs\contesto-wow-aht.md` e riusa lo stesso strato di ingestione: il grosso
   del lavoro è già fatto.
5. **`NormKey` del VBA allineato alle formule.** Oggi il VBA normalizza
   `à è é ì ò ù` e gli manca `á í ó ú`, che le formule invece gestiscono. Un nome
   con quegli accenti farebbe match nelle formule e non nel VBA. Nessun agente
   attuale è in quel caso — ma è una mina, e va disinnescata col golden test in
   piedi.

## Le cose da non fare

- **Non "sistemare" un dato nella pipeline.** Se una fonte è sbagliata, si
  corregge la fonte. Una correzione nascosta nel codice diventa invisibile in tre
  settimane.
- **Non aggiungere un valore di default per far passare un controllo.** Il
  controllo ti sta dicendo qualcosa.
- **Non togliere un `BLOCCA` perché dà fastidio.** Se blocca troppo spesso, il
  problema è nei dati o nel livello scelto: discutilo, non zittirlo.
- **Non toccare `Helper Turni` a mano.** Sono 15.128 formule: è motore, non dati.
- **Non salvare il template come `.xlsx`.** Perdi tutto il VBA, e non c'è modo di
  recuperarlo se non da git.
