# Transfer Data Submission — agosto 2026 (MBR settembre)

Materiale per compilare `August_Transfer_Data_Template_v2.xlsx`, richiesto dall'upper
management entro il **1 settembre 2026**.

**Questa cartella non tocca il programma.** Non modifica `src/`, `config/`, `template/`:
serve a verificare che i transfer si possano misurare con l'export che l'Omni Report gia'
scarica. Se la verifica regge, in un secondo momento si valutera' se portare il calcolo
dentro l'Omni Report come foglio nuovo — e in quel caso servira' aggiungere al contratto in
`config/columns.yml` le 12 colonne elencate sotto.

## La sorgente giusta e' il CSV, non il workbook generato

L'Omni Report **generato dalla pipeline non va bene**: il foglio `SF_DATABASE` ha tutte e 143
le intestazioni, ma la pipeline scrive solo gli 11 campi dichiarati nel contratto
(`config/columns.yml:154-292`) e lascia vuote le altre 132 — comprese tutte quelle che
servono qui: `Case Origin` (T), `Case Record Type` (W), `Case Resolution Category` (X),
`case_status` (AH), `Channel Group` (AJ), `Date (Range)` (AW), `Has Child Cases` (BE),
`Manager Name` (BL), `Omni Level Indicator` (BN), `parentid` (BQ), `Secondary Category` (CH),
`work_function` (DE).

Va usato l'export **`SF DATABASE*.csv`**, quello che si mette in `input/` prima di lanciare
l'Omni Report. Lo script se ne accorge da solo e si ferma con un errore esplicito se gli si
passa un workbook generato. (`samples/omni-report/Omni Report W30.xlsm` funziona lo stesso
perche' e' un workbook legacy, riempito incollando il CSV intero.)

## Come si usa

```
python analisi/transfer-mbr-agosto/estrai_transfer.py "input/SF DATABASE W34.csv" \
    --out analisi/transfer-mbr-agosto --etichetta W34
```

Accetta piu' file in un colpo solo (per sommare le settimane di agosto) e in quel caso
deduplica per `case_number`:

```
python analisi/transfer-mbr-agosto/estrai_transfer.py input/SF*.csv --out . --etichetta agosto2026
```

Aggancia le colonne **per nome**, mai per lettera o posizione: l'export SF e' gia' cresciuto
in passato e le colonne si spostano. Se manca una colonna obbligatoria si ferma dicendo
quale. Legge senza openpyxl e senza Excel, come `tools/audit_workbook.py`.

Produce:
- `breakdown_<etichetta>.md` — la tabella della sezione 1, l'annex della gamba cedente, la
  popolazione da scrubbare e il controllo per case type;
- `transfer_casi_<etichetta>.csv` — l'export **case-level** dei transfer, ordinato col
  campione in cima, con `nel_campione`, `valid_invalid` e `note_scrub` da compilare a mano.
  Separatore `;` e virgola decimale: si apre in Excel italiano senza conversioni.
  Il campione e' deterministico: rilanciare lo script sullo stesso periodo ridà lo stesso
  campione, così lo scrub gia' fatto si riaggancia.

Poi si compila il template:

```
python analisi/transfer-mbr-agosto/compila_template.py \
    --advanced-voice 18 --advanced-bof 0 --basic-voice 112 --basic-bof 9 \
    --periodo "..." --out .../August_Transfer_Data_Template_v2_COMPILATO.xlsx
```

Riscrive il foglio dentro una **copia** dello zip, lasciando intatti stili e formule.

## I numeri — W34 (17–23 agosto 2026)

2762 casi, **139 transfer = 5,03%**.

| Queue type | Voice | BOF | TOTALE |
|---|---:|---:|---:|
| Advanced | 18 | 0 | 18 |
| Basic | 112 | 9 | 121 |
| **TOTALE** | **130** | **9** | **139** |

Gamba cedente (annex, non riconcilia): `Closed - Transferred` 72 · `Case Transfer` 73 ·
`Call Transfer` 33, tutti su canale Phone.

Da scrubbare: 44 casi ad alto rischio (32 bounce, 11 che generano figli, 5 con AHT sotto il
10° percentile; **zero misrouted** questa settimana). Campione dimensionato a 103 su 139.

## Stato

- [x] Fattibilita' verificata e sorgente corretta individuata (il CSV, non il workbook).
- [x] Script scritto e validato: riproduce i numeri contati a mano su W30 e su W20+W21.
- [x] Numeri di W34 estratti.
- [x] Template compilato per la sezione 1 e per `Total transfers` della sezione 2.
- [x] Rilievi sul template documentati in `rilievi_template.md` (30 punti).
- [ ] **Scrub manuale**: marcare `valid_invalid` in `transfer_casi_W34.csv`.
- [ ] **Le altre settimane**: W34 copre 7 giorni su 31. Servono gli export di W31–W33 e
      W35–W36, oppure si dichiara W34 come settimana campione.
- [ ] Compilare le righe 17–19 del template dopo lo scrub e consegnare.

## File

| File | Cosa |
|---|---|
| `estrai_transfer.py` | estrae i transfer da CSV o workbook |
| `compila_template.py` | compila una copia del template coi numeri |
| `breakdown_W34.md` | i numeri di W34 |
| `transfer_casi_W34.csv` | i 139 transfer da scrubbare |
| `breakdown_W30_validazione.md`, `transfer_casi_W30_validazione.csv` | la validazione su W30 |
| `rilievi_template.md` | i 30 rilievi sul template ricevuto, con la decisione presa su ciascuno |
| `metodologia.md` | la nota metodologica da allegare alla submission |
| `August_Transfer_Data_Template_v2.xlsx` | il template ricevuto, intatto |
| `August_Transfer_Data_Template_v2_COMPILATO.xlsx` | la copia compilata |
