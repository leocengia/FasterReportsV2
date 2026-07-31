# Piano di automazione — Omni Report (brief per Claude Code)

> Documento operativo da mettere nel repo e passare a Claude Code.
> Obiettivo: da CSV scaricati → **un pulsante** → report generato, in modo robusto ai cambi/spostamenti di colonna, **senza AI**, con la strada già aperta verso il download via **SQL**.

---

## 0. Regola d'oro e nota sulle fonti

- **Non indirizzare mai le colonne per posizione. Indirizzarle per NOME di intestazione.** Tutto il resto discende da qui.
- Il contratto colonne di questo documento è stato **reverse-engineered dalle formule vive dei fogli e dalla macro VBA**. 
- **`Definizioni KPI` è obsoleto: ignoralo.** Se serve ri-derivare qualcosa, si scansionano le formule (vedi §11 test), non quel foglio.

---

## 1. Vincoli

1. **Un click**: l'utente scarica i 4 CSV in una cartella `input/` con nomi fissi, lancia il "pulsante", ottiene il workbook finito.
2. **Robustezza colonne**: se un CSV cambia ordine colonne o rinomina leggermente un'intestazione, la pipeline deve continuare a funzionare (match per nome + alias + normalizzazione). Se una colonna richiesta manca o è ambigua → **errore parlante, niente numeri sbagliati silenziosi**.
3. **Niente AI**: tutto deterministico.
4. **Excel obbligatorio**: il motore usa array dinamici di Excel 365/2021 (`SORT/UNIQUE/FILTER/ANCHORARRAY/LET/XLOOKUP`) + una macro VBA. Il ricalcolo fedele richiede **Excel vero** (Win/mac) pilotato via `xlwings`. Vedi §10.
5. **Futuro SQL**: sostituire la sorgente CSV con query SQL senza toccare il resto (vedi §13).

---

## 2. Com'è fatto il workbook oggi (mappa rapida)

Pipeline in due stadi:

- **Input** (incollati dai CSV): `AT_DATASET` (~130k righe), `ATwi_DATASET` (~15k), `SF_DATABASE` (143 col.), `PSAT_DATASET` (121 col.).
- **Motore** (Excel, resta invariato): fogli di supporto/output pieni di formule — `Helper Turni`, `Report Agenti`, `Analisi Status`, `Malpractice Recap`, `Dettaglio Malpractice`, `Outbound Exploitation`, `Presentazione Malpractice`, `Recap PSAT Positive`, `PSAT Positive (export x email)` — **più una macro VBA** `Refresh_Dettaglio_Malpractice` (modulo `CreaMalpractice`) che ricostruisce `Dettaglio Malpractice` applicando 6 regole di malpractice.
- **Config** (semi-persistente): `Anagrafica`, `Email Agenti`, `Turni`, `Slot Only Cases`, `Helper Malpractice`.

**Causa della fragilità**: sia le formule (`SF_DATABASE!$BB`, `$DY`, `AT_DATASET!$K`, `PSAT_DATASET!$DP`…) sia il VBA (`Cells(r,"BB")`, `Cells(r,"K")`…) leggono le colonne **per lettera fissa**. Se l'export sposta una colonna, la lettera punta al dato sbagliato → numeri falsi senza errore.

---

## 3. Strategia

**Non riscrivere il motore.** Aggiungere davanti un **layer di ingestione** che:

1. legge ogni CSV **per nome** di colonna (con alias + normalizzazione),
2. valida (fail-loud) che tutte le colonne richieste ci siano e non siano ambigue,
3. scrive i dati **nelle lettere canoniche** che formule e macro si aspettano.

Poi **Excel fa da motore**, pilotato da `xlwings`: apre una copia del template, riceve i dati, ricalcola, lancia la macro, salva. Così le formule 365 e il VBA restano intatti e non vengono mai toccati/corrotti da librerie esterne.

```
input/*.csv (ordine colonne variabile)
      │  [layer ingestione Python: match per NOME + validazione]
      ▼
DATASET (colonne sempre nelle lettere canoniche)
      │  [xlwings: calculate() + Refresh_Dettaglio_Malpractice]
      ▼
output/Omni_Report_W{n}.xlsm (+ eventuale PDF/valori sezioni email)
      ▲
      └── run_report.bat / .command  ← "il pulsante"
```

---

## 4. Contratto colonne (PRIMARIO) — ricavato da formule + VBA

Regole di match per ogni campo: **1)** nome canonico esatto (dopo normalizzazione) → **2)** alias (dopo normalizzazione) → **3)** se più colonne sorgente collassano sullo stesso nome normalizzato **e** non c'è match esatto che disambigua → **ERRORE (ambiguo)**. Mai indovinare in silenzio. Loggare ogni decisione di match.

Normalizzazione nome: `lower` → `strip` → collassa spazi multipli → tratta `_` e spazio come equivalenti → rimuovi/normalizza punteggiatura `()`, `%`, `–`, `/`.

### AT_DATASET
Foglio: intestazioni in riga 1; **colonna A = indice riga**; dati veri da colonna B.

| Lettera canonica | Header atteso | Ruolo | Alias plausibili | Note |
|---|---|---|---|---|
| B | `Agent Email` | input | `agent_email`, `Agent E-mail` | chiave agente (join con Anagrafica / Email Agenti) |
| F | `Agent State` | input | `agent_state`, `State` | filtra Break / Available Cases / Login / Offline |
| I | `Start Time` | input | `start_time`, `Interval Start` | **sorgente di P e Q** |
| K | `Total Time in seconds` | input | `total_time_seconds`, `Total Time (s)` | durata stato, secondi |
| L | `Productive Aux Flag (Yes / No)` | input | `productive_aux_flag`, `Productive Aux` | valori `Yes`/`No` |
| P | `Data Milano` | **derivato** | — | `=INT($I + 'Helper Malpractice'!$B$7/24)` |
| Q | `Ora Milano` | **derivato** | — | frazione del giorno dello stesso calcolo |

Consumato da: `Report Agenti`, `Analisi Status`, macro VBA (che legge B, F, K, P, Q).
Colonne A, C, D, E, G, H, J, M, N, O: non usate dal motore → possono restare vuote/passanti.

**Nota P/Q**: due scelte equivalenti — (a) mantenere P e Q come **formule** nel template e riempirle fino all'ultima riga (semplice, fedele, ma volatile su 130k righe); (b) **precalcolarle in Python** e scrivere valori (più veloce). Default consigliato: (a) nell'MVP, (b) come ottimizzazione. L'offset viene da `Helper Malpractice!B7` (= 9, "Seattle→Milano ore").

### ATwi_DATASET
Foglio: colonna A = indice; dati da B.

| Lettera | Header atteso | Ruolo | Alias | Note |
|---|---|---|---|---|
| C | `Agent Email` | input | `agent_email` | chiave agente |
| H | `Talk Time in seconds` | input | `talk_time_seconds` | |
| I | `Wrap-up Time in seconds` | input | `wrap_up_time_seconds`, `ACW seconds` | |
| M | `Initiation Method` | input | `initiation_method` | filtro valore `OUTBOUND` |

Consumato da: `Outbound Exploitation`. Le **soglie** (talk breve, wrap lungo, wrap>5min) vivono sul foglio in `C46/C47/C48`, **non** sono input.

### SF_DATABASE
Foglio: dati da colonna A.

| Lettera | Header atteso | Ruolo | Alias | Note |
|---|---|---|---|---|
| U | `Case Origin (group)` | input | `case_origin_group` | valori `Phone` / `Other` |
| AC | `case_has_outgoing_email` | input | — | `True`/`False` (usata per % Non-live senza email uscita) |
| AE | `case_number` | input (VBA) | `case_no`, `Case Number` | citata nei dettagli malpractice |
| BB | `Employee Name` | input | `employee_name` | chiave agente |
| DY | `Case AHT (mins)` | input | `case_aht_mins`, `AHT (min)` | ⚠ non confondere con altre colonne AHT presenti |
| EC | `Co-Browse Usage %` | input | `cobrowse_usage_pct` | |
| EI | `Misrouted Cases` | input | `misrouted_cases` | |

⚠ **Duplicati da evitare**: l'export SF ha molte colonne quasi omonime (varianti di `origin`, `AHT`, `employee/agent`). Perciò match **esatto sul nome canonico prima**, alias solo come fallback, ed **errore se ambiguo**.

### PSAT_DATASET
Foglio: dati da colonna A.

| Lettera | Header atteso | Ruolo | Alias | Note |
|---|---|---|---|---|
| I | `Agent Name` | input | — | ⚠ esiste anche `agent_name` (col. BO): usare **`Agent Name`** |
| DO | `Case Number` | input | `case_number` | |
| DP | `psat_score` | input | — | ⚠ esiste anche `PSAT Score` (col. EK): usare **`psat_score`** (valori 0/1) |
| DQ | `response_text` | input | `comment`, `verbatim` | |

**Nota "Elogio della settimana"**: nel `Recap PSAT Positive` alcune celle puntano a una **riga fissa 130** (`PSAT_DATASET!DQ130/I130/DO130`) → è una **selezione manuale** del commento migliore, fuori dal contratto. Vedi domanda aperta §14.

---

## 5. Contratto colonne (SECONDARIO — fogli config)

Questi fogli sono letti **per lettera** da formule/VBA. Oggi sono mantenuti a mano o incollati da export. Se vengono rigenerati da export, sono anch'essi a rischio "colonne che si spostano" → in tal caso includerli nel layer di ingestione con lo stesso meccanismo. Altrimenti l'ingestione li lascia stare.

- **Turni**: `A` Nome agente · `C` Contratto · `D` Ore/gg · `E` Data · `F` Stato (`LAVORA`/`FERIE-OFF`/…) · `G` Inizio turno.
- **Slot Only Cases**: `A` chiave (nome BO normalizzato) · `B` Data · `C` Slot inizio · `D` Slot fine · `E` Stato BO (`BOT`/`NO BOT`).
- **Anagrafica**: `A` Nome · `B` Email (A2 è un array dinamico che pesca i nomi da `SF_DATABASE` Employee Name).
- **Email Agenti**: `A` Nome · `B` Email (mappa persistente nome↔email).
- **Helper Malpractice**: parametri in `B2..B8` (limite break simultanei, break max/gg, percentile AHT, soglia caso rapido, tolleranza Available Cases, **offset fuso B7**, ora inizio default) · tabella alias nomi in `D:E`.

---

## 6. Architettura del codice (repo)

```
omni-report/
  pyproject.toml
  README.md
  config/
    columns.yml         # IL CONTRATTO (unica fonte di verità della mappatura) — vedi §8
    settings.yml        # percorsi, offset, nomi file, sorgente csv|sql
  template/
    Omni_Report_TEMPLATE.xlsm   # workbook con formule+VBA+config, DATASET vuoti
  input/                # inbox CSV: AT.csv, ATwi.csv, SF.csv, PSAT.csv
  output/               # Omni_Report_W{n}.xlsm (+ eventuale PDF)
  src/omni_report/
    __init__.py
    config.py           # carica/valida columns.yml + settings.yml
    csvsource.py        # legge un CSV -> DataFrame (sniff encoding/separatore, header)
    matcher.py          # header canonico/alias -> colonna sorgente; normalizza; errori mancanza/ambiguità
    transform.py        # costruisce il DataFrame canonico per dataset (rinomina, ordina, P/Q opz.)
    writer.py           # xlwings: scrive i 4 dataset nel workbook aperto
    orchestrate.py      # copia template -> apri Excel -> scrivi -> calcola -> macro -> salva -> export
    preflight.py        # report leggibile di validazione (trovato/mappato/mancante/ambiguo)
    cli.py              # es. `omni-report build --week 29`
  tests/
    test_matcher.py
    test_transform.py
    test_golden.py      # confronto con baseline W29
    fixtures/           # csv di esempio + varianti "shuffled" e "renamed"
    baseline/W29_baseline.xlsm
  run_report.bat        # "pulsante" Windows  -> python -m omni_report.cli build
  run_report.command    # "pulsante" macOS
```

---

## 7. Spec dei moduli

**config.py** — carica `columns.yml` e `settings.yml`, valida lo schema (ogni dataset: `sheet`, `header_row`, `data_start_col`, lista campi con `canonical`, `target_col`, `role` ∈ {input, derived}, `aliases`, `dtype`). Espone oggetti tipizzati.

**csvsource.py** — `read_csv(path) -> DataFrame`: rileva encoding (fallback `utf-8`/`cp1252`), separatore (`,`/`;`/`\t`), legge la sola riga header + dati; nessuna assunzione sull'ordine. Restituisce anche la lista header originale per il preflight.

**matcher.py** — cuore anti-fragilità. `resolve(canonical, aliases, source_headers) -> source_col | raise`:
- normalizza (vedi §4) header sorgente e candidati;
- priorità: match esatto canonico → alias → altrimenti errore;
- se più colonne sorgente matchano lo stesso target senza un esatto disambiguante → `AmbiguousColumnError` con l'elenco dei candidati;
- se nessun match → `MissingColumnError` con nome atteso + alias provati;
- funzione `resolve_all(dataset)` che ritorna la mappa completa e un log strutturato delle decisioni.

**transform.py** — `build_canonical(df, dataset, mapping) -> DataFrame`:
- seleziona/rinomina le colonne input alle lettere canoniche (indicizzate per `target_col`);
- coercizioni `dtype` (numeri per K/H/I/DY/EC/EI…, datetime per `Start Time`);
- per AT: se scelta la modalità (b), calcola `Data Milano`/`Ora Milano` da `Start Time` + offset; se modalità (a), non calcola (le formule del template lo faranno);
- restituisce un blocco 2D allineato alle colonne target (celle non usate = None).

**writer.py** — via `xlwings`:
- riceve il `book` già aperto + i DataFrame canonici;
- per ogni dataset: pulisce i dati sotto l'header (`sht.range((header_row+1,1)).expand('table').clear_contents()`), scrive i valori in blocco (`sht.range((header_row+1, first_col)).value = values`);
- **non** riscrive le formule del template (P/Q, ecc.); se modalità (a), estende P/Q fino all'ultima riga dati;
- non usa mai openpyxl per scrivere il file finale (vedi §10, rischio corruzione array dinamici).

**orchestrate.py** — flusso end-to-end:
1. copia `template/Omni_Report_TEMPLATE.xlsm` → `output/Omni_Report_W{n}.xlsm`;
2. `app = xw.App(visible=False)`; apri la copia (macro abilitate);
3. `preflight` sui 4 CSV → se fallisce, chiudi e esci con codice ≠ 0;
4. `writer` scrive i 4 dataset;
5. `book.app.calculate()`;
6. esegui la macro: `book.macro('Refresh_Dettaglio_Malpractice')()` (in **modalità silenziosa**, vedi §12);
7. `book.save()`; opzionale: `ExportAsFixedFormat` PDF dei fogli email-ready;
8. `finally`: `book.close()`, `app.quit()`.

**preflight.py** — per ogni dataset produce una tabella leggibile `campo canonico | colonna sorgente | via (esatto/alias) | stato`, verifica dataset non vuoto e coercibilità tipi, scrive `output/preflight_W{n}.txt`, e **solleva/torna codice ≠ 0** al primo problema bloccante.

**cli.py** — `omni-report build --week N [--input DIR] [--source csv|sql]`. Log su stdout + file.

---

## 8. `config/columns.yml` (bozza pronta, da rivedere)

```yaml
# Contratto colonne — SORGENTE DI VERITÀ della mappatura.
# Derivato da formule vive + VBA. NON allineare a "Definizioni KPI" (obsoleto).
normalize:
  lower: true
  strip: true
  collapse_spaces: true
  underscore_eq_space: true
  strip_punct: ["(", ")", "%", "–", "/"]

datasets:
  AT_DATASET:
    sheet: AT_DATASET
    header_row: 1
    data_start_col: B          # A = indice riga
    fields:
      - {canonical: "Agent Email",                    target_col: B, role: input,   dtype: str,      aliases: ["agent_email", "Agent E-mail"]}
      - {canonical: "Agent State",                    target_col: F, role: input,   dtype: str,      aliases: ["agent_state", "State"]}
      - {canonical: "Start Time",                     target_col: I, role: input,   dtype: datetime, aliases: ["start_time", "Interval Start"]}
      - {canonical: "Total Time in seconds",          target_col: K, role: input,   dtype: float,    aliases: ["total_time_seconds", "Total Time (s)"]}
      - {canonical: "Productive Aux Flag (Yes / No)", target_col: L, role: input,   dtype: str,      aliases: ["productive_aux_flag", "Productive Aux"]}
      - {canonical: "Data Milano",                    target_col: P, role: derived, source: "Start Time"}   # INT(Start + offset/24)
      - {canonical: "Ora Milano",                     target_col: Q, role: derived, source: "Start Time"}   # frazione
    derive_offset_from: "'Helper Malpractice'!B7"

  ATwi_DATASET:
    sheet: ATwi_DATASET
    header_row: 1
    data_start_col: B
    fields:
      - {canonical: "Agent Email",            target_col: C, role: input, dtype: str,   aliases: ["agent_email"]}
      - {canonical: "Talk Time in seconds",   target_col: H, role: input, dtype: float, aliases: ["talk_time_seconds"]}
      - {canonical: "Wrap-up Time in seconds",target_col: I, role: input, dtype: float, aliases: ["wrap_up_time_seconds", "ACW seconds"]}
      - {canonical: "Initiation Method",      target_col: M, role: input, dtype: str,   aliases: ["initiation_method"]}

  SF_DATABASE:
    sheet: SF_DATABASE
    header_row: 1
    data_start_col: A
    fields:
      - {canonical: "Case Origin (group)",      target_col: U,  role: input, dtype: str,   aliases: ["case_origin_group"]}
      - {canonical: "case_has_outgoing_email",  target_col: AC, role: input, dtype: str,   aliases: []}
      - {canonical: "case_number",              target_col: AE, role: input, dtype: str,   aliases: ["Case Number", "case_no"]}
      - {canonical: "Employee Name",            target_col: BB, role: input, dtype: str,   aliases: ["employee_name"]}
      - {canonical: "Case AHT (mins)",          target_col: DY, role: input, dtype: float, aliases: ["case_aht_mins"]}   # NON confondere con altre AHT
      - {canonical: "Co-Browse Usage %",        target_col: EC, role: input, dtype: float, aliases: ["cobrowse_usage_pct"]}
      - {canonical: "Misrouted Cases",          target_col: EI, role: input, dtype: float, aliases: ["misrouted_cases"]}

  PSAT_DATASET:
    sheet: PSAT_DATASET
    header_row: 1
    data_start_col: A
    fields:
      - {canonical: "Agent Name",   target_col: I,  role: input, dtype: str,   aliases: []}   # NON "agent_name" (BO)
      - {canonical: "Case Number",  target_col: DO, role: input, dtype: str,   aliases: ["case_number"]}
      - {canonical: "psat_score",   target_col: DP, role: input, dtype: float, aliases: []}   # NON "PSAT Score" (EK)
      - {canonical: "response_text",target_col: DQ, role: input, dtype: str,   aliases: ["comment", "verbatim"]}
```

---

## 9. Validazione & fail-loud (comportamento richiesto)

- Header mancante → messaggio: *"Dataset X: colonna richiesta 'Case AHT (mins)' non trovata. Alias provati: [...]. Header presenti: [...]"* → **stop, exit≠0**.
- Header ambiguo → elenco candidati e richiesta di aggiungere un alias esatto → stop.
- Dataset vuoto o header row non riconoscibile → stop.
- Tipi non coercibili (es. AHT non numerico) → warning aggregato + stop se supera soglia.
- Report di preflight sempre scritto su file, anche in caso di successo (tracciabilità delle mappature scelte).

---

## 10. Prerequisiti / ambiente (leggere prima di scegliere le librerie)

- **Serve Excel desktop** (Windows o macOS) sulla macchina dove gira il pulsante. Motivo: gli array dinamici (`ANCHORARRAY/SORT/UNIQUE/FILTER/LET/XLOOKUP`) e la macro VBA vanno valutati da Excel vero.
- **Non usare openpyxl per SCRIVERE il file finale**: mangia/riscrive le formule ad array dinamico e perde i valori in cache → corrompe il motore. openpyxl solo per **letture offline** (estrazione contratto, baseline nei test con `data_only=True`).
- **Non usare LibreOffice headless per il ricalcolo**: non valuta `XLOOKUP/FILTER/UNIQUE/SORT` → risultati vuoti/errati.
- Quindi: **scrittura + ricalcolo + macro via `xlwings` (Excel)**.
- Dipendenze Python: `xlwings`, `pandas`, `pyyaml`, `openpyxl` (solo test), `chardet`/`charset-normalizer` (encoding).
- **Il "pulsante"**: `run_report.bat` (Win) / `run_report.command` (mac) che invoca `python -m omni_report.cli build`. In alternativa un pulsante dentro Excel con `xlwings` RunPython, ma il .bat è più semplice.

---

## 11. Testing

- **Baseline golden**: conservare l'attuale `Omni_Report_W29.xlsm` come `tests/baseline/W29_baseline.xlsm`. Estrarre con openpyxl (`data_only=True`) i valori chiave: `Report Agenti` righe agenti col A:V, totali `Malpractice Recap`, `Outbound Exploitation`, conteggi `Recap PSAT Positive`. Ricostruire dai CSV della stessa settimana e **confrontare entro tolleranza** (float).
- **Resilienza colonne**: partire da un fixture CSV e (a) mescolare l'ordine colonne, (b) rinominare un paio di colonne in loro alias → l'output deve restare **identico**. (c) rimuovere una colonna richiesta → il preflight deve **fallire con messaggio chiaro**.
- **Unit**: `matcher` (normalizzazione, priorità esatto>alias, errore ambiguità), `transform` (coercizioni, P/Q).
- **Auto-verifica del contratto**: uno script di test che ri-scansiona le formule del template e verifica che ogni colonna dataset consumata sia coperta da `columns.yml` (evita derive quando il workbook evolve). *(È così che si "ri-deriva" il contratto — non da `Definizioni KPI`.)*

---

## 12. Modifica minima al VBA (modalità silenziosa)

La macro oggi mostra `MsgBox` a fine esecuzione (bloccante in automazione). Nel template aggiungere un flag pubblico, es. `Public SilentMode As Boolean`, e mettere ogni `MsgBox` dietro `If Not SilentMode Then ...`. Prima di lanciarla da `xlwings`, impostare `SilentMode = True` (o passare via cella parametro). Nessun'altra modifica al VBA nell'MVP.

---

## 13. Futuro: sorgente SQL

- Aggiungere `sqlsource.py` (SQLAlchemy) che produce **lo stesso DataFrame "largo" per dataset** delle CSV. Una query per dataset che ritorni **almeno** i campi richiesti (colonne extra ignorate).
- Il contratto `columns.yml` diventa la mappa **alias SQL → nome canonico** (basta che le `SELECT ... AS` usino i nomi canonici, o aggiungere gli alias SQL alla lista `aliases`).
- `matcher`/`transform`/`writer`/`orchestrate` **non cambiano**. `settings.yml` sceglie `source: csv|sql`.

---

## 14. Milestone (ordine consigliato per Claude Code)

- **M0 — Setup**: ricavare `Omni_Report_TEMPLATE.xlsm` (copia dell'attuale con i 4 fogli DATASET svuotati, formule/VBA/config intatti), aggiungere `SilentMode` al VBA, scaffolding repo, scrivere `columns.yml` da §8, fissare la baseline.
- **M1 — Ingestione**: `csvsource` + `matcher` + `preflight` + unit test (incl. resilienza shuffle/rename/missing).
- **M2 — Scrittura & orchestrazione**: `writer` + `orchestrate` via xlwings; build end-to-end che produce un workbook; smoke test.
- **M3 — Golden**: confronto con W29 entro tolleranza; test contratto auto-verificante.
- **M4 (opz.) — Output**: export PDF/valori delle sezioni email (`Recap PSAT Positive`, `PSAT Positive (export x email)`, `Presentazione Malpractice`); ergonomia CLI; logging.
- **M5 (futuro) — SQL**: `sqlsource` dietro lo stesso contratto; toggle in `settings.yml`.
- **M6 (opz., profondo) — Elimina la fragilità dentro Excel**: convertire i 4 range DATASET in **tabelle con nome** + riferimenti strutturati nelle formule + VBA che trova le colonne per header (`ListColumns("...").Index`). Grosso lavoro di riscrittura, separato dall'MVP; a quel punto le lettere fisse spariscono del tutto.

---

## 15. Domande aperte (da decidere con l'utente)

1. **Template**: OK ricavare l'`Omni_Report_TEMPLATE.xlsm` svuotando i DATASET dell'attuale W29 (posso prepararlo), oppure ne fornisci uno pulito?
2. **Output**: solo `.xlsm`, o anche PDF/valori delle sezioni pronte-email?
3. **"Elogio della settimana"** (riga 130 PSAT): resta selezione manuale, o vuoi una regola deterministica (es. commento positivo più lungo)?
4. **Turni / Slot Only Cases**: sono incollati da export (quindi anch'essi a rischio colonne → li includo nel contratto) o li mantieni a mano?
5. **Ambiente**: confermi che dove girerà il pulsante c'è **Excel desktop** (Win/mac)? È il requisito su cui poggia tutto il ricalcolo.
