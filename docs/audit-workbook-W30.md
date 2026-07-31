# Audit di `Omni Report W30.xlsm` — 31/07/2026

Verifica del contratto colonne del piano contro il file vero. Tutto quanto segue
e' **misurato**, non dedotto, e rigenerabile:

```bash
python tools/audit_workbook.py "samples/omni-report/Omni Report W30.xlsm" --all
python tools/extract_vba.py    "samples/omni-report/Omni Report W30.xlsm"
```

Metodo: lettura diretta dell'XML dentro il file (nessun Excel, nessun openpyxl) +
decompressione dei moduli VBA con oletools.

---

## 1. Esito: il contratto del piano §4 e' corretto, ma incompleto

Tutte le lettere canoniche del piano combaciano con le intestazioni reali —
tutte e 23, nessuna eccezione. Il reverse engineering era giusto.

Scansionando pero' **tutte** le formule del workbook sono emerse **tre colonne
consumate dal motore e assenti dal piano**. Sono state aggiunte a
`config/columns.yml`:

| Dataset | Col. | Header | Letta da |
|---|---|---|---|
| `AT_DATASET` | **H** | `Number of Active Contacts` | `Verifica AHT` |
| `ATwi_DATASET` | **J** | `Handle Time in seconds` | `Verifica AHT` |
| `PSAT_DATASET` | **AU** | `Survey Date (Exact)` | `Recap PSAT Positive` |

Perche' mancavano: il piano descrive i fogli `Helper Turni`, `Report Agenti`,
`Analisi Status`, `Malpractice Recap`, `Dettaglio Malpractice`,
`Outbound Exploitation`, `Presentazione Malpractice`, `Recap PSAT Positive`,
`PSAT Positive (export x email)`. Il file W30 ne ha **quattro in piu'**, nati
dopo la stesura del piano:

- `AHT Outliers`
- `Verifica AHT` ← consuma due delle tre colonne mancanti
- `AHT Outliers Export`
- `Profilo Colonne SF`

e `PSAT Positive Export` (il piano lo chiama `PSAT Positive (export x email)`).
Manca invece `Presentazione Malpractice`, citato dal piano ma non presente.

Senza quelle tre colonne la pipeline avrebbe prodotto un `Verifica AHT` vuoto:
nessun errore, solo numeri assenti.

> Morale, ed e' il motivo per cui `tools/audit_workbook.py` sta nel repo: il
> contratto va **ri-derivato dalle formule**, non riletto da un documento. Il
> test `tests/test_contract_reale.py` lo fa a ogni run.

## 2. Un errore facile nella scansione delle formule

`AT_DATASET` e' una sottostringa di `PSAT_DATASET`. Cercando
`AT_DATASET!$<lettera>` senza confine a sinistra, ad `AT_DATASET` vengono
attribuite tutte le colonne di `PSAT_DATASET` — nel primo giro sono comparse
`AT_DATASET!AU/DO/DP/DQ`, che non esistono. Il tool usa un lookbehind
`(?<![A-Za-z0-9_])`.

Vale anche al contrario, per chiunque scriva script su questi file: qualunque
`grep` su nomi di foglio che si contengono a vicenda va ancorato.

## 3. `SF_DATABASE` e' una tabella, e la tabella e' disallineata

Il foglio ospita un ListObject `AHT_Data`, 143 colonne. Nel file consegnato:

| | |
|---|---|
| intervallo della tabella | `A1:EM3389` |
| dati effettivi del foglio | `A1:EM3568` |
| **righe fuori dalla tabella** | **179** |

Conseguenza misurabile: `Profilo Colonne SF` calcola `ROWS(AHT_Data[])` e usa
`SF_DATABASE!$A$2:$EM$3389` come intervallo, quindi **profila 179 righe in meno
di quelle presenti**. Chi ha incollato i dati della W30 ha allargato il foglio
ma non la tabella.

Due ricadute sul codice:

1. `writer.py` ridimensiona il ListObject dopo la scrittura (`_resize_table`).
   Scrivere valori dentro l'intervallo di una tabella non la allarga da solo.
2. `Profilo Colonne SF` ha i limiti di riga scritti a mano nelle formule
   (`$EM$3389`): resta comunque da aggiornare a ogni settimana con un numero di
   righe diverso. Non tocca i numeri del report, ma il foglio dice il falso.
   Vedi domanda aperta §6.

## 4. VBA — cosa legge davvero il modulo `CreaMalpractice`

Codice decompresso e letto (18.641 caratteri, un solo modulo di sostanza).
Il piano §5 e' confermato, con una precisazione.

| Foglio | Colonne lette | Dove |
|---|---|---|
| `AT_DATASET` | B, F, K, P, Q | `AddATRules` |
| `SF_DATABASE` | BB, DY, AE | `AddSFRules` |
| `Anagrafica` | A, B | `LoadAnagrafica` |
| `Helper Turni` | A, B, E | `LoadHelperTurni` |
| `Turni` | A, E, F, G | `LoadTurniStarts` |
| `Slot Only Cases` | A, B, C, D, E | `LoadSlots` |
| `Helper Malpractice` | B2, B3, B4, B5, B6, B8 + tabella alias D:E | parametri |

Precisazione sul piano §5: di `Turni` il VBA legge **A, E, F, G** — non `C`
(Contratto) ne' `D` (Ore/gg), che il piano elenca. E di `Helper Malpractice`
legge `B2..B6` e `B8`, **non `B7`**: l'offset fuso (`B7` = 9 ore, Seattle→Milano)
serve solo alle formule `P`/`Q` di `AT_DATASET`, non alla macro.

Tre dettagli che vincolano il writer:

- **Ultima riga per `End(xlUp)`**: `AT_DATASET!B` e `SF_DATABASE!BB`. Righe
  della settimana precedente rimaste sotto i dati nuovi finirebbero nei
  conteggi. Da qui la pulizia sempre-prima-di-scrivere in `writer._clear_data`.
- **Due `MsgBox`** (riga 202 fine esecuzione, riga 207 errore). In automazione
  bloccano il processo a tempo indeterminato: serve il flag `SilentMode` del
  piano §12. Vedi `docs/architettura.md` §6.
- **`.Value2` sulle celle formattate** data/ora: il VBA lo usa perche'
  `IsNumeric` su una cella formattata come data ritorna `False`. Il writer deve
  quindi scrivere veri numeri/date, non stringhe.

## 5. Struttura misurata dei quattro dataset

| Foglio | Dimensione | Colonne header | Dati da | Note |
|---|---|---|---|---|
| `AT_DATASET` | `A1:Q130000` | 16 (B–Q) | B | A = indice riga (1..N). `P`/`Q` sono formule, materializzate fino a **riga 130000** |
| `ATwi_DATASET` | `A1:P15342` | 15 (B–P) | B | A = indice riga |
| `SF_DATABASE` | `A1:EM3568` | 143 (A–EM) | A | tabella `AHT_Data`, vedi §3 |
| `PSAT_DATASET` | `A1:ED132` | 134 (A–ED) | A | il piano diceva 121 colonne: l'export e' cresciuto |

Formule verbatim di `AT_DATASET`, riga 2:

```excel
P2: =IF($I2="","",INT($I2+'Helper Malpractice'!$B$7/24))
Q2: =IF($I2="","",($I2+'Helper Malpractice'!$B$7/24)-INT($I2+'Helper Malpractice'!$B$7/24))
```

Il limite di 130000 righe conta: se un export ne portasse di piu', le righe in
eccesso resterebbero senza `Data Milano`/`Ora Milano` e il VBA le **scarterebbe
in silenzio** (`If Not IsNumeric(d) ... GoTo NextR`). `writer.py` estende le
formule, o si ferma se non riesce. Il limite e' in `columns.yml`
(`max_template_row`).

`PSAT_DATASET` ha solo 131 righe di dati — e la riga fissa 130 dell'"Elogio
della settimana" (piano §4) e' quindi **quasi in fondo ai dati**. Con un export
piu' corto quella cella punta al vuoto; con uno piu' lungo pesca un commento a
caso. Vedi §6.

## 6. Collisioni di nome reali (perche' esiste `match: exact`)

Normalizzando i nomi (minuscole, `_` = spazio, punteggiatura via), alcune
colonne **diverse** collassano sulla stessa stringa. Nel PSAT sono due, e in
entrambi i casi le due colonne hanno semantica diversa:

| Normalizzato | Colonne che collidono | Quale vuole il motore |
|---|---|---|
| `agent name` | `Agent Name` (I) · `agent_name` (BO) | **I** |
| `psat score` | `psat_score` (DP, valori 0/1) · `PSAT Score` (EK, altra scala) | **DP** |

Due conseguenze di progetto:

1. Il matcher tenta il **nome grezzo prima del normalizzato**. Partendo dal
   normalizzato, un file perfettamente valido risulterebbe ambiguo.
2. Su questi due campi il contratto impone `match: exact`. Se l'export rinomina
   `Agent Name`, il preflight si **ferma** invece di ripiegare su `agent_name`,
   che sembrerebbe funzionare e darebbe numeri diversi. Verificato: rinominando
   `Agent Name` in `Agent Full Name`, la pipeline si blocca e propone
   `'Agent Full Name'` fra i nomi simili.

Non collidono invece, contrariamente al sospetto del piano: `Case Origin` (T) vs
`Case Origin (group)` (U), e nessuna delle altre colonne contenenti "AHT"
(`AHT due to Misroutes` DU, `aht_due_to_misrouted_hours` DV,
`Overall Performance to AHT Target` EJ, `logged_hrs_incld_in_aht` EH) collassa
su `case aht mins`.

## 7. Il quarto file: `WoW CaseType Deppdive W30.xlsm`

Nel repo c'era un terzo workbook che **nessuno dei due piani menziona**. Otto
fogli: `CSV DATASET`, `DATASET`, `Pivot`, `Monthly`, `tblClass`,
`Monthly with classification`, `manualEXPORT`, `formatCopy(Macro)`. Contiene VBA.

Sta in `samples/omni-report/` per non perderlo, ma il suo ruolo **non e' stato
indagato**: la struttura (un `DATASET`, un `Pivot`, una classificazione per
case type) somiglia a quella del WOW AHT Trend, e `manualEXPORT` +
`formatCopy(Macro)` suggeriscono un altro passaggio manuale settimanale. Da
chiarire prima di dargli un posto nell'architettura.

## 8. Domande aperte che questo audit aggiunge

Le domande dei due piani restano in piedi (piano §15, contesto WOW §11). Questo
audit ne aggiunge quattro, tutte da girare a Leonardo:

1. **`Verifica AHT`, `AHT Outliers`, `AHT Outliers Export`, `Profilo Colonne SF`
   fanno parte del ciclo settimanale**, o sono fogli di lavoro nati per
   un'indagine e poi rimasti? Cambia se le tre colonne nuove vanno considerate
   obbligatorie (oggi lo sono) oppure opzionali.
2. **`Profilo Colonne SF` va aggiornato automaticamente?** Ha i limiti di riga
   scritti a mano nelle formule e oggi e' disallineato di 179 righe.
3. **La tabella `AHT_Data` serve a qualcosa oltre a `Profilo Colonne SF`?** Se
   e' solo di supporto, il ridimensionamento e' cosmetico; se un giorno le
   formule del motore passassero ai riferimenti strutturati (piano §M6),
   diventa il punto di partenza.
4. **`WoW CaseType Deppdive W30.xlsm`**: che ruolo ha (vedi §7)?
