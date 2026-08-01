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

## 8. `Turni` e `Slot Only Cases` sono fonti, non config

Nella prima passata li avevo classificati come "config mantenuta a mano". Sono
invece **dati settimanali**, che arrivano da due export in forma di **matrice
larga** (un agente per riga, una colonna per giorno). Il foglio del workbook è in
forma lunga: serve un unpivot.

Rigenerabile con:

```bash
python tools/audit_shift_forms.py "samples/omni-report/sorgenti/Turni_W30.xlsx"
python tools/golden_turni.py --workbook "samples/omni-report/Omni Report W30.xlsm" \
    --roster "samples/omni-report/sorgenti/Turni_W30.xlsx" \
    --backoffice "samples/omni-report/sorgenti/Back_Office_Time_Final.xlsx" --monday 46223
```

### 8.1 Le due sorgenti

| | `Turni_W30.xlsx` (foglio `publish`) | `Back_Office_Time_Final.xlsx` (`Only Cases Shifts`) |
|---|---|---|
| dimensione | `A1:VP132`, 130 righe agente | `A1:FH78`, 74 righe |
| formule | 4 (irrilevanti) | **2873** — è un file di lavoro vivo, non un export |
| colonne-giorno | 494, dal 20/07/2026 al 15/09/2026 | 150, dal 04/05/2026 |
| intestazione data | testo `gg/mm/aaaa` | seriale Excel |
| contenuto cella | `0900_1331_1431_1800` (inizio, inizio pausa, fine pausa, fine) | `1000_1030`, `NO BOT`, `REQUEST`, numeri |

Il roster ha **due blocchi affiancati** (`A`–`F` + date, poi `KI`–`KN` + date) con
popolazioni diverse: 130 e 107 agenti, **96 in comune e 65 di questi con skill
diversa fra i due blocchi**. Il blocco 2 non contiene alcun `HPO`, quindi per
l'Omni Report si ignora — ma contiene `HPS`, che somiglia a `HPO` quanto basta a
fare danni: il confronto deve essere per uguaglianza, mai per "contiene". Se un
giorno un agente fosse `HPO` in entrambi i blocchi produrrebbe righe doppie, e
`Helper Turni` gli conterebbe le ore due volte: c'è un controllo che blocca.

Il foglio back-office ha **sezioni** (`HPO`, `Part-Time 6h`, `Part-Time 5h / 4h`)
e righe di totale (`Total BO Hrs`, `PSP BO Hrs`, `intervals`, e l'etichetta
`Agents in Only Cases per Interval`): non è una tabella piatta.

### 8.2 Le 26 forme delle celle-turno

Enumerate prima di scrivere il parser, perché diverse non erano prevedibili:

| Occorrenze | Forma | Interpretazione |
|---|---|---|
| 48985 | cella vuota | non schedulato |
| 3761 + 3352 | `0900_1331_1431_1800` (± spazio finale) | turno con pausa |
| 5481 | `off` / `Off` / `OFF` | riposo |
| 1530 + altri | `0900_1300_1400_1800 O` | turno, **marcatore ` O`** |
| 294 + 202 | `0800_1200_␣␣␣␣_␣␣␣␣` | turno **senza pausa** |
| 144 + 69 | `s0900_1230_1300_1700` | turno, **prefisso `s`** |
| 57 + 24 + 21 | `_1000_1239_1339_1800` | turno, **underscore iniziale** |
| 101 | `␣␣␣_␣␣␣_␣␣␣_␣␣␣ O` | non schedulato, con marcatore |
| 79 | `OFF␣␣␣␣␣␣␣␣␣␣␣O` | riposo con marcatore |
| **28** | **`Training`** | **né turno né riposo** |
| **13** | **`Flessibilità`** | **né turno né riposo** |
| 2 | `o1400 1430` | turno, prefisso `o`, **separatore spazio** |

`Training` e `Flessibilità` non compaiono nel W30: non si sa come il processo
manuale li tratti, quindi la pipeline **salta** quelle righe e le segnala.
Trattarle come riposo falserebbe le ore previste.

I marcatori (` O` 1839 volte, prefisso `s` 220, underscore iniziale 230,
prefisso `o` 2) hanno significato **non documentato**. Il parser li conserva
invece di scartarli: se un giorno si scopre che ` O` vuol dire straordinario, il
dato non è già stato buttato.

### 8.3 La semantica, ricavata dal confronto

Il W30 contiene il risultato prodotto da queste stesse sorgenti, quindi la
trasformazione si verifica invece di indovinarla.

| `Turni` | Da | Regola |
|---|---|---|
| `A` Nome agente | `Name` + `Surname` | concatenati con uno spazio |
| `B` Team/Skill | `Skill` | copia; solo `HPO` **esatto** entra |
| `C` Contratto | — | non derivabile, vedi 8.5 |
| `D` Ore/gg | cella | `(fine − inizio) − pausa`, in ore, **2 decimali** |
| `E` Data | intestazione colonna | seriale Excel |
| `F` Stato | cella | `LAVORA` / `FERIE-OFF` |
| `G`/`H` Inizio/Fine turno | cella | 1° e ultimo `HHMM`, frazione di giorno |
| `I` chiave | `A` | `normalize_name`, **solo sulle righe LAVORA** |

Verifica aritmetica (Ahmed Afifi): `0900_1300_1330_1630` → 7,5h − 0,5h = **7h**,
e il workbook dice 7. `0800_1300_1330_1630` → **8h**, workbook 8.
`0900_1400_␣␣␣_␣␣␣` → **5h**, workbook 5.

L'arrotondamento a 2 decimali viene da **un solo campione**: Viktoriia Lavrinets,
20/07, turno 08:00–13:11 = 5,1833h, e il workbook ha `5.18`. Non si distingue da
un troncamento; con altre settimane va riverificato, e il golden test lo
intercetterebbe.

Esito del confronto: **252 righe ricostruite = 252 nel workbook**, chiavi
identiche, e **259 = 259** per gli slot. Le sole differenze, tutte dichiarate in
anticipo:

- sulle righe `FERIE-OFF` il workbook ha valori spazzatura (`Ore/gg`=8,
  `Inizio turno`=1447); la pipeline scrive celle vuote. Nessuna formula le legge:
  ogni SUMIFS/MINIFS di `Helper Turni` filtra `Stato="LAVORA"`.
- `Stato BO`: la pipeline scrive `NO BOT` dove il processo manuale lascia vuoto.
  Il VBA ha un ramo `If status <> "NO BOT"` che oggi **non scatta mai**; l'esito
  numerico non cambia (senza orari la riga viene scartata comunque), ma
  l'intenzione diventa leggibile.

### 8.4 Il guasto: un agente che sparisce a metà

`Turni` ha 36 agenti, `Slot Only Cases` ne ha 37. Il differenziale è **Nora Ed
Dahir**, e la causa è misurabile:

```
sorgente, riga 43:  Skill = 'HPO                *'
Helper Turni!A2  =  FILTER(Turni!$A, ..., Turni!$B="HPO")   <- uguaglianza esatta
```

Oggi, nel report della W30:

- **è in `Slot Only Cases`** → la regola "Available Cases fuori turno" la valuta
- **non è in `Turni`** → nessuna ora prevista, nessun orario di inizio turno,
  quindi "Login in ritardo" la giudica contro il `defaultStart` del VBA
  (`Helper Malpractice!B8` = 0,375 = 09:00) invece del suo turno reale

Nessun errore, nessuna cella rossa: solo una riga che manca. Lo stesso marcatore
esiste su `VRBO␣␣*` (2 agenti) e `RELO␣␣*` (10 nel blocco 2): è sistematico, non
un errore di battitura.

Il default della pipeline è **fedele a oggi** (l'agente resta fuori) ma il
preflight **blocca**, da due direzioni indipendenti: il controllo sulle skill
quasi-uguali e quello sugli insiemi di agenti. Per includerla:
`sources.include_marked_skills: true` in `settings.yml` — cambia i numeri, quindi
è una scelta esplicita.

### 8.5 La tabella alias mescola due direzioni

`Helper Malpractice!D:E` (righe 3–10, letta da `CreaMalpractice.LoadSlots`)
contiene 7 alias utili, e non sono tutti dello stesso tipo:

| Alias | Direzione |
|---|---|
| `alessandro passierello` → `alessandro passariello` | grafia back-office → **grafia roster** |
| `asia chirrullo` → `asia chirullo` | idem |
| `kaotar garoui` → `kaotar garaoui` | idem |
| `victoriia lavrinets` → `viktoriia lavrinets` | idem |
| `eleonora rosa sissa` → `eleonora sissa` | grafia roster → **grafia Salesforce** |
| `glenda martina medola` → `glenda medola` | idem |
| `nadia ariefieva` → `nadiia ariefieva` | idem |

Il foglio `Slot Only Cases` usa la grafia del **roster** (verificato: le sue 37
chiavi coincidono esattamente con gli agenti HPO del roster). Applicare la
tabella alla cieca porterebbe quei 3 agenti *fuori* da quello spazio di nomi, e
le loro 21 righe sparirebbero.

Quindi la regola: l'alias si usa per **raggiungere** lo spazio dei nomi del
roster, non per lasciarlo — se la chiave è già valida si tiene, altrimenti si
prova l'alias.

### 8.6 `Contratto` non è derivabile — misurato

`Expected hours` → `Contratto` sui 36 agenti:

| Expected hours | Contratto | Agenti |
|---|---|---|
| `0400` | PT | 1 |
| `0500` | PT | 1 |
| `0600` | **FT** | 5 |
| `0600` | **PT** | 2 |
| `0800` | FT | 27 |

Con `0600` esistono sia FT sia PT: nessuna regola può produrlo dalla sorgente.
È un dato HR esterno, e sta in `config/contratti.yml`.

Non entra in alcun calcolo: `Helper Turni!C` lo espone, ma di `Helper Turni` il
motore legge solo `B`, `N`, `R`. Un nome mancante lascia la cella vuota e il
preflight lo segnala, senza bloccare.

### 8.7 Un buco nell'audit tool, corretto

Nelle formule i nomi di foglio con spazi sono **fra apici**:
`'Slot Only Cases'!$A$2`. Il pattern originale pretendeva `Cases!` e non li
vedeva: `Email Agenti` risultava "letto da nessuna formula", che è falso (lo
leggono `Anagrafica` e `Helper Turni`). Serve anche un confine a sinistra che
escluda lo spazio, altrimenti `Turni` cattura le colonne di `Helper Turni`.

## 9. Domande aperte che questo audit aggiunge

Le domande dei due piani restano in piedi (piano §15, contesto WOW §11). Questo
audit ne aggiunge otto, tutte da girare a Leonardo. Le prime quattro nascono dai
fogli nuovi, le altre dai dataset:

0. **Nora Ed Dahir deve entrare nel report?** Il suo `Skill` è `HPO␣␣␣*`.
   Oggi è dentro per metà delle regole (§8.4). Cambia i numeri della W30.
1. **Cosa segna l'asterisco** in `Skill` (`HPO␣*`, `VRBO␣*`, `RELO␣*`)?
   È sistematico, non un errore.
2. **Cosa segna il suffisso ` O`** nelle celle-turno (1839 occorrenze), e i
   prefissi `s` (220) e `o` (2)? Oggi si conservano ma non si usano.
3. **`Training` (28 celle) e `Flessibilità` (13)**: contano come ore previste,
   come riposo, o si ignorano? Oggi si ignorano e si segnalano.
4. **`REQUEST` nel back office**: richiesta di cambio turno pendente? Oggi vale
   "nessuno slot", come fa il processo manuale.

E poi:

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
