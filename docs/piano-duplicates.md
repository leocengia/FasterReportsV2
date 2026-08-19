# Piano: allineare il programma al template WIP (sezione Duplicate Cases)

> **Stato al 2026-08-19: il lato Python è fatto. Restano le fasi 6 e 7, che sono
> lavoro in Excel.**
>
> | fase | contenuto | stato |
> |---|---|---|
> | 1 | `tools/audit_workbook.py`: quattro difetti del lettore XML | ✅ `289dd86` |
> | 2 | template promosso, fixture rigenerata | ✅ `b68f0eb` |
> | 3 | contratto `DUP_DATASET`, preambolo, coda dei totali | ✅ `0198d0d` |
> | 4 | capienze, settimana, limite di riga a cella singola | ✅ `1311014` |
> | 5 | scrittura del preambolo, sezioni vuote attese | ✅ `991c0ba` |
> | **6** | **template: `M2`/`N2` di `Duplicates Helper`** | **da fare — §4.2** |
> | **7** | **template: capienze, collegamento esterno, banner** | **da fare — §4.3–4.5** |
> | 8 | documentazione | ✅ |
>
> **617 test passati, 33 nuovi.** La fase 6 è necessaria prima del primo build: il
> contratto ora scrive date vere, e finché `Duplicates Helper!M2` le parsa con
> `FIND("/")` quelle colonne danno `#VALUE!`. La fase 7 no — il build gira comunque.
>
> Misurato il 2026-08-19 sul template WIP (commit `54f8fd7`), 29 fogli, 8,08 MB.
> Ogni numero in questo documento è letto dal file, non stimato.

---

## 0. Cosa è cambiato, e cosa invece era già a posto

Dei sette fogli elencati nella richiesta, **quattro erano già allineati**:
`Helper CaseType`, `CaseType Deepdive`, `AHT Trend WoW` e `AHT History` sono
entrati nel template tracciato il 2026-08-18 (`docs/template-modifiche-wip.md`,
tutti e cinque i punti chiusi). Il confronto foglio per foglio fra il WIP e
`Omni_Report_TEMPLATE.xlsm` dà **cinque fogli nuovi e zero rimossi**:

| foglio | dimensione | ruolo |
|---|---|---|
| `DUP_DATASET` | `A1:Q194` | export grezzo del report SF duplicati |
| `Duplicates Helper` | `A1:W1001` | 13 120 formule che rimettono in forma DUP_DATASET |
| `DC Dashboard` | `A1:BE80` | 220 formule + 6 grafici |
| `DC Agents & Categories` | `A1:M42` | 188 formule |
| `DC Timing & Quality` | `A1:I46` | 89 formule |

Ed è verificato che **il WIP è un soprainsieme del template tracciato**, quindi
promuoverlo non perde nulla di quanto fatto ad agosto:

- `vbaProject.bin` presente, `SetSilentMode` e `Refresh_Dettaglio_Malpractice` dentro;
- `AHT Trend WoW!AB4` è ancora la riga di `week_key` e `B4` è ancora `MOD(AB4#,100)`;
- i limiti di riga vincolanti sono identici: `AT_DATASET` 130000, `SF_DATABASE`
  50000, `PSAT_DATASET` 1000, `Turni` 10000 — quindi
  `test_limiti_di_riga_del_template_reale` continua a passare;
- zero celle di errore in tutto il workbook.

### La catena di dipendenze dei fogli nuovi è pulita

Scansionando tutte le formule del workbook, chi legge chi:

```
DUP_DATASET  ──►  Duplicates Helper  ──┬─►  DC Dashboard  (── DC Agents & Categories)
                                       ├─►  DC Agents & Categories
                                       └─►  DC Timing & Quality
```

**Nessun foglio preesistente legge i cinque nuovi, e i cinque nuovi non leggono
nessun foglio preesistente.** È la notizia migliore del piano: la sezione
duplicati è un blocco isolato, e sbagliarla non può sporcare i numeri di
malpractice, AHT o PSAT.

---

## 1. La forma vera di `DUP_DATASET` — ed è diversa da tutti gli altri dataset

I sei dataset esistenti hanno le intestazioni in **riga 1** e i dati subito
sotto. `DUP_DATASET` no: è il **report Salesforce formattato incollato tale e
quale**, preambolo compreso.

```
riga  2   "Leo's Orchidea Dup Cases (date interval)"        ← titolo
riga  3   "As of 2026-08-13 12:41:48 Central European..."   ← quando è stato scaricato
riga  6   "Filtered By"
riga  7.. "Show: All users" / "Manager: Full Name equals Emilie McKenzie" /
          "Office Location equals Milan" /
          "Status equals Closed - Duplicate Case,Closed - Spam" /
          "Date/Time Closed greater or equal 8/3/2026" /
          "Date/Time Closed less or equal 8/9/2026 11:30 PM"
riga 14   INTESTAZIONI  (colonne B..Q, con C senza nome)
riga 15   primo record
riga 192  ultimo record            ← 178 righe di dati
riga 193  "Total" | "Sum"  | J=72,13
riga 194            "Count" | D=178      ← due righe di totali in coda
```

Tre conseguenze, tutte da gestire:

1. **`header_row: 14`, non 1.** Il writer scrive già a `header_row + 1`, quindi i
   dati finiscono in riga 15 senza modifiche: questa parte funziona da sola.
   Ma il **lettore** no — vedi §3.3.
2. **La colonna C non ha intestazione**: è la colonna in cui il report mette le
   etichette `Sum`/`Count` del totale. Va lasciata non mappata, ma va **pulita**
   ogni settimana, quindi deve stare dentro `data_end_col`.
3. **Le due righe di coda sono dati finti.** Se entrassero nel blocco, `Total`
   diventerebbe un agente. Il template si difende già
   (`IF(OR(DUP_DATASET!B15="",DUP_DATASET!B15="Total"),"",...)`), ma è sbagliato
   scriverle: vanno scartate a monte, e va detto quante.

### Il vincolo che nessun altro dataset ha: `Duplicates Helper` legge cella per cella

`Duplicates Helper!A2` è `IF(OR(DUP_DATASET!B15="",...),"",DUP_DATASET!B15)`;
`A3` punta a `B16`; e così via fino a `A1001` → `B1014`. **Offset fisso +13.**
Non è un `SUMIFS` su un intervallo: è una griglia di riferimenti puntuali, 13 120
formule.

Due cose ne seguono, e sono entrambe importanti:

- **Capienza di `DUP_DATASET`: 1000 righe** (righe 15..1014). Questa settimana
  sono 178, il 18%: c'è margine. Ma il controllo che il progetto fa per tutti gli
  altri dataset (`templatescan.scan_row_limits`) **non la vede**, perché cerca
  intervalli (`DS!$A$2:$A$999`) e qui ci sono celle singole. Va esteso — §3.9.
- **Se qualcuno inserisce una riga sopra la 14, tutto slitta di uno in silenzio.**
  Il preflight non lo vedrebbe, i numeri sarebbero di un'altra riga. Serve un
  test che inchiodi l'offset — §5.

### Le date sono testo, e questo va cambiato

Nel download `Date/Time Opened` e `Date/Time Closed` arrivano come **stringhe**
(`8/4/2026 3:59 PM`, verificato: celle `t="s"` in DUP_DATASET). Il template le
riparse a mano:

```
Duplicates Helper!M2 = IF($K2="","",DATE(VALUE(MID(K2,FIND("/",K2,FIND("/",K2)+1)+1,4)),
                          VALUE(LEFT(K2,FIND("/",K2)-1)), VALUE(MID(K2,FIND("/",K2)+1,…))))
```

— e le intestazioni `K1`/`L1` dicono infatti `Opened (text)` / `Closed (text)`.

**Tenere questo assetto è un rischio concreto, non teorico.** Se la pipeline
scrive `"8/4/2026 3:59 PM"` come testo in una cella in formato Generale, Excel in
locale italiano la converte in **data** (8 aprile 2026) al momento della
scrittura: `FIND("/")` su una data dà `#VALUE!`, e le colonne `Opened`, `Closed`,
`Day`, `Weekday`, `Open hour`, `TTC (hours)`, `TTC bucket` cadono tutte insieme —
cioè metà dei fogli DC. Nessun modo di garantirlo dal lato Python.

**Scelta consigliata: la pipeline scrive date vere, e il template smette di
parsare il testo.** Costa due celle da riscrivere e trascinare (§4.2), e in
cambio:

- `date_format: "%m/%d/%Y %I:%M %p"` nel contratto rende la lettura **decisa**:
  è la stessa lezione di `Date Viewpoint`, dove `8/10/2026` senza formato dichiarato
  produceva la settimana 41 invece della 33, in silenzio. Verificato:
  `8/4/2026 3:59 PM` senza formato è **`Uncoercible`** (la lista dei formati noti
  non lo copre) e con il formato dà `2026-08-04 15:59`;
- il preflight conta e segnala le date non convertibili, che oggi passerebbero
  come testo qualunque;
- si può controllare che il periodo coperto dai duplicati sia la settimana del
  report (§3.10);
- sparisce un livello di parsing fragile da 2000 formule.

---

## 2. Le decisioni prese

| # | decisione | scelta |
|---|---|---|
| D1 | formato del download | **Excel formattato `.xlsx`**, con preambolo e totali |
| D2 | se il file manca | **si procede, segnalando** (come `PSAT_DATASET`) |
| D3 | righe 1–13 di `DUP_DATASET` | **le riscrive Python dal download** |
| D4 | capienze dei fogli DC | **si allargano nel template + controllo nel preflight** |
| D5 | colonne data | **date vere + due formule del template semplificate** (§1) |
| D6 | settimana dei duplicati | **la stessa del resto del report**, e il controllo **BLOCCA** (§3.10) |

---

## 3. Lavoro lato Python

### 3.1 `config/columns.yml` — il dataset `DUP_DATASET`

```yaml
  DUP_DATASET:
    sheet: DUP_DATASET
    header_row: 14          # NON 1: il report SF porta 13 righe di preambolo
    data_start_col: B
    data_end_col: Q         # oltre l'ultimo campo mappato: serve a PULIRE C, P, Q
    optional: true          # D2 — nessun campo è letto dal VBA
    reader: table
    key_field: "Full Name"  # riga senza questo valore = non è un record
    stop_values: ["Total"]  # e questo chiude la tabella (riga dei totali SF)
    read_by_row: true       # 'Duplicates Helper' lo legge cella per cella
    max_consumer_row: null  # si RICAVA dal template, non si scrive qui (§3.9)
    fields:
      - canonical: "Full Name"                  target_col: B  dtype: str
      - canonical: "Case Number"                target_col: D  dtype: str
      - canonical: "Parent Case: Case Number"   target_col: E  dtype: str
      - canonical: "Status"                     target_col: F  dtype: str
      - canonical: "Case Origin"                target_col: G  dtype: str
      - canonical: "Date/Time Opened"           target_col: H  dtype: datetime
        date_format: "%m/%d/%Y %I:%M %p"
      - canonical: "Date/Time Closed"           target_col: I  dtype: datetime
        date_format: "%m/%d/%Y %I:%M %p"
      - canonical: "Case Age (days)"            target_col: J  dtype: float
      - canonical: "Case Record Type"           target_col: K  dtype: str
      - canonical: "Type"                       target_col: L  dtype: str
      - canonical: "Primary Category"           target_col: M  dtype: str
      - canonical: "Language: Language Name"    target_col: N  dtype: str
      - canonical: "Email"                      target_col: O  dtype: str
      - canonical: "Last Modified By: Full Name" target_col: P dtype: str
      - canonical: "Comments/Remarks"           target_col: Q  dtype: str
```

(forma abbreviata: nel file vero ogni campo va scritto con `role: input`,
`aliases`, `consumers` e le note, come gli altri.)

I `consumers` sono misurati, non ipotizzati — sono le colonne che
`Duplicates Helper` cita davvero:

| DUP | intestazione | letta da `Duplicates Helper` |
|---|---|---|
| B | Full Name | `A` (Agent) |
| C | *(senza nome)* | **nessuno** — colonna delle etichette di totale |
| D | Case Number | `C` |
| E | Parent Case: Case Number | `D` |
| F | Status | `E` |
| G | Case Origin | `F` |
| H | Date/Time Opened | `K` → `M`,`O`,`P`,`Q`,`S` |
| I | Date/Time Closed | `L` → `N`,`S` |
| J | Case Age (days) | `R` |
| K | Case Record Type | `G` |
| L | Type | `H` |
| M | Primary Category | `I` |
| N | Language: Language Name | `J` |
| O | Email | `B` |
| P | Last Modified By: Full Name | **nessuno** |
| Q | Comments/Remarks | **nessuno** (vuota in tutte le 178 righe) |

`P` e `Q` si mappano comunque: costano niente, e il preflight che le dichiara al
100% vuote è più utile di una colonna che nessuno sa se esista.

### 3.2 `config/settings.yml` — pattern del file

```yaml
input_files:
  DUP_DATASET: "*Dup*Cases*"
```

Da confermare sul nome vero del download (il report si chiama
`Leo's Orchidea Dup Cases (date interval)`). Come per gli altri, il pattern deve
corrispondere a **un solo** file.

Aggiungere anche il dataset all'elenco dei sei nel commento di testa della
sezione, che oggi dice «quattro file».

### 3.3 `core/tablesource.py` — trovare le intestazioni in riga 14

`find_header_row` scandisce solo le **prime 10 righe** (`limit: int = 10`), e le
intestazioni sono in riga 14: oggi solleverebbe
«nessuna riga di intestazione riconoscibile».

L'euristica che c'è già funziona benissimo su questo file, va solo lasciata
guardare più in basso: le righe di preambolo hanno **una cella** ciascuna, la 14
ne ha **16** — quindi `n * 2 < larga` scarta tutto il preambolo e la 14 vince
subito. Verificato a mano sul foglio.

- alzare il limite di default (20 basta e avanza), **con il commento che dice
  perché**: un report SF formattato porta titolo + `As of` + il blocco
  `Filtered By`, e quel blocco cresce con il numero di filtri;
- `_open_source` in `orchestrate.py` chiama `read_table(path)` senza argomenti:
  va bene, il riconoscimento resta automatico. **Non** passare
  `dataset.header_row`: quello è la riga nel *foglio di destinazione*, che
  coincide con 14 per caso. Se un giorno il download avesse 15 righe di
  preambolo, i dati devono comunque finire in riga 15 del foglio.

### 3.4 `core/tablesource.py` + `core/transform.py` — scartare le righe di coda

`read_table` tiene le righe 193 e 194 (non sono vuote), e diventerebbero due
record con agente `Total` e agente vuoto.

- `TableSource` restituisce anche `preamble` (le righe sopra l'intestazione) e le
  righe dati come oggi;
- il filtro vive in `build_block`, che è dove il contratto è già in mano: si
  scarta ogni riga in cui il campo `key_field` è vuoto o è in `stop_values`, e da
  quel punto in poi si smette di leggere;
- il conteggio va nel preflight: «2 righe di coda scartate (`Total`, `Count`)».
  Scartare in silenzio è quello che questo progetto non fa.

`key_field`/`stop_values` sono generici, non un caso speciale: **tutti** gli
export SF formattati finiscono così.

### 3.5 `omni/writer.py` — riscrivere il preambolo (D3)

Oggi le righe 1–13 sono testo fisso del template: a dicembre direbbero ancora
`As of 2026-08-13` e `Date/Time Closed greater or equal 8/3/2026`. Una data
sbagliata che sembra giusta è il difetto che questo progetto insegue da mesi.

Nuova funzione `write_preamble(sht, dataset, preamble)`:

- pulisce `A1:Q{header_row - 1}` e ci scrive il preambolo del download;
- **non tocca la riga 14**: le intestazioni restano quelle del template, ed è
  giusto — `Duplicates Helper` legge per posizione, quindi la riga 14 del
  template *è* il contratto, e un rename nell'export lo becca il matcher;
- se il preambolo del download è più lungo di 13 righe, scrive le prime 13 e
  **avvisa**: significa che il report ha filtri in più, e il numero di righe di
  preambolo è cambiato;
- se il download non ha preambolo (CSV pulito), le righe 1–13 restano pulite.

Va chiamata da `write_block` quando `dataset.header_row > 1`, prima della
scrittura dei dati.

### 3.6 `optional: true` e il file che manca (D2)

Il meccanismo esiste già e funziona: `_empty_block` + `write_block` **pulisce**
il foglio e non ci scrive nulla, così non restano i duplicati della settimana
prima. Ma qui c'è una conseguenza che `PSAT_DATASET` non aveva:

`Duplicates Helper!W1` è `COUNTIF(A2:A1001,"?*")`, e i fogli DC ci dividono
**56 volte** (42 su `DC Agents & Categories`, 14 su `DC Timing & Quality`). Con
zero righe, `W1 = 0` e quelle 56 celle diventano `#DIV/0!`. `scan_error_cells`
le trova, `BuildResult.ok` diventa `False`, e il programma dichiara **fallito**
un build che invece è andato bene: il file mancava, e il report lo dice.

Rimedio, lato codice (obbligatorio):

- nel contratto, `DUP_DATASET` dichiara i suoi `dependent_sheets`:
  `["Duplicates Helper", "DC Dashboard", "DC Agents & Categories", "DC Timing & Quality"]`;
- quando un dataset `optional` viene **saltato**, `build` esclude quei fogli dal
  verdetto sulle celle di errore e mette **una riga sola** nel risultato:
  «export duplicati assente: i 3 fogli DC sono vuoti (56 celle `#DIV/0!`,
  atteso)». Le celle di errore di **tutti gli altri** fogli continuano a contare
  come prima.

Così l'assenza è dichiarata invece che scoperta, e non contamina il giudizio sul
resto del workbook.

### 3.7 `core/coherence.py` — controllo delle capienze dei fogli DC (D4)

I fogli DC hanno formule **riga per riga** con l'ultima riga scritta nel foglio.
Superarle non dà nessun errore: l'agente in eccesso compare nell'elenco (che è
un array dinamico, e cresce) con **nessun numero accanto**. È esattamente il
difetto tolto a `Helper CaseType` ad agosto.

Misurato: ultima riga con formula, per colonna.

| cosa | foglio · colonne | righe | capienza | questa settimana | pieno |
|---|---|---|---|---|---|
| agenti | `DC Agents & Categories` B,C,D | 7–40 | **34** | 28 | **82%** |
| record type | `DC Agents & Categories` G,H | 7–18 | **12** | 7 | 58% |
| Type | `DC Agents & Categories` K,L | 7–32 | **26** | 19 | **73%** |
| parent con >1 duplicato | `DC Timing & Quality` G,H,I | 17–36 | **20** | 9 | 45% |
| agenti | `DC Dashboard` AB | 21–65 | 45 | 28 | 62% |
| record type | `DC Dashboard` BB | 21–40 | 20 | 7 | 35% |
| Type | `DC Dashboard` BE | 21–60 | 40 | 19 | 48% |
| Case Origin distinti | `DC Dashboard` spill di `AA5` | 5–19 | **15** | 8 | 53% |
| righe dati | `DUP_DATASET` via Helper | 15–1014 | 1000 | 178 | 18% |

**Gli agenti sono già all'82%.** Il team HPO oggi è di ~36 persone: basta una
settimana in cui ne finiscono 35 nei duplicati e il 35° compare senza numeri.

Nuovo controllo, che segue la stessa forma di `_check_row_limits` (SEGNALA
all'80%, BLOCCA sopra il 100%):

- le capienze si **leggono dal template**, non si scrivono nel codice: nuova
  `templatescan.scan_formula_extent(path, [(foglio, colonna, prima_riga), ...])`
  che restituisce l'ultima riga con formula. Se domani allarghi le formule, il
  controllo lo segue da solo — è lo stesso principio dei formati numerici di
  `AHT History`, che vivono nel template;
- quanti valori distinti ci sono lo calcola Python dal blocco `DUP_DATASET`
  appena letto (agenti distinti, record type distinti, `Type` distinti,
  `Case Origin` distinti, parent con più di un duplicato);
- l'elenco degli "scaffali" (foglio, colonne, prima riga, e quale conteggio dei
  dati li riempie) sta in un modulo nuovo `core/duplicati.py`, una tabella con
  una riga di commento per voce: è dato, ma dato che ha bisogno di una
  spiegazione accanto.

Nota su `AA5` di `DC Dashboard`: lì il limite non è una formula per riga ma uno
**spill** che sbatte contro `AA20` (`Agent`). Con più di 15 `Case Origin`
distinti diventa `#SPILL!` e il menu a tendina del filtro Origin smette di
funzionare. Va nella stessa tabella, con la sua nota.

### 3.8 `core/coherence.py` — nomi degli agenti dei duplicati

`_case_owners` confronta già i nomi delle fonti caso-per-caso con `Email Agenti`.
`DUP_DATASET!Full Name` è un'altra fonte caso-per-caso: aggiungerla alla mappa
(`"DUP_DATASET": "Full Name"`) fa scattare gratis il controllo «chi ha lavorato
duplicati e non è in `Email Agenti`». Una riga.

### 3.9 `core/templatescan.py` — il limite di riga che oggi non si vede

`scan_row_limits` cerca `DS!$A$2:$A$999`. `Duplicates Helper` usa
`DUP_DATASET!B15`: celle singole. Risultato: `DUP_DATASET` non ha alcun limite
noto, e il controllo che protegge tutti gli altri dataset non lo copre.

- estendere la scansione ai **riferimenti a cella singola**, prendendo la riga
  massima per dataset;
- farlo **solo** per i dataset che dichiarano `read_by_row: true`. Non è
  pignoleria: `Recap PSAT Positive` punta alla riga **fissa** `DQ130`
  (l'"elogio della settimana", scelto a mano). Una scansione indiscriminata
  leggerebbe quel 130 come il limite di `PSAT_DATASET` e **bloccherebbe** ogni
  settimana con più di 130 risposte al sondaggio.

Fatto questo, `_check_row_limits` copre `DUP_DATASET` senza una riga in più:
1014 come limite, 192 come ultima riga → 19%, nessun avviso.

### 3.10 `core/coherence.py` — la settimana coperta dai duplicati

L'export è filtrato su `Date/Time Closed` (`>= 8/3/2026`, `<= 8/9/2026 11:30 PM`
nel paste): con `dtype: datetime` (§1) il periodo si ricava **dai dati**, min e
max della colonna I, e si può confrontare con la settimana del report.

**Livello BLOCCA** (deciso il 2026-08-19, D6): l'export duplicati deve coprire la
**stessa** settimana del resto del report. È la stessa regola che già vale per
`Date Viewpoint` quando l'export SF è di una settimana diversa, e per la stessa
ragione: un report con l'etichetta sbagliata **viene archiviato**, ed è peggio di
un report che manca.

Nel WIP il paste dei duplicati copre **3–9 agosto (W32)** mentre
`SF_DATABASE!Date Viewpoint` dice **10 agosto (W33)**. Non è un processo
sfasato: è che al momento del montaggio del template quello della W32 era il
solo export duplicati disponibile. Quando il report lo genera il programma, le
due settimane coincidono — quindi lo sfasamento non va tollerato, va bloccato.

Il messaggio deve dire entrambi i periodi e da dove vengono, perché è l'unico
modo di capire quale dei due file è quello sbagliato:

```
BLOCCA  settimana dei duplicati diversa dal resto del report
        duplicati (Date/Time Closed):  03/08/2026 → 09/08/2026  (W32)
        report    (Date Viewpoint):    10/08/2026 → 16/08/2026  (W33)
        Scarica di nuovo il report duplicati con l'intervallo della settimana
        giusta, oppure controlla di non aver lasciato in input/ il file della
        settimana scorsa.
```

Nota di attuazione: il perimetro dei duplicati è `Date/Time Closed`, cioè si
ricava da min/max della colonna `I`, **non** dalle righe del preambolo. Il
preambolo dice l'intervallo *richiesto* al report; le date dei casi dicono
l'intervallo *ottenuto*. Se una settimana non ha duplicati chiusi in un certo
giorno, i due non coincidono — e quello che conta per l'allineamento è il
secondo. Vale la pena calcolare la settimana ISO dalla **moda** dei giorni, non
dal min/max, per lo stesso motivo per cui `_resolve_week` la calcola così su
`AT_DATASET`: un caso chiuso a cavallo della mezzanotte del lunedì non deve
spostare la settimana di tutto l'export.

### 3.11 `omni/doctor.py`

`_check_template_sheets` costruisce `wanted` da `{ds.sheet for ds in contract.datasets}`:
`DUP_DATASET` entra da solo appena è nel contratto, niente da fare.

Da aggiungere invece un elenco a livello **SEGNALA** (non `MANCA`) con
`Duplicates Helper`, `DC Dashboard`, `DC Agents & Categories`,
`DC Timing & Quality`: senza quei fogli l'Omni Report resta valido e la macro
gira, quindi bloccare sarebbe sbagliato — ma un template che li ha persi va
saputo.

### 3.12 `tools/audit_workbook.py` — due bug da correggere PRIMA di rigenerare la fixture

**Bug 1, e morde adesso.** Il regex delle celle è
`<c r="([A-Z]+\d+)"([^>]*)>(.*?)</c>`: non gestisce le celle **auto-chiuse**
(`<c r="A14" s="396"/>`, cioè formattate e vuote). Su una cella così il match
non finisce lì ma **si mangia il valore della cella successiva**. Finora non si
vedeva perché tutti i fogli dati hanno le intestazioni in riga 1 a partire da A
senza buchi. La riga 14 di `DUP_DATASET` ha `A14` e `C14` vuote-ma-formattate, e
il risultato è misurato:

```
$ audit_workbook.header_row('DUP_DATASET', 14)
  B -> "(senza intestazione)"      # in realtà 'Full Name'
  D -> "(senza intestazione)"      # in realtà 'Case Number'
```

Cioè la fixture nascerebbe con le intestazioni **spostate di una colonna**, e
`test_ogni_campo_aggancia_la_colonna_giusta` verificherebbe una mappatura falsa.
Va corretto (`(?:/>|>(.*?)</c>)`) e va lasciato un test di regressione: un foglio
con una cella vuota-formattata prima di una piena.

**Bug 2, cosmetico ma confondente.** `Workbook.sheets` non fa unescape dei nomi:
i fogli si chiamano `DC Agents &amp; Categories`. `templatescan._sheet_targets`
invece fa unescape. Due comportamenti diversi per la stessa cosa; allineare al
secondo.

**Poi:**

- `DEFAULT_DATASETS += "DUP_DATASET"`;
- `header_row(sheet, row=1)` deve poter usare una riga diversa per foglio: nuova
  opzione `--header-row DUP_DATASET=14`, o meglio la riga letta dal contratto.

### 3.13 `tests/fixtures/workbook_template.json`

Da rigenerare **dopo** §3.12, dal template promosso:

```
python tools/audit_workbook.py template/Omni_Report_TEMPLATE.xlsm \
    --headers --usage --tables --json > tests/fixtures/workbook_template.json
```

`test_ogni_colonna_consumata_da_formule_e_nel_contratto` va nella direzione
"consumata ⇒ nel contratto", quindi `C`, `P`, `Q` (non consumate) non lo fanno
fallire, e `B`, `D`..`O` sono tutte nel contratto: passa. `IGNORE_USAGE` resta
vuoto.

---

## 4. Lavoro nel template (in Excel)

Quattro interventi. Il primo e il secondo sono **necessari**, il terzo e il
quarto sono le migliorie decise in D4 e D2.

### 4.1 Promuovere il WIP a template tracciato — *necessario*

`config/settings.yml` punta a `template/Omni_Report_TEMPLATE.xlsm`, che non ha i
cinque fogli nuovi. Il WIP diventa il template:

```
git mv template/Omni_Report_TEMPLATE_WIP.xlsm template/Omni_Report_TEMPLATE.xlsm
```

(verificato in §0 che non si perde nulla). Da valutare a parte se tenere in git i
due backup `_prima_dei_limiti` e `_prima_della_patch`, 13 MB in due file che il
programma non usa.

### 4.2 Togliere il parsing del testo da `Duplicates Helper` — *necessario* (D5)

Due celle, poi si trascinano fino a riga **1001**:

| cella | prima | dopo |
|---|---|---|
| `M2` | `IF($K2="","",DATE(VALUE(MID(K2,FIND(…` | `=IF($K2="","",$K2)` |
| `N2` | `IF($L2="","",DATE(VALUE(MID(L2,FIND(…` | `=IF($L2="","",$L2)` |

I formati numerici di `M`/`N` mostrano già data e ora, quindi **l'aspetto del
foglio non cambia**. Tutto quello che sta a valle continua a funzionare senza
modifiche: `O` (`INT(M2)`), `P` (`WEEKDAY`), `Q` (`HOUR`), `S` (`(N2-M2)*24`),
`T` (i bucket), e `DC Dashboard!AF21` (`MAX(O2:O1001)`).

Facoltativo, per pulizia: rinominare `K1`/`L1` da `Opened (text)` / `Closed (text)`
a `Opened` / `Closed`, che dopo questa modifica è quello che sono.

### 4.3 Allargare le capienze dei fogli DC — *miglioria decisa* (D4)

Trascinare le formule più in basso. Nessun intervallo di ranking da riallargare:
tutte queste formule leggono già `'Duplicates Helper'!$X$2:$X$1001`, che è
l'intervallo pieno.

**`DC Agents & Categories`**

| colonne | oggi | portare a | perché |
|---|---|---|---|
| `B`,`C`,`D` | 40 | **80** | agenti: 28 su 34, l'82%. Il team è ~36 persone |
| `G`,`H` | 18 | **40** | record type |
| `K`,`L` | 32 | **60** | `Type`: 19 su 26, il 73% |

**`DC Timing & Quality`**

| colonne | oggi | portare a | perché |
|---|---|---|---|
| `G`,`H`,`I` | 36 | **80** | parent con più di un duplicato: 9 su 20 |

**`DC Dashboard`** (aree di appoggio dei grafici, colonne nascoste)

| colonne | oggi | portare a |
|---|---|---|
| `AB` | 65 | **100** |
| `BB` | 40 | **60** |
| `BE` | 60 | **100** |

Attenzione a **`AA5`**: lì il limite non è una formula da trascinare, è lo spill
della lista `Case Origin` che va da riga 5 e sbatte su `AA20` (`Agent`) — 15
posti, 8 usati. Se un giorno i `Case Origin` distinti superassero 15, quella
cella diventa `#SPILL!` e il filtro Origin del dashboard non funziona più. Il
rimedio è spostare il blocco `AA20:AB…` più in basso (per esempio da riga 30),
**aggiornando anche le serie del grafico 1** (`'DC Dashboard'!$AA$21:$AA$30` e
`$AB$21:$AB$30`). Non urgente al 53% di riempimento — ma va saputo, e il
controllo di §3.7 lo dirà prima che morda.

Le sei serie dei grafici sono finestre **top-N** volute (`AA21:AB30` = top 10
agenti, `AG21:AH34` = 14 giorni, `AM21:AN24` = top 3 origin + `Other origins`,
`AP21:AQ26` = top 5 record type + `Other record types`, `AS21:AT30` = top 10
`Type`, `AD21:AE25` = i 5 bucket): allargare le colonne di appoggio **non** le
cambia, ed è giusto così.

### 4.4 Spezzare il collegamento esterno — *miglioria*

Il pacchetto contiene `xl/externalLinks/externalLink1.xml`, che punta a

```
L:\Admin\Individuali\Leonardo\Higiene Reports\Duplicates\Duplicates Report.xlsx
```

con i fogli `SF_DOWNLOAD`, `Data`, `Dashboard`, `Agents & Categories`,
`Timing & Quality` — cioè il workbook da cui i fogli nuovi sono stati copiati.

**Nessuna formula lo usa** (verificato: zero riferimenti `[1]` in tutto il
workbook), quindi non c'è niente da riparare: è un residuo. Ma sopravvive nel
file, e chi apre il report da un PC che non vede l'unità `L:` si prende il
dialogo «Questa cartella di lavoro contiene collegamenti a origini dati esterne».
In automazione non blocca (`app.display_alerts = False`), a mano sì.

In Excel: **Dati → Modifica collegamenti → Interrompi collegamento**. Verifica
dopo: `xl/externalLinks/` non deve più esistere nello zip.

### 4.5 Banner "nessun duplicato" — *facoltativo* (D2)

Se si vuole che la settimana senza export duplicati sia leggibile anche a occhio
e non solo nel preflight, tre celle (una per foglio DC), per esempio in `A2`:

```
=IF('Duplicates Helper'!$W$1=0,"Export duplicati assente per questa settimana","")
```

Non sostituisce il lavoro di §3.6 — le 56 celle `#DIV/0!` restano —, ma dice a
chi guarda perché sono lì.

---

## 5. Test da aggiungere

| test | cosa inchioda |
|---|---|
| `test_dup_contract` | il contratto carica `DUP_DATASET`, target e dtype giusti, `optional` ammesso (nessun consumer `VBA:`) |
| `test_header_row_14` | `find_header_row` trova la 14 su un foglio SF formattato sintetico (`tests/xlsxbuild.py`) |
| `test_righe_di_coda_scartate` | `Total`/`Count` non entrano nel blocco, e il preflight dice quante ne ha scartate |
| `test_preambolo_riscritto` | righe 1–13 dal download; preambolo più lungo di 13 → avviso |
| `test_dup_assente` | fonte mancante → blocco vuoto, foglio pulito, `#DIV/0!` dei fogli DC classificati come attesi, `BuildResult` non dichiarato fallito |
| `test_capienze_dc` | SEGNALA all'80%, BLOCCA oltre il 100%, capienze lette dal template |
| `test_limite_dup_dataset` | la scansione a cella singola trova 1014, e **non** trova 130 su `PSAT_DATASET` |
| `test_settimana_duplicati` | settimana uguale → OK; W32 contro W33 → **BLOCCA** con entrambi i periodi nel messaggio; un caso chiuso a cavallo della mezzanotte del lunedì **non** sposta la settimana (moda, non min/max) |
| `test_audit_celle_autochiuse` | regressione del bug §3.12: una cella vuota-formattata non ruba il valore alla successiva |
| `test_offset_helper_duplicates` | **sul template reale**: `Duplicates Helper!A2` referenzia `DUP_DATASET!B15` e il contratto dice `header_row: 14`. Se qualcuno inserisce una riga nel preambolo, questo test lo grida |

L'ultimo è il più importante dei nove: è l'unico guasto di questa sezione che
sposterebbe **tutti** i dati di una riga senza produrre un solo errore.

---

## 6. Documentazione

- `MANUALE.md`: «i **sei** file che scarichi ogni settimana» → sette (righe 3,
  49, 122, 237); la tabella dei file di input al §2 prende la riga dei duplicati
  con la nota **opzionale**; un capitolo nuovo sulla sezione Duplicate Cases (che
  cosa mostra, che cosa vuol dire "parent case", perché una settimana può uscire
  vuota);
- `docs/architettura.md`: mappa dei fogli e la catena
  `DUP_DATASET → Duplicates Helper → 3 fogli DC`;
- `docs/estendere.md`: questo dataset è il **primo con il preambolo e con la
  lettura cella per cella**, cioè l'esempio pratico più utile che il documento
  possa avere;
- `docs/template-modifiche-wip.md`: chiuderlo rimandando qui, o aggiungere un
  §6 con i quattro interventi di §4 e le spunte.

---

## 7. Ordine di lavoro

Le fasi sono pensate per essere **verificabili una per una** senza Excel: fino
alla 4 tutto gira con `pytest` e con `omni-report check`/`preflight`.

| fase | contenuto | verifica |
|---|---|---|
| **1** ✅ | §3.12 (bug di `audit_workbook`) + test di regressione | `pytest` |
| **2** ✅ | §4.1 promozione del template, §3.13 fixture rigenerata | i 584 test passano ancora |
| **3** ✅ | §3.1 contratto, §3.2 settings, §3.3 header 14, §3.4 righe di coda | `omni-report contratto`, preflight su un export vero |
| **4** ✅ | §3.9 limite a cella singola, §3.7 capienze, §3.8 nomi, §3.10 settimana | `pytest`, e il preflight che stampa 28/34 agenti |
| **5** ✅ | §3.5 preambolo, §3.6 fogli dipendenti + assenza dichiarata | `pytest` (il pezzo xlwings resta non eseguibile qui) |
| **6** | §4.2 template: `M2`/`N2` | build vero su Windows |
| **7** | §4.3, §4.4, §4.5 template: capienze, collegamento, banner | build vero + `scan_error_cells` a zero |
| **8** ✅ | §6 documentazione | rilettura |

La fase 6 va **dopo** la 3: se il contratto scrive date vere prima che il
template smetta di parsare il testo, `Duplicates Helper` esce a `#VALUE!`. Se
invece si fa il template prima del contratto, il primo build dopo la modifica
mostra le date del paste vecchio ma nient'altro si rompe — quindi in caso di
dubbio è quest'ordine il più sicuro.

---

## 8. Nessuna domanda aperta

La sola che c'era — se i duplicati coprano la stessa settimana del resto del
report o quella prima — è chiusa: **stessa settimana** (D6). Lo sfasamento
W32/W33 visibile nel template è un artefatto del montaggio, non del processo:
quello della W32 era il solo export duplicati disponibile in quel momento.
Conseguenza, in §3.10: il controllo **BLOCCA**.

Il primo build vero è anche la prima occasione di verificarlo, e va guardata:
se quel giorno il controllo blocca, l'ipotesi qui sopra era sbagliata e va
riaperta — non aggirata alzando la soglia.

---

## 9. Rigenerare la W33 — runbook

### Il controllo che decide quale strada prendere

**Se si promuove il template e si lancia il build di oggi, la W33 esce con i
duplicati della W32 dentro, e nulla lo dice.**

`build()` scrive solo i fogli che il contratto conosce
(`for name, dataset in contract.datasets.items()`). `DUP_DATASET` non è nel
contratto → `write_block` non viene mai chiamato su quel foglio → il paste del
3–9 agosto resta dov'è, e i tre fogli DC calcolano 178 duplicati della **W32**
dentro `Omni_Report_W33.xlsm`. `scan_error_cells` trova zero celle di errore,
ed è vero: non c'è nessun errore. Ci sono dati della settimana sbagliata.

È lo stesso difetto delle 130 righe residue di `PSAT_DATASET` — quello per cui
`optional` scrive il foglio **vuoto** invece di saltarlo.

### Stato di partenza, misurato

| cosa | valore | conseguenza |
|---|---|---|
| `data/aht_history.csv` | W22 … **W32** (546 righe) | la W33 **non è mai stata chiusa** dalla pipeline: questo è il primo giro vero, non una ri-generazione |
| `SF_DATABASE` nel template | Date Viewpoint = **20/07/2026** (W30) | residuo, sovrascritto dal build |
| `DUP_DATASET` nel template | 3–9 agosto (**W32**), 178 righe | residuo, **non** sovrascritto finché non c'è il contratto |
| template tracciato | 24 fogli | il WIP ne ha 29 |

### Strada B — la W33 subito, rischio zero

Se la W33 serve adesso e la sezione duplicati può aspettare:

1. **non** promuovere il template: si usa quello tracciato, 24 fogli, senza i
   fogli DC;
2. sei file in `input/`, `preflight --week 33`, `build --week 33`.

Si ottiene la W33 identica a come l'avrebbe prodotta prima di questo lavoro. La
sezione duplicati si collauda sulla **W34**, con calma e su dati che non servono
a nessuno nell'immediato. Nessuna riga di codice, nessuna modifica al template.

### Strada A — la W33 completa, duplicati compresi

Serve prima:

- **fasi 1–5** del §7 (solo codice; tutte verificabili qui con `pytest` e
  `preflight`, nessuna richiede Excel);
- le due modifiche al template **necessarie**: §4.1 (promozione) e §4.2 (`M2`,
  `N2` di `Duplicates Helper`). In quest'ordine: prima il contratto, poi il
  template — §7 spiega perché;
- il report duplicati **riscaricato con l'intervallo della W33**:
  `Date/Time Closed` da **10/08/2026** a **16/08/2026 23:59**. Con il file della
  W32 il build **BLOCCA** (D6), ed è quello che deve fare.

Le §4.3 (capienze), §4.4 (collegamento esterno) e §4.5 (banner) **non**
impediscono il build: si possono fare dopo. Ma la §4.3 va fatta presto — gli
agenti sono all'82%.

### I passi del run, in entrambe le strade

1. **`input/` deve contenere solo i file della W33.** I pattern devono
   corrispondere a **un solo** file ciascuno: se resta lì `SF DATABASE W32.csv`,
   `input_path` solleva «2 file corrispondono a `SF DATABASE*`, non so quale
   usare». È voluto.

   | pattern | file |
   |---|---|
   | `AT DATASET*` | `AT DATASET W33.xlsx` |
   | `ATwi DATASET*` | `ATwi DATASET W33.xlsx` |
   | `SF DATABASE*` | `SF DATABASE W33.csv` |
   | `PSAT DATASET*` | `PSAT DATASET W33.csv` (opzionale) |
   | `Turni*` | il roster |
   | `Back*Office*` | il back office |
   | `*Dup*Cases*` | l'export duplicati — **solo strada A** |

2. **`omni-report check --no-excel`** — o senza `--no-excel` sulla macchina con
   Excel. Deve dire «fogli attesi OK» e «fonti in input/ 6/6» (7/7 in strada A).

3. **`omni-report preflight --week 33`** → `output/preflight_W33.txt`. Da
   leggere, non da saltare: è il posto in cui si vede la settimana ricavata dai
   dati (`Settimana dai dati: W33 2026`), quali colonne sono state agganciate a
   quali lettere, e i case type nuovi. Non apre Excel: si può lanciare appena
   arrivano gli export.

4. **`omni-report build --week 33`** → `output/Omni_Report_W33.xlsm`. Richiede
   Excel desktop e `pip install -e ".[excel]"` (`check` dice che xlwings manca).
   Oppure `run_report.bat`, che fa 3 e 4 di fila chiedendo la settimana.

5. **Committare `data/aht_history.csv`.** Il build vi aggiunge la W33 (~49
   righe): è l'unica cosa del progetto che non si ricostruisce rilanciando il
   programma. `unisci` è idempotente, quindi ri-lanciare la W33 due volte
   riscrive le stesse righe — ma se il file non viene committato, la settimana è
   persa.

### Cosa guardare nel primo report W33

- `AHT Trend WoW`: con W22…W33 in storico la finestra mostra **W23…W33** — la
  W22 esce, ed è giusto (`TAKE(...,11)`);
- `preflight_W33.txt`, riga `Settimana dai dati`: deve dire W33;
- in strada A: `DC Dashboard!B9` (numero di duplicati) e la riga 3 di
  `DUP_DATASET` (`As of ...`) devono parlare della **W33**. Se la riga 3 dice
  ancora agosto 13, il preambolo non è stato riscritto — §3.5 non funziona;
- celle di errore: `BuildResult.error_cells` deve essere vuoto. Se compare
  `#SPILL!` su `DC Agents & Categories` o `DC Dashboard`, è una capienza finita:
  §4.3.
