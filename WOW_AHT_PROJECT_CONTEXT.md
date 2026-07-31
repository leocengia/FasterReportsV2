# WOW AHT Trend by CT — contesto di progetto

Documento di consegna per il progetto Claude Code. Raccoglie tutto ciò che è stato
ricavato dall'ispezione diretta del workbook, i vincoli tecnici scoperti, la skill già
costruita, l'architettura proposta e le domande ancora aperte.

**Come leggerlo:** la sezione *Provenienza* in fondo distingue ciò che è stato **misurato**
da ciò che è **ipotesi**. Non trattare le ipotesi come fatti: alcune vanno confermate da
Leonardo, altre si validano da sé eseguendo il confronto descritto in *Piano di
migrazione*.

Rilevamento effettuato su `WOW_AHT_Trend_by_CT_-_Backup.xlsx`, 59 fogli, settimane W5–W30.

---

## 1. Cos'è il file e come scorrono i dati

```
DATASET_Wnn  (nascosto, 5–23 MB, ~143 colonne, una riga per case)
     |  pivot table, filtro rapporto su 'Case Origin (group)'
     v
Pivot_nn     A = case_type | B = Media di Case AHT (mins) | C = Volume
     |  XLOOKUP + INDIRECT — riga 1 della destinazione contiene il NOME del foglio pivot
     v
WOW AHT Trend v2   (vista live: mostra un solo canale alla volta)
     |  <<< COPIA MANUALE DI VALORI — il collo di bottiglia settimanale
     v
Heat Map PHONE  /  Heat Map NON-LIVE   (valori statici, zero formule)
```

### Fogli

| Foglio | Ruolo |
|---|---|
| `DATASET_Wnn` (×24 visibili come tali) | dati grezzi settimanali, nascosti |
| `Pivot_nn` (×26, W5–W30) | pivot della settimana, con filtro rapporto sul canale |
| `WOW AHT Trend v2` | tabella grezza, 2028 formule array |
| `Heat Map PHONE` | heatmap canale telefonico, valori statici |
| `Heat Map NON-LIVE` | heatmap canale non-live, valori statici |
| `tempTables` | nessuna formula — **ruolo non indagato** |
| `NEW! HM Selection`, `NEW! HM Selection (sorted)`, `NEW! HM Selection (export)` | nessuna formula — **ruolo non indagato** |

Nota: i pivot esistono per W5–W30 (26), i fogli `DATASET_Wnn` presenti sono 26 pure, ma
alcuni dataset compaiono in ordine non sequenziale nel `workbook.xml` (W23–W26 sono fuori
sequenza). Non è un problema funzionale, ma spiega perché l'ordine delle schede non
corrisponde all'ordine cronologico.

---

## 2. Il meccanismo centrale: il filtro sul canale

`Case Origin (group)` è un **filtro rapporto** (page field) con esattamente due valori:

- `Phone` → alimenta `Heat Map PHONE`
- `Other` → alimenta `Heat Map NON-LIVE`

Un filtro rapporto accetta **un solo valore per volta**. Questa è la ragione strutturale
per cui le heatmap non possono essere formule vive: la tabella grezza è una vista che
mostra un canale alla volta, quindi le heatmap devono essere istantanee statiche. Non è
pigrizia di chi l'ha costruito, è una conseguenza del design.

Corollario operativo: qualunque automazione che legga la tabella grezza **deve** commutare
il filtro e ricalcolare, oppure aggirare del tutto i pivot (vedi §7).

---

## 3. `WOW AHT Trend v2` — struttura esatta

| Riga | Contenuto |
|---|---|
| 1 | nome del foglio pivot, ripetuto su due colonne (`Pivot_5` in B e C … `Pivot_30` in AZ e BA) |
| 2 | etichetta settimana (`W5`…`W30`), su due colonne |
| 3 | `avg_aht` / `volume` alternati |
| 4 | `A4 = 'case_type'`, poi l'etichetta settimana ogni due colonne |
| 5 | `avg_aht` / `volume` alternati (intestazione visiva) |
| 6–43 | i 38 `case_type`, colonna A = nome |
| 44 | `totals` |

Colonne B…BA = 26 settimane × 2 metriche. `AZ`/`BA` = W30.

Formule, verbatim da `AZ6` e `BA6`:

```excel
=IFERROR(_xlfn.XLOOKUP($A6,INDIRECT(AZ$1&"!$A:$A"),INDIRECT(AZ$1&"!$B:$B"),""),"")
=IFERROR(_xlfn.XLOOKUP($A6,INDIRECT(BA$1&"!$A:$A"),INDIRECT(BA$1&"!$C:$C"),0),0)
```

Il default è `""` per l'AHT e `0` per il volume: un `case_type` assente dal pivot della
settimana risulta vuoto, non in errore. Comodo, ma vedi la trappola in §5 sulle righe fisse.

---

## 4. Heatmap — struttura esatta

Righe 2–39 = gli stessi 38 `case_type`, **nello stesso ordine** della tabella grezza.
Offset di 4: riga 6 del trend ↔ riga 2 della heatmap. Verificato identico su entrambe le
heatmap.

Due blocchi separati da una colonna vuota:

| Foglio | blocco `Wnn avg AHT` | blocco `Wnn volume` |
|---|---|---|
| `Heat Map PHONE` | B…AA = W5…W30 | AC…BB = W5…W30 |
| `Heat Map NON-LIVE` | B…AA = W5…W30 | AC…AY = **W8**…W30 |

**I due blocchi non partono dalla stessa settimana.** In NON-LIVE il volume comincia da W8
mentre l'AHT da W5: buco storico, non un errore da riempire. Da qui la regola che vale per
qualunque script su questo file: **cercare le colonne per nome di intestazione, mai per
offset fisso.** Con offset fissi si sbaglia foglio per foglio.

---

## 5. Vincoli e trappole (letti col sangue, non dedotti)

### Mai `openpyxl` in scrittura su questo workbook
- Va in **out-of-memory** caricando 60 MB in modalità scrittura (verificato: processo ucciso).
- Il salvataggio **distrugge le 26 pivot table** e le loro cache.
- Usare solo `read_only=True` per ispezionare.

### Mai `recalc.py` / LibreOffice
Le formule usano `_xlfn.XLOOKUP` in array. LibreOffice non le valuta: un recalc le
trasformerebbe in `#NAME?` permanenti in 2028 celle.

### `INDIRECT` rende le formule volatili
Dopo aver cambiato il filtro del pivot serve `CalculateFullRebuild()`. Un `calculate()`
semplice può non propagare il cambio.

### La formattazione condizionale non si estende oltre il bordo destro
La scala di colori di un blocco non copre automaticamente una colonna aggiunta alla sua
destra. La colonna di una settimana nuova va formattata copiando quella precedente,
altrimenti resta grigia — e in una heatmap è l'unica cosa che conta.

### L'elenco dei `case_type` è fisso a 38 righe
Un tipo nuovo che appare nel pivot viene **silenziosamente ignorato** da `XLOOKUP` (torna
`""`/`0`, non un errore). Va aggiunta la riga a mano in tutti e tre i fogli, mantenendo
l'ordine. Qualunque automazione dovrebbe **segnalare** i `case_type` presenti nel pivot ma
assenti dall'elenco.

---

## 6. Dove stanno i 60 MB

| Componente | XML non compresso | Quota |
|---|---|---|
| 24 fogli `DATASET_Wnn` nascosti | 287 MB | 60,9% |
| `pivotCache` — **gli stessi dati una seconda volta** | 180 MB | 38,2% |
| `sharedStrings` | 3 MB | 0,6% |
| Tutto il resto (trend, heatmap, selection) | **0,9 MB** | **0,2%** |
| Totale XML non compresso | 471 MB | |

Fogli più grossi: `DATASET_W11` 23,1 MB, `DATASET_W9` 21,9 MB, poi W20/W22/W24 ~13 MB.
W11 e W9 sono fuori scala rispetto agli altri: **da indagare** se contengono dati
duplicati o un intervallo di date più ampio.

**Lettura del dato:** il 99% del file è dati grezzi più una loro copia integrale nella
cache dei pivot. Ciò che si guarda davvero sta sotto il megabyte. Il workbook è un
database che si finge un report — ed è la premessa di tutta la sezione seguente.

---

## 7. Architettura proposta

I tre obiettivi di Leonardo — alleggerire il file, automatizzare meglio, far interagire
report diversi — **non sono tre progetti ma uno in ordine obbligato.** Spostando fuori da
Excel lo strato dati, gli altri due cadono da sé, ed Excel torna a fare solo presentazione.

```
wow-aht/
├── data/raw/              export settimanali, mai modificati
├── data/warehouse.duckdb  tabella cases: una riga per case, tutte le settimane
├── scripts/
│   ├── ingest.py          raw -> warehouse, matching per nome colonna
│   ├── aggregate.py       cases -> aht/volume per case_type × canale × settimana
│   └── refresh_heatmap.py
├── reports/               workbook di output, leggeri
└── .claude/skills/        le skill, versionate insieme al codice
```

### Perché questo sblocca l'interoperabilità

Una volta che `cases` è una tabella, l'Omni Report e il WOW AHT pescano dalla stessa
fonte: sono entrambi aggregazioni di dati caso per caso, con lo stesso problema di colonne
che cambiano posizione tra un export e l'altro. Un solo ingest, molti report.

Oggi ogni report si porta dietro la propria copia dei dati grezzi, ed è esattamente per
questo che non comunicano.

### Perché DuckDB (e non solo Parquet)

Parquet basta per archiviare, ma serve interrogare al volo (`GROUP BY case_type, canale,
settimana` su milioni di righe) e incrociare report diversi. DuckDB legge Parquet
direttamente e non richiede un server. Se si preferisce SQLite per uniformità con
RunwaySurfer, funziona, ma sarà più lento sulle aggregazioni analitiche.

---

## 8. Piano di migrazione — l'ordine conta

**Non svuotare il workbook come primo passo.** Lo storico della tabella grezza dipende dai
pivot, che dipendono dai dataset. Le heatmap invece sono già valori statici: quella storia
sopravvive comunque, qualunque cosa succeda ai dataset.

1. **Estrarre** i 26 `DATASET_Wnn` nel warehouse (streaming con `read_only=True` +
   `iter_rows`; non caricare tutto in memoria).
2. **Validare la semantica del pivot** — vedi sotto, è il passaggio critico.
3. **Solo allora** alleggerire: archiviare i dataset storici fuori dal workbook, tenendo
   nel file Excel le sole settimane ancora vive.
4. Ricostruire trend e heatmap alimentandoli dal warehouse.

### Il test di regressione esiste già

Il punto delicato è replicare **esattamente** quali righe il pivot include nella media. Nel
dataset ci sono almeno `standard_exclusions`, `Is_ClosedCase`, `Handle Time Outlier`,
`AHT due to Misroutes` — e dal file **non è deducibile** quali di questi la pivot applichi.

Ma ci sono 26 settimane di risultati già congelati nelle heatmap. Quindi:

> Ricalcola W5–W30 dal grezzo, confronta con le colonne esistenti delle heatmap. Se
> tornano, la semantica è giusta. Dove divergono, hai isolato l'esclusione che manca.

È un test di regressione già scritto. Usalo prima di fidarti di qualunque numero nuovo, e
non presumere le esclusioni: chiedile a Leonardo e verificale.

---

## 9. La skill già costruita: `wow-aht-heatmap`

Automatizza il passaggio manuale attuale, **senza** riprogettare nulla. Utile subito e
indipendente dalla migrazione.

```
wow-aht-heatmap/
├── SKILL.md
├── scripts/
│   ├── refresh_heatmap.py    xlwings, scrive
│   └── inspect_workbook.py   openpyxl read_only, audit senza Excel
└── references/workbook-structure.md
```

Per ciascun canale: imposta il filtro → `CalculateFullRebuild()` → legge le due colonne
della settimana dal trend (righe 6–43) → scrive i valori nella heatmap (righe 2–39),
trovando le colonne per nome di intestazione. Scrive **solo** la settimana nuova, lo
storico resta congelato.

Ha: `--dry-run`, `--no-create-columns`, `--overwrite`, verifica per rilettura, controllo di
allineamento delle righe, segnalazione dei `case_type` mancanti.

### Stato dei test

- **Verificato** contro le intestazioni reali del file: risoluzione delle colonne del trend
  (W30 → colonne 52/53, coerente con le formule lette in `AZ6`/`BA6`), risoluzione delle
  colonne heatmap (PHONE W30 volume → colonna 54), errore corretto per una settimana
  assente, `inspect_workbook.py` eseguito integralmente sul file da 60 MB.
- **NON verificato**: tutta la parte `xlwings` — l'ambiente di sviluppo non aveva Excel.
  In particolare l'impostazione di `CurrentPage`, il `CalculateFullRebuild`, e
  l'inserimento colonna con copia del formato.

**Prima esecuzione reale: fare una copia di backup del workbook**, poi `--dry-run`, e
controllare a occhio che la colonna nuova sia colorata come le altre.

---

## 10. Backlog — miglioramenti e skill candidate

Ordinate per rapporto valore/rischio.

| # | Intervento | Note |
|---|---|---|
| 1 | **Validare `wow-aht-heatmap` in reale** | sblocca tutto il resto; è già scritta |
| 2 | **`ingest.py` + warehouse** | primo passo della migrazione; include la validazione §8 |
| 3 | **Skill "chiusura settimana"** | orchestra: ingest → aggregate → refresh → verifica, un comando |
| 4 | **Gestione `case_type` nuovi** | rilevamento + inserimento riga nei 3 fogli mantenendo l'ordine |
| 5 | **Estendere il trend a una settimana nuova** | oggi manuale: aggiungere la coppia di colonne + ricopiare le formule. Da chiarire con Leonardo (vedi §11) |
| 6 | **Archiviazione dataset storici** | dopo la validazione; il grosso dell'alleggerimento |
| 7 | **Skill condivisa di ingest multi-report** | il punto di contatto con l'Omni Report |
| 8 | **Indagare `tempTables` e i 3 fogli `NEW! HM Selection`** | ruolo ignoto; potrebbero essere già export automatizzabili |
| 9 | **Capire `DATASET_W9` e `W11`** | 22–23 MB contro ~12 di media: duplicati? intervallo più ampio? |

### Nota per chi implementa la #3

L'orchestrazione va scritta come skill, non come script monolitico: i passaggi hanno modi
di fallire diversi e vanno riportati in modo distinguibile. Un `ingest` che salta una
colonna e un `refresh` che scrive nella colonna sbagliata richiedono reazioni opposte, e
un unico exit code li appiattisce.

---

## 11. Domande aperte per Leonardo

1. **Come arrivano gli export?** Un CSV a settimana da Tableau, o si estraggono dai
   `DATASET_Wnn` già nel file? Determina la forma di `ingest.py`.
2. **Chi estende la tabella grezza?** Leonardo dice che il pivot "alimenta la tabella
   grezza automaticamente", ma la riga 1 arriva esattamente a `Pivot_30`: per la W31 le due
   colonne nel trend **non esistono ancora**. Quindi le aggiunge a mano e considera il
   passaggio parte della creazione del pivot — da confermare, e da decidere se
   automatizzarlo (backlog #5).
3. **Quali esclusioni applica il pivot?** `standard_exclusions`, `Is_ClosedCase`,
   `Handle Time Outlier`, `AHT due to Misroutes`: quali sono attive? Sbloccabile anche da
   solo col test di regressione (§8), ma partire dalla sua risposta fa risparmiare tempo.
4. **`tempTables` e i tre `NEW! HM Selection` servono ancora?** Se sono residui, si
   eliminano; se sono output verso qualcun altro, vanno nel piano.
5. **Quanto storico deve restare nel file Excel?** Determina l'aggressività
   dell'archiviazione al passo 3 della migrazione.

---

## 12. Provenienza — cosa è misurato e cosa è ipotesi

**Misurato** (ispezione diretta del file, riproducibile):
struttura dei 59 fogli · le formule verbatim di `AZ6`/`BA6` · i due valori del filtro
`Case Origin (group)`, letti dalla `pivotCacheDefinition` · layout di `WOW AHT Trend v2`
(righe 1–44, colonne B–BA) · intestazioni e blocchi di entrambe le heatmap, incluso il
buco W5–W7 nel volume NON-LIVE · identità dell'ordine dei 38 `case_type` nei tre fogli ·
la ripartizione dei 60 MB · l'OOM di openpyxl in scrittura · layout di `Pivot_nn`
(A/B/C) · presenza nel dataset delle colonne `Case Origin (group)`, `Case Type`,
`Case AHT (mins)` e delle quattro colonne di esclusione citate.

**Ipotesi da confermare:**
la mappatura `Phone`→PHONE e `Other`→NON-LIVE (i nomi corrispondono e non ci sono altri
valori, ma non è stato osservato l'atto di copiare) · che il workflow settimanale sia
esattamente "congela/aggiungi colonna, commuta filtro, incolla valori ×2" · che
`tempTables` e i fogli Selection non siano coinvolti nel ciclo settimanale · che i pivot
usino la stessa configurazione di esclusioni in tutte le settimane.

**Non indagato:**
contenuto e scopo di `tempTables` e dei tre `NEW! HM Selection` · perché `DATASET_W9` e
`W11` siano molto più grandi · se esista una macro o un `Personal.xlsb` collegato (il file
consegnato è `.xlsx`, ma la genesi del workbook potrebbe includere VBA altrove) · se le
formattazioni condizionali delle heatmap siano scale di colori native o formule.

Il materiale di partenza includeva anche una registrazione dello schermo di 3'23'' del
workflow manuale, letta per fotogrammi via OCR. Da lì viene l'individuazione della formula
`INDIRECT`; la sequenza precisa dei copia-incolla **non** è stata ricostruita in modo
affidabile e resta tra le ipotesi.
