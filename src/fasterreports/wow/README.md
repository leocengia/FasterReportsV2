# WOW AHT Trend by CT — stato e prossimi passi

Contesto completo: `docs/contesto-wow-aht.md`. Qui c'e' solo cosa manca e in che
ordine.

## Perche' il package e' vuoto

Il backlog del contesto (§10) mette al primo posto **validare in reale la skill
`wow-aht-heatmap`**, che esiste gia' e ha la parte `xlwings` dichiarata *non
verificata* — l'ambiente in cui e' stata scritta non aveva Excel. Questo
ambiente non lo ha nemmeno.

Accumulare altro codice xlwings non verificabile sopra a codice xlwings non
verificato non fa avanzare il progetto: sposta solo il collaudo piu' in la' e lo
rende piu' grosso. Quindi qui non c'e' codice Excel.

## Ordine dei passi (dal contesto §8, e l'ordine conta)

1. **Validare `wow-aht-heatmap`** su una macchina con Excel. Prima una copia di
   backup del workbook, poi `--dry-run`, poi controllare a occhio che la colonna
   nuova sia colorata come le altre — la formattazione condizionale non si
   estende da sola oltre il bordo destro del blocco (§5).
2. **`ingest.py` → warehouse**: estrarre i 26 `DATASET_Wnn` in streaming
   (`read_only=True` + `iter_rows`, mai openpyxl in scrittura su questo file).
   **Il matching degli header non va riscritto**: si usa `fasterreports.core`,
   che fa esattamente questo per l'Omni Report ed e' testato.
3. **Validare la semantica del pivot** — il passaggio critico. Nel dataset ci
   sono almeno `standard_exclusions`, `Is_ClosedCase`, `Handle Time Outlier`,
   `AHT due to Misroutes`, e dal file **non e' deducibile** quali il pivot
   applichi. Ma le heatmap contengono 26 settimane di risultati congelati:
   ricalcolare W5–W30 dal grezzo e confrontare **e' il test di regressione, ed
   e' gia' scritto**. Dove diverge, si e' isolata l'esclusione mancante.
4. **Solo dopo**: alleggerire il workbook archiviando i dataset storici.

Non svuotare il workbook come primo passo: lo storico della tabella grezza
dipende dai pivot, che dipendono dai dataset.

## Vincoli da non riscoprire col sangue

Tutti misurati, dettagli nel contesto §5:

- **Mai openpyxl in scrittura** su questo file: out-of-memory a 60 MB e il
  salvataggio **distrugge le 26 pivot table** e le loro cache.
- **Mai LibreOffice** per il ricalcolo: non valuta `_xlfn.XLOOKUP` in array,
  le trasformerebbe in `#NAME?` permanenti in 2028 celle.
- `INDIRECT` rende le formule volatili: dopo aver cambiato il filtro del pivot
  serve `CalculateFullRebuild()`, non un `calculate()`.
- **Cercare le colonne per nome di intestazione, mai per offset fisso**: i due
  blocchi delle heatmap non partono dalla stessa settimana (in NON-LIVE il
  volume comincia da W8, l'AHT da W5). Con offset fissi si sbaglia foglio per
  foglio. E' lo stesso principio del contratto colonne dell'Omni Report.
- L'elenco dei `case_type` e' **fisso a 38 righe**: un tipo nuovo viene
  ignorato in silenzio da `XLOOKUP` (torna `""`/`0`, non un errore). Qualunque
  automazione deve **segnalarlo**.

## Punto di contatto con l'Omni Report

Entrambi i report leggono export Salesforce caso-per-caso con colonne in comune
(`Case Origin (group)`, `Case Type`, `Case AHT (mins)`) e hanno lo stesso
problema di colonne che si spostano. Quando si arrivera' al passo 2, il
contratto WOW va scritto come `config/columns.yml` dell'Omni: stesso
`fasterreports.core`, un file di contratto per report.

## Da chiarire con Leonardo

Le domande del contesto §11 sono ancora aperte, e la #1 e la #3 bloccano il
passo 2:

1. **Come arrivano gli export?** Un CSV a settimana da Tableau, o si estraggono
   dai `DATASET_Wnn` gia' nel file? Determina la forma di `ingest.py`.
2. Chi estende la tabella grezza alla settimana nuova?
3. **Quali esclusioni applica il pivot?** Sbloccabile anche col test di
   regressione del passo 3, ma partire dalla sua risposta fa risparmiare tempo.
4. `tempTables` e i tre `NEW! HM Selection` servono ancora?
5. Quanto storico deve restare nel file Excel?
