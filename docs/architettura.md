# Architettura — decisioni e loro motivo

Cosa e' stato deciso mettendo in pratica i due piani
(`piano-omni-report.md`, `contesto-wow-aht.md`), e perche'. Dove si e' deviato
dai piani, e' detto e argomentato.

---

## 1. Un repo, due report, un layer di ingestione in comune

I due documenti descrivono due progetti separati, ognuno col suo albero
(`omni-report/`, `wow-aht/`). Qui sono un solo pacchetto con tre parti:

```
src/fasterreports/
  core/    ingestione: lettura CSV, matching header, coercizioni, preflight
  omni/    Omni Report: writer/orchestrate/CLI via xlwings
  wow/     WOW AHT Trend (da fare)
```

Il motivo e' scritto nel contesto WOW §7, e vale letteralmente:

> Una volta che `cases` e' una tabella, l'Omni Report e il WOW AHT pescano dalla
> stessa fonte: sono entrambi aggregazioni di dati caso per caso, **con lo stesso
> problema di colonne che cambiano posizione tra un export e l'altro**. Un solo
> ingest, molti report.

Il problema anti-fragilita' e' identico, e i due report leggono export
Salesforce con colonne in comune (`Case Origin (group)`, `Case Type`,
`Case AHT (mins)`). Scrivere due matcher significherebbe correggere due volte
ogni bug e vederli divergere. Quindi `core` non conosce Excel, non importa
xlwings e si testa su qualunque macchina; solo `omni` e (in futuro) `wow`
sanno di Excel.

E' anche il backlog #7 del WOW ("skill condivisa di ingest multi-report")
realizzato subito, perche' costa meno farlo ora che unificare dopo.

## 2. Il contratto e' l'unica verita', ed e' verificabile

`config/columns.yml` dice per ogni campo: nome canonico, lettera di destinazione,
alias, tipo, **e chi lo legge nel workbook** (`consumers`).

Il campo `consumers` non e' documentazione decorativa: `tests/test_contract_reale.py`
scansiona le formule del workbook vero e verifica che ogni colonna consumata sia
coperta dal contratto. E' l'auto-verifica del piano §11, e serve — nel W30 ha
trovato **tre colonne che il piano non aveva**
(`docs/audit-workbook-W30.md` §1).

## 3. Priorita' di match: prima il nome grezzo, poi il normalizzato

Il piano §4 mette come primo criterio il "nome canonico esatto (dopo
normalizzazione)". Nel workbook vero questo non basta: in `PSAT_DATASET`
esistono sia `Agent Name` (col. I) sia `agent_name` (col. BO), che normalizzano
sulla stessa stringa. Partendo dal normalizzato, ogni run sarebbe **ambiguo su
un file perfettamente valido**.

Quindi quattro livelli, in ordine:

1. nome canonico == header grezzo (a meno di BOM e spazi ai bordi)
2. nome canonico == header normalizzato
3. alias == header grezzo
4. alias == header normalizzato

e in aggiunta un `match: exact` per campo, che disattiva i livelli 2–4. Serve
dove ripiegare sarebbe peggio che fermarsi: `Agent Name` e `psat_score` hanno
entrambi un omonimo con semantica diversa, e agganciarlo darebbe numeri
plausibili e sbagliati. Dettagli e misure in `audit-workbook-W30.md` §6.

## 4. Niente pandas

Il piano lo elenca fra le dipendenze. Non e' stato usato:

- l'unico consumatore a valle e' xlwings, che vuole una lista 2D di valori Python:
  un DataFrame va comunque disfatto;
- l'inferenza di tipo di pandas e' l'opposto di quello che serve. Qui la
  coercizione e' esplicita e **fail-loud**: se il 90% di `Case AHT (mins)` non e'
  numerico, la colonna agganciata e' quella sbagliata e il run deve fermarsi, non
  ritrovarsi una colonna `object`;
- su 130k righe evita un giro di copie.

Le dipendenze obbligatorie sono `pyyaml` e nient'altro. `xlwings` serve solo per
scrivere (`.[excel]`), `openpyxl`/`oletools` solo per gli audit offline
(`.[audit]`). Se in futuro servisse pandas per analisi, si aggiunge a valle senza
toccare il core.

## 5. Niente openpyxl in scrittura, mai

Vale per entrambi i report, ed e' misurato (contesto WOW §5): openpyxl in
scrittura riscrive le formule ad array dinamico perdendo i valori in cache — cioe'
corrompe il motore — e sul WOW da 60 MB va in out-of-memory.

`tools/audit_workbook.py` non usa nemmeno openpyxl in lettura: legge l'XML dentro
lo zip. Costa poche righe in piu' e gira su qualunque file, incluso quello da
60 MB, senza dipendenze.

## 6. La patch al VBA (da fare, una volta sola, sul template)

Il modulo `CreaMalpractice` ha due `MsgBox`: uno a fine esecuzione, uno nel
gestore d'errore. In automazione bloccano il processo a tempo indeterminato.
Serve la modifica minima del piano §12. Nel modulo:

```vb
Public SilentMode As Boolean

Public Sub SetSilentMode(ByVal value As Boolean)
    SilentMode = value
End Sub
```

e i due `MsgBox` diventano:

```vb
If Not SilentMode Then MsgBox "Dettaglio Malpractice rigenerato: " & outRows.Count & " righe.", vbInformation
...
If Not SilentMode Then MsgBox "Errore durante la rigenerazione: " & Err.Description, vbExclamation
```

Attenzione al secondo: silenziare l'errore non basta, va anche **propagato**,
altrimenti la macro fallisce senza che la pipeline lo sappia. Nel ramo
`CleanFail`, dopo il messaggio:

```vb
If SilentMode Then Err.Raise Err.Number, , Err.Description
```

`orchestrate._run_macro` chiama `SetSilentMode(True)` prima della macro e, se il
template non espone quella Sub, si ferma con un errore parlante invece di
appendersi su un dialogo invisibile.

## 6-bis. `Turni` e `Slot Only Cases`: sei fonti, non quattro

Nella prima passata li avevo messi fra la "config mantenuta a mano". Sbagliato:
sono dati settimanali come i 4 CSV. Ora sono nel contratto, e i loro guasti nel
workbook reale sono documentati in `audit-workbook-W30.md` §8.

Le loro sorgenti non sono tabelle: sono **matrici larghe** (un agente per riga,
una colonna per giorno). Servono quindi due adattatori, in
`core/wfmsource.py`, che le riducono a righe tidy e restituiscono la stessa
coppia `(headers, rows)` di `csvsource.read_csv`. Da lì in poi la pipeline non
cambia: matcher, coercizioni, preflight e writer sono gli stessi.

Quale lettore usare sta nel contratto (`reader: csv | wfm_roster |
wfm_backoffice`), non nel codice. È la stessa astrazione prevista per il futuro
SQL (§7).

**Semplificazione rispetto al piano.** Il piano proponeva `role: computed` e
`dtype: time` per far calcolare al contratto la chiave e le frazioni di giorno.
Non servono: gli adattatori emettono già i valori finali (float e datetime), che
il core gestisce. Meno macchinari, stesso risultato — e la logica di
trasformazione sta dove è specifica, nell'adattatore, invece di diventare un
meccanismo generico usato una volta sola.

**Ordine di lettura.** I dataset non sono indipendenti:
`AT_DATASET` → da lui si ricava la settimana a cui ritagliare le sorgenti (che
coprono mesi) → il roster → il back office, che si filtra sugli agenti del
roster. `orchestrate._dataset_order` lo impone, e `--only` aggiunge le
dipendenze implicite invece di lasciar fallire il run.

**Una sola verità per la settimana.** Non si scrive in due posti: si ricava
dall'intervallo di `AT_DATASET!Start Time`. `sources.monday_serial` esiste solo
come scavalco.

**Una sola verità per gli alias nomi.** La tabella sta in
`Helper Malpractice!D:E`, dov'è anche il VBA che la usa: la pipeline legge
quella, non una copia in config. Se il template non c'è ancora, il preflight lo
dice — senza gli alias quattro agenti risulterebbero orfani per un motivo falso.

## 7. Sorgente SQL

`settings.yml` ha `source: csv|sql`; `sql` oggi rifiuta con un messaggio che dice
cosa manca. Quando servira' bastera' un `sqlsource.py` che restituisca la stessa
coppia (header, righe) di `csvsource.py`: matcher, transform, writer e
orchestrate non cambiano, e gli alias SQL si aggiungono al contratto. E' il
piano §13, e la forma del core e' stata scelta per renderlo vero.

## 8. Cosa e' provato e cosa non lo e'

Distinzione importante, nello spirito della sezione *Provenienza* del contesto WOW.

**Provato, con test che girano** (296 test, nessuna dipendenza da Excel):

- normalizzazione e priorita' di match, incluse le collisioni reali del W30;
- resilienza: colonne mescolate e rinominate negli alias → output **identico**;
  colonna rimossa → preflight fallito con messaggio azionabile (piano §11 a/b/c);
- coercizioni, comprese quelle che devono **rifiutare** (`1 2`, `1.234,56`);
- caricamento e validazione del contratto;
- il contratto contro le intestazioni e le formule reali del W30;
- il preflight end-to-end su sei fonti, con CSV a 143 colonne in ordine stravolto;
- il parser dei turni su **tutte le 26 forme** del roster reale, piu' i casi che
  devono fallire;
- **il golden test dei due fogli WFM**: ricostruiti dalle sorgenti, 252 = 252
  righe per `Turni` e 259 = 259 per `Slot Only Cases`, chiavi identiche, con le
  sole tre differenze dichiarate in anticipo. E' il test piu' importante del
  repo: il risultato del processo manuale e' la specifica;
- i controlli di coerenza, uno per uno, con la distinzione fra blocco e
  segnalazione — che e' la sostanza del modulo.

**Scritto ma mai eseguito** — questo ambiente non ha Excel:

- `omni/writer.py` e la parte Excel di `omni/orchestrate.py`. Sono modellati sulla
  struttura misurata del W30, ma nessuna riga e' passata da un Excel vero.

Checklist per la prima esecuzione reale:

1. `pip install -e '.[excel,dev]'` su una macchina con Excel desktop.
2. Preparare `template/Omni_Report_TEMPLATE.xlsm`: copia di una settimana chiusa
   con i 4 fogli DATASET svuotati (formule `P`/`Q` e VBA **intatti**), piu' la
   patch `SilentMode` del §6.
3. `omni-report preflight --week NN` sui CSV veri. Deve passare **prima** di
   provare `build`.
4. `omni-report build --week NN --visible`, guardando cosa fa. Da controllare in
   particolare: il ridimensionamento di `AHT_Data`, che `Dettaglio Malpractice`
   si popoli, che nessun dialogo compaia.
5. Confronto con la settimana chiusa: e' il golden test del piano §11 (M3),
   ancora da scrivere perche' richiede un run reale da cui partire.

Il punto 5 e' il vero collaudo. Fino a quel confronto, nessun numero prodotto
dalla pipeline va considerato buono.

## 9. WOW AHT: perche' non e' stato iniziato

Non e' una dimenticanza. Il contesto WOW §10 mette al primo posto del backlog
*"validare `wow-aht-heatmap` in reale"*, e la skill esiste gia' con la parte
xlwings **non verificata** perche' l'ambiente di sviluppo non aveva Excel.
Questo ambiente non lo ha nemmeno: scrivere altro codice xlwings non verificabile
sopra a codice xlwings non verificato aumenterebbe solo la quantita' di cose da
collaudare in una volta.

Il passo che si puo' fare senza Excel e' `ingest.py` → warehouse (backlog #2), e
il contesto §8 e' chiaro sull'ordine: estrarre i 26 `DATASET_Wnn`, **validare la
semantica del pivot** contro le heatmap esistenti, e solo dopo alleggerire. Il
test di regressione esiste gia' (26 settimane di risultati congelati). E' il
prossimo lavoro sensato, e usera' `fasterreports.core` per il matching.

Dettagli in `src/fasterreports/wow/README.md`.
