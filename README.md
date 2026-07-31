# FasterReportsV2

Automazione dei report settimanali. Due report, un layer di ingestione in comune.

- **Omni Report** — da 4 CSV a workbook finito, con un comando.
  Piano: `docs/piano-omni-report.md`.
- **WOW AHT Trend by CT** — non ancora iniziato, e il perche' e' spiegato:
  `src/fasterreports/wow/README.md`. Contesto: `docs/contesto-wow-aht.md`.

Il problema che risolve: sia le formule sia il VBA leggono le colonne **per
lettera fissa** (`SF_DATABASE!$BB`, `Cells(r, "DY")`). Se l'export sposta una
colonna, la lettera punta al dato sbagliato e i numeri sono falsi **senza alcun
errore**. Davanti al motore c'e' quindi un layer che aggancia le colonne **per
nome**, e che si ferma con un messaggio chiaro se un nome manca o e' ambiguo.

Il motore Excel (formule + VBA) non viene riscritto: resta nel workbook e viene
pilotato.

---

## Com'e' fatto

```
config/columns.yml     IL CONTRATTO: nome colonna -> lettera di destinazione
config/settings.yml    percorsi, nomi file, scelte di esecuzione
src/fasterreports/
  core/                ingestione condivisa — non sa niente di Excel
  omni/                Omni Report: writer, orchestrate, CLI (xlwings)
  wow/                 WOW AHT Trend (da fare)
tools/                 audit dei workbook senza aprire Excel
tests/                 133 test, nessuno richiede Excel
docs/                  i due piani + architettura + audit del W30
samples/               i workbook di riferimento (W30)
input/  output/        inbox CSV e prodotti (non versionati)
template/              Omni_Report_TEMPLATE.xlsm (da preparare, vedi sotto)
```

## Uso

```bash
pip install -e '.[dev]'                       # core + test, senza Excel
pip install -e '.[excel,audit,dev]'           # tutto (richiede Excel desktop)

omni-report contract                          # stampa il contratto colonne
omni-report preflight --week 31               # valida i CSV, NON apre Excel
omni-report build --week 31                   # il "pulsante"
```

I 4 CSV vanno in `input/` con i nomi indicati in `config/settings.yml`
(`AT.csv`, `ATwi.csv`, `SF.csv`, `PSAT.csv`). L'ordine delle colonne dentro i
file **non conta**.

Exit code: `0` fatto · `1` bloccato da un problema nei dati · `2` uso sbagliato.

### `preflight` prima di `build`, sempre

`preflight` gira in un secondo, non apre Excel e produce
`output/preflight_W{n}.txt`: per ogni campo, a quale colonna del CSV e' stato
agganciato e come. Viene scritto **anche quando tutto va bene**, e serve — fra sei
mesi, davanti a un numero strano, e' l'unico posto che dice a cosa era agganciata
quella formula.

```
campo canonico          | col. | colonna sorgente     | via          | note
------------------------+------+----------------------+--------------+------
Agent Email             | C    | agent_email          | normalizzato |
Wrap-up Time in seconds | I    | wrap_up_time_seconds | alias esatto |
Agent Name              | I    | Agent Name           | esatto       | disambiguata da match esatto; scartate: 'agent_name'
```

## Prima di poterlo usare davvero

Due cose mancano, ed entrambe richiedono una macchina con Excel desktop:

1. **`template/Omni_Report_TEMPLATE.xlsm`** — copia di una settimana chiusa con i
   4 fogli DATASET svuotati (formule `P`/`Q` e VBA intatti) e la patch
   `SilentMode` al VBA. Istruzioni: `docs/architettura.md` §6.
2. **Il collaudo del writer.** Ingestione e preflight hanno test che girano; la
   parte Excel (`omni/writer.py`, `omni/orchestrate.py`) e' scritta sulla
   struttura misurata del W30 ma **non e' mai stata eseguita**: qui non c'e'
   Excel. Checklist della prima esecuzione: `docs/architettura.md` §8.

Il collaudo vero e' il confronto con una settimana chiusa (golden test, piano
§11). Fino a quel confronto nessun numero prodotto dalla pipeline va considerato
buono.

## Audit dei workbook

Il contratto colonne si **ri-deriva dal file**, non si rilegge da un documento:

```bash
python tools/audit_workbook.py "samples/omni-report/Omni Report W30.xlsm" --all
python tools/extract_vba.py "samples/omni-report/Omni Report W30.xlsm" --columns
```

`audit_workbook.py` legge l'XML dentro lo zip: nessun Excel, nessun openpyxl,
quindi funziona anche sul WOW da 60 MB. `--usage` dice quali colonne dei dataset
sono consumate da quali formule — ed e' cosi' che si sono trovate **tre colonne
che il piano non aveva** (`docs/audit-workbook-W30.md`).

## Test

```bash
python -m pytest            # 133 test, ~0,2 s, nessuna dipendenza da Excel
```

Fra questi, i tre scenari del piano §11: colonne mescolate e rinominate negli
alias → output identico; colonna rimossa → blocco con messaggio azionabile.
E `tests/test_contract_reale.py` verifica il contratto contro le intestazioni e
le formule del workbook vero: se il workbook evolve, lo dice.

## Documenti

| File | Cosa contiene |
|---|---|
| `docs/piano-omni-report.md` | il piano originale dell'Omni Report |
| `docs/contesto-wow-aht.md` | il contesto originale del WOW AHT |
| `docs/architettura.md` | decisioni prese, deviazioni dai piani e loro motivo; cosa e' provato e cosa no |
| `docs/audit-workbook-W30.md` | misure sul workbook vero: contratto verificato, tre colonne mancanti, VBA, tabella `AHT_Data` |
