# FasterReportsV2

Automazione dei report settimanali. Due report, un layer di ingestione in comune.

- **Omni Report** — da sei fonti a workbook finito, con un comando.
  Piano: `docs/piano-omni-report.md`.
- **WOW AHT Trend by CT** — non ancora iniziato, e il perche' e' spiegato:
  `src/fasterreports/wow/README.md`. Contesto: `docs/contesto-wow-aht.md`.

Il problema che risolve: sia le formule sia il VBA leggono le colonne **per
lettera fissa** (`SF_DATABASE!$BB`, `Cells(r, "DY")`). Se l'export sposta una
colonna, la lettera punta al dato sbagliato e i numeri sono falsi **senza alcun
errore**. Davanti al motore c'e' quindi un layer che aggancia le colonne **per
nome**, e che si ferma con un messaggio chiaro se un nome manca o e' ambiguo.

Sei fonti a settimana: 4 CSV piu' due export WFM in forma di matrice larga
(`Turni`, `Slot Only Cases`), che vengono riportati in forma lunga. Oltre a
risparmiare il copia-incolla, i **controlli di coerenza fra le fonti** fanno
emergere guasti che oggi non si vedono guardando il report — vedi
`docs/audit-workbook-W30.md` §8.4.

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
tools/                 audit dei workbook e golden test, senza aprire Excel
tests/                 346 test, nessuno richiede Excel
docs/                  i due piani + architettura + audit del W30
samples/               workbook e sorgenti di riferimento (W30)
input/  output/        inbox delle fonti e prodotti (non versionati)
template/              Omni_Report_TEMPLATE.xlsm + come prepararlo
```

## Uso

```bash
pip install -e '.[dev]'                       # core + test, senza Excel
pip install -e '.[excel,audit,dev]'           # tutto (richiede Excel desktop)

omni-report check                             # ambiente e template pronti?
omni-report patch-template                    # applica la patch VBA (una volta sola)
omni-report contract                          # stampa il contratto colonne
omni-report preflight --week 31               # valida le fonti, NON apre Excel
omni-report build --week 31                   # il "pulsante"
```

Le sei fonti vanno in `input/` con i nomi indicati in `config/settings.yml`:
i 4 CSV (`AT.csv`, `ATwi.csv`, `SF.csv`, `PSAT.csv`) piu' `Turni.xlsx` e
`Back_Office_Time.xlsx`. L'ordine delle colonne dentro i file **non conta**, e
le sorgenti WFM possono coprire mesi: la settimana viene ritagliata da sola,
ricavandola dalle date di `AT_DATASET`.

Le sorgenti WFM contengono molti piu' giorni e molte piu' persone del necessario
(il roster del W30 copre due mesi e 130 agenti): vengono **ritagliate alla
settimana di `AT_DATASET`**, che e' quella in cui i suoi dati stanno per la gran
parte. Il numero passato a `--week` non la decide, la **controlla**: se non
coincide con quella dei dati il preflight blocca e dice quale usare, cosi' non si
producono i numeri di una settimana con il nome di un'altra.

Non si usa il min/max delle date: l'export di AT e' per data Seattle e
`Data Milano` lo sposta di 9 ore, quindi nel W30 sborda al 27 luglio — otto
giorni invece di sette.

`--only` ricarica un sottoinsieme, per quando i turni cambiano in corsa:

```bash
omni-report build --week 31 --only Turni
```

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

1. **`template/Omni_Report_TEMPLATE.xlsm`** — copia di un workbook di una
   settimana chiusa, piu' la patch `SilentMode` al VBA. **Non serve svuotare i
   DATASET**: il build li pulisce da se'. Istruzioni passo per passo in
   `template/COME_PREPARARE_IL_TEMPLATE.md`, e `omni-report check` ti dice se
   e' a posto senza tentare un build.
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

```bash
# le 26 forme delle celle-turno del roster, e se il parser le regge tutte
python tools/audit_shift_forms.py "samples/omni-report/sorgenti/Turni_W30.xlsx"

# il golden test: ricostruisce Turni e Slot Only Cases dalle sorgenti
python tools/golden_turni.py --monday 46223 \
  --workbook "samples/omni-report/Omni Report W30.xlsm" \
  --roster "samples/omni-report/sorgenti/Turni_W30.xlsx" \
  --backoffice "samples/omni-report/sorgenti/Back_Office_Time_Final.xlsx"
```

`audit_workbook.py` legge l'XML dentro lo zip: nessun Excel, nessun openpyxl,
quindi funziona anche sul WOW da 60 MB. `--usage` dice quali colonne dei dataset
sono consumate da quali formule — ed e' cosi' che si sono trovate **tre colonne
che il piano non aveva** (`docs/audit-workbook-W30.md`).

## Test

```bash
python -m pytest            # 346 test, <2 s, nessuna dipendenza da Excel
```

Fra questi, i tre scenari del piano §11: colonne mescolate e rinominate negli
alias → output identico; colonna rimossa → blocco con messaggio azionabile.
`tests/test_contract_reale.py` verifica il contratto contro le intestazioni e le
formule del workbook vero: se il workbook evolve, lo dice.

Il piu' importante e' `tests/test_golden_wfm.py`: ricostruisce `Turni` e
`Slot Only Cases` dalle sorgenti e li confronta col W30 — 252 = 252 righe e
259 = 259, chiavi identiche, con solo tre differenze dichiarate in anticipo.
Il risultato del processo manuale e' la specifica.

## Documenti

| File | Cosa contiene |
|---|---|
| `docs/piano-omni-report.md` | il piano originale dell'Omni Report |
| `docs/contesto-wow-aht.md` | il contesto originale del WOW AHT |
| `docs/architettura.md` | decisioni prese, deviazioni dai piani e loro motivo; cosa e' provato e cosa no |
| `docs/audit-workbook-W30.md` | misure sul workbook vero: contratto verificato, tre colonne mancanti, VBA, tabella `AHT_Data`, le sorgenti WFM e l'agente che sparisce a meta' |
| `config/contratti.yml` | mappa nome → contratto: dato HR non derivabile dalle sorgenti |
