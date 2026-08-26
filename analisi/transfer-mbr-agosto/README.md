# Transfer Data Submission — agosto 2026 (MBR settembre)

Materiale per compilare `August_Transfer_Data_Template_v2.xlsx`, richiesto dall'upper
management entro il **1 settembre 2026**.

**Questa cartella non tocca il programma.** Non modifica `src/`, `config/`, `template/`:
serve a verificare che i transfer si possano misurare con l'export che l'Omni Report gia'
scarica. Se la verifica regge, in un secondo momento si valutera' se portare il calcolo
dentro l'Omni Report come foglio nuovo.

## Come si usa

```
python analisi/transfer-mbr-agosto/estrai_transfer.py "input/Omni Report W34.xlsm" \
    --out analisi/transfer-mbr-agosto --etichetta W34
```

Accetta piu' workbook in un colpo solo (per sommare le settimane di agosto) e in quel caso
deduplica per `case_number`:

```
python analisi/transfer-mbr-agosto/estrai_transfer.py input/*.xlsm --out . --etichetta agosto2026
```

Legge il foglio `SF_DATABASE` (o `CSV DATASET` / `DATASET`) aprendo il workbook come zip,
senza openpyxl e senza Excel — come `tools/audit_workbook.py`. Aggancia le colonne **per
nome**, mai per lettera: l'export SF e' gia' cresciuto in passato e le lettere si spostano.
Se manca una colonna obbligatoria si ferma con un errore che dice quale, invece di produrre
zeri silenziosi.

Produce due file:
- `breakdown_<etichetta>.md` — la tabella della sezione 1, l'annex della gamba cedente,
  la popolazione da scrubbare e il controllo per case type;
- `transfer_casi_<etichetta>.csv` — l'export **case-level** di tutti i transfer, ordinato
  per rischio decrescente, con le colonne `valid_invalid` e `note_scrub` da compilare a
  mano. Separatore `;` e virgola decimale: si apre in Excel italiano senza conversioni.

## Stato

- [x] Fattibilita' verificata su `samples/omni-report/Omni Report W30.xlsm` (3388 casi) e
      sul workbook `MulticaseBounces_Analysis_W20W21.xlsx` (8570 casi).
- [x] Script scritto e validato: riproduce i numeri misurati a mano su entrambi.
- [ ] Girare sull'Omni Report **W34** (17–23 agosto), che l'utente deve generare.
- [ ] Decidere se coprire tutto agosto (W31–W36) o consegnare W34 come settimana campione.
- [ ] Scrub manuale del campione.
- [ ] Compilare il template e mandare le domande al mittente (`domande_al_mittente.md`).

## Validazione (W30, 23–29 luglio 2026)

| Queue type | Voice | BOF | TOTALE |
|---|---:|---:|---:|
| Advanced | 10 | 4 | 14 |
| Basic | 104 | 19 | 123 |
| Altro | 2 | 0 | 2 |
| **TOTALE** | **116** | **23** | **139** |

139 transfer su 3388 casi = **4,10%**. Dettagli in `breakdown_W30_validazione.md`.

Controprova sul workbook W20+W21 (8570 casi): 776 transfer = 9,05%, con
Advanced 73/24 e Basic 572/106. Gli stessi numeri erano stati contati a mano prima di
scrivere lo script.
