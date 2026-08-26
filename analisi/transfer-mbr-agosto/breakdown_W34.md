# Transfer breakdown — W34

Casi totali nell'export: **2762**
Transfer in ingresso: **139** → transfer rate **5.03%**

## Sezione 1 — destinazione x canale di provenienza (gamba ricevente)

| Queue type | Voice | BOF | TOTALE |
|---|---:|---:|---:|
| Advanced | 18 | 0 | 18 |
| Basic | 112 | 9 | 121 |
| **TOTALE** | **130** | **9** | **139** |

Voice = `Case Origin = Transferred from Phone`; BOF = `Case Origin = Transferred from Live Agent`.
La riga *Altro* raccoglie i `work_function` che non sono un livello (`Call Assignment`, `UNMAPPED`, `Project`): non va sommata ad Advanced o Basic.

## Annex — gamba cedente (NON riconcilia con la ricevente)

- **case_status = Closed - Transferred**: 72 casi · canale {'Phone': 72} · livello {'Basic': 61, 'Advanced': 11}
- **Case Resolution Category = Case Transfer**: 73 casi · canale {'Phone': 73} · livello {'Basic': 62, 'Advanced': 11}
- **Case Resolution Category = Call Transfer**: 33 casi · canale {'Phone': 33} · livello {'Altro': 33}

Le due gambe sono righe diverse dello stesso export, senza chiave che le colleghi: i totali non tornano e non devono tornare.

## Sezione 2 — popolazione da scrubbare

- transfer da scrubbare: **139**
- di cui segnalati ad alto rischio (da scrubbare al 100%): **44**
    - `bounce`: 32
    - `genera_figli`: 11
    - `aht_basso`: 5
- soglia `aht_basso` (10° percentile dei transfer): 0.60 min
- casi padre con piu' di un transfer figlio (bounce): 14
- **campione da marcare a mano: 103** su 139 (74%) — alto rischio al 100% piu' un casuale stratificato per case type, per ±5% al 95%
- NOTA: mancano solo **36** casi al 100% di copertura. A questi volumi conviene scrubbare tutto: la mail cita Legazpi proprio come esempio di sito al 100%, e un campione va spiegato mentre il 100% no.

## Controllo per case type

| Case type | transfer | di cui Advanced | di cui a rischio |
|---|---:|---:|---:|
| Supplier Initiated Traveler Contact | 49 | 0 | 14 |
| Property Settings | 12 | 0 | 6 |
| Rates & Inventory Changes | 10 | 10 | 4 |
| Expedia Collect Invoice | 10 | 0 | 2 |
| Booking Information | 10 | 0 | 2 |
| Hotel Collect Issue | 8 | 0 | 1 |
| Promotion | 7 | 0 | 3 |
| Room Type/Rate Plan | 7 | 7 | 4 |
| EVC | 6 | 0 | 2 |
| Content Update | 5 | 0 | 2 |
| Partner Central Access | 4 | 0 | 0 |
| Supplier Initiated Relocation | 4 | 0 | 1 |
| Contact Update | 2 | 0 | 2 |
| Refund Request | 2 | 0 | 0 |
| Live Site Rates & Inventory Issue | 1 | 1 | 1 |
| Property Details | 1 | 0 | 0 |
| Pre-Onboarding | 1 | 0 | 0 |
