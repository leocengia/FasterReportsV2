# Transfer breakdown — W30 (validazione)

Casi totali nell'export: **3388**
Transfer in ingresso: **139** → transfer rate **4.10%**

## Sezione 1 — destinazione x canale di provenienza (gamba ricevente)

| Queue type | Voice | BOF | TOTALE |
|---|---:|---:|---:|
| Advanced | 10 | 4 | 14 |
| Basic | 104 | 19 | 123 |
| Altro | 2 | 0 | 2 |
| **TOTALE** | **116** | **23** | **139** |

Voice = `Case Origin = Transferred from Phone`; BOF = `Case Origin = Transferred from Live Agent`.
La riga *Altro* raccoglie i `work_function` che non sono un livello (`Call Assignment`, `UNMAPPED`, `Project`): non va sommata ad Advanced o Basic.

## Annex — gamba cedente (NON riconcilia con la ricevente)

- **case_status = Closed - Transferred**: 64 casi · canale {'Phone': 64} · livello {'Advanced': 13, 'Basic': 51}
- **Case Resolution Category = Case Transfer**: 67 casi · canale {'Phone': 67} · livello {'Advanced': 14, 'Basic': 53}
- **Case Resolution Category = Call Transfer**: 76 casi · canale {'Phone': 76} · livello {'Altro': 76}

Le due gambe sono righe diverse dello stesso export, senza chiave che le colleghi: i totali non tornano e non devono tornare.

## Sezione 2 — popolazione da scrubbare

- transfer da scrubbare: **139**
- di cui segnalati ad alto rischio (da scrubbare al 100%): **54**
    - `bounce`: 29
    - `genera_figli`: 21
    - `aht_basso`: 7
    - `misrouted`: 2
- soglia `aht_basso` (10° percentile dei transfer): 0.60 min
- casi padre con piu' di un transfer figlio (bounce): 12

## Controllo per case type

| Case type | transfer | di cui Advanced | di cui a rischio |
|---|---:|---:|---:|
| Supplier Initiated Traveler Contact | 39 | 0 | 7 |
| Booking Information | 13 | 0 | 6 |
| Supplier Initiated Relocation | 13 | 0 | 6 |
| Promotion | 12 | 0 | 6 |
| Property Settings | 11 | 0 | 3 |
| Hotel Collect Issue | 8 | 0 | 4 |
| Partner Central Access | 6 | 0 | 3 |
| Room Type/Rate Plan | 6 | 6 | 4 |
| Rates & Inventory Changes | 6 | 6 | 4 |
| EVC | 5 | 0 | 2 |
| Property Details | 4 | 0 | 2 |
| Expedia Collect Invoice | 4 | 0 | 4 |
| Content Update | 4 | 0 | 2 |
| Proactive Outreach Project | 2 | 0 | 0 |
| Connectivity Questions | 2 | 0 | 1 |
| New Contract | 1 | 1 | 0 |
| Customer Review Removal | 1 | 0 | 0 |
| Partner Tool Issue | 1 | 1 | 0 |
| Refund Request | 1 | 0 | 0 |
