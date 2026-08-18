# Modifiche da applicare al template WIP

Destinatario: chi lavora sul template dentro Excel (l'agente Claude collegato al
file, o Leonardo a mano). Riferimento: `Omni_Report_TEMPLATE_WIP.xlsm` con i
quattro fogli nuovi `AHT History`, `AHT Trend WoW`, `CaseType Deepdive`,
`Helper CaseType`, misurato il 2026-08-18.

Il lato Python è fatto: la pipeline scrive `AHT History` da uno storico
persistente (`data/aht_history.csv`) e appende in fondo a `Helper CaseType` i
case type che il template non conosce. Restano cinque interventi sul template,
in ordine di importanza. Solo il primo è **necessario alla correttezza**; gli
altri sono migliorie con un motivo misurato.

Nessuno di questi tocca `CaseType Deepdive` a parte una cella di parametro: la
formattazione fatta a mano lì è al sicuro.

---

## 1. `AHT History` ha due colonne nuove, e `AHT Trend WoW` deve usarne una

**Perché.** Python scrive sette colonne, non cinque:

| col | contenuto | note |
|---|---|---|
| A | `week` | il numero di settimana ISO, come oggi. Formato `"W"0` → si legge `W33` |
| B | `channel` | `Phone` / `Non-live` |
| C | `case_type` | |
| D | `volume` | numero di casi |
| E | `aht` | media dei soli AHT numerici, formato `0.00` |
| **F** | **`iso_year`** | **nuova** — 2026 |
| **G** | **`week_key`** | **nuova** — `iso_year * 100 + week`, es. `202633` |

Le settimane ISO ripartono da 1 ogni gennaio: è la convenzione dell'azienda, ed
è giusta. Il problema è che le formule di `AHT Trend WoW` oggi ordinano e
filtrano sulla **sola** settimana (colonna A), e questo produce due guasti a
cavallo d'anno:

- **Gennaio 2027, visibile.** `SORT(UNIQUE(week), 1, -1)` ordina numericamente:
  la W1 del 2027 finisce *sotto* la W52 del 2026, e `TAKE(...,11)` prende
  `52…42`. Il trend mostrerebbe agosto–dicembre 2026 e **mai** le settimane
  nuove, fino a marzo.
- **Dicembre 2027, silenzioso.** Nello storico ci sarebbero due righe con
  `week = 50`, quella del 2026 e quella del 2027. `SUMIFS`/`AVERAGEIFS`
  filtrano su `week = 50` e **le sommano insieme**: volume raddoppiato, AHT
  mediata fra due anni. Nessun errore, nessun `#REF!`. Solo numeri sbagliati.

La colonna `week_key` esiste per questo: `week` resta l'etichetta da leggere,
`week_key` è quella su cui si ordina e si confronta.

### Cosa cambiare — la buona notizia: `B4` è l'unica formula che conta

`N4`, `B44`, `N44` sono tutte `=B4:L4`, cioè riferimenti allo spill di `B4`.
C'è quindi **un solo posto** in cui si scelgono le 11 settimane.

**a) Nuova riga di chiavi, fuori dalla vista.** In `AB4`:

```
=TRANSPOSE(SORT(TAKE(SORT(UNIQUE(FILTER('AHT History'!$G$2:$G$100000,'AHT History'!$G$2:$G$100000<>"")),1,-1),11),1,1))
```

Spilla `AB4:AL4` con le 11 `week_key` più recenti, dalla più vecchia alla più
recente — lo stesso ordine che ha oggi `B4:L4`. Poi **nascondi le colonne
AB:AL**.

**b) `B4` diventa l'etichetta, ricavata dalle chiavi.** Sostituisci la formula
di `B4` con:

```
=MOD(AB4#,100)
```

Spilla `B4:L4` con i numeri di settimana (33, 32, …), e il formato `"W"0` che
c'è già continua a farli leggere `W33`. **L'aspetto del foglio non cambia.**

**c) I criteri passano dalla colonna A alla colonna G.** Sono quattro formule,
poi si trascinano come oggi. In ciascuna, `'AHT History'!$A:$A,B$4` diventa
`'AHT History'!$G:$G,AB$4` (e l'equivalente per gli altri blocchi):

| cella | prima | dopo |
|---|---|---|
| `B5` (AHT Phone) | `…,'AHT History'!$A:$A,B$4,…` | `…,'AHT History'!$G:$G,AB$4,…` |
| `N5` (Volume Phone) | `…,'AHT History'!$A:$A,N$4,…` | `…,'AHT History'!$G:$G,AB$4,…` |
| `B45` (AHT Non-live) | `…,'AHT History'!$A:$A,B$44,…` | `…,'AHT History'!$G:$G,AB$4,…` |
| `N45` (Volume Non-live) | `…,'AHT History'!$A:$A,N$44,…` | `…,'AHT History'!$G:$G,AB$4,…` |

Gli offset tornano da soli trascinando: `B→AB` è +26 colonne e il blocco AHT è
`B..L` → `AB..AL`; `N→AB` è +14 e il blocco volume è `N..X` → `AB..AL`. La riga
va tenuta assoluta (`AB$4`), la colonna relativa.

**d) L'ordinamento dei case type.** In `A5` e `A45`, dentro il `LET`:
`w` passa da `$L$4` a `$AL$4`, e il `SUMIFS` interno da
`'AHT History'!$A:$A,w` a `'AHT History'!$G:$G,w`.

### Alternativa più economica, se preferisci zero modifiche alle formule

Se accetti che l'intestazione legga `2026-W33` invece di `W33`, si può far
scrivere a Python la colonna A come **testo** `2026-W33`: le stringhe con anno
davanti e settimana a due cifre si ordinano correttamente per sempre, e nessuna
formula di `AHT Trend WoW` va toccata (via il formato `"W"0`, che non serve
più). Dimmelo e cambio il lato Python in cinque minuti.

Ho consigliato la versione con `week_key` perché mantiene le etichette
identiche a quelle che confronti con le heat map che usi già.

---

## 2. `Helper CaseType`: fare spazio ai case type nuovi

**Perché.** Oggi le formule arrivano a riga 63 (31 case type × 2 canali) e
l'elenco è scritto a mano. Nell'export della W33 i case type distinti erano 32,
e **cinque combinazioni (canale, case type) non erano in lista**:

| canale | case type | casi in W33 |
|---|---|---|
| Phone | Call Assignment | 30 |
| Phone | Specialty Functions | 8 |
| Non-live | Specialty Functions | — |
| Phone | Collections | 1 |
| Non-live | Live Site Property Settings Issue | 1 |

`Call Assignment` con 30 casi supera anche la soglia di volume: era un risultato
vero che non compariva in nessun foglio, senza nessun errore.

Python ora li appende in fondo, **dalla prima riga libera in colonna B, senza
toccare le righe 2..63** — così `CaseType Deepdive`, che punta alle righe per
posizione, continua a mostrare esattamente i case type che mostra oggi, con la
tua formattazione intatta. Ma Python si ferma dove finiscono le formule del
template: una coppia scritta oltre riga 63 avrebbe nome e canale e nessun numero
accanto, cioè sarebbe presente e invisibile insieme — lo stesso difetto che
stiamo togliendo. Se lo spazio manca, il build lo dice e non scrive.

**Cosa cambiare.**

1. **Trascina le formule delle colonne C..AG** dalla riga 63 fino alla riga
   **163** (100 righe in più, ~50 per canale: nella W33 servivano 5).
2. **Allarga gli intervalli di ranking** da `$A$2:$A$63` a `$A$2:$A$163` — e
   analogamente per gli intervalli di valori nella stessa formula. Sono le
   colonne **P, Q, S, T, U, V**, che fanno tutte la stessa cosa:
   ```
   P:  COUNTIFS($A$2:$A$163,$A2,$N$2:$N$163,"<"&$N2)/(COUNTIF($A$2:$A$163,$A2)-1)
   Q:  … $O$2:$O$163 …
   S:  … $F$2:$F$163 …
   T:  … $E$2:$E$163 …
   U:  … $M$2:$M$163 …
   V:  … $G$2:$G$163 …
   ```
   Le righe vuote non falsano niente: questi `COUNTIFS` filtrano sul canale in
   colonna A, e una riga vuota non corrisponde né a `Phone` né a `Non-live`.
   È anche il motivo per cui i case type nuovi possono stare in fondo mescolati
   fra i due canali — la posizione non conta, conta la colonna A.
3. **Facoltativo, per pulizia:** avvolgi le formule delle colonne D..X in
   `IF($B2="","",…)` così le ~100 righe di riserva restano visivamente vuote
   invece di mostrare zeri e `Process Driven`. Non è necessario: `CaseType
   Deepdive` non le legge.
   ⚠️ **Non toccare la colonna D in modo che risulti del tutto vuota**: Python
   misura la capienza del foglio dall'ultima cella non vuota di quella colonna.
   Una formula che restituisce `""` va benissimo (la cella contiene comunque una
   formula); cancellare la formula no.

---

## 3. `CaseType Deepdive`: soglia di volume da 20 a 10

Una cella per canale, niente formule da toccare:

- `C6` (Phone): `20` → **`10`**
- `C7` (Non-live): `20` → **`10`**

**Perché.** Il deepdive originale (`WoW CaseType Deppdive`) lavorava su un
periodo **mensile**; qui il perimetro è la settimana, cioè circa un quarto dei
casi. Misurato sull'export della W33:

| Min_Volume | combinazioni analizzate | forzate a `Process Driven` |
|---|---|---|
| 20 (oggi) | 17 | 36 |
| 15 | 22 | 31 |
| **10** | **26** | **27** |
| 5 | 32 | 21 |

Con 20, su 53 combinazioni presenti solo 17 venivano classificate davvero.

---

## 4. Cancellare `Profilo Colonne SF`

**Si può, senza rompere niente: nessun altro foglio lo legge.** È un foglio
diagnostico — per ognuna delle 143 colonne di `AHT_Data` conta celle popolate,
% riempimento, valori distinti e i più frequenti — ed è lo strumento con cui a
suo tempo si è capito quali colonne dell'export servissero.

Due motivi per toglierlo:

- **Costa in ricalcolo**: 435 formule di cui **143 ad array** che fanno
  `UNIQUE(FILTER(...))` su una colonna intera della tabella, cioè circa mezzo
  milione di letture di cella a ogni ricalcolo, per un foglio che nessuno apre.
- **È già sostituito**: `tools/audit_workbook.py` fa la stessa cosa offline in
  mezzo secondo, ed è esattamente lo strumento che il 2026-08-18 ha scoperto che
  l'export SF ha cambiato forma (una colonna nuova, `case_has_outbound_call`,
  inserita in AC; una persa, `claim_subtype_list`).

Se lo cancelli, dimmelo: ci sono otto file nel repo che lo citano (test,
fixture, `tools/fix_profilo_colonne_sf.py`, `tools/golden_report.py`, docs) e li
allineo io.

---

## 5. Nascondere `AHT History`

`Visible = xlSheetVeryHidden` (o semplicemente Nascondi). Le formule di
`AHT Trend WoW` continuano a leggerlo senza problemi.

Serve solo a togliere una linguetta di rumore: il foglio è dati grezzi, non
qualcosa da guardare. **Non serve a alleggerire il file** — misurato, pesa
18,7 KB compressi, lo 0,31% di un file da 6,5 MB. Il peso è tutto in
`AT_DATASET` (3311 KB, il 54,9%) e `SF_DATABASE` (1488 KB, 24,7%).

Il foglio **deve restare nel workbook**: `AHT Trend WoW` sono 1590 formule che
leggono `'AHT History'!$A$2:$A$100000`. Lo storico *persistente* invece vive
fuori, in `data/aht_history.csv`; nel workbook ce n'è solo la copia che le
formule interrogano.

---

## Cosa fa Python, per sapere cosa NON fare a mano

- Riscrive `AHT History` **da zero** a ogni build, intestazione compresa, con lo
  storico **completo** (non solo 11 settimane: la finestra la sceglie `B4` con
  `TAKE`, e tagliare prima butterebbe dati senza mostrare niente di più). Le
  righe scritte a mano lì vengono sovrascritte: le correzioni vanno fatte in
  `data/aht_history.csv`.
- Propaga alle righe nuove i formati numerici **letti dalla riga 2** del
  template. Il formato vive nel template, non nel codice: se cambi `"W"0` o
  `0.00` lì, la pipeline lo segue. Rovescio della medaglia: se svuoti
  completamente il foglio, la riga 2 non ha più un formato da copiare.
- Appende a `Helper CaseType` solo colonne **A** e **B**, solo sotto le righe
  occupate, e mai riordina.
- Ridimensiona già `AHT_Data` a ogni scrittura di `SF_DATABASE` (lo faceva
  prima di questa richiesta).
- Il preflight **SEGNALA** i case type nuovi *prima* del build, e **BLOCCA** se
  `Date Viewpoint` non cade di lunedì o se l'export SF è di una settimana
  diversa dal resto del report.
