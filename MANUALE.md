# Omni Report — manuale d'uso

Questo programma prende i sei file che scarichi ogni settimana e produce
`Omni_Report_W<NN>.xlsm`, lo stesso workbook che prima si costruiva a mano
incollando i dati nei fogli.

Non serve saper programmare. Servono tre comandi, e la capacità di **leggere il
rapporto che il programma stampa** — che è la parte importante, ed è quella
spiegata meglio qui.

**Regola unica da ricordare:** se il programma si ferma, non ha fallito. Si è
fermato *invece* di produrre numeri sbagliati. Il messaggio dice sempre tre cose:
cosa non torna, dove l'ha visto, e come si sistema. Non c'è nessun caso in cui la
risposta giusta sia forzarlo.

---

## Indice

1. [Cosa fa, in due minuti](#1-cosa-fa-in-due-minuti)
2. [Preparare il computer, una volta sola](#2-preparare-il-computer-una-volta-sola)
3. [Il giro settimanale](#3-il-giro-settimanale)
4. [Leggere il preflight](#4-leggere-il-preflight)
5. [Le segnalazioni, una per una](#5-le-segnalazioni-una-per-una)
6. [I blocchi, una per uno](#6-i-blocchi-uno-per-uno)
7. [Controllare il workbook prodotto](#7-controllare-il-workbook-prodotto)
8. [Tutti i comandi](#8-tutti-i-comandi)
9. [Cosa si può cambiare senza toccare il codice](#9-cosa-si-può-cambiare-senza-toccare-il-codice)
10. [Quando qualcosa va storto](#10-quando-qualcosa-va-storto)
11. [Cosa il programma NON fa](#11-cosa-il-programma-non-fa)

---

## 1. Cosa fa, in due minuti

Sei file in ingresso, un workbook in uscita.

| file (il nome può cambiare) | cosa contiene | finisce in |
|---|---|---|
| `AT DATASET W<NN>.xlsx` | stati agente da Amazon Connect | foglio `AT_DATASET` |
| `ATwi DATASET W<NN>.xlsx` | tempi di gestione contatti | foglio `ATwi_DATASET` |
| `SF DATABASE W<NN>.csv` | casi Salesforce | foglio `SF_DATABASE` |
| `PSAT DATASET W<NN>.csv` | sondaggi di soddisfazione | foglio `PSAT_DATASET` |
| `Turni ....xlsx` | roster WFM | foglio `Turni` |
| `Back Office ....xlsx` | slot di back office | foglio `Slot Only Cases` |

Il programma:

1. legge i sei file e **controlla che le colonne che servono ci siano**;
2. confronta le fonti fra loro, cercando le incoerenze che a mano non si vedono;
3. se tutto torna, copia il template, scrive i sei fogli, ricalcola e lancia la
   macro che genera `Dettaglio Malpractice`;
4. salva in `output\Omni_Report_W<NN>.xlsm`.

Le due cose che rendono il programma più affidabile del copia-incolla:

- **Le colonne si cercano per nome, non per posizione.** Se un export sposta una
  colonna, il programma la ritrova. A mano, un incollaggio disallineato produce
  numeri plausibili e sbagliati, e nessuno se ne accorge.
- **Confronta le fonti fra loro.** Un agente che ha slot di back office ma non ha
  turni entra in metà delle regole di malpractice e resta fuori dall'altra metà.
  Nel report finito non si vede: non c'è nessun errore, solo una riga che manca.
  Il programma lo trova e si ferma.

---

## 2. Preparare il computer, una volta sola

Serve un PC **Windows con Excel desktop installato**. Non funziona su Excel
online né su un Mac senza Excel: le formule ad array e il VBA li deve valutare
Excel vero.

### La via semplice: `installa.bat`

Doppio clic su **`installa.bat`**, nella cartella del progetto. Controlla se
Python c'è, se non c'è prova a installarlo da solo, e poi installa quello che
serve al programma. Va fatto **una volta per PC**.

Se durante l'installazione di Python compare la richiesta "chiudi e rilancia":
chiudi davvero la finestra e fai di nuovo doppio clic su `installa.bat` — Windows
deve aggiornare le sue impostazioni prima che il resto funzioni.

Se qualcosa va storto, la finestra lo dice in chiaro e non si chiude da sola:
leggi il messaggio, e se resti bloccato manda uno screenshot di quella finestra
a chi ti ha dato il programma.

### La via manuale, se preferisci la riga di comando

```bash
git clone <indirizzo del repository>
cd "Faster Reports v2"
pip install -e ".[excel]"
```

### In entrambi i casi, verifica

```bash
omni-report check
```

Devi vedere `OK` su tutte le righe. Se qualcosa manca, la riga stessa dice come
sistemarlo. La voce che merita attenzione è **`VBA compilabile`**: se dice
`MANCA`, il modulo VBA del template ha le dichiarazioni nell'ordine sbagliato e
Excel lo rifiuterebbe **a metà del lavoro**, con una finestra che in automazione
nessuno può chiudere. Le istruzioni sono nel messaggio.

> **La cartella `template\` è la cosa da non perdere.** Contiene
> `Omni_Report_TEMPLATE.xlsm`, cioè tutto il motore: 15.000 formule, il VBA, la
> tabella degli alias dei nomi, il foglio `Email Agenti`. I dati si riscaricano;
> il template no. È versionato nel repository — se lo modifichi, committalo.

---

## 3. Il giro settimanale

### Passo 1 — Metti i sei file in `input\`

**Non rinominarli.** I nomi si riconoscono per schema, quindi
`AT DATASET W31.xlsx` va bene com'è. Nemmeno il formato va dichiarato: `.csv` e
`.xlsx` vengono riconosciuti dall'estensione.

**Togli i file della settimana scorsa.** Se in cartella ci sono
`AT DATASET W30.xlsx` e `AT DATASET W31.xlsx`, il programma si ferma e ti dice
quali ha trovato. Non sceglie "il più recente": scegliere darebbe il risultato
giusto per caso, e sbagliato in silenzio la volta che la data del file inganna.

I file di `Turni` e `Back Office` possono contenere **molte più persone e molti
più giorni** del necessario, anche di altri team: il programma ritaglia la
settimana e tiene solo gli agenti del roster HPO. Non serve pulirli.

### Passo 2 — Controlla, senza aprire Excel

```bash
omni-report preflight --week 31
```

Dura pochi secondi e non apre Excel. Stampa il rapporto a schermo e lo salva in
`output\preflight_W31.txt`. **Leggilo prima di andare avanti** — il capitolo 4
spiega come.

Il numero dopo `--week` è un **riscontro**, non un comando: la settimana la
ricavano i dati. Se scrivi `--week 30` e i dati sono della 31, il programma
blocca. Serve a evitare di produrre `Omni_Report_W30.xlsm` pieno di dati della
31 — un file con l'etichetta sbagliata è peggio di un file che manca, perché
viene archiviato.

### Passo 3 — Genera

```bash
omni-report build --week 31
```

La prima volta che lo usi aggiungi `--visible`, così vedi Excel lavorare:

```bash
omni-report build --week 31 --visible
```

Aspettati **qualche minuto**: sono ~24.000 righe più il ricalcolo completo del
workbook. Non toccare Excel mentre lavora.

Se `build` trova qualcosa che non torna si ferma **prima** di aprire Excel, e non
produce niente. Il template resta intatto in ogni caso: il programma lavora
sempre su una copia.

### Passo 4 — Controlla il risultato

Vedi capitolo 7. Sono tre controlli, due minuti.

---

## 4. Leggere il preflight

Il rapporto ha tre parti. La prima riga è il verdetto:

```
Esito complessivo: OK
Settimana dai dati: W31 2026 (27/07/2026 → 02/08/2026, 7 giorni)
```

**Controlla sempre la seconda riga.** Sono i sette giorni su cui verrà costruito
tutto. Se non sono quelli che ti aspetti, hai scaricato l'export sbagliato.

### Parte 1 — Una sezione per fonte

```
### AT_DATASET
  sorgente: C:\...\input\AT DATASET W31.xlsx
  encoding: xlsx · separatore: '(foglio)' · colonne nel sorgente: 14
  STATO: OK
  campo canonico                 | col. | colonna sorgente               | via    | note
  Agent Email                    | B    | Agent Email                    | esatto |
  Agent State                    | F    | Agent State                    | esatto | 42 vuote
  ...
  righe scritte: 23919 · intervallo colonne: B:L
```

Cosa guardare:

- **`STATO`**: `OK` o `BLOCCATO`. Se è bloccato, il perché è sotto.
- **`via`**: come è stata trovata la colonna. `esatto` è il caso normale.
  `alias` o `normalizzato` significa che l'export ha cambiato il nome della
  colonna e il programma l'ha ritrovata comunque: funziona, ma **è un
  cambiamento nell'export che vale segnalare a chi lo produce**.
- **`righe scritte`**: confrontalo con la settimana precedente. Un calo del 40%
  senza una ragione nota vuol dire export incompleto.
- **`note`**: `42 vuote` su 23.919 righe è normale; su 200 righe non lo è.

### Parte 2 — Cosa i lettori hanno scartato

Solo per `Turni` e `Slot Only Cases`, che nascono da fogli larghi:

```
  righe sorgente saltate: 26
      r50: nome incompleto ('intervals')
  agenti della sorgente esclusi (fuori dal target): mario mecca, lucia bianchi
  alias nomi applicati: 4
      'alessandro passierello' -> 'alessandro passariello'
```

- **righe saltate**: totali e etichette del foglio di lavoro, non agenti.
  Normale.
- **agenti esclusi**: persone nel file di back office che non sono nel roster
  HPO — altri team. Normale.
- **alias applicati**: nomi scritti diversamente fra le due fonti, riconciliati
  con la tabella in `Helper Malpractice!D:E` del template. Se un nome nuovo non
  viene riconosciuto, comparirà fra le segnalazioni e la riga va aggiunta lì.

### Parte 3 — `### COERENZA FRA LE FONTI`

È la parte che a mano non si può fare. Ogni riga inizia con `[SEGNALA]` o
`[BLOCCA]`:

- **`[SEGNALA]`** = da sapere, non ferma niente. **Va letto, non ignorato**:
  quasi tutte le segnalazioni descrivono qualcosa che potresti voler sistemare.
- **`[BLOCCA]`** = il programma non produce il workbook finché non decidi tu.

I due capitoli seguenti le elencano tutte.

---

## 5. Le segnalazioni, una per una

### `skill nel roster`
L'elenco dei valori di `Skill` trovati e quante persone per ciascuno. Serve a
vedere a occhio se il roster è quello giusto.

### `skill che somigliano a quelle richieste`
Skill come `HPO   *` — con l'asterisco. Vedi il capitolo 6: se
`include_marked_skills` è attivo diventa una segnalazione, altrimenti blocca.

### `skill riscritte alla forma canonica`
Compare solo con `include_marked_skills: true`. È **l'unico posto in cui il
valore scritto non è quello letto**: `HPO   *` diventa `HPO`, perché il filtro di
`Helper Turni` confronta per uguaglianza esatta e altrimenti quegli agenti
resterebbero fuori dai calcoli pur essendo nel foglio.

### `valori in cache (roster turni)`
Il file contiene formule e non è stato ricalcolato prima di salvarlo, quindi si
leggono i valori dell'ultimo salvataggio. Se hai dubbi: aprilo in Excel, premi
`F9`, risalva.

### `richieste di cambio turno pendenti`
Celle `REQUEST` nel back office. Trattate come "nessuno slot", come fa il
processo manuale.

### `agenti di SF_DATABASE senza email`
Una persona ha chiuso casi ma non è nel foglio `Email Agenti` del template. In
`Anagrafica` comparirà con la email vuota, e le regole che passano dalla email
(login in ritardo, pause, Available fuori turno) non la riguardano, mentre quelle
sui casi sì.

**Da sistemare** aggiungendo una riga in `Email Agenti` — a meno che quella
persona non debba essere nel report, e allora non dovrebbe avere casi qui.

### `agenti con casi ma senza turno`
Persone con casi chiusi e nessun turno nel roster. Il messaggio riporta **quanti
casi**:

```
lucia bianchi: 1 casi in SF_DATABASE
```

- **fino a ~5 casi: normale, non va corretto.** Sono agenti assenti che avevano
  casi in coda, chiusi mentre erano via. Non hanno ore previste perché davvero
  non lavoravano.
- **molti casi: da guardare.** Quella persona ha lavorato per davvero e la sua
  riga nel roster manca.

In `Report Agenti` avranno `Ore previste = 0` e nessuna produttività.

### `agenti con turni ma senza slot`
Ha turni ma nessuno slot di back office. Di solito è chi non fa back office.

### `agenti senza contratto`
`Contratto` non si può dedurre dal roster (con `Expected hours` 0600 esistono sia
FT sia PT) ed è una colonna **informativa: nessun calcolo la usa**. Se vuoi
riempirla, la mappa è `config\contratti.yml`. Se non ti interessa, ignora questa
riga per sempre.

### `formule vicine al limite per <fonte>`
Alcune formule del workbook leggono intervalli con la riga finale scritta dentro
(`SUMIFS(AT_DATASET!$P$2:$P$130000, ...)`). Questa riga dice che i dati stanno
arrivando a quel limite — l'80% o piu'.

**Non c'e' niente di rotto adesso.** Ma il giorno in cui il limite verra' superato,
le righe in eccesso resteranno **fuori dai calcoli senza nessun errore**: medie e
conteggi su un sottoinsieme, numeri plausibili e piu' bassi del vero. Conviene
allargare quando la segnalazione compare, non quando i numeri sono già sbagliati.

Se il limite e' già superato la riga diventa `[BLOCCA] formule troppo corte`, ed
elenca quali formule vanno allargate.

### `settimana dedotta dai dati` / `settimana scelta`
Conferme. La seconda dice anche quante righe di `AT_DATASET` cadono dentro la
settimana: devono essere tutte.

### `giorni scoperti`
Un giorno presente in una fonte e non nell'altra. Con un export parziale è il
primo sintomo.

---

## 6. I blocchi, uno per uno

Sono pochi, e ognuno è una domanda a cui solo una persona può rispondere.

### `settimana dichiarata` diversa da quella dei dati
Hai scritto `--week 30` e i dati sono della 31. Rilancia col numero giusto, o
riscarica l'export.

### `skill che somigliano a quelle richieste`
Nel roster ci sono skill come `HPO   *`. Quelle persone non entrano in `Turni`
(il filtro vuole `HPO` esatto) ma sono nel back office, quindi metà delle regole
le valuta e metà no.

**La decisione:** in `config\settings.yml`,

```yaml
sources:
  include_marked_skills: true    # entrano nel report
  # include_marked_skills: false # restano fuori (e il programma blocca)
```

Con `true` entrano davvero: la skill scritta diventa `HPO`, e nella settimana 31
questo ha fatto passare tre agenti da `Ore previste = 0` a 24, 8 e 40 ore.

Con `false` restano fuori, ma allora **vanno esclusi anche dal back office**,
altrimenti l'ingresso a metà si ripresenta dall'altro lato.

### `skill che il FILTER non riconoscerà`
Una riga di `Turni` ha un `Team/Skill` diverso da quelli richiesti: finirebbe nel
foglio senza entrare nei calcoli. È la difesa che impedisce al caso precedente di
tornare da un'altra strada.

### `agenti con slot ma senza turni`
Qualcuno è nel back office e non nel roster. Cause tipiche: la skill con
l'asterisco, o un nome scritto diversamente fra le due fonti — e allora la riga
va aggiunta alla tabella alias in `Helper Malpractice!D:E`.

### `agenti senza email`
Un agente con turni non è in `Email Agenti`. Blocca perché il VBA lavora per
email: senza email **nessuna** regola di malpractice lo riguarda. Aggiungi la
riga nel foglio del template.

### `agente in più blocchi`
Il roster ha blocchi di colonne affiancati e la stessa persona compare come HPO
in due: produrrebbe righe doppie, quindi ore contate due volte.

### `settimana di Turni` fuori dall'intervallo
Il file dei turni copre una settimana diversa da `AT_DATASET`. Di solito è il
file della settimana scorsa rimasto in cartella.

### Turni implausibili
Fine prima dell'inizio, pausa fuori dal turno, ore ≤ 0 o > 16. Il messaggio dice
cella e agente: si corregge nel file di origine.

### Una cella di turno che il programma non capisce
Il parser conosce tutte le forme viste nei file reali (`0900_1331_1431_1800`,
`Off`, turni senza pausa, il suffisso ` O`). Se ne arriva una nuova **blocca
dicendo valore, cella e agente**, invece di indovinare: un turno interpretato
male diventa ore previste sbagliate.

---

## 7. Controllare il workbook prodotto

Apri `output\Omni_Report_W<NN>.xlsm` e guarda tre cose. Due minuti.

**1. `Helper Turni`, colonna `N` (Ore previste settimana).** È il controllo più
importante. Se è **tutta a zero**, i nomi non hanno fatto match e ogni numero di
malpractice è falso. Qualche zero isolato è normale: è chi ha la settimana intera
di riposo — verificalo in `Turni`, dovrebbe avere 7 righe `FERIE-OFF`.

**2. `Turni`, colonna `E` (Data).** Devono essere i sette giorni della settimana
giusta. Se vedi le date della settimana scorsa, il programma non ha scritto.

**3. `Dettaglio Malpractice`.** Deve avere righe, con le categorie di sempre
(AHT alto, caso chiuso veloce, break, login in ritardo, Available fuori turno).
Confronta il totale con la settimana precedente: uno scostamento grosso ha una
causa, e vale trovarla.

Due cose che sembrano errori e non lo sono:

- nella colonna `Data` di `Dettaglio Malpractice` molte righe hanno `-`: sono le
  regole basate sui casi, dove una data non c'è. Era così anche prima.
- in `AT_DATASET` la numerazione della colonna `A` continua oltre l'ultima riga
  di dati: è un residuo del template, inerte — niente lo legge.

### Le celle di errore, che il programma controlla da sé

Dopo il salvataggio il programma rilegge il workbook e cerca le celle con un
errore di calcolo. Se ne trova, lo stampa e **esce con errore**, anche se il file
e' stato prodotto:

```
ATTENZIONE: 116 celle contengono un errore di calcolo.
  AHT Outliers Export: 73 #REF! (es. L3, N3, L4, N4, L5)
  Verifica AHT: 37 #REF! (es. H29, H30, H31, H32, H33)
  AHT Outliers: 6 #VALUE! (es. K11, M11, N11, O11, P11)
```

Cosa significano:

| errore | causa tipica |
|---|---|
| `#SPILL!` | un array dinamico non ha spazio per espandersi — succede **quando i dati crescono** |
| `#REF!` | una formula legge il risultato di un'altra che e' rimasta vuota |
| `#VALUE!` | una media o un quartile su un insieme vuoto |
| `#N/D` | un XLOOKUP che non trova: un nome che non fa match, o una colonna sorgente non riempita |

**Guarda prima il foglio con piu' errori**: gli altri di solito sono conseguenze
sue. Nell'esempio sopra tutte e 116 le celle venivano da due colonne di
`SF_DATABASE` che la pipeline non riempiva.

### Il confronto con un file fatto a mano

Se hai anche la versione costruita a mano:

```bash
python tools/golden_report.py "output/Omni_Report_W31.xlsm" "il tuo file a mano.xlsm"
```

Elenca ogni differenza cella per cella. Le differenze **attese** sono due, sulle
righe `FERIE-OFF`: il programma lascia vuote `Ore/gg` e `chiave`, il processo
manuale le riempiva. Nessuna formula le legge (filtrano tutte
`Stato="LAVORA"`), quindi non cambiano un numero.

Qualunque **altra** differenza è da spiegare, non da approvare.

---

## 8. Tutti i comandi

### `omni-report check`
Verifica ambiente e template. Non tocca i dati. Da lanciare quando qualcosa non
funziona e non si capisce perché.

```
--no-excel     salta la prova di apertura di Excel
--config DIR   usa un'altra cartella di configurazione
```

### `omni-report preflight --week NN`
Valida le sei fonti senza aprire Excel. Salva `output\preflight_W<NN>.txt`.

```
--only DATASET [DATASET ...]   controlla solo alcune fonti
--input DIR / --output DIR     altre cartelle
```

### `omni-report build --week NN`
Il "pulsante". Richiede Excel.

```
--visible                      mostra Excel mentre lavora
--only DATASET [DATASET ...]   ricarica SOLO alcuni fogli in un workbook già
                               prodotto (vedi sotto)
```

`--only` serve quando i turni cambiano a giro iniziato:

```bash
omni-report build --week 31 --only Turni "Slot Only Cases"
```

Scrive nel workbook che esiste già, non riparte dal template. Se il workbook non
c'è, si rifiuta: ripartendo dal template i fogli non ricaricati resterebbero
pieni dei dati della settimana del template.

### `omni-report contract`
Stampa quali colonne il programma cerca in ogni fonte, dove le scrive e **chi le
legge nel workbook**. Utile per rispondere a "questa colonna serve ancora?".

### `omni-report patch-template`
Applica al VBA del template il flag che silenzia i due `MsgBox` — serve una volta
sola, per template. Richiede la spunta *"Considera attendibile l'accesso al
modello a oggetti dei progetti VBA"* nel Centro protezione di Excel. Se non vuoi
abilitarla, `python tools/make_vba_patch.py <template>` genera il modulo da
incollare a mano.

---

## 9. Cosa si può cambiare senza toccare il codice

Tutto quello che sta in `config\`. Sono file di testo: aprili con Notepad.

### `config\settings.yml`

| voce | a cosa serve |
|---|---|
| `input_files` | gli schemi dei nomi dei file. Cambiali se cambia il nome degli export |
| `sources.skills` | quali skill entrano nel report (`["HPO"]`) |
| `sources.include_marked_skills` | le skill con l'asterisco entrano? Vedi capitolo 6 |
| `sources.roster_sheet` | quale foglio leggere nel file dei turni (`publish`) |
| `sources.backoffice_sheet` | quale foglio nel back office (`Only Cases Shifts`) |
| `output.workbook_name` | come si chiama il file prodotto |
| `validation.max_uncoercible_ratio` | quanti valori illeggibili tollerare per colonna (0.02 = 2%) |
| `excel.visible` | mostrare sempre Excel |

### `config\contratti.yml`
La mappa nome → contratto (FT/PT), che il roster non contiene. Solo informativa.

### `config\columns.yml`
**Il contratto delle colonne**: dice per ogni fonte quali colonne servono, come
si chiamano, dove vanno nel workbook e chi le legge. Si tocca quando un export
cambia i nomi delle colonne o quando serve una colonna nuova — vedi
`docs\estendere.md`.

### Nel template, non in un file di configurazione

- **`Email Agenti`**: nome → email. Il programma non può inventarle. Quando il
  preflight segnala un agente senza email, la riga si aggiunge qui.
- **`Helper Malpractice!D:E`**: la tabella degli alias dei nomi, per riconciliare
  le grafie diverse fra roster e Salesforce.
- **`Helper Malpractice!B2:B8`**: i parametri delle regole. Ogni riga ha la sua
  etichetta accanto, in colonna `A`:

  | cella | parametro | valore oggi |
  |---|---|---|
  | `B2` | limite agenti in break simultanei | 4 |
  | `B3` | break massimo giornaliero (min) | 24 |
  | `B4` | percentile AHT "alto" | 0,95 |
  | `B5` | soglia caso chiuso rapido (min) | 2,5 |
  | `B6` | tolleranza Available Cases fuori turno (min) | 5 |
  | `B7` | **offset fuso Seattle–Milano (ore)** | 9 |
  | `B8` | ora inizio prevista di default | 0,375 (09:00) |

  **Cambiare un parametro cambia i numeri**: fallo di proposito e annota quando.

  `B7` merita un discorso a parte: **non è una soglia**, è il fuso. Vale 9 ore
  con l'ora legale (Seattle PDT → Milano CEST) e diventa **8** in inverno, quando
  i due paesi cambiano ora in date diverse. Da lì escono `Data Milano` e
  `Ora Milano` di `AT_DATASET`, quindi sbagliarlo sposta ogni evento di un'ora e
  può spostare di giorno quelli a cavallo della mezzanotte. Lo legge anche la
  pipeline, per sapere quali sono i sette giorni della settimana: una verità sola
  in un posto solo.

---

## 10. Quando qualcosa va storto

### «Python non è stato trovato; eseguire senza argomenti da installare dal Microsoft Store...»
Non è un errore del programma: significa che su questo PC **non c'è un Python
vero**. Fai doppio clic su `installa.bat` — prova a installarlo da solo. Se non
l'hai ancora fatto su questo PC, è sempre il primo passo (capitolo 2).

### «nessun file corrisponde a `Turni*`»
Il file non è in `input\`, o si chiama diversamente da come lo aspetta
`settings.yml`. Il messaggio elenca i file che ha trovato.

### «2 file corrispondono a `AT DATASET*`»
Due settimane nella stessa cartella. Togli quella vecchia.

### «xlwings non è installato»
`pip install -e ".[excel]"`, su una macchina con Excel.

### Excel resta aperto e il comando è appeso
Chiudi il processo Excel dal Task Manager e rilancia. Se succede sempre, lancia
`omni-report check`: la voce `VBA compilabile` o `MsgBox silenziati` ti dirà se il
template ha un problema che ferma la macro con una finestra invisibile.

### Il preflight dice OK ma i numeri sembrano strani
Fai i tre controlli del capitolo 7. Se `Helper Turni!N` è a zero il problema è il
match dei nomi; se le date sono sbagliate hai l'export della settimana
sbagliata. Se non è nessuno dei due, il preflight salvato in `output\` è la prima
cosa da mandare a chi ti aiuta.

### «Il modulo CreaMalpractice del template non compila»
Il VBA ha le dichiarazioni dopo una procedura. Il messaggio dice quali righe:
vanno spostate sopra la prima `Sub`. Il modulo corretto è già pronto in
`template\CreaMalpractice_patched_da_incollare.vb`.

### Ho rotto qualcosa e non so cosa
```bash
git status          # cosa ho modificato
git diff            # in che modo
git checkout .      # butta via le mie modifiche non committate
python -m pytest    # se i test sono verdi, il programma è integro
```

I dati in `input\` e `output\` non sono versionati: `git checkout` non li tocca.

---

## 11. Cosa il programma NON fa

Per evitare aspettative sbagliate:

- **Non inventa dati.** Se una email manca, lo dice; non la deduce.
- **Non corregge le fonti.** Un turno scritto male va corretto nel file di
  origine.
- **Non genera i due report WOW** (AHT Trend, CaseType Deepdive). Restano
  manuali, per ora.
- **Non scrive `Email Agenti`, `Helper Turni`, `Anagrafica`, `Helper
  Malpractice`.** I primi sono anagrafiche curate a mano, gli altri si calcolano
  da soli.
- **Non gira senza Excel desktop.** Il `preflight` sì; il `build` no.
- **Non decide al posto tuo.** Le domande che bloccano sono domande vere:
  cambiano i numeri, e la responsabilità è di chi firma il report.

---

## In caso di dubbio

Nell'ordine:

1. rileggi il messaggio d'errore: contiene già il cosa, il dove e il come;
2. `omni-report check`;
3. `omni-report preflight --week NN` e leggi `output\preflight_W<NN>.txt`;
4. `python -m pytest` — se i test sono verdi il programma è integro e il problema
   è nei dati o nella configurazione.

E la regola con cui abbiamo costruito tutto: **meglio un report che manca di un
report sbagliato**. Un file mancante te ne accorgi subito. Un numero sbagliato
viene archiviato.
