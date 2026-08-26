# Rilievi sul template `August_Transfer_Data_Template_v2.xlsx`

Audit fatto sull'XML del file. Il file e' stato creato il **2026-08-24 con openpyxl**,
ultimo salvataggio da Excel il **2026-08-26 alle 07:22 UTC**, ultimo autore registrato
**Annie Ng**. Un solo foglio, `August Data Submission`, area usata `A1:F30`.

Non e' possibile chiedere chiarimenti al mittente in tempo utile, quindi per ogni rilievo
qui sotto c'e' anche **la decisione presa** per compilare comunque. Questo documento va
allegato alla submission: copre il sito su ogni scelta fatta al posto di una definizione
mancante.

---

## A. Date e periodo

**1. Le intestazioni dicono "August 2025".**
`A6` = *"1. Transfer destination breakdown — August 2025"*, `A13` = *"2. Transfer validity —
August 2025"*. Il file e' del 24 agosto 2026 e la scadenza e' il 1 settembre 2026.
→ **Decisione:** si compila per **agosto 2026** e lo si dichiara nelle note.

**2. `A16` dice solo "Total transfers (August)", senza anno**, mentre le due intestazioni di
sezione l'anno ce l'hanno (sbagliato). Incoerenza interna al file.

**3. Il template non dice quale periodo copra "August".** Mese di calendario (1–31) o
settimane ISO complete? Agosto 2026 tocca sei settimane ISO, W31–W36, perche' il 1 cade di
sabato e il 31 di lunedi'.
→ **Decisione:** si dichiara esattamente il periodo coperto dai dati consegnati.

## B. Struttura e numerazione

**4. La sezione 3 non esiste.** Il foglio passa da `A13` *"2. Transfer validity"* ad `A23`
*"4. Action plan"*. Le righe 21 e 22 sono **due righe spaziatrici consecutive** alte 7,5 pt,
mentre ovunque altrove la spaziatura e' di una riga sola (riga 2, riga 5, riga 12, riga 30).
Due spaziatrici adiacenti sono la traccia tipica di un blocco rimosso — indizio, non prova:
la `calcChain` non contiene riferimenti orfani.

**5. La mail parla di "four areas" ma ne descrive due** (destination breakdown, validity) e
poi un paragrafo sull'accountability che corrisponde alla sezione 4. L'area 3 manca sia dal
file sia dalla mail.
→ **Decisione:** si consegna quello che c'e' e si segnala il buco.

**6. Le due tabelle non sono allineate fra loro.** In sezione 1 la colonna `D` e' **Notes**;
in sezione 2 la colonna `D` e' **Combined** e Notes si sposta in `F`. Chi compila trova la
stessa colonna con due significati diversi a dieci righe di distanza.

**7. `Voice` e `BOF` compaiono come intestazione in entrambe le tabelle** (`B8`/`C8` e
`B15`/`C15`) **senza essere mai definiti**, e nelle due sezioni significano cose diverse:
in sezione 1 e' una ripartizione della destinazione, in sezione 2 e' il canale del transfer.

**8. Il template non usa mai il termine "flow type"** che la mail mette al centro della
richiesta ("*by flow type to Advance or Basic queue*"), e non dice se Voice/BOF sia il canale
**di partenza** o **di arrivo** del transfer.
→ **Decisione:** si usa il canale **di partenza** sulla gamba ricevente, e si dichiara.
Sui nostri dati e' l'unica lettura che riempie tutta la 2×2: la gamba cedente e' al 100% su
canale Phone, quindi lascerebbe la colonna BOF a zero.

**9. La mail scrive "Advance queue", il template "Advanced queue"** (`A9`). Refuso minore, ma
sono due nomi diversi per la stessa cosa in due documenti che viaggiano insieme.

## C. Formule

**10. `D20` e' sbagliata.** Sta nella colonna **Combined** ma contiene
`=IFERROR(B19/B16,"—")`, cioe' *invalid Voice / totale Voice*. Dovrebbe essere `D19/D16`.
Cosi' com'e', il "combined invalid rate" di ogni sito e' in realta' il rate del solo Voice.
→ **Decisione:** nella copia compilata la formula viene corretta in `D19/D16`, e la
correzione viene dichiarata.

**11. `B20` e `C20` sono celle di input vuote formattate in percentuale** (stile con formato
`0.0%` e sfondo crema, lo stesso delle celle da compilare). Nelle righe 16–19 l'utente inserisce
`B` e `C` e la `D` si calcola; nella riga 20 invece l'utente dovrebbe **digitare a mano** due
percentuali che sono derivabili da righe gia' presenti (`B19/B16` e `C19/C16`). Incoerente
con il resto della tabella e a rischio di disallineamento con `D20`.
→ **Decisione:** si compilano con il valore calcolato, non a occhio.

**12. `A7` promette totali che non esistono.** Dice *"Row totals are calculated
automatically"*, ma le uniche formule della sezione 1 sono `B11=SUM(B9:B10)` e
`C11=SUM(C9:C10)`: sono totali **di colonna**, sulla riga TOTAL. **Non esiste nessun totale di
riga**, perche' la sezione 1 non ha una colonna Combined — al suo posto c'e' Notes.
→ **Decisione:** i totali di riga si riportano nelle note.

**13. `IFERROR(...,0)` maschera errori veri, e in modo asimmetrico.** Su `B11`/`C11` e'
inutile (`SUM` ignora il testo e non puo' dare errore). Su `D16:D19` e' pericoloso:
`B16+C16` restituisce `#VALUE!` se una delle due celle contiene testo — un sito che scrive
`n/a`, `TBC` o `~450` vede **Combined = 0**, senza nessun avviso, e il consolidatore somma
zeri credendoli dati.

**14. Niente collega la sezione 1 alla sezione 2.** `B11` (totale transfer Voice della
sezione 1) e `B16` (*Total transfers*, Voice, sezione 2) devono essere lo stesso numero, ma
sono due input indipendenti senza formula ne' controllo incrociato. E' l'incoerenza piu'
insidiosa per un confronto fra siti: due sezioni dello stesso file possono raccontare due
volumi diversi.
→ **Decisione:** si compilano con lo stesso numero e la coerenza si verifica a mano.

**15. Nessun vincolo interno alla sezione 2.** Non c'e' niente che imponga
`Marked valid + Marked invalid = Transfers scrubbed`, ne' `Transfers scrubbed ≤ Total
transfers`. Tutte e quattro le righe sono input liberi.

**16. `D20` restituisce testo in una cella formattata percentuale.** Il fallback e' la
stringa `"—"` in una cella con formato `0.0%`. Chi consolida i file dei vari siti con una
`AVERAGE` si trova dei non-numeri.

**17. Formati numerici incoerenti fra celle dello stesso tipo.** Gli input (`B9`, `C9`,
`B10`, `C10`, `B16:C19`) sono in formato *Generale*; i totali `B11`/`C11` sono `#,##0`; i
totali calcolati `D16:D19` sono di nuovo *Generale*. Due celle calcolate con lo stesso ruolo,
due formati diversi.

## D. Target e criterio della sezione 4

**18. Il target `< 5%` (`E20`) e' agganciato all'invalid rate, ma la mail parla di target sul
transfer rate.** *"Sites performing below target on **transfer rate**..."*. Un target sul
transfer rate nel template non compare da nessuna parte.
→ **Decisione:** si assume che il criterio sia l'invalid rate sopra il 5% e lo si dichiara.
Da notare: il nostro transfer rate di W34 e' **5,03%**, cioe' esattamente sulla soglia del
solo numero "5%" che compare nel file. Se il target fosse davvero sul transfer rate, saremmo
sopra — motivo in piu' per far chiarire il punto.

**19. La direzione del target e' invertita nel testo.** `A23` dice *"Complete if your site is
below target on any transfer metric"*, ma il target e' scritto come `< 5%`: essere **sotto**
il 5% significa andare **bene**. Letteralmente, il template chiede il piano d'azione a chi non
ne ha bisogno.

**20. "any transfer metric" non ha referente.** L'unica metrica con un target e' l'invalid
rate: non esistono altre metriche con target, quindi "any" e' vuoto.

**21. Le celle Target `E16:E19` sono formattate come target ma vuote.** In particolare manca
un target sulla **copertura di scrubbing**, che la mail invece chiede esplicitamente
("*if your site is already scrubbing 100%...*").

## E. Cosa la mail chiede e il template non permette di scrivere

**22. Non c'e' nessun campo per la metodologia.** La mail chiede *"please note your current
methodology and coverage"*, ma l'unico spazio e' la colonna `Notes` (`F16:F20`), larga 22
caratteri, **senza a capo automatico** e con righe ad altezza fissa 21,75 pt. Un testo di
metodologia non ci sta e non si vede.
→ **Decisione:** la metodologia va in un allegato separato (`metodologia.md`), citato nelle note.

**23. Non c'e' nessun campo di intestazione oltre a "Site name".** Mancano periodo coperto,
chi compila, data di compilazione, perimetro di canale. Per un file che viene consolidato su
piu' di venti siti sono i primi campi che servono.

**24. Nessuna convalida dati in tutto il foglio.** L'elemento `dataValidations` non esiste:
`Status` (`F25:F29`), `Owner`, le date e le note sono testo libero. Venti siti risponderanno
con venti vocabolari diversi, e il consolidamento va fatto a mano.

**25. Nessuna protezione del foglio.** L'elemento `sheetProtection` non esiste: le celle con
formula (`B11`, `C11`, `D16:D20`) sono sovrascrivibili. Un sito che incolla dei numeri sopra
`D16` distrugge la formula e nessuno se ne accorge.

## F. Difetti minori del file

**26. `Site name` e' una cella unita** (`B4:F4`). Innocuo a mano, fastidioso per chi
consolida con uno script: incollare su celle unite da errore.

**27. `A15` e' vuota ma formattata come intestazione con testo bianco su sfondo grigio
chiaro.** E' la cella d'angolo della tabella della sezione 2: se ci si scrive, il testo e'
illeggibile.

**28. Due stili identici duplicati** (le celle di input usano uno stile, `B4:F4` un secondo
stile con attributi identici). Traccia della generazione via openpyxl senza deduplica.
Irrilevante per l'uso.

**29. La cella attiva salvata nel file e' `N23`**, fuori dall'area usata `A1:F30`. Segno che
l'ultimo salvataggio e' avvenuto mentre si lavorava fuori dalla tabella.

**30. Il file porta un'etichetta di riservatezza Microsoft** (`docMetadata/LabelInfo.xml`,
`enabled=1`). La copia compilata la eredita: va tenuto conto di dove la si inoltra.
