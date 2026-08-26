# Domande da mandare al mittente del template

Il template `August_Transfer_Data_Template_v2.xlsx` e' stato creato il 2026-08-24 con
openpyxl; ultimo autore registrato nelle proprieta' del file: Annie Ng.

Conviene mandarle **subito**: la scadenza e' il 1 settembre e la mail stessa invita a
chiedere chiarimenti in tempo utile.

## Difetti del file

1. **Le intestazioni dicono "August 2025".** Le sezioni 1 e 2 si intitolano
   *"— August 2025"*, ma la richiesta e' per agosto 2026. Presumibilmente un refuso, ma va
   confermato prima di compilare.
2. **La sezione 3 non esiste.** Il foglio passa da *"2. Transfer validity"* (riga 13) a
   *"4. Action plan"* (riga 23). La mail parla di "four areas" ma ne descrive due. Manca un
   pezzo del template, oppure la numerazione e' sbagliata.
3. **La formula dell'invalid rate e' sbagliata.** `D20` sta nella colonna **Combined** ma
   contiene `=IFERROR(B19/B16,"—")`, cioe' calcola il rate del **solo Voice**. Dovrebbe
   essere `D19/D16`. Le celle `B20` e `C20` (rate per Voice e per BOF) sono vuote.

## Definizioni

4. **Serve la "agreed definition" di transfer valido/invalido.** La mail la cita ma il
   template non la contiene. Senza, ogni sito marca con criteri propri e i numeri non sono
   confrontabili fra siti — che e' esattamente lo scopo dell'esercizio.
5. **Come va letto "by flow type to Advance or Basic queue, across voice and BOF"?**
   Nel nostro export si puo' leggere in due modi che danno numeri diversi:
   - *gamba ricevente* — il caso nato dal transfer, con il canale di provenienza
     (Voice = trasferito dalla voce, BOF = trasferito da chat) e la coda di destinazione
     dedotta dal `work_function` del caso ricevente. E' quella che stiamo usando;
   - *gamba cedente* — il caso chiuso come `Closed - Transferred`, che pero' nel nostro
     export e' **sempre** su canale Phone, quindi lascerebbe la colonna BOF a zero.
   Le due gambe non riconciliano (nella nostra settimana di riferimento: 64 contro 139) e non
   hanno una chiave che le colleghi. Quale delle due si aspettano?
6. **I "call transfer" voce→voce vanno contati?** Sono una popolazione a parte
   (`Case Resolution Category = "Call Transfer"`, 76 casi nella settimana di riferimento):
   non hanno destinazione Advanced/Basic perche' il `work_function` e' `Call Assignment`.
   Oggi li riportiamo solo in annex.
7. **Qual e' il target sul transfer rate?** Il template indica `< 5%` ma accanto
   all'*invalid rate*, non al transfer rate. La sezione 4 va compilata solo "if your site is
   below target": senza sapere qual e' il target sul transfer rate non si puo' stabilire se
   ricadiamo nel caso.
8. **Che periodo copre "August"?** Mese di calendario (1–31 agosto) o settimane ISO complete?
   Agosto 2026 tocca sei settimane ISO, W31–W36, e i nostri export sono settimanali.
