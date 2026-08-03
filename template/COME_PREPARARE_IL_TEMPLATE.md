# Come preparare il template

Il template è il **motore**: formule, VBA e i fogli di configurazione
(`Anagrafica`, `Email Agenti`, `Turni`, `Helper Malpractice`). La pipeline non lo
apre mai in scrittura — ne fa una copia a ogni run e lavora sulla copia. Se un
giro va male, il template resta intatto.

Il file finito deve stare qui e chiamarsi:

    template/Omni_Report_TEMPLATE.xlsm

Quando l'hai fatto, verifica con:

    omni-report check

che ti dice riga per riga cosa è a posto e cosa manca, senza provare un build.

---

## Passo 1 — copia

Copia un workbook di una settimana chiusa (es. `Omni Report W30.xlsm`) e
rinominalo `Omni_Report_TEMPLATE.xlsm`.

**Non serve svuotare i fogli DATASET.** Il `build` li pulisce prima di scriverli,
quindi un template con i dati della W30 ancora dentro funziona identico. Pesa
6,5 MB invece di pochi KB, e basta. Svuotarli è igiene, non un requisito — e
farlo a mano è un'occasione in più per rompere le formule ad array.

## Passo 2 — la patch al VBA (l'unica cosa obbligatoria)

La macro `Refresh_Dettaglio_Malpractice` mostra due `MsgBox`: uno a fine
esecuzione, uno nel gestore d'errore. In automazione **bloccano il processo a
tempo indeterminato**, in attesa di un clic che nessuno darà — e con Excel
invisibile non si vede nemmeno il dialogo.

Apri l'editor VBA (`Alt+F11`), modulo **`CreaMalpractice`**.

### 2a. In testa al modulo, sotto `Option Explicit`

```vb
Public SilentMode As Boolean

Public Sub SetSilentMode(ByVal value As Boolean)
    SilentMode = value
End Sub
```

### 2b. Il `MsgBox` di fine esecuzione

Cerca questa riga (verso la fine di `Refresh_Dettaglio_Malpractice`):

```vb
    MsgBox "Dettaglio Malpractice rigenerato: " & outRows.Count & " righe.", vbInformation
```

e mettila dietro il flag:

```vb
    If Not SilentMode Then MsgBox "Dettaglio Malpractice rigenerato: " & outRows.Count & " righe.", vbInformation
```

### 2c. Il `MsgBox` di errore — e la propagazione

Il blocco `CleanFail` oggi è:

```vb
CleanFail:
    Application.Calculation = savedCalc
    Application.ScreenUpdating = True
    MsgBox "Errore durante la rigenerazione: " & Err.Description, vbExclamation
End Sub
```

Diventa:

```vb
CleanFail:
    Application.Calculation = savedCalc
    Application.ScreenUpdating = True
    If Not SilentMode Then MsgBox "Errore durante la rigenerazione: " & Err.Description, vbExclamation
    If SilentMode Then Err.Raise Err.Number, , Err.Description
End Sub
```

**La seconda riga conta quanto la prima.** Silenziare il messaggio senza
propagare l'errore farebbe fallire la macro senza che la pipeline lo sappia: il
workbook verrebbe salvato con un `Dettaglio Malpractice` incompleto e nessuno se
ne accorgerebbe. Con `Err.Raise`, il run si ferma e te lo dice.

### 2d. Salva

Salva **mantenendo il formato `.xlsm`** (con macro). Se Excel propone `.xlsx`,
rifiuta: perderesti tutto il VBA.

## Passo 3 — macro abilitate

xlwings può aprire il file e lanciare la macro solo se Excel non blocca le macro.
La via più solida è dichiarare la cartella come **percorso attendibile**:

    File → Opzioni → Centro protezione → Impostazioni Centro protezione
      → Percorsi attendibili → Aggiungi nuovo percorso
      → scegli la cartella del progetto, spunta "Anche le sottocartelle"

Alternativa più larga (meno consigliata): *Impostazioni macro → Abilita tutte le
macro VBA*.

## Passo 4 — verifica

```bash
omni-report check
```

Controlla, senza aprire Excel: che il template ci sia, che abbia i fogli attesi,
che il VBA esponga `SetSilentMode`, che la tabella alias di
`Helper Malpractice`!D:E ci sia e che `Email Agenti` sia popolato. Poi, se
xlwings è installato, verifica anche che Excel risponda.

---

## Perché il template non è nel repository

Contiene dati reali degli agenti. Se lo aggiungi al repo, sappi che ci finiscono
anche quelli. Se preferisci tenerlo fuori, scommenta la riga `template/*.xlsm`
in `.gitignore` — ma allora ricordati di conservarne una copia altrove, perché
la patch al VBA è l'unica cosa non ricostruibile da questo repository.
