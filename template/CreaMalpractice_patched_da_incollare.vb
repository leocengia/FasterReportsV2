Option Explicit

' ==== Modalita' silenziosa, per l'esecuzione automatica ====
' In automazione un MsgBox blocca il processo a tempo indeterminato, in attesa
' di un clic che nessuno dara', e con Excel invisibile il dialogo non si vede
' nemmeno. La pipeline chiama SetSilentMode(True) prima di lanciare la macro.
Public SilentMode As Boolean

Public Sub SetSilentMode(ByVal value As Boolean)
    SilentMode = value
End Sub
' ==== Etichette categoria: unica fonte di verita' ====
Private mAHT As String, mFast As String, mBreakDay As String
Private mBreakSim As String, mLogin As String, mAC As String

Public Sub Refresh_Dettaglio_Malpractice()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsAT As Worksheet, wsSF As Worksheet, wsDet As Worksheet
    Dim wsAn As Worksheet, wsHT As Worksheet, wsSlot As Worksheet, wsHM As Worksheet
    Dim wsTurni As Worksheet
    Set wsAT = wb.Worksheets("AT_DATASET")
    Set wsSF = wb.Worksheets("SF_DATABASE")
    Set wsDet = wb.Worksheets("Dettaglio Malpractice")
    Set wsAn = wb.Worksheets("Anagrafica")
    Set wsHT = wb.Worksheets("Helper Turni")
    Set wsSlot = wb.Worksheets("Slot Only Cases")
    Set wsHM = wb.Worksheets("Helper Malpractice")
    Set wsTurni = wb.Worksheets("Turni")
    Dim savedCalc As XlCalculation
    savedCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanFail
    Dim limitBreak As Long, maxBreakMin As Double, pctAHT As Double
    Dim fastCaseMin As Double, tolMin As Double, defaultStart As Double
    limitBreak = CLng(wsHM.Range("B2").Value)
    maxBreakMin = CDbl(wsHM.Range("B3").Value)
    pctAHT = CDbl(wsHM.Range("B4").Value)
    fastCaseMin = CDbl(wsHM.Range("B5").Value)
    tolMin = CDbl(wsHM.Range("B6").Value)
    defaultStart = CDbl(wsHM.Range("B8").Value)
    ' Etichette costruite UNA sola volta dai parametri
    mAHT = "AHT alto (>" & NumStr(pctAHT * 100) & "° pct)"
    mFast = "Caso chiuso <" & NumStr(fastCaseMin) & " min"
    mBreakDay = "Break giornaliero >" & NumStr(maxBreakMin) & " min"
    mBreakSim = "Break simultaneo (>" & limitBreak & ")"
    mLogin = "Login in ritardo"
    mAC = "Available Cases fuori turno"
    Dim nameToEmail As Object, emailToName As Object, startByEmail As Object
    Dim aliasMap As Object, slotMap As Object, startByDay As Object
    Set nameToEmail = CreateObject("Scripting.Dictionary")
    Set emailToName = CreateObject("Scripting.Dictionary")
    Set startByEmail = CreateObject("Scripting.Dictionary")
    Set aliasMap = CreateObject("Scripting.Dictionary")
    Set slotMap = CreateObject("Scripting.Dictionary")
    Set startByDay = CreateObject("Scripting.Dictionary")
    LoadAnagrafica wsAn, nameToEmail, emailToName
    LoadHelperTurni wsHT, emailToName, startByEmail
    LoadAliases wsHM, aliasMap
    LoadSlots wsSlot, aliasMap, slotMap
    LoadTurniStarts wsTurni, startByDay
    Dim outRows As Collection: Set outRows = New Collection
    AddSFRules wsSF, nameToEmail, outRows, pctAHT, fastCaseMin
    AddATRules wsAT, emailToName, startByEmail, startByDay, slotMap, outRows, _
               limitBreak, maxBreakMin, tolMin, defaultStart
    WriteDettaglio wsDet, outRows
    SyncRecapHeaders wb          ' allinea intestazioni Malpractice Recap
    Application.Calculation = savedCalc
    Application.ScreenUpdating = True
    If Not SilentMode Then MsgBox "Dettaglio Malpractice rigenerato: " & outRows.Count & " righe.", vbInformation
    Exit Sub
CleanFail:
    Application.Calculation = savedCalc
    Application.ScreenUpdating = True
    If Not SilentMode Then MsgBox "Errore durante la rigenerazione: " & Err.Description, vbExclamation
    ' Senza propagare l'errore la macro fallirebbe in silenzio:
    ' il workbook verrebbe salvato con un Dettaglio Malpractice
    ' incompleto e nessuno se ne accorgerebbe.
    If SilentMode Then Err.Raise Err.Number, , Err.Description
End Sub

Private Sub LoadAnagrafica(ws As Worksheet, nameToEmail As Object, emailToName As Object)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim r As Long, nm As String, em As String
    For r = 2 To lastRow
        nm = Trim(CStr(ws.Cells(r, "A").Value))
        em = LCase(Trim(CStr(ws.Cells(r, "B").Value)))
        If nm <> "" Then
            nameToEmail(nm) = em
            If em <> "" Then emailToName(em) = nm
        End If
    Next r
End Sub

Private Sub LoadHelperTurni(ws As Worksheet, emailToName As Object, startByEmail As Object)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim r As Long, nm As String, em As String
    For r = 2 To lastRow
        nm = Trim(CStr(ws.Cells(r, "A").Value))
        em = LCase(Trim(CStr(ws.Cells(r, "B").Value)))
        If em <> "" Then
            If Not emailToName.Exists(em) And nm <> "" Then emailToName(em) = nm
            ' E "Ora inizio prevista" e' formattata h:mm -> usare .Value2 (IsNumeric(Date)=False)
            If IsNumeric(ws.Cells(r, "E").Value2) Then startByEmail(em) = CDbl(ws.Cells(r, "E").Value2)
        End If
    Next r
End Sub

' Orario di inizio atteso per ogni coppia AGENTE + GIORNO, letto dai turni LAVORA
Private Sub LoadTurniStarts(ws As Worksheet, startByDay As Object)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim r As Long, key As String, st As Double, dnum As Long, stato As String
    For r = 2 To lastRow
        stato = UCase(Trim(CStr(ws.Cells(r, "F").Value)))
        If stato = "LAVORA" Then
            ' E = Data, G = Inizio turno: formattate data/ora -> .Value2
            If IsNumeric(ws.Cells(r, "E").Value2) And IsNumeric(ws.Cells(r, "G").Value2) Then
                dnum = CLng(ws.Cells(r, "E").Value2)
                st = CDbl(ws.Cells(r, "G").Value2)
                key = NormKey(ws.Cells(r, "A").Value) & "|" & dnum
                If Not startByDay.Exists(key) Then
                    startByDay(key) = st                 ' primo turno del giorno
                ElseIf st < startByDay(key) Then
                    startByDay(key) = st                 ' turno piu' precoce se spezzato
                End If
            End If
        End If
    Next r
End Sub

Private Sub LoadAliases(ws As Worksheet, aliasMap As Object)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    Dim r As Long, src As String, dst As String
    For r = 3 To lastRow
        src = NormKey(ws.Cells(r, "D").Value)
        dst = NormKey(ws.Cells(r, "E").Value)
        If src <> "" And dst <> "" Then aliasMap(src) = dst
    Next r
End Sub

Private Sub LoadSlots(ws As Worksheet, aliasMap As Object, slotMap As Object)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim r As Long, k As String, d As Variant, status As String
    Dim arr(1 To 3) As Variant, mapKey As String
    For r = 2 To lastRow
        k = NormKey(ws.Cells(r, "A").Value)
        If aliasMap.Exists(k) Then k = aliasMap(k)
        d = ws.Cells(r, "B").Value2                  ' Data (formato data)
        status = UCase(Trim(CStr(ws.Cells(r, "E").Value)))
        If k <> "" And IsNumeric(d) Then
            arr(1) = ws.Cells(r, "C").Value2          ' Slot inizio (ora)
            arr(2) = ws.Cells(r, "D").Value2          ' Slot fine (ora)
            arr(3) = status
            mapKey = k & "|" & CLng(d)
            If Not slotMap.Exists(mapKey) Then Set slotMap(mapKey) = New Collection
            slotMap(mapKey).Add arr        ' multi-slot: accumula piu' turni/giorno
        End If
    Next r
End Sub

Private Sub AddSFRules(ws As Worksheet, nameToEmail As Object, outRows As Collection, _
                       pctAHT As Double, fastCaseMin As Double)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "BB").End(xlUp).Row
    Dim vals() As Double, n As Long, r As Long
    Dim aht As Variant, nm As String, em As String, caseNo As String
    Dim threshold As Double
    ReDim vals(1 To lastRow - 1)
    For r = 2 To lastRow
        aht = ws.Cells(r, "DY").Value
        nm = Trim(CStr(ws.Cells(r, "BB").Value))
        If nm <> "" And IsNumeric(aht) Then
            n = n + 1: vals(n) = CDbl(aht)
        End If
    Next r
    If n = 0 Then Exit Sub
    ReDim Preserve vals(1 To n)
    threshold = PercentileInc(vals, pctAHT)
    For r = 2 To lastRow
        nm = Trim(CStr(ws.Cells(r, "BB").Value))
        aht = ws.Cells(r, "DY").Value
        caseNo = Trim(CStr(ws.Cells(r, "AE").Value))
        em = ""
        If nameToEmail.Exists(nm) Then em = nameToEmail(nm)
        If nm <> "" And IsNumeric(aht) Then
            If CDbl(aht) > threshold Then
                AddRow outRows, nm, em, "-", mAHT, _
                       "Caso " & caseNo & ": AHT " & Round(CDbl(aht), 1) & " min", _
                       ">" & Round(threshold, 0) & " min"
            End If
            If CDbl(aht) < fastCaseMin Then
                AddRow outRows, nm, em, "-", mFast, _
                       "Caso " & caseNo & ": AHT " & Round(CDbl(aht), 1) & " min", _
                       "<" & NumStr(fastCaseMin) & " min"
            End If
        End If
    Next r
End Sub

Private Sub AddATRules(ws As Worksheet, emailToName As Object, startByEmail As Object, _
                       startByDay As Object, slotMap As Object, outRows As Collection, _
                       limitBreak As Long, maxBreakMin As Double, tolMin As Double, defaultStart As Double)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    Dim breakDaily As Object, firstStart As Object, acOutside As Object, breakByDate As Object
    Set breakDaily = CreateObject("Scripting.Dictionary")
    Set firstStart = CreateObject("Scripting.Dictionary")
    Set acOutside = CreateObject("Scripting.Dictionary")
    Set breakByDate = CreateObject("Scripting.Dictionary")
    Dim r As Long, em As String, nm As String, state As String
    Dim secs As Double, d As Variant, t As Variant, key As String
    Dim outsideSecs As Double
    For r = 2 To lastRow
        em = LCase(Trim(CStr(ws.Cells(r, "B").Value)))
        If em = "" Or Not emailToName.Exists(em) Then GoTo NextR
        nm = emailToName(em)
        state = Trim(CStr(ws.Cells(r, "F").Value))
        secs = ws.Cells(r, "K").Value2               ' Total Time in seconds
        d = ws.Cells(r, "P").Value2                  ' Data Milano (formula formattata data)
        t = ws.Cells(r, "Q").Value2                  ' Ora Milano
        If Not IsNumeric(d) Or Not IsNumeric(t) Or secs <= 0 Then GoTo NextR
        key = em & "|" & CLng(d)
        If Not firstStart.Exists(key) Then
            firstStart(key) = CDbl(t)
        ElseIf CDbl(t) < firstStart(key) Then
            firstStart(key) = CDbl(t)
        End If
        If state = "Break" Then
            AddSum breakDaily, key, secs
            AddBreakInterval breakByDate, CLng(d), em, CDbl(t), CDbl(t) + secs / 86400#
        End If
        If state = "Available Cases" Then
            outsideSecs = AvailableOutsideSeconds(nm, CLng(d), CDbl(t), secs, slotMap)
            If outsideSecs > 0 Then AddSum acOutside, key, outsideSecs
        End If
NextR:
    Next r
    AddBreakDailyRows breakDaily, emailToName, outRows, maxBreakMin
    AddLoginRows firstStart, emailToName, startByEmail, startByDay, outRows, defaultStart
    AddBreakSimRows breakByDate, emailToName, outRows, limitBreak
    AddAvailableRows acOutside, emailToName, outRows, tolMin
End Sub

Private Sub AddBreakDailyRows(d As Object, emailToName As Object, outRows As Collection, maxBreakMin As Double)
    Dim k As Variant, parts() As String
    For Each k In d.Keys
        If d(k) > maxBreakMin * 60 Then
            parts = Split(CStr(k), "|")
            AddRow outRows, emailToName(parts(0)), parts(0), CLng(parts(1)), _
                   mBreakDay, Round(d(k) / 60, 1) & " min di break", _
                   ">" & NumStr(maxBreakMin) & " min/gg"
        End If
    Next k
End Sub

Private Sub AddLoginRows(d As Object, emailToName As Object, startByEmail As Object, _
                         startByDay As Object, outRows As Collection, defaultStart As Double)
    Dim k As Variant, parts() As String, expected As Double
    Dim nm As String, dayKey As String
    For Each k In d.Keys
        parts = Split(CStr(k), "|")
        nm = emailToName(parts(0))
        dayKey = NormKey(nm) & "|" & parts(1)            ' agente | giorno
        If startByDay.Exists(dayKey) Then
            expected = startByDay(dayKey)                ' orario atteso DEL GIORNO
        ElseIf startByEmail.Exists(parts(0)) Then
            expected = startByEmail(parts(0))            ' fallback: inizio piu' precoce
        Else
            expected = defaultStart                      ' fallback finale
        End If
        If d(k) > expected Then
            AddRow outRows, nm, parts(0), CLng(parts(1)), _
                   mLogin, "Primo stato " & TimeLabel(d(k)) & " (previsto " & TimeLabel(expected) & ")", _
                   "Ora inizio turno (giorno)"
        End If
    Next k
End Sub

Private Sub AddBreakSimRows(breakByDate As Object, emailToName As Object, outRows As Collection, limitBreak As Long)
    Dim d As Variant, intervals As Collection, i As Long, j As Long
    Dim marked As Object, active As Collection, t As Double, it As Variant
    For Each d In breakByDate.Keys
        Set intervals = breakByDate(d)
        Set marked = CreateObject("Scripting.Dictionary")
        For i = 1 To intervals.Count
            t = intervals(i)(1)
            Set active = New Collection
            For j = 1 To intervals.Count
                If intervals(j)(1) <= t And intervals(j)(2) > t Then active.Add intervals(j)
            Next j
            If active.Count > limitBreak Then
                For Each it In active
                    marked(it(0)) = active.Count
                Next it
            End If
        Next i
        For Each it In marked.Keys
            AddRow outRows, emailToName(CStr(it)), CStr(it), CLng(d), _
                   mBreakSim, "In pausa con " & marked(it) & " agenti contemporaneamente", _
                   ">" & limitBreak & " in contemporanea"
        Next it
    Next d
End Sub

Private Sub AddAvailableRows(d As Object, emailToName As Object, outRows As Collection, tolMin As Double)
    Dim k As Variant, parts() As String
    For Each k In d.Keys
        If d(k) > tolMin * 60 Then
            parts = Split(CStr(k), "|")
            AddRow outRows, emailToName(parts(0)), parts(0), CLng(parts(1)), _
                   mAC, Round(d(k) / 60, 0) & " min di AC fuori slot", _
                   "Slot Only Cases, tol " & NumStr(tolMin) & " min"
        End If
    Next k
End Sub

Private Function AvailableOutsideSeconds(nm As String, d As Long, startT As Double, secs As Double, slotMap As Object) As Double
    Dim k As String: k = NormKey(nm) & "|" & d
    If Not slotMap.Exists(k) Then AvailableOutsideSeconds = 0: Exit Function
    Dim endT As Double: endT = startT + secs / 86400#
    Dim insideDays As Double: insideDays = 0
    Dim it As Variant, sl As Variant, status As String, sS As Variant, sE As Variant
    For Each it In slotMap(k)
        sl = it
        status = CStr(sl(3)): sS = sl(1): sE = sl(2)
        ' i turni "NO BOT" non sono finestre valide: il tempo li' resta "fuori"
        If status <> "NO BOT" And IsNumeric(sS) And IsNumeric(sE) Then
            insideDays = insideDays + Application.Max(0, Application.Min(endT, CDbl(sE)) - Application.Max(startT, CDbl(sS)))
        End If
    Next it
    If insideDays > secs / 86400# Then insideDays = secs / 86400#
    AvailableOutsideSeconds = Application.Max(0, secs - insideDays * 86400#)
End Function

Private Sub WriteDettaglio(ws As Worksheet, outRows As Collection)
    ws.Range("A:I").ClearContents
    ws.Range("A1:F1").Value = Array("Agente", "Email", "Data", "Tipo malpractice", "Dettaglio / valore", "Soglia / riferimento")
    ws.Range("H1:I1").Value = Array("Riepilogo per tipo", "")
    If outRows.Count > 0 Then
        Dim arr() As Variant, i As Long, r As Variant
        ReDim arr(1 To outRows.Count, 1 To 6)
        For i = 1 To outRows.Count
            r = outRows(i)
            arr(i, 1) = r(0): arr(i, 2) = r(1): arr(i, 3) = r(2)
            arr(i, 4) = r(3): arr(i, 5) = r(4): arr(i, 6) = r(5)
        Next i
        ws.Range("A2").Resize(outRows.Count, 6).Value = arr
    End If
    ws.Range("H2:H8").Value = Application.Transpose(Array( _
        mAHT, mFast, mBreakDay, mBreakSim, mLogin, mAC, "TOTALE"))
    Dim rr As Long
    For rr = 2 To 7
        ws.Range("I" & rr).Formula = "=COUNTIF($D:$D,H" & rr & ")"
    Next rr
    ws.Range("I8").Formula = "=SUM(I2:I7)"
    ws.Columns("A:I").AutoFit
End Sub

Private Sub SyncRecapHeaders(wb As Workbook)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = wb.Worksheets("Malpractice Recap")
    If ws Is Nothing Then Exit Sub
    ws.Range("B2:G2").Value = Array(mAHT, mFast, mBreakDay, mBreakSim, mLogin, mAC)
    On Error GoTo 0
End Sub

Private Sub AddRow(outRows As Collection, nm As String, em As String, d As Variant, typ As String, detail As String, threshold As String)
    Dim r(0 To 5) As Variant
    r(0) = nm: r(1) = em: r(2) = d: r(3) = typ: r(4) = detail: r(5) = threshold
    outRows.Add r
End Sub

Private Sub AddSum(d As Object, k As String, v As Double)
    If d.Exists(k) Then d(k) = d(k) + v Else d(k) = v
End Sub

Private Sub AddBreakInterval(d As Object, dayKey As Long, em As String, s As Double, e As Double)
    Dim item(0 To 2) As Variant
    item(0) = em: item(1) = s: item(2) = e
    If Not d.Exists(CStr(dayKey)) Then Set d(CStr(dayKey)) = New Collection
    d(CStr(dayKey)).Add item
End Sub

Private Function PercentileInc(vals() As Double, p As Double) As Double
    Dim i As Long, j As Long, tmp As Double
    For i = LBound(vals) To UBound(vals) - 1
        For j = i + 1 To UBound(vals)
            If vals(j) < vals(i) Then tmp = vals(i): vals(i) = vals(j): vals(j) = tmp
        Next j
    Next i
    Dim n As Long, k As Double, f As Long, c As Long
    n = UBound(vals) - LBound(vals) + 1
    k = (n - 1) * p + 1
    f = Int(k)
    c = Application.WorksheetFunction.RoundUp(k, 0)
    If f = c Then PercentileInc = vals(f) Else PercentileInc = vals(f) + (k - f) * (vals(c) - vals(f))
End Function

Private Function TimeLabel(v As Double) As String
    TimeLabel = Format(v, "hh:mm")
End Function

Private Function NumStr(v As Double) As String
    NumStr = Replace(Format(v, "0.####"), ",", ".")
End Function

Private Function NormKey(v As Variant) As String
    Dim s As String
    s = LCase(Trim(CStr(v)))
    s = Replace(s, "à", "a"): s = Replace(s, "è", "e"): s = Replace(s, "é", "e")
    s = Replace(s, "ì", "i"): s = Replace(s, "ò", "o"): s = Replace(s, "ù", "u")
    s = Replace(s, "'", ""): s = Replace(s, "’", "")
    s = WorksheetFunction.Trim(s)
    NormKey = s
End Function
