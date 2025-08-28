Option Explicit

' =========================
' Entry point (shows in Macros dialog)
' =========================
Public Sub SplitTransactionsToSheets_Run()
    Const START_ROW As Long = 6
    Const MARKER1 As String = "Kundedokumenter totalt"
    Const MARKER2 As String = "Kunde dokumenter totalt"
    Const COL_I As Long = 9   ' I = Beløp
    Const COL_J As Long = 10  ' J = Saldo
    Const MAX_LOOKAHEAD_TOTALS As Long = 6
    Const ALLOW_MINUS_AS_NEGATIVE As Boolean = True  ' set False to require parentheses-only
    
    Dim wsSrc As Worksheet
    Set wsSrc = ActiveSheet
    If wsSrc.name = "Negativ" Or wsSrc.name = "Avvik" Or wsSrc.name = "Logg" Then
        MsgBox "Kjør makroen fra arket med rådata.", vbExclamation
        Exit Sub
    End If
    
    Dim lastRow As Long, lastCol As Long
    lastRow = STS_LastUsedRow(wsSrc)
    lastCol = STS_LastUsedCol(wsSrc)
    If lastRow < START_ROW Then
        MsgBox "Ingen data fra rad " & START_ROW & " og nedover.", vbInformation
        Exit Sub
    End If
    
    Dim wsNeg As Worksheet, wsAvv As Worksheet, wsLog As Worksheet
    Set wsNeg = STS_SheetOrCreate("Negativ"): wsNeg.Cells.Clear
    Set wsAvv = STS_SheetOrCreate("Avvik"):   wsAvv.Cells.Clear
    Set wsLog = STS_SheetOrCreate("Logg"):    wsLog.Cells.Clear
    wsLog.Range("A1:J1").Value = Array("Blokk#", "Rad-range", "Marker-rad", "Kontout-label", "Eval-rad", _
                                       "I.Text", "J.Text", "I.Val", "J.Val", "Resultat")
    
    ' Find marker rows that end blocks
    Dim marks As Collection: Set marks = New Collection
    Dim r As Long
    For r = START_ROW To lastRow
        If STS_RowContainsAny(wsSrc, r, lastCol, MARKER1, MARKER2) Then
            marks.Add r
        End If
    Next r
    
    Dim destNeg As Long: destNeg = 1
    Dim destAvv As Long: destAvv = 1
    Dim logRow As Long: logRow = 2
    Dim blockStart As Long: blockStart = START_ROW
    Dim i As Long, blockIdx As Long: blockIdx = 0
    
    If marks.Count = 0 Then
        blockIdx = 1
        STS_ProcessBlock wsSrc, blockIdx, blockStart, lastRow, lastCol, _
                         COL_I, COL_J, MAX_LOOKAHEAD_TOTALS, ALLOW_MINUS_AS_NEGATIVE, _
                         wsNeg, destNeg, wsAvv, destAvv, wsLog, logRow
    Else
        For i = 1 To marks.Count
            blockIdx = blockIdx + 1
            STS_ProcessBlock wsSrc, blockIdx, blockStart, CLng(marks(i)), lastCol, _
                             COL_I, COL_J, MAX_LOOKAHEAD_TOTALS, ALLOW_MINUS_AS_NEGATIVE, _
                             wsNeg, destNeg, wsAvv, destAvv, wsLog, logRow
            blockStart = CLng(marks(i)) + 1
        Next i
        If blockStart <= lastRow Then
            blockIdx = blockIdx + 1
            STS_ProcessBlock wsSrc, blockIdx, blockStart, lastRow, lastCol, _
                             COL_I, COL_J, MAX_LOOKAHEAD_TOTALS, ALLOW_MINUS_AS_NEGATIVE, _
                             wsNeg, destNeg, wsAvv, destAvv, wsLog, logRow
        End If
    End If
    
    MsgBox "Ferdig. Sjekk arkene 'Negativ', 'Avvik' og 'Logg' for detaljer."
End Sub

' =========================
' Per-block logic
' =========================
Private Sub STS_ProcessBlock(wsSrc As Worksheet, ByVal blockNo As Long, _
                             ByVal firstRow As Long, ByVal lastRow As Long, ByVal lastCol As Long, _
                             ByVal colI As Long, ByVal colJ As Long, _
                             ByVal maxTotalsLookahead As Long, ByVal allowMinus As Boolean, _
                             wsNeg As Worksheet, ByRef destNeg As Long, _
                             wsAvv As Worksheet, ByRef destAvv As Long, _
                             wsLog As Worksheet, ByRef logRow As Long)
    ' 1) Find "Kontoutskrift total(t)" label row (fuzzy: must contain "kontoutskrift" + ("totalt" or "total"))
    Dim kontoLabelRow As Long
    kontoLabelRow = STS_FindKontoutRow(wsSrc, firstRow, lastRow, lastCol)
    If kontoLabelRow = 0 Then
        ' Log and skip
        wsLog.Cells(logRow, 1).Resize(1, 10).Value = _
            Array(blockNo, firstRow & "-" & lastRow, "-", "-", "-", "-", "-", "-", "-", "Ingen 'Kontoutskrift'")
        logRow = logRow + 1
        Exit Sub
    End If
    
    ' 2) Pick the evaluation row (same row or up to +N rows if numbers are below)
    Dim evalRow As Long: evalRow = kontoLabelRow
    Dim k As Long
    For k = 0 To maxTotalsLookahead
        Dim rr As Long: rr = kontoLabelRow + k
        If rr > lastRow Then Exit For
        If (STS_LooksNumeric(wsSrc.Cells(rr, colI).Text) Or IsNumeric(wsSrc.Cells(rr, colI).Value2)) _
        And (STS_LooksNumeric(wsSrc.Cells(rr, colJ).Text) Or IsNumeric(wsSrc.Cells(rr, colJ).Value2)) Then
            evalRow = rr
            Exit For
        End If
    Next k
    
    ' 3) Read values/text
    Dim txtI As String, txtJ As String
    Dim vI As Variant, vJ As Variant
    txtI = STS_ToSafeString(wsSrc.Cells(evalRow, colI).Text)
    txtJ = STS_ToSafeString(wsSrc.Cells(evalRow, colJ).Text)
    vI = wsSrc.Cells(evalRow, colI).Value2
    vJ = wsSrc.Cells(evalRow, colJ).Value2
    
    ' 4) Negativ?
    Dim isNeg As Boolean
    isNeg = STS_HasParentheses(txtJ)
    If Not isNeg And allowMinus And IsNumeric(vJ) Then isNeg = (CDbl(vJ) < 0#)
    
    ' 5) Avvik? (compare totals I vs J)
    Dim isAvvik As Boolean
    If IsNumeric(vI) And IsNumeric(vJ) Then
        isAvvik = (Abs(CDbl(vI) - CDbl(vJ)) > 0.005)
    Else
        Dim pI As Variant, pJ As Variant
        pI = STS_ParseAmount(txtI)
        pJ = STS_ParseAmount(txtJ)
        If IsNumeric(pI) And IsNumeric(pJ) Then
            isAvvik = (Abs(CDbl(pI) - CDbl(pJ)) > 0.005)
        Else
            isAvvik = (STS_NormalizeForCompare(txtI) <> STS_NormalizeForCompare(txtJ))
        End If
    End If
    
    ' 6) Copy blocks independently
    If isNeg Then STS_CopyBlock wsSrc, wsNeg, firstRow, lastRow, lastCol, destNeg
    If isAvvik Then STS_CopyBlock wsSrc, wsAvv, firstRow, lastRow, lastCol, destAvv
    
    ' 7) Log
    Dim res As String
    res = IIf(isNeg, "NEG", "") & IIf(isAvvik, IIf(isNeg, " + ", "") & "AVVIK", "")
    If res = "" Then res = "(ingen)"
    wsLog.Cells(logRow, 1).Resize(1, 10).Value = _
        Array(blockNo, firstRow & "-" & lastRow, _
              STS_FindMarkerRowInRange(wsSrc, firstRow, lastRow, lastCol), _
              kontoLabelRow, evalRow, _
              txtI, txtJ, _
              IIf(IsNumeric(vI), CDbl(vI), "n/a"), _
              IIf(IsNumeric(vJ), CDbl(vJ), "n/a"), res)
    logRow = logRow + 1
End Sub

' =========================
' Helpers (namespaced STS_)
' =========================

' Find the row containing "kontoutskrift" and ( "totalt" OR "total" ) anywhere on the row
Private Function STS_FindKontoutRow(ws As Worksheet, ByVal rStart As Long, ByVal rEnd As Long, ByVal lastCol As Long) As Long
    Dim rr As Long, rowText As String
    For rr = rStart To rEnd
        rowText = STS_ConcatRowText(ws, rr, lastCol)
        If rowText <> "" Then
            If InStr(1, rowText, "kontoutskrift", vbTextCompare) > 0 _
               And (InStr(1, rowText, "totalt", vbTextCompare) > 0 _
                 Or InStr(1, rowText, "total", vbTextCompare) > 0) Then
                STS_FindKontoutRow = rr
                Exit Function
            End If
        End If
    Next rr
    STS_FindKontoutRow = 0
End Function

' Return the first marker row (for logging only)
Private Function STS_FindMarkerRowInRange(ws As Worksheet, ByVal rStart As Long, ByVal rEnd As Long, ByVal lastCol As Long) As Variant
    Dim rr As Long
    For rr = rStart To rEnd
        If STS_RowContainsAny(ws, rr, lastCol, "Kundedokumenter totalt", "Kunde dokumenter totalt") Then
            STS_FindMarkerRowInRange = rr
            Exit Function
        End If
    Next rr
    STS_FindMarkerRowInRange = "-"
End Function

' Concatenate normalized text of a whole row
Private Function STS_ConcatRowText(ws As Worksheet, ByVal r As Long, ByVal lastCol As Long) As String
    Dim c As Long, s As String
    For c = 1 To lastCol
        s = s & " " & STS_NormalizeText(ws.Cells(r, c).Text)
    Next c
    s = Trim$(s)
    STS_ConcatRowText = s
End Function

' Safe string coercion
Private Function STS_ToSafeString(ByVal v As Variant) As String
    On Error Resume Next
    STS_ToSafeString = CStr(v & vbNullString)
End Function

' Normalize cell text (strip NBSP/thin spaces, punctuation; VBA Trim)
Private Function STS_NormalizeText(ByVal s As Variant) As String
    Dim t As String
    t = STS_ToSafeString(s)
    If LenB(t) = 0 Then
        STS_NormalizeText = ""
        Exit Function
    End If
    t = Replace(t, ChrW(160), " ")  ' NBSP
    t = Replace(t, ChrW(8239), " ") ' NARROW NBSP
    t = Replace(t, ChrW(8201), " ") ' THIN SPACE
    t = Replace(t, ":", "")
    t = Replace(t, ".", "")
    t = Trim$(t)
    STS_NormalizeText = t
End Function

Private Function STS_NormalizeForCompare(ByVal s As Variant) As String
    Dim t As String
    t = STS_NormalizeText(s)
    t = Replace(t, " ", "")
    t = Replace(t, vbTab, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, "kr", "", 1, -1, vbTextCompare)
    t = Replace(t, "nok", "", 1, -1, vbTextCompare)
    t = Replace(t, ".", "")
    STS_NormalizeForCompare = t
End Function

' Split markers
Private Function STS_RowContainsAny(ws As Worksheet, ByVal r As Long, ByVal lastCol As Long, ParamArray needles() As Variant) As Boolean
    Dim c As Long, cellText As String, nrmCell As String
    Dim needle As Variant, nrmNeedle As String
    For c = 1 To lastCol
        cellText = STS_ToSafeString(ws.Cells(r, c).Text)
        If LenB(cellText) > 0 Then
            nrmCell = STS_NormalizeText(cellText)
            For Each needle In needles
                nrmNeedle = STS_NormalizeText(needle)
                If LenB(nrmNeedle) > 0 Then
                    If InStr(1, nrmCell, nrmNeedle, vbTextCompare) > 0 Then
                        STS_RowContainsAny = True
                        Exit Function
                    End If
                End If
            Next needle
        End If
    Next c
    STS_RowContainsAny = False
End Function

' Number-ish text?
Private Function STS_LooksNumeric(ByVal s As String) As Boolean
    Dim t As String: t = Trim$(s)
    t = Replace(t, ChrW(160), " ")
    t = Replace(t, ChrW(8239), " ")
    t = Replace(t, ChrW(8201), " ")
    If t = "" Then Exit Function
    Dim i As Long, ch As String, hasDigit As Boolean
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch Like "[0-9]" Then hasDigit = True
    Next i
    STS_LooksNumeric = hasDigit
End Function

' Parentheses detection (accounting format)
Private Function STS_HasParentheses(ByVal s As String) As Boolean
    Dim t As String: t = s
    t = Replace(t, ChrW(160), " ")
    t = Replace(t, " ", "")
    STS_HasParentheses = (InStr(1, t, "(") > 0 And InStr(1, t, ")") > 0)
End Function

' Parse European-style amounts; respects parentheses anywhere in the text
Private Function STS_ParseAmount(ByVal s As Variant) As Variant
    Dim t As String: t = STS_ToSafeString(s)
    If LenB(t) = 0 Then STS_ParseAmount = Empty: Exit Function
    
    t = Replace(t, ChrW(160), " ")
    t = Replace(t, ChrW(8239), " ")
    t = Replace(t, "kr", "", 1, -1, vbTextCompare)
    t = Replace(t, "nok", "", 1, -1, vbTextCompare)
    
    Dim parens As Boolean
    parens = (InStr(1, t, "(") > 0 And InStr(1, t, ")") > 0)
    t = Replace(t, "(", ""): t = Replace(t, ")", "")
    
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "," Or ch = "." Or ch = "-" Then out = out & ch
    Next i
    
    Dim hasComma As Boolean, hasDot As Boolean
    hasComma = (InStr(1, out, ",") > 0)
    hasDot = (InStr(1, out, ".") > 0)
    If hasComma And hasDot Then
        out = Replace(out, ".", "")   ' thousands dot
        out = Replace(out, ",", ".")  ' decimal comma
    ElseIf hasComma Then
        out = Replace(out, ",", ".")
    End If
    
    If out = "" Or out = "-" Then STS_ParseAmount = Empty: Exit Function
    
    On Error GoTo notnum
    Dim v As Double: v = CDbl(out)
    If parens Then v = -Abs(v)
    STS_ParseAmount = v
    Exit Function
notnum:
    STS_ParseAmount = Empty
End Function

' Copy a whole block to a destination sheet (values only)
Private Sub STS_CopyBlock(wsSrc As Worksheet, wsDst As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal lastCol As Long, ByRef destRow As Long)
    Dim rowsCount As Long: rowsCount = lastRow - firstRow + 1
    If rowsCount <= 0 Then Exit Sub
    wsDst.Cells(destRow, 1).Value = "--- Rader " & firstRow & "–" & lastRow & " ---"
    destRow = destRow + 1
    wsDst.Cells(destRow, 1).Resize(rowsCount, lastCol).Value = _
        wsSrc.Range(wsSrc.Cells(firstRow, 1), wsSrc.Cells(lastRow, lastCol)).Value
    destRow = destRow + rowsCount + 1
End Sub

' Last used row/col
Private Function STS_LastUsedRow(ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    STS_LastUsedRow = IIf(c Is Nothing, 1, c.Row)
End Function

Private Function STS_LastUsedCol(ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    STS_LastUsedCol = IIf(c Is Nothing, 1, c.Column)
End Function

' Create-or-get sheet
Private Function STS_SheetOrCreate(ByVal name As String) As Worksheet
    On Error Resume Next
    Set STS_SheetOrCreate = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If STS_SheetOrCreate Is Nothing Then
        Set STS_SheetOrCreate = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        STS_SheetOrCreate.name = name
    End If
End Function


