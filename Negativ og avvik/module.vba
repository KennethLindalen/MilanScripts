Option Explicit

' ======================================================
' Entry point
' ======================================================
Public Sub SplitTransactionsToSheets_FinalV6()
    Const START_ROW As Long = 6
    Const MARKER1 As String = "Kundedokumenter totalt"
    Const MARKER2 As String = "Kunde dokumenter totalt"
    Const COL_I As Long = 9   ' I = Beløp
    Const COL_J As Long = 10  ' J = Saldo
    Const LOOK_ABOVE As Long = 3
    Const LOOK_BELOW As Long = 6
    Const TOL As Double = 0.005
    Const ALLOW_MINUS_AS_NEGATIVE As Boolean = True  ' set False to require parentheses-only
    
    Dim wsSrc As Worksheet
    Set wsSrc = ActiveSheet
    If wsSrc.Name = "Negativ" Or wsSrc.Name = "Avvik" Or wsSrc.Name = "Logg" Then
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
    wsLog.Range("A1:N1").Value = Array( _
        "Blokk#", "Rad-range", "Marker-rad", _
        "Kontout-label", "EvalRow (I&J)", _
        "I.Text", "J.Text", "I.Num", "J.Num", _
        "Negativ?", "Avvik?", "Why(I)", "Why(J)", "FallbackNegScan" _
    )
    
    ' Find split markers
    Dim marks As Collection: Set marks = New Collection
    Dim r As Long
    For r = START_ROW To lastRow
        If STS_RowContainsAny(wsSrc, r, lastCol, MARKER1, MARKER2) Then marks.Add r
    Next r
    
    Dim destNeg As Long: destNeg = 1
    Dim destAvv As Long: destAvv = 1
    Dim logRow As Long: logRow = 2
    Dim blockStart As Long: blockStart = START_ROW
    Dim i As Long, blockIdx As Long: blockIdx = 0
    
    If marks.Count = 0 Then
        blockIdx = 1
        STS_ProcessBlock wsSrc, blockIdx, blockStart, lastRow, lastCol, _
                         COL_I, COL_J, LOOK_ABOVE, LOOK_BELOW, TOL, ALLOW_MINUS_AS_NEGATIVE, _
                         wsNeg, destNeg, wsAvv, destAvv, wsLog, logRow
    Else
        For i = 1 To marks.Count
            blockIdx = blockIdx + 1
            STS_ProcessBlock wsSrc, blockIdx, blockStart, CLng(marks(i)), lastCol, _
                             COL_I, COL_J, LOOK_ABOVE, LOOK_BELOW, TOL, ALLOW_MINUS_AS_NEGATIVE, _
                             wsNeg, destNeg, wsAvv, destAvv, wsLog, logRow
            blockStart = CLng(marks(i)) + 1
        Next i
        If blockStart <= lastRow Then
            blockIdx = blockIdx + 1
            STS_ProcessBlock wsSrc, blockIdx, blockStart, lastRow, lastCol, _
                             COL_I, COL_J, LOOK_ABOVE, LOOK_BELOW, TOL, ALLOW_MINUS_AS_NEGATIVE, _
                             wsNeg, destNeg, wsAvv, destAvv, wsLog, logRow
        End If
    End If
    
    MsgBox "Ferdig. Sjekk 'Negativ', 'Avvik' og 'Logg'."
End Sub

' ======================================================
' Per-block logic (use the SAME totals row for Avvik & Negativ)
' ======================================================
Private Sub STS_ProcessBlock(wsSrc As Worksheet, ByVal blockNo As Long, _
                             ByVal firstRow As Long, ByVal lastRow As Long, ByVal lastCol As Long, _
                             ByVal colI As Long, ByVal colJ As Long, _
                             ByVal lookAbove As Long, ByVal lookBelow As Long, ByVal tol As Double, _
                             ByVal allowMinus As Boolean, _
                             wsNeg As Worksheet, ByRef destNeg As Long, _
                             wsAvv As Worksheet, ByRef destAvv As Long, _
                             wsLog As Worksheet, ByRef logRow As Long)
    ' Anchor: kontoutskrift total(t)
    Dim kontoLabelRow As Long
    kontoLabelRow = STS_FindKontoutRow(wsSrc, firstRow, lastRow, lastCol)
    
    ' Find ONE evaluation row where BOTH I and J are numeric (preferred totals row)
    Dim evalRow As Long
    evalRow = STS_FindTotalsRowBoth(wsSrc, kontoLabelRow, firstRow, lastRow, colI, colJ, lookAbove, lookBelow)
    
    ' Read text / numbers from evalRow (if any)
    Dim txtI As String, txtJ As String
    Dim numI As Double, numJ As Double
    Dim okI As Boolean, okJ As Boolean
    Dim whyI As String, whyJ As String
    If evalRow > 0 Then
        txtI = STS_ToSafeString(wsSrc.Cells(evalRow, colI).Text)
        txtJ = STS_ToSafeString(wsSrc.Cells(evalRow, colJ).Text)
        okI = STS_GetNumeric(wsSrc.Cells(evalRow, colI), txtI, numI, whyI)
        okJ = STS_GetNumeric(wsSrc.Cells(evalRow, colJ), txtJ, numJ, whyJ)
    Else
        txtI = "": txtJ = ""
        okI = False: okJ = False
        whyI = "no evalRow": whyJ = "no evalRow"
    End If
    
    ' -------- Negativ (use SAME totals row first) --------
    Dim isNeg As Boolean, negFallbackNote As String
    negFallbackNote = ""
    
    If evalRow > 0 Then
        ' Parentheses OR (value < 0 when allowed)
        isNeg = STS_HasParentheses(txtJ)
        If Not isNeg And allowMinus And okJ Then isNeg = (numJ < 0#)
    End If
    
    ' Fallback: if still not negative (or no evalRow), scan a small window around the label for a J that is negative/parenthesized
    If Not isNeg Then
        Dim rowForNegJ As Long
        rowForNegJ = STS_FindRowForNegJ(wsSrc, kontoLabelRow, firstRow, lastRow, colJ, lookAbove, lookBelow)
        If rowForNegJ > 0 Then
            Dim fTxtJ As String, fNumJ As Double, fOkJ As Boolean, fWhy As String
            fTxtJ = STS_ToSafeString(wsSrc.Cells(rowForNegJ, colJ).Text)
            fOkJ = STS_GetNumeric(wsSrc.Cells(rowForNegJ, colJ), fTxtJ, fNumJ, fWhy)
            If STS_HasParentheses(fTxtJ) Or (allowMinus And fOkJ And fNumJ < 0#) Then
                isNeg = True
                negFallbackNote = "fallback@" & CStr(rowForNegJ)
            End If
        End If
    End If
    
    ' -------- Avvik (only if BOTH numeric on the SAME totals row) --------
    Dim isAvvik As Boolean
    If okI And okJ Then
        isAvvik = (Abs(numI - numJ) > tol)
    Else
        isAvvik = False
    End If
    
    ' Copy independently
    If isNeg Then STS_CopyBlock wsSrc, wsNeg, firstRow, lastRow, lastCol, destNeg
    If isAvvik Then STS_CopyBlock wsSrc, wsAvv, firstRow, lastRow, lastCol, destAvv
    
    ' Log
    wsLog.Cells(logRow, 1).Resize(1, 14).Value = _
        Array(blockNo, firstRow & "-" & lastRow, _
              STS_FindMarkerRowInRange(wsSrc, firstRow, lastRow, lastCol), _
              IIf(kontoLabelRow > 0, kontoLabelRow, "-"), _
              IIf(evalRow > 0, evalRow, "-"), _
              txtI, txtJ, _
              IIf(okI, numI, "n/a"), IIf(okJ, numJ, "n/a"), _
              IIf(isNeg, "Ja", "Nei"), IIf(isAvvik, "Ja", "Nei"), _
              whyI, whyJ, negFallbackNote)
    logRow = logRow + 1
End Sub

' ======================================================
' Totals row finders
' ======================================================
' Find a row near the label where BOTH I and J are numeric (used for BOTH Avvik and Negativ)
Private Function STS_FindTotalsRowBoth(ws As Worksheet, ByVal labelRow As Long, ByVal blockStart As Long, ByVal blockEnd As Long, _
                                       ByVal colI As Long, ByVal colJ As Long, _
                                       ByVal lookAbove As Long, ByVal lookBelow As Long) As Long
    If labelRow = 0 Then STS_FindTotalsRowBoth = 0: Exit Function
    Dim rFrom As Long, rTo As Long, r As Long
    rFrom = IIf(labelRow - lookAbove < blockStart, blockStart, labelRow - lookAbove)
    rTo = IIf(labelRow + lookBelow > blockEnd, blockEnd, labelRow + lookBelow)
    For r = rFrom To rTo
        If (IsNumeric(ws.Cells(r, colI).Value2) Or STS_LooksNumeric(ws.Cells(r, colI).Text)) _
        And (IsNumeric(ws.Cells(r, colJ).Value2) Or STS_LooksNumeric(ws.Cells(r, colJ).Text)) Then
            STS_FindTotalsRowBoth = r
            Exit Function
        End If
    Next r
    STS_FindTotalsRowBoth = 0
End Function

' Find any nearby row where J is clearly negative or parenthesized (fallback just for Negativ)
Private Function STS_FindRowForNegJ(ws As Worksheet, ByVal labelRow As Long, ByVal blockStart As Long, ByVal blockEnd As Long, _
                                    ByVal colJ As Long, ByVal lookAbove As Long, ByVal lookBelow As Long) As Long
    If labelRow = 0 Then STS_FindRowForNegJ = 0: Exit Function
    Dim rFrom As Long, rTo As Long, r As Long
    rFrom = IIf(labelRow - lookAbove < blockStart, blockStart, labelRow - lookAbove)
    rTo = IIf(labelRow + lookBelow > blockEnd, blockEnd, labelRow + lookBelow)
    For r = rFrom To rTo
        Dim t As String: t = ws.Cells(r, colJ).Text
        If STS_HasParentheses(t) Then STS_FindRowForNegJ = r: Exit Function
        If IsNumeric(ws.Cells(r, colJ).Value2) Then
            If CDbl(ws.Cells(r, colJ).Value2) < 0# Then STS_FindRowForNegJ = r: Exit Function
        End If
        If STS_LooksNumeric(t) Then
            Dim vv As Variant: vv = STS_ParseAmount(t)
            If IsNumeric(vv) Then If CDbl(vv) < 0# Then STS_FindRowForNegJ = r: Exit Function
        End If
    Next r
    STS_FindRowForNegJ = 0
End Function

' ======================================================
' Helpers (namespaced STS_)
' ======================================================
Private Function STS_FindKontoutRow(ws As Worksheet, ByVal rStart As Long, ByVal rEnd As Long, ByVal lastCol As Long) As Long
    Dim rr As Long, rowText As String
    For rr = rStart To rEnd
        rowText = STS_ConcatRowText(ws, rr, lastCol)
        If rowText <> "" Then
            If InStr(1, rowText, "kontoutskrift", vbTextCompare) > 0 _
               And (InStr(1, rowText, "totalt", vbTextCompare) > 0 _
                 Or InStr(1, rowText, "total", vbTextCompare) > 0) Then
                STS_FindKontoutRow = rr: Exit Function
            End If
        End If
    Next rr
    STS_FindKontoutRow = 0
End Function

Private Function STS_FindMarkerRowInRange(ws As Worksheet, ByVal rStart As Long, ByVal rEnd As Long, ByVal lastCol As Long) As Variant
    Dim rr As Long
    For rr = rStart To rEnd
        If STS_RowContainsAny(ws, rr, lastCol, "Kundedokumenter totalt", "Kunde dokumenter totalt") Then
            STS_FindMarkerRowInRange = rr: Exit Function
        End If
    Next rr
    STS_FindMarkerRowInRange = "-"
End Function

Private Function STS_ConcatRowText(ws As Worksheet, ByVal r As Long, ByVal lastCol As Long) As String
    Dim c As Long, s As String
    For c = 1 To lastCol
        s = s & " " & STS_NormalizeText(ws.Cells(r, c).Text)
    Next c
    STS_ConcatRowText = Trim$(s)
End Function

Private Function STS_ToSafeString(ByVal v As Variant) As String
    On Error Resume Next
    STS_ToSafeString = CStr(v & vbNullString)
End Function

Private Function STS_NormalizeText(ByVal s As Variant) As String
    Dim t As String
    t = STS_ToSafeString(s)
    If LenB(t) = 0 Then STS_NormalizeText = "": Exit Function
    t = Replace(t, ChrW(160), " ")
    t = Replace(t, ChrW(8239), " ")
    t = Replace(t, ChrW(8201), " ")
    t = Replace(t, ":", "")
    t = Replace(t, ".", "")
    STS_NormalizeText = Trim$(t)
End Function

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
                        STS_RowContainsAny = True: Exit Function
                    End If
                End If
            Next needle
        End If
    Next c
    STS_RowContainsAny = False
End Function

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

Private Function STS_HasParentheses(ByVal s As String) As Boolean
    Dim t As String: t = s
    t = Replace(t, ChrW(160), " ")
    t = Replace(t, " ", "")
    STS_HasParentheses = (InStr(1, t, "(") > 0 And InStr(1, t, ")") > 0)
End Function

' Prefer Value2 if numeric; else parse (handles parentheses & trailing minus)
Private Function STS_GetNumeric(ByVal cell As Range, ByVal txt As String, ByRef outVal As Double, ByRef why As String) As Boolean
    If IsNumeric(cell.Value2) Then
        outVal = CDbl(cell.Value2)
        STS_GetNumeric = True
        why = "Value2"
        Exit Function
    End If
    Dim p As Variant
    p = STS_ParseAmount(txt)
    If IsNumeric(p) Then
        outVal = CDbl(p)
        STS_GetNumeric = True
        why = "parsed"
    Else
        STS_GetNumeric = False
        why = "non-numeric"
    End If
End Function

' Parse European amounts; handles parentheses and trailing minus ("123,45-")
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
    
    ' trailing minus -> leading minus
    If Right$(out, 1) = "-" Then out = "-" & Left$(out, Len(out) - 1)
    
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

' Copy whole block (values only)
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
        STS_SheetOrCreate.Name = name
    End If
End Function
