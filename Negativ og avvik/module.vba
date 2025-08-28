Option Explicit

Public Sub SplitTransactionsToSheets()
    Const START_ROW As Long = 6
    Const MARKER As String = "Kunde dokumenter totalt:"
    Const KONTO As String = "Kontoutskrift totalt"
    Const DEBUG_LOG As Boolean = True
    
    Dim wsSrc As Worksheet
    Set wsSrc = ActiveSheet
    If wsSrc.Name = "Negativ" Or wsSrc.Name = "Avvik" Or wsSrc.Name = "Logg" Then
        MsgBox "Run this macro from the sheet that contains the raw data.", vbExclamation
        Exit Sub
    End If
    
    Dim lastRow As Long, lastCol As Long
    lastRow = LastUsedRow(wsSrc)
    lastCol = LastUsedCol(wsSrc)
    If lastRow < START_ROW Then
        MsgBox "No data at or after row " & START_ROW, vbInformation
        Exit Sub
    End If
    
    Dim wsNeg As Worksheet, wsAvv As Worksheet, wsLog As Worksheet
    Set wsNeg = SheetOrCreate("Negativ"): wsNeg.Cells.Clear
    Set wsAvv = SheetOrCreate("Avvik"):   wsAvv.Cells.Clear
    If DEBUG_LOG Then
        Set wsLog = SheetOrCreate("Logg"): wsLog.Cells.Clear
        wsLog.Range("A1:H1").Value = Array("Blokk#", "Rad-range", "Konto-label-rad", "Sammenligningsrad", _
                                           "I (Text)", "J (Text)", "I (Value2)", "J (Value2)")
    End If
    
    Dim destNeg As Long: destNeg = 1
    Dim destAvv As Long: destAvv = 1
    Dim logRow As Long: logRow = 2
    
    Dim blockStart As Long: blockStart = START_ROW
    Dim r As Long, blockIdx As Long: blockIdx = 0
    
    For r = START_ROW To lastRow
        If RowContains(wsSrc, r, lastCol, MARKER) Then
            blockIdx = blockIdx + 1
            ProcessBlock wsSrc, blockIdx, blockStart, r, lastCol, KONTO, _
                         wsNeg, destNeg, wsAvv, destAvv, wsLog, logRow, DEBUG_LOG
            blockStart = r + 1
        End If
    Next r
    
    If blockStart <= lastRow Then
        blockIdx = blockIdx + 1
        ProcessBlock wsSrc, blockIdx, blockStart, lastRow, lastCol, KONTO, _
                     wsNeg, destNeg, wsAvv, destAvv, wsLog, logRow, DEBUG_LOG
    End If
    
    MsgBox "Ferdig." & vbCrLf & _
           "Skrevet -> Negativ: " & (destNeg - 1) & " rader, Avvik: " & (destAvv - 1) & " rader.", vbInformation
End Sub

Private Sub ProcessBlock(wsSrc As Worksheet, ByVal blockNo As Long, _
                         ByVal firstRow As Long, ByVal lastRow As Long, ByVal lastCol As Long, _
                         ByVal kontoNeedle As String, _
                         wsNeg As Worksheet, ByRef destNeg As Long, _
                         wsAvv As Worksheet, ByRef destAvv As Long, _
                         wsLog As Worksheet, ByRef logRow As Long, ByVal doLog As Boolean)
    Const MAX_LOOKAHEAD As Long = 6     ' rows to scan below the label to find the numeric pair
    Const COL_I As Long = 9            ' column I
    Const COL_J As Long = 10           ' column J
    
    Dim kontoLabelRow As Long
    kontoLabelRow = FindFirstRowInRangeContains(wsSrc, firstRow, lastRow, lastCol, kontoNeedle)
    If kontoLabelRow = 0 Then
        If doLog Then
            wsLog.Cells(logRow, 1).Resize(1, 8).Value = Array(blockNo, firstRow & "-" & lastRow, "(ikke funnet)", "", "", "", "", "")
            logRow = logRow + 1
        End If
        Exit Sub
    End If
    
    ' Find the actual row with numbers in I and J (could be the label row or a few rows below)
    Dim evalRow As Long
    evalRow = FindNumericPairRow(wsSrc, kontoLabelRow, WorksheetFunction.Min(kontoLabelRow + MAX_LOOKAHEAD, lastRow), COL_I, COL_J)
    If evalRow = 0 Then
        ' fallback: use the label row if nothing numeric is found
        evalRow = kontoLabelRow
    End If
    
    Dim txtI As String, txtJ As String
    txtI = CStr(wsSrc.Cells(evalRow, COL_I).Text)
    txtJ = CStr(wsSrc.Cells(evalRow, COL_J).Text)
    
    Dim vI As Variant, vJ As Variant, iIsNum As Boolean, jIsNum As Boolean
    vI = wsSrc.Cells(evalRow, COL_I).Value2
    vJ = wsSrc.Cells(evalRow, COL_J).Value2
    iIsNum = IsNumeric(vI)
    jIsNum = IsNumeric(vJ)
    
    ' NEGATIV: visible parentheses or underlying numeric < 0
    Dim isNeg As Boolean
    isNeg = HasParentheses(txtJ)
    If Not isNeg And jIsNum Then isNeg = (CDbl(vJ) < 0#)
    
    ' AVVIK: prefer Value2 numeric compare; else robust text-parsed compare
    Dim isAvvik As Boolean
    If iIsNum And jIsNum Then
        isAvvik = (Abs(CDbl(vI) - CDbl(vJ)) > 0.005) ' ~half a øre tolerance
    Else
        Dim pI As Variant, pJ As Variant
        pI = ParseAmountToDouble(txtI)
        pJ = ParseAmountToDouble(txtJ)
        If IsNumeric(pI) And IsNumeric(pJ) Then
            isAvvik = (Abs(CDbl(pI) - CDbl(pJ)) > 0.005)
        Else
            ' final fallback: normalized text (remove spaces, thin spaces, thousands sep, currency)
            isAvvik = (NormalizeForCompare(txtI) <> NormalizeForCompare(txtJ))
        End If
    End If
    
    If isNeg Then
        CopyBlock wsSrc, wsNeg, firstRow, lastRow, lastCol, destNeg
    End If
    If isAvvik Then
        CopyBlock wsSrc, wsAvv, firstRow, lastRow, lastCol, destAvv
    End If
    
    If doLog Then
        wsLog.Cells(logRow, 1).Resize(1, 8).Value = _
            Array(blockNo, firstRow & "-" & lastRow, kontoLabelRow, evalRow, txtI, txtJ, _
                  IIf(iIsNum, CDbl(vI), "not numeric"), IIf(jIsNum, CDbl(vJ), "not numeric"))
        logRow = logRow + 1
    End If
End Sub

Private Function FindNumericPairRow(ws As Worksheet, ByVal rStart As Long, ByVal rEnd As Long, ByVal colI As Long, ByVal colJ As Long) As Long
    Dim r As Long
    For r = rStart To rEnd
        If LooksNumeric(ws.Cells(r, colI).Text) And LooksNumeric(ws.Cells(r, colJ).Text) Then
            FindNumericPairRow = r
            Exit Function
        End If
        If IsNumeric(ws.Cells(r, colI).Value2) And IsNumeric(ws.Cells(r, colJ).Value2) Then
            FindNumericPairRow = r
            Exit Function
        End If
    Next r
    FindNumericPairRow = 0
End Function

Private Function LooksNumeric(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    t = ReplaceWeirdSpaces(t)
    If t = "" Then Exit Function
    Dim i As Long, ch As String, hasDigit As Boolean
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch Like "[0-9]" Then hasDigit = True
    Next i
    LooksNumeric = hasDigit
End Function

Private Sub CopyBlock(wsSrc As Worksheet, wsDst As Worksheet, _
                      ByVal firstRow As Long, ByVal lastRow As Long, ByVal lastCol As Long, _
                      ByRef destRow As Long)
    Dim rowsCount As Long: rowsCount = lastRow - firstRow + 1
    If rowsCount <= 0 Then Exit Sub
    
    wsDst.Cells(destRow, 1).Value = "--- Rader " & firstRow & "–" & lastRow & " ---"
    destRow = destRow + 1
    
    wsDst.Cells(destRow, 1).Resize(rowsCount, lastCol).Value = _
        wsSrc.Range(wsSrc.Cells(firstRow, 1), wsSrc.Cells(lastRow, lastCol)).Value
    
    destRow = destRow + rowsCount + 1
End Sub

Private Function RowContains(ws As Worksheet, ByVal r As Long, ByVal lastCol As Long, ByVal needle As String) As Boolean
    Dim c As Long
    For c = 1 To lastCol
        If InStr(1, CStr(ws.Cells(r, c).Text), needle, vbTextCompare) > 0 Then
            RowContains = True
            Exit Function
        End If
    Next c
End Function

Private Function FindFirstRowInRangeContains(ws As Worksheet, ByVal rStart As Long, ByVal rEnd As Long, ByVal lastCol As Long, ByVal needle As String) As Long
    Dim r As Long
    For r = rStart To rEnd
        If RowContains(ws, r, lastCol, needle) Then
            FindFirstRowInRangeContains = r
            Exit Function
        End If
    Next r
    FindFirstRowInRangeContains = 0
End Function

Private Function HasParentheses(ByVal s As String) As Boolean
    Dim t As String: t = ReplaceWeirdSpaces(s)
    t = Replace(t, " ", "")
    HasParentheses = (InStr(1, t, "(", vbBinaryCompare) > 0 And InStr(1, t, ")", vbBinaryCompare) > 0)
End Function

Private Function ReplaceWeirdSpaces(ByVal s As String) As String
    Dim t As String: t = s
    t = Replace(t, ChrW(160), " ")    ' NBSP
    t = Replace(t, ChrW(8239), " ")   ' NARROW NBSP (U+202F)
    t = Replace(t, ChrW(8201), " ")   ' THIN SPACE (U+2009)
    t = Replace(t, ChrW(8199), " ")   ' FIGURE SPACE (U+2007)
    ReplaceWeirdSpaces = t
End Function

Private Function NormalizeForCompare(ByVal s As String) As String
    Dim t As String
    t = ReplaceWeirdSpaces(s)
    t = Replace(t, vbTab, "")
    t = Replace(t, " ", "")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Replace(t, ".", "")        ' drop thousands dot
    t = Replace(t, "kr", "", 1, -1, vbTextCompare)
    t = Replace(t, "nok", "", 1, -1, vbTextCompare)
    NormalizeForCompare = t
End Function

' Parse European-style number strings to Double, honoring parentheses anywhere in the text.
Private Function ParseAmountToDouble(ByVal s As String) As Variant
    Dim t As String: t = Trim$(ReplaceWeirdSpaces(s))
    If t = "" Then ParseAmountToDouble = Empty: Exit Function
    
    Dim hasParens As Boolean
    hasParens = (InStr(1, t, "(", vbTextCompare) > 0 And InStr(1, t, ")", vbTextCompare) > 0)
    
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
        out = Replace(out, ",", ".")  ' comma decimal
    ElseIf hasComma And Not hasDot Then
        out = Replace(out, ",", ".")
    End If
    
    If out = "" Or out = "-" Then ParseAmountToDouble = Empty: Exit Function
    
    On Error GoTo notnum
    Dim v As Double: v = CDbl(out)
    If hasParens Then v = -Abs(v)
    ParseAmountToDouble = v
    Exit Function
notnum:
    ParseAmountToDouble = Empty
End Function

Private Function LastUsedRow(ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    LastUsedRow = IIf(c Is Nothing, 1, c.Row)
End Function

Private Function LastUsedCol(ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    LastUsedCol = IIf(c Is Nothing, 1, c.Column)
End Function

Private Function SheetOrCreate(ByVal name As String) As Worksheet
    On Error Resume Next
    Set SheetOrCreate = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If SheetOrCreate Is Nothing Then
        Set SheetOrCreate = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        SheetOrCreate.Name = name
    End If
End Function
