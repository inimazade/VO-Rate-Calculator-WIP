Attribute VB_Name = "StandardRateCalc"
Public ms As Worksheet
' ============================================================
'  MAIN ENTRY POINT
' ============================================================
Sub RunProgram1()
    Dim hs As Worksheet
    Set hs = GetOrCreateSheet("HiddenSheet")

    ' --- Import script and detect format ---
    Dim fmtForm As MainFormatSelector
    Set fmtForm = New MainFormatSelector
    fmtForm.Show

    ' Check if user cancelled
    If fmtForm.Tag = "cancelled" Then
        Set fmtForm = Nothing
        Exit Sub
    End If

    Dim fmt As Integer
    fmt = fmtForm.SelectedFormat
    Set fmtForm = Nothing

    If fmt = 0 Then
        MsgBox "That format is not yet implemented.", vbInformation
        Exit Sub
    End If
    
    ' --- Manual quotation mark mode ---
    If fmt = 98 Then
        RunManual
        Exit Sub
    End If
    
    ' --- If two-line, show sub-format selector to pick name style ---
    If fmt = 1 Then
        Dim subForm As TwoSubFormatSelector
        Set subForm = New TwoSubFormatSelector
        subForm.Show

        If subForm.Tag = "cancelled" Then
            Set subForm = Nothing
            Exit Sub
        End If

        fmt = subForm.SelectedStyle
        Set subForm = Nothing

        If fmt = 0 Then
            MsgBox "No name format selected.", vbExclamation
            Exit Sub
        End If
    End If

    ImportScript hs, fmt

    ' --- Extract unique character names from column A ---
    Dim CharNames() As String
    CharNames = GetCharacterNames(hs)
    If UBound(CharNames) < 0 Then
        MsgBox "No character names found. Check your script format.", vbExclamation
        Exit Sub
    End If

    ' --- Let user pick a character or all ---
    Dim CharChoice As String
    CharChoice = PickCharacter(CharNames)
    If CharChoice = "" Then Exit Sub

    ' --- Rate setup ---
    Dim userChoice As String
    userChoice = InputBox("Choose rate basis:" & vbCrLf & _
                          "1: Per Line" & vbCrLf & _
                          "2: Per Word", "Rate Type")
    Dim RateType As String
    Select Case userChoice
        Case "1": RateType = "Line"
        Case "2": RateType = "Word"
        Case "": MsgBox "User cancelled.": Exit Sub
        Case Else: MsgBox "Invalid input.": Exit Sub
    End Select

    ' --- Currency selection ---
    Dim CurrSymbol As String
    CurrSymbol = GetCurrency()

    Dim Rate As Double
    Rate = CDbl(InputBox("Enter rate per " & RateType & " (no symbol):", "Rate Input", "0.00"))

    Dim MinFee As Double
    MinFee = CDbl(InputBox("Enter minimum session fee (no symbol):" & vbCrLf & _
                           "This is the minimum the actor is paid regardless of " & RateType & " count.", _
                           "Minimum Fee", "0.00"))

    Dim Threshold As Long
    Threshold = CLng(InputBox("Enter the " & RateType & " threshold before per-unit rate applies:" & vbCrLf & _
                              "Below this number the actor receives the minimum fee only.", _
                              "Threshold", "10"))

    ' --- Build output ---
    If CharChoice = "__ALL__" Then
        BuildFullCastSummary hs, CharNames, RateType, Rate, MinFee, Threshold, CurrSymbol
    Else
        hs.Columns("C").Clear
        ExtractLinesToColumnC hs, CharChoice
        BuildCharacterSummary hs, CharChoice, RateType, Rate, MinFee, Threshold, CurrSymbol
    End If
End Sub

' ============================================================
'  GET OR CREATE A SHEET BY NAME
' ============================================================
Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

' ============================================================
'  IMPORT SCRIPT — handles both formats
'  Format 1: Name on its own line ending in colon, line below
'  Format 2: Name: Line  (inline, line may or may not be quoted)
' ============================================================
Private Sub ImportScript(hs As Worksheet, fmt As Integer)
    Dim FilePath As Variant
    FilePath = Application.GetOpenFilename( _
        "Text & Word Files (*.txt;*.docx), *.txt;*.docx", , _
        "Select Script File (UTF-8 TXT or DOCX)")
    If FilePath = False Then Exit Sub

    Dim FileExt As String
    FileExt = LCase(Mid(FilePath, InStrRev(FilePath, ".") + 1))

    Dim FileContent As String
    Select Case FileExt
        Case "txt":  FileContent = ReadTxtFile(CStr(FilePath))
        Case "docx": FileContent = ReadDocxFile(CStr(FilePath))
        Case Else
            MsgBox "Unsupported file type.", vbExclamation
            Exit Sub
    End Select
    If FileContent = "" Then Exit Sub

    hs.Cells.Clear

    Dim Lines() As String
    Lines = SplitLines(FileContent)

    Dim Row As Long
    Row = 1
    Dim i As Long, j As Long

    Select Case fmt

        Case 1, 2, 3
            ' --- Styles 1/2/3: name identified by line content, dialogue on next non-empty line ---
            For i = LBound(Lines) To UBound(Lines)
                Dim lineVal As String
                lineVal = Trim(Lines(i))
                If lineVal <> "" Then
                    ' Peek at next non-empty line to pass into heuristic
                    Dim nextLine As String
                    nextLine = ""
                    For j = i + 1 To UBound(Lines)
                        If Trim(Lines(j)) <> "" Then
                            nextLine = Trim(Lines(j))
                            Exit For
                        End If
                    Next j
                    If IsCharacterName(lineVal, fmt, nextLine) Then
                        Dim charPart As String
                        If Right(lineVal, 1) = ":" Then
                            charPart = Trim(Left(lineVal, Len(lineVal) - 1))
                        Else
                            charPart = Replace(lineVal, " ", "")
                        End If
                        ' Write name and next non-empty line as dialogue
                        If nextLine <> "" Then
                            hs.Cells(Row, 1).Value = charPart
                            hs.Cells(Row, 2).Value = StripQuotes(nextLine)
                            Row = Row + 1
                        End If
                    End If
                End If
            Next i

        Case 4
            ' --- Style 4: name line, blank line, dialogue line ---
            For i = LBound(Lines) To UBound(Lines) - 2
                lineVal = Trim(Lines(i))
                If lineVal <> "" Then
                    Dim blankLine As String
                    Dim dialogLine As String
                    blankLine = Trim(Lines(i + 1))
                    dialogLine = Trim(Lines(i + 2))
                    If blankLine = "" And dialogLine <> "" Then
                        hs.Cells(Row, 1).Value = Replace(lineVal, " ", "")
                        hs.Cells(Row, 2).Value = StripQuotes(dialogLine)
                        Row = Row + 1
                        i = i + 2
                    End If
                End If
            Next i

        Case 5
            ' --- Style 5: name line, vocal cue line, dialogue line ---
            For i = LBound(Lines) To UBound(Lines) - 2
                lineVal = Trim(Lines(i))
                If lineVal <> "" Then
                    Dim cueLine As String
                    cueLine = Trim(Lines(i + 1))
                    dialogLine = Trim(Lines(i + 2))
                    Dim looksLikeCue As Boolean
                    looksLikeCue = (InStr(cueLine, "(") > 0 Or _
                                    InStr(cueLine, "[") > 0 Or _
                                    InStr(cueLine, "*") > 0 Or _
                                    (Len(cueLine) < 40 And InStr(cueLine, """") = 0))
                    If looksLikeCue And dialogLine <> "" Then
                        hs.Cells(Row, 1).Value = Replace(lineVal, " ", "")
                        hs.Cells(Row, 2).Value = StripQuotes(dialogLine)
                        hs.Cells(Row, 3).Value = cueLine
                        Row = Row + 1
                        i = i + 2
                    End If
                End If
            Next i

        Case 99
            ' --- Inline format: Name: "Line", Name (cue): "Line", Name [cue]: "Line" ---
            For i = LBound(Lines) To UBound(Lines)
                lineVal = Trim(Lines(i))
                If lineVal <> "" Then
                    ' Find the colon that separates name from dialogue
                    ' Must skip colons that appear inside brackets or parentheses
                    Dim colonPos As Long
                    colonPos = FindNameColon(lineVal)
                    If colonPos > 1 Then
                        Dim namePart As String
                        Dim dialogPart As String
                        namePart = Trim(Left(lineVal, colonPos - 1))
                        dialogPart = Trim(Mid(lineVal, colonPos + 1))
                        ' Clean cue markers out of the name before storing
                        namePart = StripCueFromName(namePart)
                        If namePart <> "" And dialogPart <> "" Then
                            hs.Cells(Row, 1).Value = Replace(namePart, " ", "")
                            hs.Cells(Row, 2).Value = StripQuotes(dialogPart)
                            Row = Row + 1
                        End If
                    End If
                End If
            Next i

    End Select

    hs.Columns("A:C").WrapText = False
    MsgBox "Import complete. " & (Row - 1) & " lines imported.", vbInformation
End Sub

' ============================================================
'  HEURISTIC: IS THIS LINE A CHARACTER NAME? (Format 1)
'  Adjust the logic here to match your script conventions.
' ============================================================
Private Function IsCharacterName(lineVal As String, style As Integer, nextLine As String) As Boolean
    Select Case style
        Case 1 ' Ends with colon
            IsCharacterName = (Right(Trim(lineVal), 1) = ":")

        Case 2 ' ALL CAPS
            IsCharacterName = (lineVal = UCase(lineVal) And lineVal <> LCase(lineVal))

        Case 3
            ' Must be short
            If Len(Trim(lineVal)) >= 40 Then Exit Function
            ' Must contain no dialogue markers
            If InStr(lineVal, """") > 0 Then Exit Function
            If InStr(lineVal, "*") > 0 Then Exit Function
            If InStr(lineVal, "(") > 0 Then Exit Function
            ' Must contain no punctuation that dialogue would have
            If InStr(lineVal, "!") > 0 Then Exit Function
            If InStr(lineVal, "?") > 0 Then Exit Function
            If InStr(lineVal, ".") > 0 Then Exit Function
            If InStr(lineVal, ",") > 0 Then Exit Function
            If InStr(lineVal, "-") > 0 Then Exit Function
            If InStr(lineVal, "…") > 0 Then Exit Function
            If InStr(lineVal, "|") > 0 Then Exit Function
            ' Must contain only letters, spaces, ampersands, and apostrophes
            Dim c As String
            Dim k As Integer
            For k = 1 To Len(lineVal)
                c = Mid(lineVal, k, 1)
                If Not (c Like "[A-Za-z ]" Or c = "&" Or c = "'") Then
                    Exit Function
                End If
            Next k
            IsCharacterName = True
    End Select
End Function

' ============================================================
'  STRIP SURROUNDING QUOTES FROM A STRING
' ============================================================
Private Function StripQuotes(s As String) As String
    s = Trim(s)
    ' Handle smart quotes and straight quotes
    's = Replace(s, Chr(8220), """")
    's = Replace(s, Chr(8221), """")
    If Left(s, 1) = """" And Right(s, 1) = """" Then
        s = Mid(s, 2, Len(s) - 2)
    End If
    StripQuotes = Trim(s)
End Function

Private Function StripCueFromName(namePart As String) As String
    Dim result As String
    result = namePart

    ' Strip bracket cues e.g. "John [whispering]" -> "John"
    Do While InStr(result, "[") > 0
        Dim openB As Long, closeB As Long
        openB = InStr(result, "[")
        closeB = InStr(result, "]")
        If closeB > openB Then
            result = Left(result, openB - 1) & Mid(result, closeB + 1)
        Else
            Exit Do ' Unpaired bracket, stop
        End If
    Loop

    ' Strip parenthetical cues e.g. "John (whispering)" -> "John"
    Do While InStr(result, "(") > 0
        Dim openP As Long, closeP As Long
        openP = InStr(result, "(")
        closeP = InStr(result, ")")
        If closeP > openP Then
            result = Left(result, openP - 1) & Mid(result, closeP + 1)
        Else
            Exit Do ' Unpaired parenthesis, stop
        End If
    Loop

    StripCueFromName = Trim(result)
End Function

Private Function FindNameColon(lineVal As String) As Long
    Dim depth As Integer
    Dim k As Long
    Dim c As String
    depth = 0

    For k = 1 To Len(lineVal)
        c = Mid(lineVal, k, 1)
        Select Case c
            Case "(", "["
                depth = depth + 1
            Case ")", "]"
                If depth > 0 Then depth = depth - 1
            Case ":"
                If depth = 0 Then
                    FindNameColon = k
                    Exit Function
                End If
        End Select
    Next k

    FindNameColon = 0 ' No valid colon found
End Function

' ============================================================
'  GET UNIQUE CHARACTER NAMES FROM COLUMN A OF HIDDENSHEET
' ============================================================
Private Function GetCharacterNames(hs As Worksheet) As String()
    Dim lastRow As Long
    lastRow = hs.Cells(hs.Rows.count, "A").End(xlUp).Row

    Dim names() As String
    ReDim names(0)
    Dim count As Long
    count = 0

    Dim i As Long
    For i = 1 To lastRow
        Dim nm As String
        nm = Trim(hs.Cells(i, 1).Value)
        If nm <> "" Then
            ' Check if already in list
            Dim found As Boolean
            found = False
            Dim k As Long
            For k = 0 To count - 1
                If LCase(names(k)) = LCase(nm) Then
                    found = True
                    Exit For
                End If
            Next k
            If Not found Then
                ReDim Preserve names(count)
                names(count) = nm
                count = count + 1
            End If
        End If
    Next i

    If count = 0 Then
        GetCharacterNames = Array()
    Else
        GetCharacterNames = names
    End If
End Function

' ============================================================
'  LET USER PICK A CHARACTER FROM THE EXTRACTED LIST
'  Returns "__ALL__" if user picks the all-characters option
'  Returns "" if cancelled
' ============================================================
Private Function PickCharacter(CharNames() As String) As String
    Dim prompt As String
    prompt = "Select a character by number, or enter 0 for full cast breakdown:" & vbCrLf & vbCrLf
    prompt = prompt & "0: ALL CHARACTERS (director view)" & vbCrLf

    Dim i As Long
    For i = 0 To UBound(CharNames)
        prompt = prompt & (i + 1) & ": " & CharNames(i) & vbCrLf
    Next i

    Dim ans As String
    ans = InputBox(prompt, "Select Character")

    If ans = "" Then
        PickCharacter = ""
    ElseIf ans = "0" Then
        PickCharacter = "__ALL__"
    Else
        Dim idx As Long
        idx = CLng(ans) - 1
        If idx >= 0 And idx <= UBound(CharNames) Then
            PickCharacter = CharNames(idx)
        Else
            MsgBox "Invalid selection.", vbExclamation
            PickCharacter = ""
        End If
    End If
End Function

' ============================================================
'  COPY A SINGLE CHARACTER'S LINES INTO COLUMN C
'  (used when building single-character summary)
' ============================================================
Private Sub ExtractLinesToColumnC(hs As Worksheet, charName As String)
    Dim lastRow As Long
    lastRow = hs.Cells(hs.Rows.count, "A").End(xlUp).Row
    Dim j As Long
    j = 1
    Dim i As Long
    For i = 1 To lastRow
        If LCase(Trim(hs.Cells(i, 1).Value)) = LCase(Trim(charName)) Then
            hs.Cells(j, 3).Value = hs.Cells(i, 2).Value
            j = j + 1
        End If
    Next i
End Sub

Private Function GetCurrency() As String
    Dim msg As String
    msg = "Select your currency:" & vbCrLf & vbCrLf & _
          "1:  USD - US Dollar ($)" & vbCrLf & _
          "2:  EUR - Euro (€)" & vbCrLf & _
          "3:  GBP - British Pound (£)" & vbCrLf & _
          "4:  CAD - Canadian Dollar (CA$)" & vbCrLf & _
          "5:  AUD - Australian Dollar (A$)" & vbCrLf & _
          "6:  JPY - Japanese Yen (¥)" & vbCrLf & _
          "7:  CNY - Chinese Yuan (¥)" & vbCrLf & _
          "8:  INR - Indian Rupee (?)" & vbCrLf & _
          "9:  BRL - Brazilian Real (R$)" & vbCrLf & _
          "10: KRW - South Korean Won (?)" & vbCrLf & _
          "11: MXN - Mexican Peso (MX$)" & vbCrLf & _
          "12: SEK - Swedish Krona (kr)" & vbCrLf & _
          "13: NOK - Norwegian Krone (kr)" & vbCrLf & _
          "14: DKK - Danish Krone (kr)" & vbCrLf & _
          "15: CHF - Swiss Franc (Fr)" & vbCrLf & _
          "16: NZD - New Zealand Dollar (NZ$)" & vbCrLf & _
          "17: ZAR - South African Rand (R)" & vbCrLf & _
          "18: Other - Enter your own symbol"
    Dim ans As String
    ans = InputBox(msg, "Select Currency", "1")

    Select Case ans
        Case "1":  GetCurrency = "$"
        Case "2":  GetCurrency = "€"
        Case "3":  GetCurrency = "£"
        Case "4":  GetCurrency = "CA$"
        Case "5":  GetCurrency = "A$"
        Case "6":  GetCurrency = "¥"
        Case "7":  GetCurrency = "¥"
        Case "8":  GetCurrency = "?"
        Case "9":  GetCurrency = "R$"
        Case "10": GetCurrency = "?"
        Case "11": GetCurrency = "MX$"
        Case "12": GetCurrency = "kr"
        Case "13": GetCurrency = "kr"
        Case "14": GetCurrency = "kr"
        Case "15": GetCurrency = "Fr"
        Case "16": GetCurrency = "NZ$"
        Case "17": GetCurrency = "R"
        Case "18"
            Dim custom As String
            custom = InputBox("Enter your currency symbol:", "Custom Currency", "$")
            If custom = "" Then
                GetCurrency = "$"
            Else
                GetCurrency = custom
            End If
        Case ""
            GetCurrency = "$" ' Default if cancelled
        Case Else
            GetCurrency = "$"
    End Select
End Function

Private Function GetCurrencyFormat(symbol As String) As String
    ' JPY and KRW are zero-decimal currencies
    If symbol = "¥" Or symbol = "?" Then
        GetCurrencyFormat = """" & symbol & """#,##0"
    Else
        GetCurrencyFormat = """" & symbol & """#,##0.00"
    End If
End Function

' ============================================================
'  SINGLE CHARACTER SUMMARY  ?  Output sheet
' ============================================================
Sub BuildCharacterSummary(hs As Worksheet, charName As String, RateType As String, Rate As Double, MinFee As Double, Threshold As Long, CurrSymbol As String)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("Output")
    ws.Cells.Clear

    Dim lastRowB As Long
    lastRowB = hs.Cells(hs.Rows.count, "B").End(xlUp).Row

    ' --- Headers ---
    ws.Cells(1, 1).Value = "Character Summary"
    ws.Range("A2:I2").Value = Array("Character", "Lines", "Words", "Est. Time", "Priority", "Rate per " & RateType, "Threshold", "Min Fee", "Est. Cost")

    ws.Cells(3, 1).Value = charName

    Dim i As Long, words As Long, Lines As Long
    Dim txt As String
    words = 0: Lines = 0

    For i = 1 To lastRowB
        txt = hs.Cells(i, 3).Value
        If Trim(txt) <> "" Then
            Lines = Lines + 1
            words = words + CountWords(txt)
        End If
    Next i

    ws.Cells(3, 2).Value = Lines
    ws.Cells(3, 3).Value = words
    ws.Cells(3, 4).Value = Round(words / 1.875) / 86400
    ws.Cells(3, 4).NumberFormat = "h:mm:ss"
    ws.Cells(3, 5).Value = 1
    ws.Cells(3, 6).Value = Rate

    ' Col H: threshold, Col I: min fee (helper values for formula, can be hidden)
    ws.Cells(3, 7).Value = Threshold
    ws.Cells(3, 8).Value = MinFee

    ' Est. Cost: if units <= threshold, pay min fee; otherwise MAX(min fee, units * rate)
    If RateType = "Line" Then
        ws.Cells(3, 9).Formula = "=IF(B3<=G3,H3,MAX(H3,B3*F3))"
    Else
        ws.Cells(3, 9).Formula = "=IF(C3<=G3,H3,MAX(H3,C3*F3))"
    End If

    ' --- Formatting ---
    Dim cf As String
    cf = GetCurrencyFormat(CurrSymbol)
    ws.Range("A1:I1").Merge
    ws.Range("A1:I1").HorizontalAlignment = xlCenter
    ws.Range("A2:I2").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("A2:I2").Borders(xlEdgeBottom).Weight = xlMedium
    ws.Range("F3").NumberFormat = cf
    ws.Range("H3").NumberFormat = cf
    ws.Range("I3").NumberFormat = cf
    ws.Range("G3").NumberFormat = "#,##0"
    ws.Columns("A:I").AutoFit
End Sub

' ============================================================
'  FULL CAST SUMMARY  ?  Output sheet (director view)
' ============================================================
Sub BuildFullCastSummary(hs As Worksheet, CharNames() As String, RateType As String, Rate As Double, MinFee As Double, Threshold As Long, CurrSymbol As String)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("Output")
    ws.Cells.Clear

    ' --- Headers ---
    ws.Cells(1, 1).Value = "Full Cast Summary"
    ws.Range("A2:I2").Value = Array("Character", "Lines", "Words", "Est. Time", "Priority", "Rate per " & RateType, "Threshold", "Min Fee", "Est. Cost")

    Dim lastRow As Long
    lastRow = hs.Cells(hs.Rows.count, "A").End(xlUp).Row

    Dim c As Long
    For c = 0 To UBound(CharNames)
        Dim nm As String
        nm = CharNames(c)
        Dim outRow As Long
        outRow = c + 3

        Dim i As Long, words As Long, Lines As Long
        Dim txt As String
        words = 0: Lines = 0

        For i = 1 To lastRow
            If LCase(Trim(hs.Cells(i, 1).Value)) = LCase(Trim(nm)) Then
                Lines = Lines + 1
                txt = hs.Cells(i, 2).Value
                words = words + CountWords(txt)
            End If
        Next i

        ws.Cells(outRow, 1).Value = nm
        ws.Cells(outRow, 2).Value = Lines
        ws.Cells(outRow, 3).Value = words
        ws.Cells(outRow, 4).Value = Round(words / 1.875) / 86400
        ws.Cells(outRow, 4).NumberFormat = "h:mm:ss"
        ws.Cells(outRow, 5).Formula = "=RANK(B" & outRow & ",B$3:B$" & (UBound(CharNames) + 3) & ",0)"
        ws.Cells(outRow, 6).Value = Rate
        ws.Cells(outRow, 7).Value = Threshold
        ws.Cells(outRow, 8).Value = MinFee

        If RateType = "Line" Then
            ws.Cells(outRow, 9).Formula = "=IF(B" & outRow & "<=G" & outRow & ",H" & outRow & ",MAX(H" & outRow & ",B" & outRow & "*F" & outRow & "))"
        Else
            ws.Cells(outRow, 9).Formula = "=IF(C" & outRow & "<=G" & outRow & ",H" & outRow & ",MAX(H" & outRow & ",C" & outRow & "*F" & outRow & "))"
        End If
    Next c

    ' --- Totals row ---
    Dim totRow As Long
    totRow = UBound(CharNames) + 4
    ws.Cells(totRow, 1).Value = "TOTALS"
    ws.Cells(totRow, 2).Formula = "=SUM(B3:B" & (totRow - 1) & ")"
    ws.Cells(totRow, 3).Formula = "=SUM(C3:C" & (totRow - 1) & ")"
    ws.Cells(totRow, 4).Formula = "=SUM(D3:D" & (totRow - 1) & ")"
    ws.Cells(totRow, 4).NumberFormat = "h:mm:ss"
    ws.Cells(totRow, 5).Value = "---"
    ws.Cells(totRow, 6).Value = "---"
    ws.Cells(totRow, 7).Value = "---"
    ws.Cells(totRow, 8).Value = "---"
    ws.Cells(totRow, 9).Formula = "=SUM(I3:I" & (totRow - 1) & ")"
    
    ' --- Formatting ---
    Dim cf As String
    cf = GetCurrencyFormat(CurrSymbol)
    ws.Range("A1:I1").Merge
    ws.Range("A1:I1").HorizontalAlignment = xlCenter
    ws.Range("A2:I2").Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("A2:I2").Borders(xlEdgeBottom).Weight = xlMedium
    ws.Range("A" & (totRow - 1) & ":I" & (totRow - 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("A" & (totRow - 1) & ":I" & (totRow - 1)).Borders(xlEdgeBottom).Weight = xlMedium
    ws.Range("F3:F" & totRow).NumberFormat = cf
    ws.Range("H3:H" & totRow).NumberFormat = cf
    ws.Range("I3:I" & totRow).NumberFormat = cf
    ws.Range("G3:G" & totRow).NumberFormat = "#,##0"
    ws.Columns("A:I").AutoFit
End Sub

' ============================================================
'  CHARACTER EXTRACT  —  highlight name rows on HiddenSheet
' ============================================================
Sub CharacterExtract(hs As Worksheet, charName As String)
    Dim lastRowA As Long
    lastRowA = hs.Cells(hs.Rows.count, "A").End(xlUp).Row
    Dim i As Long
    For i = 1 To lastRowA
        If LCase(Trim(hs.Cells(i, 1).Value)) = LCase(Trim(charName)) Then
            hs.Cells(i, 1).Interior.Color = vbRed
        End If
    Next i
End Sub

' ============================================================
'  UTILITIES
' ============================================================
Private Function SplitLines(content As String) As String()
    Dim normalized As String
    normalized = Replace(content, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    SplitLines = Split(normalized, vbLf)
End Function

Sub RunManual()
    Set ms = ThisWorkbook.Sheets("Manual")
    ms.Cells.Clear

    ' --- File import ---
    Dim FilePath As Variant
    FilePath = Application.GetOpenFilename( _
        "Word Files (*.docx), *.docx", , _
        "Select Script File (DOCX Only)")
    If FilePath = False Then Exit Sub

    Dim FileContent As String
    FileContent = ReadDocxFile2(CStr(FilePath))
    If FileContent = "" Then Exit Sub

    ' --- Normalize Word characters to plain text ---
    FileContent = NormalizeWordText(FileContent)

    ' --- Write to worksheet ---
    Dim Row As Long
    Row = 1
    Dim i As Long
    
    ms.Cells(1, 1).Value = "NARRATOR"
    Lines = Split(FileContent, """")
    Row = 1

    For i = LBound(Lines) To UBound(Lines)
        If Trim(Lines(i)) <> "" Then
            If i Mod 2 = 0 Then
                ms.Cells(Row, 2).Value = "NARRATOR"
                ms.Cells(Row, 3).Value = Trim(Lines(i))
            Else
                ms.Cells(Row, 3).Value = Trim(Lines(i))
            End If
            If ms.Cells(Row, 2).Value = "NARRATOR" & ms.Cells(Row, 3).Value = vbCrLf Then
                ms.Cells(Row, 2).Value = ""
                ms.Cells(Row, 3).Value = ""
            Else
                Row = Row + 1
            End If
       End If
    Next i
    
    
    

    ' --- Clean up empty narrator rows ---
    Dim lastRowC As Long
    lastRowC = ms.Cells(ms.Rows.count, "C").End(xlUp).Row
    DenarrateMaunual ms, lastRowC

    ms.Columns("B:C").WrapText = False

    ' --- Launch allocation userform ---
    ms.Activate
    Dim allocForm As ManualCharNameAllocation
    Set allocForm = New ManualCharNameAllocation
    allocForm.Show vbModeless
    Set allocForm = Nothing
End Sub

Private Sub DenarrateMaunual(ms As Worksheet, lastRowC As Long)
    Dim i As Long
    Dim txt As String
    For i = lastRowC To 1 Step -1
        txt = ms.Cells(i, 3).Value
        If ms.Cells(i, 2).Value = "NARRATOR" And IsEmptyOrWhitespace(txt) Then
            ms.Rows(i).Delete
        End If
    Next i
End Sub

Private Function IsEmptyOrWhitespace(ByVal txt As String) As Boolean
    Dim i As Long
    IsEmptyOrWhitespace = True  ' Assume empty until we find a non-whitespace char
    
    For i = 1 To Len(txt)
        Dim c As String
        c = Mid(txt, i, 1)
        ' Check if character is NOT a whitespace character
        Select Case c
            Case " ", vbTab, vbCr, vbLf, Chr(13), Chr(10), Chr(32), Chr(160)
                ' These are whitespace - continue checking
            Case Else
                ' Found a real character
                IsEmptyOrWhitespace = False
                Exit Function
        End Select
    Next i
End Function

Private Function ReadTxtFile(FilePath As String) As String
    Dim FileNum As Integer
    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    ReadTxtFile = Input$(LOF(FileNum), FileNum)
    Close #FileNum
End Function

Private Function ReadDocxFile(FilePath As String) As String
    Dim wdApp As Object, wdDoc As Object
    On Error GoTo DocError
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(FilePath, ReadOnly:=True)
    ReadDocxFile = wdDoc.content.text
    wdDoc.Close SaveChanges:=False
    wdApp.Quit
    Exit Function
DocError:
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    MsgBox "Error reading Word file: " & Err.Description, vbExclamation
    ReadDocxFile = ""
End Function
Private Function ReadDocxFile2(FilePath As String) As String
    Dim wdApp As Object, wdDoc As Object
    On Error GoTo DocError
    
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(FilePath, ReadOnly:=True)
    
    ' Read the content
    ReadDocxFile2 = wdDoc.content.text
    
    wdDoc.Close SaveChanges:=False
    wdApp.Quit
    
    Exit Function
    
DocError:
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    MsgBox "Error reading Word file: " & Err.Description, vbExclamation
    ReadDocxFile2 = ""
End Function

Private Function CountWords(ByVal txt As String) As Long
    ' --- Strip parenthetical director notes e.g. (softly) ---
    Dim result As String
    result = txt
    Do While InStr(result, "(") > 0
        Dim openP As Long, closeP As Long
        openP = InStr(result, "(")
        closeP = InStr(result, ")")
        If closeP > openP Then
            result = Left(result, openP - 1) & Mid(result, closeP + 1)
        Else
            Exit Do ' Malformed parenthesis, stop to avoid infinite loop
        End If
    Loop

    ' --- Replace *foley cues* with single placeholder word ---
    Do While InStr(result, "*") > 0
        Dim firstStar As Long, secondStar As Long
        firstStar = InStr(result, "*")
        secondStar = InStr(firstStar + 1, result, "*")
        If secondStar > firstStar Then
            result = Left(result, firstStar - 1) & "FOLEY" & Mid(result, secondStar + 1)
        Else
            Exit Do ' Unpaired asterisk, stop to avoid infinite loop
        End If
    Loop

    ' --- Count remaining words ---
    result = Application.WorksheetFunction.Trim(result)
    If result = "" Then
        CountWords = 0
    Else
        CountWords = UBound(Split(result, " ")) + 1
    End If
End Function

Private Sub LetUserTagNames(hs As Worksheet)
    ' Dump column A to a visible temp sheet
    Dim ts As Worksheet
    Set ts = GetOrCreateSheet("TagNames")
    ts.Cells.Clear
    ts.Cells(1, 1).Value = "Mark character name rows yellow, then run ConfirmTags"

    Dim lastRow As Long
    lastRow = hs.Cells(hs.Rows.count, "A").End(xlUp).Row
    Dim i As Long
    For i = 1 To lastRow
        ts.Cells(i + 1, 1).Value = hs.Cells(i, 1).Value
        ts.Cells(i + 1, 2).Value = hs.Cells(i, 2).Value
    Next i

    MsgBox "Lines have been copied to the 'TagNames' sheet." & vbCrLf & _
           "Highlight any rows that are character names in yellow, then run ConfirmTags.", _
           vbInformation
End Sub

Private Sub ConfirmTags()
    ' Read back user-tagged rows and rebuild column A as name indicators
    Dim ts As Worksheet
    Dim hs As Worksheet
    Set ts = GetOrCreateSheet("TagNames")
    Set hs = GetOrCreateSheet("HiddenSheet")

    Dim lastRow As Long
    lastRow = ts.Cells(ts.Rows.count, "A").End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If ts.Cells(i, 1).Interior.Color = RGB(255, 255, 0) Then
            hs.Cells(i - 1, 1).Interior.Color = RGB(255, 255, 0) ' Mark in HiddenSheet too
        End If
    Next i
    MsgBox "Tags confirmed.", vbInformation
End Sub

Sub ClearSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Output")
    ws.Range("A1:I10000").Clear
    Dim hs As Worksheet
    Set hs = ThisWorkbook.Sheets("HiddenSheet")
    hs.Range("A1:C10000").Clear
    Dim ms As Worksheet
    Set ms = ThisWorkbook.Sheets("Manual")
    ms.Cells.Clear
End Sub

Sub FinalizeManual()
    Dim ms As Worksheet
    Set ms = GetOrCreateSheet("Manual")

    Dim lastRowB As Long
    Dim lastRowC As Long
    Dim lastRowA As Long
    lastRowB = ms.Cells(ms.Rows.count, "B").End(xlUp).Row
    lastRowC = ms.Cells(ms.Rows.count, "C").End(xlUp).Row
    lastRowA = ms.Cells(ms.Rows.count, "A").End(xlUp).Row

    ' --- Check for unassigned lines ---
    Dim unassigned As Long
    unassigned = 0
    Dim i As Long
    For i = 1 To lastRowC
        If Trim(ms.Cells(i, 3).Value) <> "" And Trim(ms.Cells(i, 2).Value) = "" Then
            unassigned = unassigned + 1
        End If
    Next i
    If unassigned > 0 Then
        Dim proceed As VbMsgBoxResult
        proceed = MsgBox(unassigned & " line(s) have no character assigned in Column B." & vbCrLf & _
                         "Continue anyway?", vbYesNo + vbExclamation)
        If proceed = vbNo Then Exit Sub
    End If

    ThisWorkbook.Sheets("Output").Activate
    ms.Visible = xlSheetHidden

    ' --- Remove column A if B and C are populated and A is shorter ---
    ' This matches your condition: col B & C same length, col A shorter
    If lastRowB >= lastRowC And lastRowA < lastRowB Then
        ms.Columns("A").Delete
        ' After deletion B becomes A, C becomes B
        lastRowB = ms.Cells(ms.Rows.count, "A").End(xlUp).Row
        lastRowC = ms.Cells(ms.Rows.count, "B").End(xlUp).Row
    End If

    ' --- Copy to HiddenSheet col A (character) and col B (dialogue) ---
    Dim hs As Worksheet
    Set hs = GetOrCreateSheet("HiddenSheet")
    hs.Cells.Clear

    Dim hsRow As Long
    hsRow = 1
    Dim colChar As Long, colDialog As Long

    ' After potential column A deletion, character is col A(1) and dialogue is col B(2)
    ' If column A was NOT deleted, character is col B(2) and dialogue is col C(3)
    If lastRowA < lastRowB Then
        colChar = 1
        colDialog = 2
    Else
        colChar = 2
        colDialog = 3
    End If

    Dim lastRowFinal As Long
    lastRowFinal = ms.Cells(ms.Rows.count, colChar).End(xlUp).Row

    For i = 1 To lastRowFinal
        Dim charVal As String
        Dim dialogVal As String
        charVal = Trim(ms.Cells(i, colChar).Value)
        dialogVal = Trim(ms.Cells(i, colDialog).Value)
        If charVal <> "" And dialogVal <> "" Then
            hs.Cells(hsRow, 1).Value = charVal
            hs.Cells(hsRow, 2).Value = dialogVal
            hsRow = hsRow + 1
        End If
    Next i

    MsgBox "Manual data finalized. " & (hsRow - 1) & " lines copied to HiddenSheet." & vbCrLf & vbCrLf & _
           "You can now run the character selection and summary steps.", vbInformation

    ' --- Hand off to character selection ---
    Dim CharNames() As String
    CharNames = GetCharacterNames(hs)
    If UBound(CharNames) < 0 Then
        MsgBox "No character names found after finalization.", vbExclamation
        Exit Sub
    End If

    Dim CharChoice As String
    CharChoice = PickCharacter(CharNames)
    If CharChoice = "" Then Exit Sub

    Dim userChoice As String
    userChoice = InputBox("Choose rate basis:" & vbCrLf & _
                          "1: Per Line" & vbCrLf & _
                          "2: Per Word", "Rate Type")
    Dim RateType As String
    Select Case userChoice
        Case "1": RateType = "Line"
        Case "2": RateType = "Word"
        Case "": MsgBox "User cancelled.": Exit Sub
        Case Else: MsgBox "Invalid input.": Exit Sub
    End Select

    On Error Resume Next
    Dim Rate As Double
    Rate = CDbl(InputBox("Enter rate per " & RateType & " (no symbol):", "Rate Input", "0.00"))
    If Err.Number <> 0 Then
        MsgBox "Invalid rate entered.", vbExclamation: Exit Sub
    End If
    Dim MinFee As Double
    MinFee = CDbl(InputBox("Enter minimum session fee (no symbol):", "Minimum Fee", "0.00"))
    If Err.Number <> 0 Then
        MsgBox "Invalid minimum fee entered.", vbExclamation: Exit Sub
    End If
    Dim Threshold As Long
    Threshold = CLng(InputBox("Enter threshold:", "Threshold", "10"))
    If Err.Number <> 0 Then
        MsgBox "Invalid threshold entered.", vbExclamation: Exit Sub
    End If
    On Error GoTo 0

    Dim CurrSymbol As String
    CurrSymbol = GetCurrency()

    If CharChoice = "__ALL__" Then
        BuildFullCastSummary hs, CharNames, RateType, Rate, MinFee, Threshold, CurrSymbol
    Else
        hs.Columns("C").Clear
        ExtractLinesToColumnC hs, CharChoice
        BuildCharacterSummary hs, CharChoice, RateType, Rate, MinFee, Threshold, CurrSymbol
    End If
End Sub




Private Function NormalizeWordText(ByVal text As String) As String
    ' Replace smart quotes with straight quotes
    text = Replace(text, ChrW(8220), """")  ' Left double smart quote
    text = Replace(text, ChrW(8221), """")  ' Right double smart quote
    text = Replace(text, ChrW(8216), "'")   ' Left single smart quote
    text = Replace(text, ChrW(8217), "'")   ' Right single smart quote
    text = Replace(text, ChrW(8243), """")  ' Double prime (sometimes used)
    
    ' Replace Word-specific line breaks
    text = Replace(text, vbCrLf, vbCr)      ' Windows line breaks
    text = Replace(text, vbCr, vbLf)        ' Normalize to line feed
    text = Replace(text, vbLf & vbLf, vbLf) ' Remove double line breaks
    
    ' Replace common Word characters
    text = Replace(text, ChrW(8230), "...") ' Ellipsis
    text = Replace(text, ChrW(8211), "-")   ' En dash
    text = Replace(text, ChrW(8212), "-")   ' Em dash
    
    ' Remove paragraph marks (Word sometimes adds these)
    text = Replace(text, Chr(13) & Chr(7), vbLf)
    
    NormalizeWordText = text
End Function

Private Function ParseDialogueLines(ByVal content As String) As String()
    ' Split by line breaks first
    Dim tempLines() As String
    tempLines = Split(content, vbLf)
    
    Dim result() As String
    ReDim result(0)
    Dim resultCount As Long
    resultCount = 0
    
    Dim i As Long
    For i = LBound(tempLines) To UBound(tempLines)
        Dim line As String
        line = Trim(tempLines(i))
        
        If line <> "" Then
            ' Check if this line starts a dialogue (has opening quote)
            If Left(line, 1) = """" Then
                ' This line is dialogue - could span multiple lines
                Dim dialogue As String
                dialogue = line
                
                ' If line doesn't end with a closing quote, keep reading
                Do While Right(Trim(dialogue), 1) <> """" And i < UBound(tempLines)
                    i = i + 1
                    dialogue = dialogue & " " & Trim(tempLines(i))
                Loop
                
                ' Add the complete dialogue line
                ReDim Preserve result(resultCount)
                result(resultCount) = dialogue
                resultCount = resultCount + 1
            Else
                ' This is narration/stage direction
                ReDim Preserve result(resultCount)
                result(resultCount) = line
                resultCount = resultCount + 1
            End If
        End If
    Next i
    
    ParseDialogueLines = result
End Function

Private Function IsDialogueLine(ByVal line As String) As Boolean
    ' Check if line starts with a quote (after trimming)
    Dim trimmed As String
    trimmed = Trim(line)
    IsDialogueLine = (Left(trimmed, 1) = """")
End Function

Private Function ExtractQuotedText(ByVal line As String) As String
    ' Extract text between first and last quote
    Dim firstQuote As Long
    Dim lastQuote As Long
    
    firstQuote = InStr(line, """")
    If firstQuote = 0 Then
        ExtractQuotedText = line
        Exit Function
    End If
    
    lastQuote = InStrRev(line, """")
    If lastQuote <= firstQuote Then
        ExtractQuotedText = Mid(line, firstQuote + 1)
    Else
        ExtractQuotedText = Mid(line, firstQuote + 1, lastQuote - firstQuote - 1)
    End If
    
    ' Clean up any remaining quote characters
    ExtractQuotedText = Replace(ExtractQuotedText, """", "")
    ExtractQuotedText = Trim(ExtractQuotedText)
End Function
