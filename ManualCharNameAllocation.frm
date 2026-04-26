VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ManualCharNameAllocation 
   Caption         =   "Manual Character Name Allocation"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15135
   OleObjectBlob   =   "ManualCharNameAllocation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ManualCharNameAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
'  MODULE-LEVEL STATE
' ============================================================
Private msSheet As Worksheet
Private dialogueRows() As Long
Private currentIndex As Long
Private totalLines As Long

Public Property Set ManualSheet(ws As Worksheet)
    Set msSheet = ws
End Property

' ============================================================
'  INITIALIZE — cosmetic only, no external data needed
' ============================================================
Private Sub UserForm_Initialize()
    
End Sub

' ============================================================
'  SETUP — called explicitly after ManualSheet is set
'  This is what Initialize used to do
' ============================================================
Public Sub Setup()
    If msSheet Is Nothing Then
        MsgBox "ManualSheet not set before calling Setup.", vbCritical
        Unload Me
        Exit Sub
    End If

    ' --- Build index of rows with dialogue in col C ---
    Dim lastRowC As Long
    lastRowC = msSheet.Cells(msSheet.Rows.count, "C").End(xlUp).Row

    Dim tempRows() As Long
    ReDim tempRows(1 To lastRowC)
    Dim count As Long
    count = 0

    Dim i As Long
    For i = 1 To lastRowC
        If Not IsEmptyOrWhitespace(CStr(msSheet.Cells(i, 3).Value)) Then
            count = count + 1
            tempRows(count) = i
        End If
    Next i

    If count = 0 Then
        MsgBox "No dialogue lines found in Column C.", vbExclamation
        Unload Me
        Exit Sub
    End If

    totalLines = count
    ReDim dialogueRows(1 To totalLines)
    For i = 1 To totalLines
        dialogueRows(i) = tempRows(i)
    Next i

    RefreshListBox
    currentIndex = 1
    UpdateDisplay
End Sub

' ============================================================
'  COMMANDBUTTON1 — Add newly typed name, assign, advance
' ============================================================
Private Sub CommandButton1_Click()
    Dim newName As String
    newName = Trim(TextBox1.Value)
    If newName = "" Then
        MsgBox "Please type a character name first.", vbExclamation
        Exit Sub
    End If

    If Not NameExistsInList(newName) Then
        Dim lastRowA As Long
        lastRowA = msSheet.Cells(msSheet.Rows.count, "A").End(xlUp).Row
        If lastRowA < 1 Then lastRowA = 1
        msSheet.Cells(lastRowA + 1, 1).Value = newName
        RefreshListBox
    End If

    AssignCurrentLine newName
    TextBox1.Value = ""

    If currentIndex < totalLines Then
        currentIndex = currentIndex + 1
        UpdateDisplay
    Else
        MsgBox "All lines assigned. Click Finalize when ready.", vbInformation
    End If
End Sub

' ============================================================
'  COMMANDBUTTON2 — Assign selected ListBox name, advance
' ============================================================
Private Sub CommandButton2_Click()
    If ListBox1.ListIndex = -1 Then
        MsgBox "Please select a character from the list.", vbExclamation
        Exit Sub
    End If

    Dim selectedName As String
    selectedName = ListBox1.List(ListBox1.ListIndex)
    AssignCurrentLine selectedName

    If currentIndex < totalLines Then
        currentIndex = currentIndex + 1
        UpdateDisplay
    Else
        MsgBox "All lines assigned. Click Finalize when ready.", vbInformation
    End If
End Sub

' ============================================================
'  COMMANDBUTTON3 — Previous line
' ============================================================
Private Sub CommandButton3_Click()
    If currentIndex > 1 Then
        currentIndex = currentIndex - 1
        UpdateDisplay
    Else
        MsgBox "Already at the first line.", vbInformation
    End If
End Sub

' ============================================================
'  COMMANDBUTTON4 — Next line without assigning
' ============================================================
Private Sub CommandButton4_Click()
    If currentIndex < totalLines Then
        currentIndex = currentIndex + 1
        UpdateDisplay
    Else
        MsgBox "Already at the last line.", vbInformation
    End If
End Sub

' ============================================================
'  COMMANDBUTTON5 — Finalize
' ============================================================
Private Sub CommandButton5_Click()
    Dim unassigned As Long
    unassigned = 0
    Dim i As Long
    For i = 1 To totalLines
        If Trim(msSheet.Cells(dialogueRows(i), 2).Value) = "" Then
            unassigned = unassigned + 1
        End If
    Next i

    If unassigned > 0 Then
        Dim proceed As VbMsgBoxResult
        proceed = MsgBox(unassigned & " line(s) have no character assigned." & vbCrLf & _
                         "Continue anyway?", vbYesNo + vbExclamation)
        If proceed = vbNo Then Exit Sub
    End If

    Me.Hide
    FinalizeManual
End Sub

' ============================================================
'  LISTBOX1
' ============================================================
Private Sub ListBox1_Click()
    If ListBox1.ListIndex >= 0 Then
        TextBox1.Value = ListBox1.List(ListBox1.ListIndex)
    End If
End Sub

' ============================================================
'  TEXTBOX1
' ============================================================
Private Sub TextBox1_Change()
    If Trim(TextBox1.Value) <> "" Then
        ListBox1.ListIndex = -1
    End If
End Sub

' ============================================================
'  PRIVATE HELPERS
' ============================================================
Private Sub UpdateDisplay()
    Label1.Caption = "Line " & currentIndex & " / " & totalLines

    Dim currentRow As Long
    currentRow = dialogueRows(currentIndex)
    Label3.Caption = msSheet.Cells(currentRow, 3).Value

    If currentIndex > 1 Then
        Label4.Caption = "Previous: " & msSheet.Cells(dialogueRows(currentIndex - 1), 3).Value
    Else
        Label4.Caption = "Previous: (none)"
    End If

    If currentIndex < totalLines Then
        Label5.Caption = "Next: " & msSheet.Cells(dialogueRows(currentIndex + 1), 3).Value
    Else
        Label5.Caption = "Next: (none)"
    End If

    ' Pre-select existing assignment if present
    Dim existing As String
    existing = Trim(msSheet.Cells(currentRow, 2).Value)
    ListBox1.ListIndex = -1
    TextBox1.Value = ""
    If existing <> "" Then
        Dim k As Long
        For k = 0 To ListBox1.ListCount - 1
            If LCase(ListBox1.List(k)) = LCase(existing) Then
                ListBox1.ListIndex = k
                Exit For
            End If
        Next k
    End If
End Sub

Private Sub AssignCurrentLine(charName As String)
    msSheet.Cells(dialogueRows(currentIndex), 2).Value = charName
End Sub

Private Sub RefreshListBox()
    ListBox1.Clear
    Dim lastRowA As Long
    lastRowA = msSheet.Cells(msSheet.Rows.count, "A").End(xlUp).Row
    Dim i As Long
    For i = 1 To lastRowA
        Dim nm As String
        nm = Trim(msSheet.Cells(i, 1).Value)
        If nm <> "" Then
            ListBox1.AddItem nm
        End If
    Next i
End Sub

Private Function NameExistsInList(nm As String) As Boolean
    Dim k As Long
    For k = 0 To ListBox1.ListCount - 1
        If LCase(ListBox1.List(k)) = LCase(nm) Then
            NameExistsInList = True
            Exit Function
        End If
    Next k
    NameExistsInList = False
End Function

Private Function IsEmptyOrWhitespace(ByVal txt As String) As Boolean
    If Len(txt) = 0 Then
        IsEmptyOrWhitespace = True
        Exit Function
    End If
    Dim i As Long
    For i = 1 To Len(txt)
        Select Case Mid(txt, i, 1)
            Case " ", vbTab, vbCr, vbLf, Chr(160)
            Case Else
                IsEmptyOrWhitespace = False
                Exit Function
        End Select
    Next i
    IsEmptyOrWhitespace = True
End Function

