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
Private dialogueRows() As Long  ' Stores row indices of lines in col C
Private currentIndex As Long    ' Current position in dialogueRows
Private totalLines As Long      ' Total number of dialogue lines


Private Sub CommandButton5_Click()
    Dim lastRowC As Long
    lastRowC = ms.Cells(ms.Rows.count, "C").End(xlUp).Row

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
        If proceed = vbYes Then
            FinalizeManual
            Me.Hide
        End If
    End If
End Sub

' ============================================================
'  INITIALIZE
' ============================================================
Private Sub UserForm_Initialize()
    ' Check if ms is set
    If ms Is Nothing Then
        MsgBox "Please set the ms variable before showing the form!", vbCritical
        Unload Me
        Exit Sub
    End If
    
    ' --- Build index of all rows that have dialogue in col C ---
    Dim lastRowC As Long
    lastRowC = ms.Cells(ms.Rows.count, "C").End(xlUp).Row

    Dim tempRows() As Long
    ReDim tempRows(1 To lastRowC)
    Dim count As Long
    count = 0

    Dim i As Long
    For i = 1 To lastRowC
        If Trim(ms.Cells(i, 3).Value) <> "" Then
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

    ' --- Load existing character names from col A into ListBox ---
    RefreshListBox

    ' --- Start at first line ---
    currentIndex = 1
    UpdateDisplay
End Sub

' ============================================================
'  COMMANDBUTTON1 — Add newly typed name, assign to current
'  line, advance to next line
' ============================================================
Private Sub CommandButton1_Click()
    Dim newName As String
    newName = Trim(TextBox1.Value)

    If newName = "" Then
        MsgBox "Please type a character name first.", vbExclamation
        Exit Sub
    End If

    ' Add to col A if not already there
    If Not NameExistsInList(newName) Then
        Dim lastRowA As Long
        lastRowA = ms.Cells(ms.Rows.count, "A").End(xlUp).Row
        ' Skip instruction rows
        If lastRowA < 3 Then lastRowA = 3
        ms.Cells(lastRowA + 1, 1).Value = newName
        RefreshListBox
    End If

    ' Assign to current line
    AssignCurrentLine newName

    TextBox1.Value = ""

    ' Advance
    If currentIndex < totalLines Then
        currentIndex = currentIndex + 1
        UpdateDisplay
    Else
        MsgBox "All lines assigned. Click OK then run FinalizeManual.", vbInformation
    End If
End Sub

' ============================================================
'  COMMANDBUTTON2 — Assign selected name from ListBox to
'  current line, advance to next
' ============================================================
Private Sub CommandButton2_Click()
    If ListBox1.ListIndex = -1 Then
        MsgBox "Please select a character from the list.", vbExclamation
        Exit Sub
    End If

    Dim selectedName As String
    selectedName = ListBox1.List(ListBox1.ListIndex)
    AssignCurrentLine selectedName

    ' Advance
    If currentIndex < totalLines Then
        currentIndex = currentIndex + 1
        UpdateDisplay
    Else
        MsgBox "All lines assigned. Click OK then run FinalizeManual.", vbInformation
    End If
End Sub

' ============================================================
'  COMMANDBUTTON3 — Go to previous line
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
'  COMMANDBUTTON4 — Go to next line without assigning
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
'  LISTBOX1 — Show currently selected name in textbox
'  for easy editing if needed
' ============================================================
Private Sub ListBox1_Click()
    If ListBox1.ListIndex >= 0 Then
        TextBox1.Value = ListBox1.List(ListBox1.ListIndex)
    End If
End Sub

' ============================================================
'  TEXTBOX1 — Clear ListBox selection when user starts typing
'  a new name so it doesn't interfere with Add & Assign
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
    ' Page counter label
    Label1.Caption = "Line " & currentIndex & " / " & totalLines

    ' Dialogue label — show the line the user is assigning
    Dim currentRow As Long
    currentRow = dialogueRows(currentIndex)
    Label3.Caption = ms.Cells(currentRow, 3).Value
    Label4.Caption = "Previous Line: " & ms.Cells(currentRow - 1, 3).Value
    Label5.Caption = "Next Line: " & ms.Cells(currentRow + 1, 3).Value

    ' Show existing assignment in col B if already set
    Dim existing As String
    existing = Trim(ms.Cells(currentRow, 2).Value)
    If existing <> "" Then
        ' Pre-select in ListBox if found
        Dim k As Long
        For k = 0 To ListBox1.ListCount - 1
            If LCase(ListBox1.List(k)) = LCase(existing) Then
                ListBox1.ListIndex = k
                Exit For
            End If
        Next k
        TextBox1.Value = ""
    Else
        ListBox1.ListIndex = -1
        TextBox1.Value = ""
    End If
End Sub

Private Sub AssignCurrentLine(charName As String)
    Dim currentRow As Long
    currentRow = dialogueRows(currentIndex)
    ms.Cells(currentRow, 2).Value = charName
End Sub

Private Sub RefreshListBox()
    ListBox1.Clear
    Dim lastRowA As Long
    lastRowA = ms.Cells(ms.Rows.count, "A").End(xlUp).Row
    Dim i As Long
    For i = 1 To lastRowA
        Dim nm As String
        nm = Trim(ms.Cells(i, 1).Value)
        ' Skip instruction text
        If nm <> "" And nm <> "TYPE character names here" And _
           nm <> "Then run RefreshDropdowns" Then
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
