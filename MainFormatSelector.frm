VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainFormatSelector 
   Caption         =   "Format Selector"
   ClientHeight    =   3885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9645.001
   OleObjectBlob   =   "MainFormatSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainFormatSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    ' OK button — validate a selection was made then close
    If Not OptionButton1.Value And Not OptionButton2.Value And Not OptionButton3.Value Then
        MsgBox "Please select a script format before continuing.", vbExclamation
        Exit Sub
    End If
    Me.Hide
End Sub

Private Sub CommandButton2_Click()
    ' Cancel button — unload without storing a result
    Me.Tag = "cancelled"
    Unload Me
End Sub

' ============================================================
'  PUBLIC PROPERTIES — read by RunProgram1 after form closes
' ============================================================
Public Property Get SelectedFormat() As Integer
    If OptionButton1.Value Then
        SelectedFormat = 1 ' Two-line
    ElseIf OptionButton2.Value Then
        SelectedFormat = 99 ' Inline
    ElseIf OptionButton3.Value Then
        SelectedFormat = 98  ' Manual Mode
    Else
        SelectedFormat = 0  ' Nothing selected
    End If
End Property
' H: 223.5
' W: 495
Private Sub UserForm_Click()

End Sub
