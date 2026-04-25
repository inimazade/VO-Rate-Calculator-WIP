VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TwoSubFormatSelector 
   Caption         =   "UserForm1"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6945
   OleObjectBlob   =   "TwoSubFormatSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TwoSubFormatSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    ' OK button — validate a selection was made then close
    If Not OptionButton1.Value And Not OptionButton2.Value And _
       Not OptionButton3.Value And Not OptionButton4.Value And _
       Not OptionButton5.Value Then
        MsgBox "Please select a name format before continuing.", vbExclamation
        Exit Sub
    End If
    Me.Hide
End Sub

Private Sub CommandButton2_Click()
    Me.Tag = "cancelled"
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "Character Name Format"
End Sub

' ============================================================
'  PUBLIC PROPERTY — read by RunProgram1 after form closes
'  Returns the case number matching ImportScript's Select Case
' ============================================================
Public Property Get SelectedStyle() As Integer
    If OptionButton1.Value Then
        SelectedStyle = 1  ' Colon
    ElseIf OptionButton2.Value Then
        SelectedStyle = 3  ' Plain name
    ElseIf OptionButton3.Value Then
        SelectedStyle = 2  ' ALL CAPS
    ElseIf OptionButton4.Value Then
        SelectedStyle = 4  ' Blank line
    ElseIf OptionButton5.Value Then
        SelectedStyle = 5  ' Vocal cue
    Else
        SelectedStyle = 0
    End If
End Property

' H: 395.25
' W: 359.25
