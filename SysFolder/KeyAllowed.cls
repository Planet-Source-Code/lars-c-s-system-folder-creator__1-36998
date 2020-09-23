VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtendedTextbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'--------------EX---------------
'Dim ExTextbox1 As New ExtendedTextbox
'Private Sub form_load()
'    ExTextbox1.BindControl Text1
'    ExTextbox1.ValidKeys = "abcdefghijklmnopqrstuvwxyzæøåABCDEFGHIJKLMNOPQRSTUVWXYZÆØÅ0123456789-_! "
'End Sub


Public ValidKeys As String

'format of the textbox
Public TextFormat As String

'backcolor when active

'Public Active_BackColor As OLE_COLOR

'backcolor when deactive

'Public Deactive_BackColor As OLE_COLOR

'this is the object that will intercept the events from the real rextbox
Private WithEvents myTextBox As textbox
Attribute myTextBox.VB_VarHelpID = -1





Private Sub Class_Initialize()
'    Active_BackColor = vbYellow
'    Deactive_BackColor = vbWhite
End Sub

Private Sub myTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
    'if the user press enter , replace that keypress with a tab press
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        SendKeys vbTab
    End If
    
End Sub

Private Sub myTextBox_KeyPress(KeyAscii As Integer)

    'if we havent applied any valid keys , just skip this event
    If ValidKeys = "" Then Exit Sub
    
    'test if the key we pressed is allowed in this textbox
    'also force the "back" key to be valid
    If InStr(ValidKeys & vbBack, Chr(KeyAscii)) = 0 Then
        '..No the key was not valid , so lets skip it..
        KeyAscii = 0
    End If
    
End Sub

Private Sub myTextBox_LostFocus()
    'the textbox was deactivated , lets set the deactive backcolor
'    myTextBox.BackColor = Deactive_BackColor
    
    'if we have applied a format to the textbox , lets aply the
    'format to the content now.
    If TextFormat <> "" Then
        myTextBox.text = Format(myTextBox.text, TextFormat)
    End If
End Sub

Private Sub myTextBox_GotFocus()
    'the textbox got focus
    'lets set the active backcolor
'    myTextBox.BackColor = Active_BackColor
End Sub

Public Sub BindControl(textbox As textbox)
    'lets set our reference to the passed textbox
    'this will allow us to intercept the events to this textbox
    Set myTextBox = textbox
End Sub
