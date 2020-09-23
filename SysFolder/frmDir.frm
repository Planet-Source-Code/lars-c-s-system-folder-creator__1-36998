VERSION 5.00
Begin VB.Form frmDir 
   BackColor       =   &H00843201&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Path selector"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "frmDir.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox wiz4 
      Caption         =   "Check6"
      Height          =   255
      Left            =   840
      TabIndex        =   23
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox wiz1 
      Caption         =   "Check6"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox wiz2 
      Caption         =   "Check6"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox wiz3 
      Caption         =   "Check6"
      Height          =   255
      Left            =   600
      TabIndex        =   20
      Top             =   6240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00843201&
      Caption         =   "Files"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00843201&
      Caption         =   "Folder"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4080
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.CheckBox T10 
      Caption         =   "Check4"
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox T11 
      Caption         =   "Check5"
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox T12 
      Caption         =   "Check6"
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox T7 
      Caption         =   "Check4"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox T8 
      Caption         =   "Check5"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox T9 
      Caption         =   "Check6"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C4C28F&
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C4C28F&
      Caption         =   "OK"
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   975
   End
   Begin VB.CheckBox T6 
      Caption         =   "Check6"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox T5 
      Caption         =   "Check5"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox T4 
      Caption         =   "Check4"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox T3 
      Caption         =   "Check3"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox T2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox T1 
      Caption         =   "Check1"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF80&
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2355
   End
   Begin VB.FileListBox File1 
      Height          =   3990
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF80&
      Height          =   3690
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   2355
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Here can you select a file to open insteed of an folder."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      TabIndex        =   19
      Top             =   4080
      Width           =   2055
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
frmmain.Show
Unload Me
End Sub

Private Sub Command11_Click()
If Option1.Value = True Then
If T1.Value = 1 Then S1
If T2.Value = 1 Then S2
If T3.Value = 1 Then S3
If T4.Value = 1 Then S4
If T5.Value = 1 Then S5
If T6.Value = 1 Then S6
If T7.Value = 1 Then S7
If T8.Value = 1 Then S8
If T9.Value = 1 Then S9
If T10.Value = 1 Then S10
If T11.Value = 1 Then S11

If wiz1.Value = 1 Then wa1
If wiz2.Value = 1 Then wa2
If wiz3.Value = 1 Then wa3
If wiz4.Value = 1 Then wa4
'If T12.Value = 1 Then S12
Else
If T1.Value = 1 Then F1
If T2.Value = 1 Then F2
If T3.Value = 1 Then F3
If T4.Value = 1 Then F4
If T5.Value = 1 Then F5
If T6.Value = 1 Then F6
If T7.Value = 1 Then F7
If T8.Value = 1 Then F8
If T9.Value = 1 Then F9
If T10.Value = 1 Then F10
If T11.Value = 1 Then F11

If wiz1.Value = 1 Then wb1
If wiz2.Value = 1 Then wb2
If wiz3.Value = 1 Then wb3
If wiz4.Value = 1 Then wb4
End If
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Public Sub T1T()
T1.Value = 1
End Sub
Public Sub T2T()
T2.Value = 1
End Sub
Public Sub T3T()
T3.Value = 1
End Sub
Public Sub T4T()
T4.Value = 1
End Sub
Public Sub T5T()
T5.Value = 1
End Sub
Public Sub T6T()
T6.Value = 1
End Sub
Public Sub T7T()
T7.Value = 1
End Sub
Public Sub T8T()
T8.Value = 1
End Sub
Public Sub T9T()
T9.Value = 1
End Sub
Public Sub T10T()
T10.Value = 1
End Sub
Public Sub T11T()
T11.Value = 1
End Sub
Public Sub T12T()
T12.Value = 1
End Sub

Public Sub Wizz1()
wiz1.Value = 1
End Sub
Public Sub Wizz2()
wiz2.Value = 1
End Sub
Public Sub Wizz3()
wiz3.Value = 1
End Sub
Public Sub Wizz4()
wiz4.Value = 1
End Sub

Public Sub S1()
File1.Path = Dir1.Path
        If Len(Dir1.Path) > 3 Then
            frmmain.Text4.Text = Dir1.Path & "\"
        Else
            frmmain.Text4.Text = Dir1.Path & "\"
        End If
    frmmain.Show
            Unload Me
End Sub
Public Sub S2()
File1.Path = Dir1.Path
        If Len(Dir1.Path) > 3 Then
            frmmain.Text6.Text = Dir1.Path & "\"
        Else
            frmmain.Text6.Text = Dir1.Path & "\"
        End If
    frmmain.Show
    frmmain.Command31.Caption = "Folder"

            Unload Me
End Sub
Public Sub S3()
File1.Path = Dir1.Path
        If Len(Dir1.Path) > 3 Then
            frmmain.Text8.Text = Dir1.Path & "\"
        Else
            frmmain.Text8.Text = Dir1.Path & "\"
        End If
    frmmain.Show
    frmmain.Command32.Caption = "Folder"

            Unload Me
End Sub
Public Sub S4()
File1.Path = Dir1.Path
        If Len(Dir1.Path) > 3 Then
            frmmain.Text10.Text = Dir1.Path & "\"
        Else
            frmmain.Text10.Text = Dir1.Path & "\"
        End If
    frmmain.Show
    frmmain.Command33.Caption = "Folder"

            Unload Me
End Sub
Public Sub S5()
File1.Path = Dir1.Path
        If Len(Dir1.Path) > 3 Then
            frmmain.Text12.Text = Dir1.Path & "\"
        Else
            frmmain.Text12.Text = Dir1.Path & "\"
        End If
    frmmain.Show
    frmmain.Command34.Caption = "Folder"

            Unload Me
End Sub
Public Sub S6()
File1.Path = Dir1.Path
        If Len(Dir1.Path) > 3 Then
            frmmain.Text14.Text = Dir1.Path & "\"
        Else
            frmmain.Text14.Text = Dir1.Path & "\"
        End If
    frmmain.Show
    frmmain.Command35.Caption = "Folder"

            Unload Me
End Sub
Public Sub S7()
File1.Path = Dir1.Path
        If Len(Dir1.Path) > 3 Then
            frmmain.Text2.Text = Dir1.Path & "\"
        Else
            frmmain.Text2.Text = Dir1.Path & "\"
        End If
    frmmain.Show
    frmmain.Command36.Caption = "Folder"

            Unload Me
End Sub
Public Sub S8()
File1.Path = Dir1.Path
        If Len(Dir1.Path) > 3 Then
            frmmain.Text15.Text = Dir1.Path & "\"
        Else
            frmmain.Text15.Text = Dir1.Path & "\"
        End If
    frmmain.Show
    frmmain.Command37.Caption = "Folder"

            Unload Me
End Sub
Public Sub S9()
File1.Path = Dir1.Path
        If Len(Dir1.Path) > 3 Then
            frmmain.Text17.Text = Dir1.Path & "\"
        Else
            frmmain.Text17.Text = Dir1.Path & "\"
        End If
    frmmain.Show
    frmmain.Command38.Caption = "Folder"
            Unload Me
End Sub
Public Sub S10()
File1.Path = Dir1.Path
        If Len(Dir1.Path) > 3 Then
            frmmain.Text19.Text = Dir1.Path & "\"
        Else
            frmmain.Text19.Text = Dir1.Path & "\"
        End If
    frmmain.Show
    frmmain.Command39.Caption = "Folder"

            Unload Me
End Sub
Public Sub S11()
File1.Path = Dir1.Path
        If Len(Dir1.Path) > 3 Then
            frmmain.Text21.Text = Dir1.Path & "\"
        Else
            frmmain.Text21.Text = Dir1.Path & "\"
        End If
    frmmain.Show
    frmmain.Command40.Caption = "Folder"

            Unload Me
End Sub

Public Sub F1()
MsgBox ("Sorry but you ,must select a folder for this path.")
'Option1.Value = True
'Option2.Value = False
End Sub
Public Sub F2()
    frmmain.Show
    frmmain.Text6.Text = Dir1.Path & "\" & File1.FileName
    frmmain.Command31.Caption = "File"
    Unload Me
End Sub
Public Sub F3()
    frmmain.Show
    frmmain.Text8.Text = Dir1.Path & "\" & File1.FileName
    frmmain.Command32.Caption = "File"
    Unload Me
End Sub
Public Sub F4()
    frmmain.Show
    frmmain.Text10.Text = Dir1.Path & "\" & File1.FileName
    frmmain.Command33.Caption = "File"
    Unload Me
End Sub
Public Sub F5()
    frmmain.Show
    frmmain.Text12.Text = Dir1.Path & "\" & File1.FileName
    frmmain.Command34.Caption = "File"
    Unload Me
End Sub
Public Sub F6()
    frmmain.Show
    frmmain.Text14.Text = Dir1.Path & "\" & File1.FileName
    frmmain.Command35.Caption = "File"
    Unload Me
End Sub
Public Sub F7()
    frmmain.Show
    frmmain.Text2.Text = Dir1.Path & "\" & File1.FileName
    frmmain.Command36.Caption = "File"
    Unload Me
End Sub
Public Sub F8()
    frmmain.Show
    frmmain.Text15.Text = Dir1.Path & "\" & File1.FileName
    frmmain.Command37.Caption = "File"
    Unload Me
End Sub
Public Sub F9()
    frmmain.Show
    frmmain.Text17.Text = Dir1.Path & "\" & File1.FileName
    frmmain.Command38.Caption = "File"
    Unload Me
End Sub
Public Sub F10()
    frmmain.Show
    frmmain.Text19.Text = Dir1.Path & "\" & File1.FileName
    frmmain.Command39.Caption = "File"
    Unload Me
End Sub
Public Sub F11()
    frmmain.Show
    frmmain.Text21.Text = Dir1.Path & "\" & File1.FileName
    frmmain.Command40.Caption = "File"
    Unload Me
End Sub




Private Sub Form_Load()
Me.Width = "2475"
End Sub

Private Sub Form_Unload(Cancel As Integer)
T1.Value = 0
T2.Value = 0
T3.Value = 0
T4.Value = 0
T5.Value = 0
T6.Value = 0
T7.Value = 0
T8.Value = 0
T9.Value = 0
T10.Value = 0
T11.Value = 0
T12.Value = 0

wiz1.Value = 0
wiz2.Value = 0
wiz3.Value = 0
wiz4.Value = 0
End Sub

Private Sub Option1_Click()
Me.Width = "2475"
End Sub

Private Sub Option2_Click()
Me.Width = "4650"
End Sub

Public Sub wa1()
    frmWiz.Show
    frmWiz.Text4.Text = Dir1.Path & "\" ' & File1.FileName
    Unload Me
End Sub
Public Sub wa2()
    frmWiz.Show
    frmWiz.Text6.Text = Dir1.Path & "\" ' & File1.FileName
    frmWiz.Command31.Caption = "Folder"
    Unload Me
End Sub
Public Sub wa3()
    frmWiz.Show
    frmWiz.Text8.Text = Dir1.Path & "\" ' & File1.FileName
    frmWiz.Command32.Caption = "Folder"
    Unload Me
End Sub
Public Sub wa4()
    frmWiz.Show
    frmWiz.Text10.Text = Dir1.Path & "\" ' & File1.FileName
    frmWiz.Command33.Caption = "Folder"
    Unload Me
End Sub
'----------------
Public Sub wb1()
MsgBox ("Sorry but you ,must select a folder for this path.")
'Option1.Value = True
'Option2.Value = False
End Sub
Public Sub wb2()
    frmWiz.Show
    frmWiz.Text6.Text = Dir1.Path & "\" & File1.FileName
    frmWiz.Command31.Caption = "File"
    Unload Me
End Sub
Public Sub wb3()
    frmWiz.Show
    frmWiz.Text8.Text = Dir1.Path & "\" & File1.FileName
    frmWiz.Command32.Caption = "File"
    Unload Me
End Sub
Public Sub wb4()
    frmWiz.Show
    frmWiz.Text10.Text = Dir1.Path & "\" & File1.FileName
    frmWiz.Command33.Caption = "File"
    Unload Me
End Sub
