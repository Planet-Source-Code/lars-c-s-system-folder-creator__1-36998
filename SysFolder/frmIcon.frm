VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFileIcon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose an icon"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIcon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShowIcons 
      Caption         =   "Show &Icons"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   3120
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picDefault 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3600
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picBuff 
      AutoRedraw      =   -1  'True
      Height          =   540
      Left            =   3000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox pic32 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3240
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   7
      Top             =   2265
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pic16 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   3105
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   3720
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse.."
      Height          =   285
      Left            =   6120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "C:\WINDOWS\SYSTEM\SHELL32.DLL"
      Top             =   120
      Width           =   4935
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   4575
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8070
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "IconPath:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5280
      Width           =   7335
   End
   Begin VB.Label lblInfo 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   7335
   End
   Begin VB.Label lblLabel 
      Caption         =   "Select a file:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmFileIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lngIcon As Long
Public strProgram As String
Public strSaveIconFile
Public NoOfIcons As Integer

Private Const ILD_TRANSPARENT = &H1       'Display transparent

'Api for extracting the first icon in a dll or exe file
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long
Dim iconpath As String

Private Sub cmdBrowse_Click()
On Error Resume Next
    ComDlg.DialogTitle = "Select File to Extract Icon"
    ComDlg.Filter = "Executable Files (*.EXE)|*.exe|Dynamic Link Libraries (*.DLL)|*.dll|Icon Files (*.ICO)|*.ico|All Files (*.*)|*.*"
    ComDlg.ShowOpen
    If Trim(ComDlg.FileName) = "" Then Exit Sub
    If FileLen(ComDlg.FileName) = 0 Then
        MsgBox "File does not exist!!", vbCritical, App.Title
        Exit Sub
    Else
        txtFile.Text = ComDlg.FileName
        Me.Caption = "Icon from - [" & txtFile.Text & "]"
        cmdShowIcons_Click
    End If
End Sub

Private Sub cmdExit_Click()
    frmmain.Show
    Unload Me
End Sub

Private Sub Initialise()
On Local Error Resume Next
    'Break the link to iml lists
    lvListView.ListItems.clear
    lvListView.Icons = Nothing
    lvListView.SmallIcons = Nothing
    'Clear the image lists
    iml32.ListImages.clear
    iml16.ListImages.clear
    strProgram = txtFile.Text
    If FileLen(txtFile.Text) = 0 Then
        lblInfo.Caption = "File does not exist!!!"
        Exit Sub
    End If
    NoOfIcons = ExtractIcon(App.hInstance, strProgram, -1)
    lblInfo.Caption = "File: " & strProgram & vbCrLf
    If NoOfIcons = 0 Then
        lblInfo.Caption = lblInfo.Caption & "No Icons Available in Selected File."
    Else
        lblInfo.Caption = lblInfo.Caption & "No of Icons: " & CStr(NoOfIcons)
    End If
End Sub
Private Function GetIcons()
On Error Resume Next
    Initialise
    If FileLen(txtFile.Text) = 0 Then
        Exit Function
    End If
    Dim hLIcon As Long    'Large & Small Icons
    Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
    Dim r As Long
    Dim i As Integer
    i = 0
    Do While i < NoOfIcons
        DestroyIcon lngIcon
        lngIcon = ExtractIcon(App.hInstance, strProgram, i)
        With pic32
            .Cls
            .AutoSize = True
            .AutoRedraw = True
            DrawIcon .hdc, 0, 0, lngIcon
            .Refresh
        End With
        Set imgObj = iml32.ListImages.Add(i + 1, , pic32.Image)
        i = i + 1
    Loop
    ShowIcons
End Function
Private Sub ShowIcons()
On Error Resume Next
    Dim i As Integer
    i = 0
    Do While i < NoOfIcons
        lvListView.ListItems.Add
        'list1.AddItem
        i = i + 1
    Loop
    Dim Item As ListItem
    With lvListView
      '.ListItems.Clear
      .Icons = iml32        'Large
      For Each Item In .ListItems
        Item.Icon = Item.Index
      Next
    End With
End Sub

Private Sub cmdSave_Click()
frmmain.text38 = iconpath
frmmain.Show
Unload Me
End Sub

Private Sub cmdShowIcons_Click()
    GetIcons
End Sub

Private Sub Form_Load()
txtFile = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\", "IcoPath", "")
    GetIcons
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\", "IcoPath", "" & txtFile & ""
End Sub

Private Sub lvListView_DblClick()
cmdSave_Click
End Sub

Private Sub lvListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    picBuff.Cls
    picBuff.AutoRedraw = True
    picBuff.AutoSize = True
    hPic = iml32.ListImages(Item.Index).ExtractIcon.Handle
    DrawIcon picBuff.hdc, 0, 0, hPic
    Label1.Caption = "IconPath: " & txtFile & "," & Item.Index - 1
    Label1.ToolTipText = Label1.Caption
    iconpath = txtFile & "," & Item.Index - 1
    picBuff.Refresh
End Sub

Private Sub txtFile_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn:
            cmdShowIcons_Click
    End Select
End Sub
Private Sub saveico()
On Error Resume Next
    Dim Res, hPic As Long
    ComDlg.DialogTitle = "Save Icon as..."
    ComDlg.CancelError = True
    ComDlg.Filter = "Bitmap Images (*.BMP) | *.bmp|Icon Files (*.ICO)|*.ico"
LinktoSaveWindow:
    ComDlg.ShowSave
    If FileLen(ComDlg.FileName) = 0 Then
        If LCase(Right(ComDlg.FileName, 3)) = "ico" Then
            SavePicture iml32.ListImages(lvListView.SelectedItem.Index).ExtractIcon, ComDlg.FileName
        Else
            SavePicture picBuff.Image, ComDlg.FileName
        End If
        MsgBox "Icon saved as " & ComDlg.FileName, vbExclamation, App.Title
    Else
        Res = MsgBox("File already exists. Overwrite it?", vbQuestion + vbYesNoCancel, App.Title)
        If Res = vbYes Then
            If LCase(Right(ComDlg.FileName, 3)) = "ico" Then
                SavePicture iml32.ListImages(lvListView.SelectedItem.Index).ExtractIcon, ComDlg.FileName
            Else
                SavePicture picBuff.Image, ComDlg.FileName
            End If
            MsgBox "Icon saved as " & ComDlg.FileName, vbExclamation, App.Title
        ElseIf Res = vbNo Then
            GoTo LinktoSaveWindow
        Else
            Exit Sub
        End If
    End If
End Sub
