VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00843201&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Folder Creator"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00843201&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      TabIndex        =   110
      Top             =   1780
      Width           =   1695
      Begin VB.OptionButton mode1 
         BackColor       =   &H00843201&
         Caption         =   "Classic Mode"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   112
         Top             =   0
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton mode2 
         BackColor       =   &H00843201&
         Caption         =   "Advanced Mode"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   111
         Top             =   260
         Width           =   1575
      End
   End
   Begin VB.TextBox FolderNumber 
      Height          =   285
      Left            =   120
      MaxLength       =   2
      TabIndex        =   109
      Top             =   4080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00C4C28F&
      Caption         =   "Delete"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00C4C28F&
      Caption         =   "Create"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox listfolder 
      Height          =   1425
      Left            =   240
      TabIndex        =   106
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00C4C28F&
      Caption         =   "&Wizard"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00843201&
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   7920
      TabIndex        =   102
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton Command18 
         BackColor       =   &H00C4C28F&
         Caption         =   "Select NOTHING"
         Height          =   205
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   60
         Width           =   2295
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   47
         Left            =   1920
         Top             =   6960
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   46
         Left            =   1920
         Top             =   6360
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   45
         Left            =   1920
         Top             =   5760
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   44
         Left            =   1920
         Top             =   5160
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   43
         Left            =   1920
         Top             =   4560
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   42
         Left            =   1920
         Top             =   3960
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   41
         Left            =   1920
         Top             =   3360
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   40
         Left            =   1920
         Top             =   2760
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   39
         Left            =   1920
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   38
         Left            =   1920
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   37
         Left            =   1920
         Top             =   960
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   36
         Left            =   1920
         Top             =   360
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   495
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   35
         Left            =   1320
         Top             =   6960
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   34
         Left            =   1320
         Top             =   6360
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   33
         Left            =   1320
         Top             =   5760
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   32
         Left            =   1320
         Top             =   5160
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   31
         Left            =   1320
         Top             =   4560
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   30
         Left            =   1320
         Top             =   3960
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   29
         Left            =   1320
         Top             =   3360
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   28
         Left            =   1320
         Top             =   2760
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   27
         Left            =   1320
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   26
         Left            =   1320
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   25
         Left            =   1320
         Top             =   960
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   24
         Left            =   1320
         Top             =   360
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   23
         Left            =   720
         Top             =   6960
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   22
         Left            =   720
         Top             =   6360
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   21
         Left            =   720
         Top             =   5760
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   20
         Left            =   720
         Top             =   5160
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   19
         Left            =   720
         Top             =   4560
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   18
         Left            =   720
         Top             =   3960
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   17
         Left            =   720
         Top             =   3360
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   16
         Left            =   720
         Top             =   2760
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   15
         Left            =   720
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   14
         Left            =   720
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   13
         Left            =   720
         Top             =   960
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   12
         Left            =   720
         Top             =   360
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   11
         Left            =   120
         Top             =   6960
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   10
         Left            =   120
         Top             =   6360
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   9
         Left            =   120
         Top             =   5760
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   8
         Left            =   120
         Top             =   5160
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   7
         Left            =   120
         Top             =   4560
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   6
         Left            =   120
         Top             =   3960
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   5
         Left            =   120
         Top             =   3360
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   4
         Left            =   120
         Top             =   2760
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   3
         Left            =   120
         Top             =   2160
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   2
         Left            =   120
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   1
         Left            =   120
         Top             =   960
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   0
         Left            =   120
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C4C28F&
      Caption         =   "&About SysFolder"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   1920
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C4C28F&
      Caption         =   "EXIT"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Frame Sys1 
      BackColor       =   &H00843201&
      Caption         =   "System folder 1"
      ForeColor       =   &H8000000E&
      Height          =   7455
      Left            =   2280
      TabIndex        =   17
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton Command17 
         BackColor       =   &H00C4C28F&
         Caption         =   ">>--->>>"
         Height          =   255
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   7080
         Width           =   735
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   7080
         Width           =   255
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00843201&
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   120
         TabIndex        =   92
         Top             =   650
         Width           =   5315
         Begin VB.OptionButton Tool1 
            BackColor       =   &H00843201&
            Caption         =   "Extra Long3"
            ForeColor       =   &H00FFC0C0&
            Height          =   195
            Index           =   6
            Left            =   3720
            TabIndex        =   100
            ToolTipText     =   ">===>>===>>>===>>>>   NAME   <<<<===<<<===<<===<"
            Top             =   630
            Width           =   1575
         End
         Begin VB.OptionButton Tool1 
            BackColor       =   &H00843201&
            Caption         =   "Extra Long2"
            ForeColor       =   &H00FFC0C0&
            Height          =   195
            Index           =   5
            Left            =   3720
            TabIndex        =   99
            ToolTipText     =   "(¯`'·.¸(¯`'·.¸(¯`'·.¸  NAME ¸.·'´¯)¸.·'´¯)¸.·'´¯)"
            Top             =   390
            Width           =   1575
         End
         Begin VB.OptionButton Tool1 
            BackColor       =   &H00843201&
            Caption         =   "Extra Long1"
            ForeColor       =   &H00FFC0C0&
            Height          =   195
            Index           =   4
            Left            =   3720
            TabIndex        =   98
            ToolTipText     =   "°º¤¤º°º¤.,¸¸,.¤º°º¤° NAME °¤º°º¤.,¸¸,.¤º°º¤°¤º°"
            Top             =   150
            Width           =   1575
         End
         Begin VB.OptionButton Tool1 
            BackColor       =   &H00843201&
            Caption         =   """¨¨°º©o.,NAME,.o©º°¨¨"""
            ForeColor       =   &H00FFC0C0&
            Height          =   195
            Index           =   3
            Left            =   1680
            TabIndex        =   97
            Top             =   630
            Width           =   2055
         End
         Begin VB.OptionButton Tool1 
            BackColor       =   &H00843201&
            Caption         =   "¯`·.¸¸.->NAME<-.¸¸.·´¯"
            ForeColor       =   &H00FFC0C0&
            Height          =   195
            Index           =   2
            Left            =   1680
            TabIndex        =   96
            Top             =   390
            Width           =   2055
         End
         Begin VB.OptionButton Tool1 
            BackColor       =   &H00843201&
            Caption         =   "(¯`'·.¸  NAME ¸.·'´¯)"
            ForeColor       =   &H00FFC0C0&
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   95
            Top             =   150
            Width           =   1935
         End
         Begin VB.OptionButton Tool1 
            BackColor       =   &H00843201&
            Caption         =   "NO Change"
            ForeColor       =   &H00FFC0C0&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   93
            Top             =   410
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.Label Label19 
            BackColor       =   &H00843201&
            Caption         =   "AutoGenerate ToolTip:"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   0
            TabIndex        =   94
            Top             =   100
            Width           =   1695
         End
      End
      Begin VB.OptionButton F1C2 
         BackColor       =   &H00843201&
         Caption         =   "My computer"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   3360
         TabIndex        =   26
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00C4C28F&
         Caption         =   "AutoDetect Folder or File"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   6180
         Width           =   5295
      End
      Begin VB.CommandButton Command40 
         BackColor       =   &H00C4C28F&
         Caption         =   "Folder"
         Height          =   255
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   5820
         Width           =   735
      End
      Begin VB.CommandButton Command39 
         BackColor       =   &H00C4C28F&
         Caption         =   "Folder"
         Height          =   255
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   5535
         Width           =   735
      End
      Begin VB.CommandButton Command38 
         BackColor       =   &H00C4C28F&
         Caption         =   "Folder"
         Height          =   255
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   5250
         Width           =   735
      End
      Begin VB.CommandButton Command37 
         BackColor       =   &H00C4C28F&
         Caption         =   "Folder"
         Height          =   255
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4965
         Width           =   735
      End
      Begin VB.CommandButton Command36 
         BackColor       =   &H00C4C28F&
         Caption         =   "Folder"
         Height          =   255
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   4680
         Width           =   735
      End
      Begin VB.CommandButton Command35 
         BackColor       =   &H00C4C28F&
         Caption         =   "Folder"
         Height          =   255
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   4380
         Width           =   735
      End
      Begin VB.CommandButton Command34 
         BackColor       =   &H00C4C28F&
         Caption         =   "Folder"
         Height          =   255
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   4095
         Width           =   735
      End
      Begin VB.CommandButton Command33 
         BackColor       =   &H00C4C28F&
         Caption         =   "Folder"
         Height          =   255
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3810
         Width           =   735
      End
      Begin VB.CommandButton Command32 
         BackColor       =   &H00C4C28F&
         Caption         =   "Folder"
         Height          =   255
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3525
         Width           =   735
      End
      Begin VB.CommandButton Command31 
         BackColor       =   &H00C4C28F&
         Caption         =   "Folder"
         Height          =   255
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3240
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4680
         Width           =   255
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4965
         Width           =   255
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   5250
         Width           =   255
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   5535
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   5820
         Width           =   255
      End
      Begin VB.TextBox Text22 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   960
         TabIndex        =   21
         Top             =   5805
         Width           =   1455
      End
      Begin VB.TextBox Text21 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   22
         Top             =   5805
         Width           =   1815
      End
      Begin VB.TextBox Text20 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   960
         TabIndex        =   19
         Top             =   5520
         Width           =   1455
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   20
         Top             =   5520
         Width           =   1815
      End
      Begin VB.TextBox Text18 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Top             =   5235
         Width           =   1455
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   18
         Top             =   5235
         Width           =   1815
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   960
         TabIndex        =   14
         Top             =   4950
         Width           =   1455
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   15
         Top             =   4950
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   4665
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   13
         Top             =   4665
         Width           =   1815
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   3810
         Width           =   1455
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   4380
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   3525
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Text            =   "Name"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Top             =   4380
         Width           =   1815
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   4095
         Width           =   1815
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   3810
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Top             =   3525
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         Text            =   "Full path"
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   4095
         Width           =   1455
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   4380
         Width           =   255
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   4095
         Width           =   255
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3810
         Width           =   255
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3525
         Width           =   255
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3240
         Width           =   255
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton F1C1 
         BackColor       =   &H00843201&
         Caption         =   "Desktop"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   2400
         TabIndex        =   25
         Top             =   1680
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton F1C3 
         BackColor       =   &H00843201&
         Caption         =   "Both"
         ForeColor       =   &H00FFC0C0&
         Height          =   215
         Left            =   4680
         TabIndex        =   27
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text38 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   600
         TabIndex        =   24
         Text            =   "Remember full path"
         Top             =   7080
         Width           =   3735
      End
      Begin VB.TextBox Text37 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   1080
         TabIndex        =   23
         Text            =   "Can be many words"
         Top             =   6600
         Width           =   4335
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   1
         Text            =   "Type full path"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Text            =   "Name"
         Top             =   240
         Width           =   3375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         X1              =   120
         X2              =   5400
         Y1              =   620
         Y2              =   620
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000009&
         X1              =   120
         X2              =   5400
         Y1              =   2670
         Y2              =   2670
      End
      Begin VB.Label Label18 
         BackColor       =   &H00843201&
         Caption         =   "Number 6:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   4665
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label Label17 
         BackColor       =   &H00843201&
         Caption         =   "Number 7:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   4950
         Width           =   855
      End
      Begin VB.Label Label16 
         BackColor       =   &H00843201&
         Caption         =   "Number 8:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   5235
         Width           =   855
      End
      Begin VB.Label Label15 
         BackColor       =   &H00843201&
         Caption         =   "Number 9:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   5535
         Width           =   855
      End
      Begin VB.Label Label14 
         BackColor       =   &H00843201&
         Caption         =   "Number 10:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   5820
         Width           =   855
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Path"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2520
         TabIndex        =   75
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   960
         TabIndex        =   74
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label30 
         BackColor       =   &H00843201&
         Caption         =   "Icon:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   7120
         Width           =   495
      End
      Begin VB.Line Line23 
         BorderColor     =   &H80000009&
         X1              =   120
         X2              =   5400
         Y1              =   6990
         Y2              =   6990
      End
      Begin VB.Label Label29 
         BackColor       =   &H00843201&
         Caption         =   "Tool tip text:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   6720
         Width           =   975
      End
      Begin VB.Line Line21 
         BorderColor     =   &H80000009&
         X1              =   120
         X2              =   5400
         Y1              =   6510
         Y2              =   6510
      End
      Begin VB.Label Label11 
         BackColor       =   &H00843201&
         Caption         =   "Number 5:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   4380
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H00843201&
         Caption         =   "Number 4:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   4095
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00843201&
         Caption         =   "Number 3:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   3810
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00843201&
         Caption         =   "Number 2:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   3525
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00843201&
         Caption         =   "Number 1:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   3240
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00843201&
         Caption         =   "Where to go when right click:"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   2760
         Width           =   2085
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000009&
         X1              =   120
         X2              =   5400
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   120
         X2              =   5400
         Y1              =   2055
         Y2              =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00843201&
         Caption         =   "Where to go when dobble click:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   2265
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackColor       =   &H00843201&
         Caption         =   "Pick where you want the folder:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1725
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00843201&
         Caption         =   "Pick a name of the folder:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   300
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00843201&
         Height          =   15
         Left            =   120
         TabIndex        =   72
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.TextBox Text52 
      Height          =   285
      Left            =   5880
      TabIndex        =   85
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text51 
      Height          =   285
      Left            =   5400
      TabIndex        =   84
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text50 
      Height          =   285
      Left            =   4920
      TabIndex        =   83
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text49 
      Height          =   285
      Left            =   4440
      TabIndex        =   82
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text48 
      Height          =   285
      Left            =   3960
      TabIndex        =   81
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text46 
      Height          =   285
      Left            =   3000
      TabIndex        =   80
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text40 
      Height          =   285
      Left            =   4560
      TabIndex        =   79
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   2400
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   1920
      Top             =   0
   End
   Begin VB.TextBox N3 
      Height          =   285
      Left            =   3960
      TabIndex        =   78
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox N2 
      Height          =   285
      Left            =   3360
      TabIndex        =   77
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox N1 
      Height          =   285
      Left            =   7080
      TabIndex        =   76
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtByte 
      Height          =   375
      Index           =   3
      Left            =   720
      TabIndex        =   71
      Text            =   "00"
      Top             =   -120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtByte 
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   70
      Text            =   "00"
      Top             =   -120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtByte 
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   69
      Text            =   "00"
      Top             =   -120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtByte 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   68
      Text            =   "00"
      Top             =   -120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C4C28F&
      Caption         =   "Delete Current folder"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C4C28F&
      Caption         =   "Make/Update  folder"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   6240
      Width           =   2055
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00843201&
      Caption         =   "System folder 3"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   55
      Top             =   960
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00843201&
      Caption         =   "System folder 2"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   54
      Top             =   600
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00843201&
      Caption         =   "System folder 1"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   53
      Top             =   240
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.Label Label28 
      BackColor       =   &H00843201&
      ForeColor       =   &H8000000E&
      Height          =   2055
      Left            =   240
      TabIndex        =   65
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   2180
      Y1              =   2380
      Y2              =   2380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00843201&
      Height          =   255
      Left            =   4800
      TabIndex        =   73
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Author: Lars Christian Solberg
' Last modified: 17 July 2002

' Note:
' (1) If you modify, make this code better or anything else please let me know:  xe0r@dr.com
' (2) Most of the code but not every line of code is written by me. Soo if you are the author of some of this code
' and you have some sort of copyright or something please let me know.

' Next version:
' In the next version of this program (2.0) I thinking of making a way that you can create items to your
' system folder dynamicly like the folder list. So it can use over 100 item with name and path.
' The only problem with that is that I have no idea of how I can dynamical create command buttons. If there anybody that can help me with that and
' maybe send me an mail or something??


' This code is quite messy and my english is not that good nether.
' I know lots of this code it NOT good but much of it is very old from the days I started with VB.. :)

' Dont be affraid to run this program becouse its write to the registry.
' If it couldnt write to the registry it wont work. If you will erase all off the registry keys it make simple delete your system folder with this program again.



'\/BEGINNING\/
Option Explicit                         'ALWAYS to this
Dim currentshape As Byte
Dim finishload As Boolean
Dim readymake As Boolean                ' Is the folder redy to be writed to registry

Dim ExTextbox1 As New ExtendedTextbox   'Thats my way soo only valid characters can be added to the textboxes
Dim ExTextbox2 As New ExtendedTextbox
Dim ExTextbox3 As New ExtendedTextbox
Dim ExTextbox4 As New ExtendedTextbox
Dim ExTextbox5 As New ExtendedTextbox
Dim ExTextbox6 As New ExtendedTextbox
Dim ExTextbox7 As New ExtendedTextbox
Dim ExTextbox8 As New ExtendedTextbox
Dim ExTextbox9 As New ExtendedTextbox
Dim ExTextbox10 As New ExtendedTextbox

Private Sub Command1_Click()       'End the program
Unload Me
End
End Sub

Private Sub Command10_Click()
frmDir.Show
frmDir.T2T
Me.Hide
End Sub

Private Sub Command11_Click()
frmDir.Show
frmDir.T3T
Me.Hide
End Sub

Private Sub Command12_Click()
frmDir.Show
frmDir.T4T
Me.Hide
End Sub

Private Sub Command13_Click()
frmDir.Show
frmDir.T5T
Me.Hide
End Sub

Private Sub Command14_Click()
frmDir.Show
frmDir.T6T
Me.Hide
End Sub

Private Sub Command15_Click()       'Auto detect folder or file

Command31.Caption = "Nothing"       'Reset all buttons first
Command32.Caption = "Nothing"
Command33.Caption = "Nothing"
Command34.Caption = "Nothing"
Command35.Caption = "Nothing"
Command36.Caption = "Nothing"
Command37.Caption = "Nothing"
Command38.Caption = "Nothing"
Command39.Caption = "Nothing"
Command40.Caption = "Nothing"

If Text6 <> "" Then Command31.Caption = "File"      'Set all buttons that contains text to file
If Text8 <> "" Then Command32.Caption = "File"
If Text10 <> "" Then Command33.Caption = "File"
If Text12 <> "" Then Command34.Caption = "File"
If Text14 <> "" Then Command35.Caption = "File"
If Text2 <> "" Then Command36.Caption = "File"
If Text15 <> "" Then Command37.Caption = "File"
If Text17 <> "" Then Command38.Caption = "File"
If Text19 <> "" Then Command39.Caption = "File"
If Text21 <> "" Then Command40.Caption = "File"

If Right$(Text6, "1") = "\" Then Command31.Caption = "Folder"       'Set all buttons that contains an "\" at the end to folder
If Right$(Text8, "1") = "\" Then Command32.Caption = "Folder"
If Right$(Text10, "1") = "\" Then Command33.Caption = "Folder"
If Right$(Text12, "1") = "\" Then Command34.Caption = "Folder"
If Right$(Text14, "1") = "\" Then Command35.Caption = "Folder"
If Right$(Text2, "1") = "\" Then Command36.Caption = "Folder"
If Right$(Text15, "1") = "\" Then Command37.Caption = "Folder"
If Right$(Text17, "1") = "\" Then Command38.Caption = "Folder"
If Right$(Text19, "1") = "\" Then Command39.Caption = "Folder"
If Right$(Text21, "1") = "\" Then Command40.Caption = "Folder"

checkff     'Call the sub that sets all buttons without text to "Nothing"

End Sub

Private Sub Command16_Click()       'Show about box
frmAbout.Show
Me.Hide
End Sub

Private Sub Command17_Click()               'Show the icons to the right
If Command17.Caption = ">>--->>>" Then
Me.Width = "10530"
Command17.Caption = "<<<---<<"
Else
Me.Width = "7965"
Command17.Caption = ">>--->>>"
End If

    Me.Top = (Screen.Height * 0.85) / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
End Sub

Private Sub Command18_Click()               'Set selected icon to nothing
If Shape1.Visible = False Then Exit Sub
text38.Text = ""
currentshape = "250"
Shape1.Visible = False
End Sub

Private Sub Command19_Click()               'Start the wizard
frmWiz.Show
Me.Hide
End Sub

Private Sub Command2_Click()        'Make the folder
On Error Resume Next

checkready                          'Check if everything is OK

If readymake = False Then Exit Sub  'If its not ready to make exit sub

If mode1.Value = True Then          'Check what mode that is used
GoTo classic
Else
Command22_Click                     'If not classic mode it will press the button for advanced mode
End If

Exit Sub
classic:
If Option1.Value = True Then delc1  'Delete the current folder. Becouse else it will dublicate the items in the folder
If Option2.Value = True Then delc2
If Option3.Value = True Then delc3

If Option1.Value = True Then m1     'Make the new folder
If Option2.Value = True Then m2
If Option3.Value = True Then m3

Command2.BackColor = &HC4C28F       'Make sure its the right color on the create button. (Becous after the wizard its red)
Exit Sub
End Sub

Private Sub Command20_Click()
frmFileIcon.Show
Me.Hide
End Sub

Private Sub Command22_Click()                'Create the folder in advanced mode
On Error Resume Next
FolderNumber = listfolder.ListIndex

If Val(FolderNumber) <= "9" Then FolderNumber = "0" & FolderNumber
If FolderNumber = "0" Then FolderNumber = "00"

FolderNumber = FolderNumber
makefolder

listfolder.clear
loadlist
Form_Load
End Sub

Private Sub Command23_Click()
FolderNumber = listfolder.ListIndex

If Val(FolderNumber) <= "9" Then FolderNumber = "0" & FolderNumber
If FolderNumber = "0" Then FolderNumber = "00"

deletefolder
listfolder.clear
loadlist

Form_Load
End Sub

Private Sub Command3_Click()            'Delete folder
Dim RetVals

RetVals = MsgBox("Sure you wanne delete this folder?", 292, "Are you sure?")
Select Case RetVals
     Case 6:
     GoTo dell
     Case 7:
     Exit Sub
End Select

dell:

If mode1.Value = True Then
GoTo classic
Else
Command23_Click
End If

Exit Sub
classic:
If Option1.Value = True Then del1
If Option2.Value = True Then del2
If Option3.Value = True Then del3

End Sub

Private Sub Command31_Click()
If Command31.Caption = "Folder" Then
Command31.Caption = "File"
Else
Command31.Caption = "Folder"
End If
End Sub

Private Sub Command32_Click()
If Command32.Caption = "Folder" Then
Command32.Caption = "File"
Else
Command32.Caption = "Folder"
End If
End Sub

Private Sub Command33_Click()
If Command33.Caption = "Folder" Then
Command33.Caption = "File"
Else
Command33.Caption = "Folder"
End If
End Sub

Private Sub Command34_Click()
If Command34.Caption = "Folder" Then
Command34.Caption = "File"
Else
Command34.Caption = "Folder"
End If
End Sub

Private Sub Command35_Click()
If Command35.Caption = "Folder" Then
Command35.Caption = "File"
Else
Command35.Caption = "Folder"
End If
End Sub

Private Sub Command36_Click()
If Command36.Caption = "Folder" Then
Command36.Caption = "File"
Else
Command36.Caption = "Folder"
End If
End Sub

Private Sub Command37_Click()
If Command37.Caption = "Folder" Then
Command37.Caption = "File"
Else
Command37.Caption = "Folder"
End If
End Sub

Private Sub Command38_Click()
If Command38.Caption = "Folder" Then
Command38.Caption = "File"
Else
Command38.Caption = "Folder"
End If
End Sub

Private Sub Command39_Click()
If Command39.Caption = "Folder" Then
Command39.Caption = "File"
Else
Command39.Caption = "Folder"
End If
End Sub

Private Sub Command4_Click()
frmDir.Show
frmDir.T11T
Me.Hide
End Sub

Private Sub Command40_Click()
If Command40.Caption = "Folder" Then
Command40.Caption = "File"
Else
Command40.Caption = "Folder"
End If
End Sub

Private Sub Command5_Click()
frmDir.Show
frmDir.T10T
Me.Hide
End Sub

Private Sub Command6_Click()
frmDir.Show
frmDir.T9T
Me.Hide
End Sub

Private Sub Command7_Click()
frmDir.Show
frmDir.T8T
Me.Hide
End Sub

Private Sub Command8_Click()
frmDir.Show
frmDir.T7T
Me.Hide
End Sub

Private Sub Command9_Click()
frmDir.Show
frmDir.T1T
Me.Hide
End Sub

Private Sub Form_Load()     'What will happend when program starts
On Error Resume Next
Dim TipStyle

keyallow                    'Load the keys that are allowed in the name box's
loadlist                    'Load the list in the list box. From 0 to 99

currentshape = "0"

Me.Caption = "System Folder Creator - Multima"      'Just to be sure :)
Me.Width = "7965"
listfolder.Top = "220"

TipStyle = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "ToolTipStyle", "")

If TipStyle = "0" Then Tool1(0).Value = True
If TipStyle = "1" Then Tool1(1).Value = True
If TipStyle = "2" Then Tool1(2).Value = True
If TipStyle = "3" Then Tool1(3).Value = True
If TipStyle = "4" Then Tool1(4).Value = True
If TipStyle = "5" Then Tool1(5).Value = True
If TipStyle = "6" Then Tool1(6).Value = True

Set TipStyle = Nothing

Text1 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "Name", "")
text37 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "Tip", "")
Text4 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "Click", "")
Text5 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P1", "")
Text6 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N1", "")
Text7 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P2", "")
Text8 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N2", "")
Text9 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P3", "")
Text10 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N3", "")
Text11 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P4", "")
Text12 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N4", "")
Text13 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P5", "")
Text14 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N5", "")
Text3 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P6", "")
Text2 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N6", "")
Text16 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P7", "")
Text15 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N7", "")
Text18 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P8", "")
Text17 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N8", "")
Text20 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P9", "")
Text19 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N9", "")
Text22 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P10", "")
Text21 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N10", "")

text38.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "Icon", "")

N1 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "Name", "")
N2 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "Name", "")
N3 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "Name", "")

Command31.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T1", "")
Command32.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T2", "")
Command33.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T3", "")
Command34.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T4", "")
Command35.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T5", "")
Command36.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T6", "")
Command37.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T7", "")
Command38.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T8", "")
Command39.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T9", "")
Command40.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T10", "")

checkff

If N1.Text = "" Then GoTo Ne1
Option1.Caption = N1
Ne1:
If N2.Text = "" Then GoTo Ne2
Option2.Caption = N2
Ne2:
If N3.Text = "" Then GoTo Ne3
Option3.Caption = N3
Ne3:
If N1 = "" Then Option1.Caption = "System folder 1"
If N2 = "" Then Option2.Caption = "System folder 2"
If N3 = "" Then Option3.Caption = "System folder 3"
Sys1.Caption = Option1.Caption

loadpic             'Load all icon in the resource file into the picture box's

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If Timer3.Enabled = True Then Exit Sub

If Text40.Text = "" Then
Label28.Caption = "Hold your mouse over a text and you will get help..."
Else
Label28.Caption = Text40.Text
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End
End Sub

Private Sub Image1_Click(Index As Integer)      'When you click an image
text38.Text = App.Path & "\" & App.EXEName & ".exe," & Index + 1    'The icon path that we will use
currentshape = Index
Shape1.Visible = True                                               'Make sure its visible
Shape1.Top = Image1(Index).Top                                      'Make its easier to see whats icon you have choosen
Shape1.Left = Image1(Index).Left
End Sub

Private Sub Image1_DblClick(Index As Integer)   'When user dbl click then select icon and minimize the form to normal again
Command17_Click
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Here can you enter the name of your system folder..."
End Sub
Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Place item name and type full path to the place you want it to go. Remember \ at the end of the path."
End Sub
Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Place item name and type full path to the place you want it to go. Remember \ at the end of the path."
End Sub
Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Place item name and type full path to the place you want it to go. Remember \ at the end of the path."
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Text45.Text = "" Then
Label28.Caption = "Place item name and type full path to the place you want it to go. Remember \ at the end of the path."
Else
Label28.Caption = Text45.Text
End If
End Sub
Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Place item name and type full path to the place you want it to go. Remember \ at the end of the path."
End Sub
Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Place item name and type full path to the place you want it to go. Remember \ at the end of the path."
End Sub
Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Place item name and type full path to the place you want it to go. Remember \ at the end of the path."
End Sub
Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "This will autogenerate a Tool Tip for you. If you like one of this.. You are free to use them!! :)"
End Sub
Private Sub Label29_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Pick the text that will bee shown when you hold your mouse over the icon."
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Hold your mouse over a text and you will get help..."
End Sub
Private Sub Label30_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Type the full path to an icon or choose one you like from the list. Get the list by clicking the button (->-->>-->>>). Else your folder will bee very ugly"
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Pick a place where the system folder will be, on the desktop or in My computer or both."
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Choose where to go when you dobble click the folder. Like a shortcut."
End Sub
Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "When you right click your folder a popup menu will be shown. Pick what that will be shown."
End Sub
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Place item name and type full path to the place you want it to go. Remember \ at the end of the path."
End Sub
Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Place item name and type full path to the place you want it to go. Remember \ at the end of the path."
End Sub
Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Place item name and type full path to the place you want it to go. Remember \ at the end of the path."
End Sub

Private Sub listfolder_Click()          'When you click the list
FolderNumber = listfolder.ListIndex
If Val(FolderNumber) <= "9" Then FolderNumber = "0" & FolderNumber
If FolderNumber = "0" Then FolderNumber = "00"
findfolder                              'Find the folder and write the output
Label28 = listfolder.Text
End Sub

Private Sub listfolder_DblClick()       'When you dbl click the list
FolderNumber = listfolder.ListIndex

If Val(FolderNumber) <= "9" Then FolderNumber = "0" & FolderNumber
If FolderNumber = "0" Then FolderNumber = "00"
End Sub

Private Sub mode1_Click()
mode1.Value = True
mode2.Value = False

listfolder.Visible = False

If Option1.Value = True Then Option1_Click
If Option2.Value = True Then Option2_Click
If Option3.Value = True Then Option3_Click

End Sub

Private Sub mode1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Press this button and you will choose the classic mode of how many folders you can have." & vbNewLine & "[3 folders]"
End Sub

Private Sub mode2_Click()
mode1.Value = False
mode2.Value = True

listfolder.Visible = True

listfolder.Selected(0) = True
listfolder_Click

End Sub

Private Sub mode2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "Press this button and you will choose the advanced mode of how many folders you can have." & vbNewLine & "[100 folders]"
End Sub

Private Sub Option1_Click()
Dim TipStyle
TipStyle = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "ToolTipStyle", "")

If TipStyle = "0" Then Tool1(0).Value = True
If TipStyle = "1" Then Tool1(1).Value = True
If TipStyle = "2" Then Tool1(2).Value = True
If TipStyle = "3" Then Tool1(3).Value = True
If TipStyle = "4" Then Tool1(4).Value = True
If TipStyle = "5" Then Tool1(5).Value = True
If TipStyle = "6" Then Tool1(6).Value = True

Set TipStyle = Nothing

Text1.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "Name", "")
text37.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "Tip", "")
Text4.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "Click", "")
Text5.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P1", "")
Text6.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N1", "")
Text7.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P2", "")
Text8.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N2", "")
Text9.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P3", "")
Text10.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N3", "")
Text11.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P4", "")
Text12.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N4", "")
Text13.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P5", "")
Text14.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N5", "")

Text3 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P6", "")
Text2 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N6", "")
Text16 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P7", "")
Text15 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N7", "")
Text18 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P8", "")
Text17 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N8", "")
Text20 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P9", "")
Text19 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N9", "")
Text22 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P10", "")
Text21 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N10", "")

text38.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "Icon", "")

Command31.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T1", "")
Command32.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T2", "")
Command33.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T3", "")
Command34.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T4", "")
Command35.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T5", "")
Command36.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T6", "")
Command37.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T7", "")
Command38.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T8", "")
Command39.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T9", "")
Command40.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T10", "")

checkshape
checkff

If Text48.Text = "" Then GoTo Nor
If Text48.Text <> "" Then GoTo Nor1
GoTo endd:

Nor1:
If Text1.Text = "" Then
    Option1.Caption = Text48 & " 1"
    Else
    Option1.Caption = Text1.Text
End If
Sys1.Caption = Option1.Caption
GoTo endd

Nor:
If Text1.Text = "" Then
    Option1.Caption = "System folder 1"
    Else
    Option1.Caption = Text1.Text
End If
Sys1.Caption = Option1.Caption
endd:
'Option2.Caption = "System folder 2"
'Option3.Caption = "System folder 3"
Command2.BackColor = &HC4C28F
End Sub

Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "" & Option1.Caption & ""
End Sub

Private Sub Option2_Click()
Dim TipStyle
TipStyle = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "ToolTipStyle", "")

If TipStyle = "0" Then Tool1(0).Value = True
If TipStyle = "1" Then Tool1(1).Value = True
If TipStyle = "2" Then Tool1(2).Value = True
If TipStyle = "3" Then Tool1(3).Value = True
If TipStyle = "4" Then Tool1(4).Value = True
If TipStyle = "5" Then Tool1(5).Value = True
If TipStyle = "6" Then Tool1(6).Value = True

Set TipStyle = Nothing


Text1.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "Name", "")
text37.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "Tip", "")
Text4.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "Click", "")
Text5.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P1", "")
Text6.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N1", "")
Text7.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P2", "")
Text8.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N2", "")
Text9.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P3", "")
Text10.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N3", "")
Text11.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P4", "")
Text12.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N4", "")
Text13.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P5", "")
Text14.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N5", "")

Text3 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P6", "")
Text2 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N6", "")
Text16 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P7", "")
Text15 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N7", "")
Text18 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P8", "")
Text17 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N8", "")
Text20 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P9", "")
Text19 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N9", "")
Text22 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P10", "")
Text21 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N10", "")

text38.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "Icon", "")

Command31.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T1", "")
Command32.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T2", "")
Command33.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T3", "")
Command34.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T4", "")
Command35.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T5", "")
Command36.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T6", "")
Command37.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T7", "")
Command38.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T8", "")
Command39.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T9", "")
Command40.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T10", "")

checkff
checkshape
'---------------------------
If Text48.Text = "" Then GoTo Nor
If Text48.Text <> "" Then GoTo Nor1
GoTo endd:
Nor1:
If Text1.Text = "" Then
    Option2.Caption = Text48 & " 2"
    Else
    Option2.Caption = Text1.Text
End If
Sys1.Caption = Option2.Caption
GoTo endd
Nor:
If Text1.Text = "" Then
    Option2.Caption = "System folder 2"
    Else
    Option2.Caption = Text1.Text
End If
Sys1.Caption = Option2.Caption
endd:

Command2.BackColor = &HC4C28F
End Sub

Private Sub Option2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "" & Option2.Caption & ""
End Sub

Private Sub Option3_Click()
Dim TipStyle
TipStyle = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "ToolTipStyle", "")

If TipStyle = "0" Then Tool1(0).Value = True
If TipStyle = "1" Then Tool1(1).Value = True
If TipStyle = "2" Then Tool1(2).Value = True
If TipStyle = "3" Then Tool1(3).Value = True
If TipStyle = "4" Then Tool1(4).Value = True
If TipStyle = "5" Then Tool1(5).Value = True
If TipStyle = "6" Then Tool1(6).Value = True

Set TipStyle = Nothing

Text1.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "Name", "")
text37.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "Tip", "")
Text4.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "Click", "")
Text5.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P1", "")
Text6.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N1", "")
Text7.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P2", "")
Text8.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N2", "")
Text9.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P3", "")
Text10.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N3", "")
Text11.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P4", "")
Text12.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N4", "")
Text13.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P5", "")
Text14.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N5", "")

Text3 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P6", "")
Text2 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N6", "")
Text16 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P7", "")
Text15 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N7", "")
Text18 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P8", "")
Text17 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N8", "")
Text20 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P9", "")
Text19 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N9", "")
Text22 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P10", "")
Text21 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N10", "")

text38.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "Icon", "")

Command31.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T1", "")
Command32.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T2", "")
Command33.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T3", "")
Command34.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T4", "")
Command35.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T5", "")
Command36.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T6", "")
Command37.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T7", "")
Command38.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T8", "")
Command39.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T9", "")
Command40.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T10", "")

checkff
checkshape

If Text1.Text = "" Then
If Text40.Text = "" Then
Label28.Caption = "Hold your mouse over a text and you will get help..."
Else
Label28.Caption = Text40.Text
End If
    Option3.Caption = "System folder 3"
    Else
    Option3.Caption = Text1.Text
End If
Sys1.Caption = Option3.Caption
'---------------------------
If Text48.Text = "" Then GoTo Nor
If Text48.Text <> "" Then GoTo Nor1
GoTo endd:
Nor1:
If Text1.Text = "" Then
    Option3.Caption = Text48 & " 3"
    Else
    Option3.Caption = Text1.Text
End If
Sys1.Caption = Option3.Caption
GoTo endd
Nor:
If Text1.Text = "" Then
    Option3.Caption = "System folder 3"
    Else
    Option3.Caption = Text1.Text
End If
Sys1.Caption = Option3.Caption
endd:

Command2.BackColor = &HC4C28F
End Sub

Private Sub m1()
Dim byTemp(3) As Byte
byTemp(0) = CByte(txtByte(0))
byTemp(1) = CByte(txtByte(1))
byTemp(2) = CByte(txtByte(2))
byTemp(3) = CByte(txtByte(3))

If Text1.Text = "" Or Text4.Text = "" Then GoTo err1



        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\", "", Text1.Text
        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\", "InfoTip", text37.Text
        
        SaveSettingByte HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\ShellFolder\", "Attributes", byTemp
        
        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\ShellEx\PropertySheetHandler\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\", "", ""
        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\DefaultIcon\", "", text38.Text
        
SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text1.Text & "\Command", "", "c:\windows\explorer /n /root," & Text4.Text & ""
        
If Command31.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text5.Text & "\Command", "", "c:\windows\explorer /n /root," & Text6.Text & ""
If Command32.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text7.Text & "\Command", "", "c:\windows\explorer /n /root," & Text8.Text & ""
If Command33.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text9.Text & "\Command", "", "c:\windows\explorer /n /root," & Text10.Text & ""
If Command34.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text11.Text & "\Command", "", "c:\windows\explorer /n /root," & Text12.Text & ""
If Command35.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text13.Text & "\Command", "", "c:\windows\explorer /n /root," & Text14.Text & ""
If Command36.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text3.Text & "\Command", "", "c:\windows\explorer /n /root," & Text2.Text & ""
If Command37.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text16.Text & "\Command", "", "c:\windows\explorer /n /root," & Text15.Text & ""
If Command38.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text18.Text & "\Command", "", "c:\windows\explorer /n /root," & Text17.Text & ""
If Command39.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text20.Text & "\Command", "", "c:\windows\explorer /n /root," & Text19.Text & ""
If Command40.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text22.Text & "\Command", "", "c:\windows\explorer /n /root," & Text21.Text & ""

If Command31.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text5.Text & "\Command", "", Text6.Text
If Command32.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text7.Text & "\Command", "", Text8.Text
If Command33.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text9.Text & "\Command", "", Text10.Text
If Command34.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text11.Text & "\Command", "", Text12.Text
If Command35.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text13.Text & "\Command", "", Text14.Text
If Command36.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text3.Text & "\Command", "", Text2.Text
If Command37.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text16.Text & "\Command", "", Text15.Text
If Command38.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text18.Text & "\Command", "", Text17.Text
If Command39.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text20.Text & "\Command", "", Text19.Text
If Command40.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\Shell\" & Text22.Text & "\Command", "", Text21.Text



If F1C1.Value = True Then 'desktop
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Desktop\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\", "" & Text1.Text & "", ""
GoTo here
End If
If F1C2.Value = True Then 'My computer
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\MyComputer\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\", "" & Text1.Text & "", ""
GoTo here
End If
If F1C3.Value = True Then 'Both
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Desktop\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\", "" & Text1.Text & "", ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\MyComputer\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\", "" & Text1.Text & "", ""
GoTo here
End If
here:

SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "Name", "" & Text1 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "Tip", "" & text37 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "Click", "" & Text4 & ""

SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P1", "" & Text5 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N1", "" & Text6 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P2", "" & Text7 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N2", "" & Text8 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P3", "" & Text9 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N3", "" & Text10 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P4", "" & Text11 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N4", "" & Text12 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P5", "" & Text13 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N5", "" & Text14 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "Icon", "" & text38 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P6", "" & Text3 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P7", "" & Text16 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P8", "" & Text18 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P9", "" & Text20 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "P10", "" & Text22 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N6", "" & Text2 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N7", "" & Text15 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N8", "" & Text17 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N9", "" & Text19 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "N10", "" & Text21 & ""

If Command31.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T1", "File"
If Command32.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T2", "File"
If Command33.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T3", "File"
If Command34.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T4", "File"
If Command35.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T5", "File"
If Command36.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T6", "File"
If Command37.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T7", "File"
If Command38.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T8", "File"
If Command39.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T9", "File"
If Command40.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T10", "File"

If Command31.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T1", "Folder"
If Command32.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T2", "Folder"
If Command33.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T3", "Folder"
If Command34.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T4", "Folder"
If Command35.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T5", "Folder"
If Command36.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T6", "Folder"
If Command37.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T7", "Folder"
If Command38.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T8", "Folder"
If Command39.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T9", "Folder"
If Command40.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "T10", "Folder"

If Tool1(0).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "ToolTipStyle", "0"
If Tool1(1).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "ToolTipStyle", "1"
If Tool1(2).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "ToolTipStyle", "2"
If Tool1(3).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "ToolTipStyle", "3"
If Tool1(4).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "ToolTipStyle", "4"
If Tool1(5).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "ToolTipStyle", "5"
If Tool1(6).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "ToolTipStyle", "6"

SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "ShapePosition", "" & currentshape & ""

If Text48.Text = "" Then
Label28.Caption = "Folder " & Text1.Text & " has been createt..."
Else
Label28.Caption = Text49.Text & " " & Text1.Text & " " & Text50.Text
End If
Timer3.Enabled = True

GoTo endd
err1:
If Text51.Text = "" Then
Label28.Caption = "You must fill in the name and where to go when you dobble click."
Else
Label28.Caption = Text51.Text
End If
endd:
End Sub
Private Sub m2()
Dim byTemp(3) As Byte
byTemp(0) = CByte(txtByte(0))
byTemp(1) = CByte(txtByte(1))
byTemp(2) = CByte(txtByte(2))
byTemp(3) = CByte(txtByte(3))

If Text1.Text = "" Or Text4.Text = "" Then GoTo err1

        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\", "", Text1.Text
        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\", "InfoTip", text37.Text
        
        SaveSettingByte HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\ShellFolder\", "Attributes", byTemp
        
        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\ShellEx\PropertySheetHandler\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\", "", ""
        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\DefaultIcon\", "", text38.Text
        
        
If Command31.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text5.Text & "\Command", "", "c:\windows\explorer /n /root," & Text6.Text & ""
If Command32.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text7.Text & "\Command", "", "c:\windows\explorer /n /root," & Text8.Text & ""
If Command33.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text9.Text & "\Command", "", "c:\windows\explorer /n /root," & Text10.Text & ""
If Command34.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text11.Text & "\Command", "", "c:\windows\explorer /n /root," & Text12.Text & ""
If Command35.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text13.Text & "\Command", "", "c:\windows\explorer /n /root," & Text14.Text & ""
If Command36.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text3.Text & "\Command", "", "c:\windows\explorer /n /root," & Text2.Text & ""
If Command37.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text16.Text & "\Command", "", "c:\windows\explorer /n /root," & Text15.Text & ""
If Command38.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text18.Text & "\Command", "", "c:\windows\explorer /n /root," & Text17.Text & ""
If Command39.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text20.Text & "\Command", "", "c:\windows\explorer /n /root," & Text19.Text & ""
If Command40.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text22.Text & "\Command", "", "c:\windows\explorer /n /root," & Text21.Text & ""

If Command31.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text5.Text & "\Command", "", Text6.Text
If Command32.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text7.Text & "\Command", "", Text8.Text
If Command33.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text9.Text & "\Command", "", Text10.Text
If Command34.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text11.Text & "\Command", "", Text12.Text
If Command35.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text13.Text & "\Command", "", Text14.Text
If Command36.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text3.Text & "\Command", "", Text2.Text
If Command37.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text16.Text & "\Command", "", Text15.Text
If Command38.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text18.Text & "\Command", "", Text17.Text
If Command39.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text20.Text & "\Command", "", Text19.Text
If Command40.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\Shell\" & Text22.Text & "\Command", "", Text21.Text

If F1C1.Value = True Then 'desktop
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Desktop\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\", "" & Text1.Text & "", ""
GoTo here
End If
If F1C2.Value = True Then 'My computer
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\MyComputer\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\", "" & Text1.Text & "", ""
GoTo here
End If
If F1C3.Value = True Then 'Both
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Desktop\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\", "" & Text1.Text & "", ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\MyComputer\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\", "" & Text1.Text & "", ""
GoTo here
End If
here:
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "Name", "" & Text1 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "Tip", "" & text37 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "Click", "" & Text4 & ""

SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P1", "" & Text5 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N1", "" & Text6 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P2", "" & Text7 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N2", "" & Text8 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P3", "" & Text9 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N3", "" & Text10 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P4", "" & Text11 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N4", "" & Text12 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P5", "" & Text13 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N5", "" & Text14 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "Icon", "" & text38 & ""

SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P6", "" & Text3 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P7", "" & Text16 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P8", "" & Text18 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P9", "" & Text20 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "P10", "" & Text22 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N6", "" & Text2 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N7", "" & Text15 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N8", "" & Text17 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N9", "" & Text19 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "N10", "" & Text21 & ""

If Command31.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T1", "File"
If Command32.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T2", "File"
If Command33.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T3", "File"
If Command34.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T4", "File"
If Command35.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T5", "File"
If Command36.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T6", "File"
If Command37.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T7", "File"
If Command38.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T8", "File"
If Command39.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T9", "File"
If Command40.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T10", "File"

If Command31.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T1", "Folder"
If Command32.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T2", "Folder"
If Command33.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T3", "Folder"
If Command34.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T4", "Folder"
If Command35.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T5", "Folder"
If Command36.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T6", "Folder"
If Command37.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T7", "Folder"
If Command38.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T8", "Folder"
If Command39.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T9", "Folder"
If Command40.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "T10", "Folder"

If Tool1(0).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "ToolTipStyle", "0"
If Tool1(1).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "ToolTipStyle", "1"
If Tool1(2).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "ToolTipStyle", "2"
If Tool1(3).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "ToolTipStyle", "3"
If Tool1(4).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "ToolTipStyle", "4"
If Tool1(5).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "ToolTipStyle", "5"
If Tool1(6).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "ToolTipStyle", "6"

SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "ShapePosition", "" & currentshape & ""

If Text48.Text = "" Then
Label28.Caption = "Folder " & Text1.Text & " has been createt..."
Else
Label28.Caption = Text49.Text & " " & Text1.Text & " " & Text50.Text
End If
Timer3.Enabled = True

GoTo endd
err1:
If Text51.Text = "" Then
Label28.Caption = "You must fill in the name and where to go when you dobble click."
Else
Label28.Caption = Text51.Text
End If
endd:
End Sub
Private Sub m3()
Dim byTemp(3) As Byte
byTemp(0) = CByte(txtByte(0))
byTemp(1) = CByte(txtByte(1))
byTemp(2) = CByte(txtByte(2))
byTemp(3) = CByte(txtByte(3))

If Text1.Text = "" Or Text4.Text = "" Then GoTo err1

        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\", "", Text1.Text
        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\", "InfoTip", text37.Text
        
        SaveSettingByte HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\ShellFolder\", "Attributes", byTemp
        
        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\ShellEx\PropertySheetHandler\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\", "", ""
        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\DefaultIcon\", "", text38.Text
        
        
If Command31.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text5.Text & "\Command", "", "c:\windows\explorer /n /root," & Text6.Text & ""
If Command32.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text7.Text & "\Command", "", "c:\windows\explorer /n /root," & Text8.Text & ""
If Command33.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text9.Text & "\Command", "", "c:\windows\explorer /n /root," & Text10.Text & ""
If Command34.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text11.Text & "\Command", "", "c:\windows\explorer /n /root," & Text12.Text & ""
If Command35.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text13.Text & "\Command", "", "c:\windows\explorer /n /root," & Text14.Text & ""
If Command36.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text3.Text & "\Command", "", "c:\windows\explorer /n /root," & Text2.Text & ""
If Command37.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text16.Text & "\Command", "", "c:\windows\explorer /n /root," & Text15.Text & ""
If Command38.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text18.Text & "\Command", "", "c:\windows\explorer /n /root," & Text17.Text & ""
If Command39.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text20.Text & "\Command", "", "c:\windows\explorer /n /root," & Text19.Text & ""
If Command40.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text22.Text & "\Command", "", "c:\windows\explorer /n /root," & Text21.Text & ""

If Command31.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text5.Text & "\Command", "", Text6.Text
If Command32.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text7.Text & "\Command", "", Text8.Text
If Command33.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text9.Text & "\Command", "", Text10.Text
If Command34.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text11.Text & "\Command", "", Text12.Text
If Command35.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text13.Text & "\Command", "", Text14.Text
If Command36.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text3.Text & "\Command", "", Text2.Text
If Command37.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text16.Text & "\Command", "", Text15.Text
If Command38.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text18.Text & "\Command", "", Text17.Text
If Command39.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text20.Text & "\Command", "", Text19.Text
If Command40.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\Shell\" & Text22.Text & "\Command", "", Text21.Text


If F1C1.Value = True Then 'desktop
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Desktop\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\", "" & Text1.Text & "", ""
GoTo here
End If
If F1C2.Value = True Then 'My computer
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\MyComputer\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\", "" & Text1.Text & "", ""
GoTo here
End If
If F1C3.Value = True Then 'Both
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Desktop\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\", "" & Text1.Text & "", ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\MyComputer\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\", "" & Text1.Text & "", ""
GoTo here
End If
here:
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "Name", "" & Text1 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "Tip", "" & text37 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "Click", "" & Text4 & ""

SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P1", "" & Text5 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N1", "" & Text6 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P2", "" & Text7 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N2", "" & Text8 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P3", "" & Text9 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N3", "" & Text10 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P4", "" & Text11 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N4", "" & Text12 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P5", "" & Text13 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N5", "" & Text14 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "Icon", "" & text38 & ""

SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P6", "" & Text3 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P7", "" & Text16 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P8", "" & Text18 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P9", "" & Text20 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "P10", "" & Text22 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N6", "" & Text2 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N7", "" & Text15 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N8", "" & Text17 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N9", "" & Text19 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "N10", "" & Text21 & ""

If Command31.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T1", "File"
If Command32.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T2", "File"
If Command33.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T3", "File"
If Command34.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T4", "File"
If Command35.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T5", "File"
If Command36.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T6", "File"
If Command37.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T7", "File"
If Command38.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T8", "File"
If Command39.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T9", "File"
If Command40.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T10", "File"

If Command31.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T1", "Folder"
If Command32.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T2", "Folder"
If Command33.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T3", "Folder"
If Command34.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T4", "Folder"
If Command35.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T5", "Folder"
If Command36.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T6", "Folder"
If Command37.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T7", "Folder"
If Command38.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T8", "Folder"
If Command39.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T9", "Folder"
If Command40.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "T10", "Folder"

If Tool1(0).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "ToolTipStyle", "0"
If Tool1(1).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "ToolTipStyle", "1"
If Tool1(2).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "ToolTipStyle", "2"
If Tool1(3).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "ToolTipStyle", "3"
If Tool1(4).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "ToolTipStyle", "4"
If Tool1(5).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "ToolTipStyle", "5"
If Tool1(6).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "ToolTipStyle", "6"

SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "ShapePosition", "" & currentshape & ""

If Text48.Text = "" Then
Label28.Caption = "Folder " & Text1.Text & " has been createt..."
Else
Label28.Caption = Text49.Text & " " & Text1.Text & " " & Text50.Text
End If
Timer3.Enabled = True

GoTo endd
err1:
If Text51.Text = "" Then
Label28.Caption = "You must fill in the name and where to go when you dobble click."
Else
Label28.Caption = Text51.Text
End If
endd:
End Sub

Public Sub del1()
On Error Resume Next
If Text49.Text = "" Or Text52.Text = "" Then
Label28.Caption = "Folder " & Text1.Text & " has been deleted..."
Else
Label28.Caption = Text49.Text & " " & Text52.Text
End If
Timer3.Enabled = True

DeleteKey HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Desktop\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\MyComputer\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\"
If Text48.Text = "" Then
Option1.Caption = "System folder 1"
Sys1.Caption = "System folder 1"
Else
Option1.Caption = Text48.Text & " 1"
Sys1.Caption = Text48.Text & " 1"
End If
clear
End Sub
Public Sub del2()

If Text49.Text = "" Or Text52.Text = "" Then
Label28.Caption = "Folder " & Text1.Text & " has been deleted..."
Else
Label28.Caption = Text49.Text & " " & Text52.Text
End If
Timer3.Enabled = True

DeleteKey HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Desktop\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\MyComputer\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\"
If Text48.Text = "" Then
Option2.Caption = "System folder 2"
Sys1.Caption = "System folder 2"
Else
Option2.Caption = Text48.Text & " 2"
Sys1.Caption = Text48.Text & " 2"
End If
clear
End Sub
Public Sub del3()
If Text49.Text = "" Or Text52.Text = "" Then
Label28.Caption = "Folder " & Text1.Text & " has been deleted..."
Else
Label28.Caption = Text49.Text & " " & Text52.Text
End If


DeleteKey HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Desktop\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\MyComputer\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\"
If Text48.Text = "" Then
Option3.Caption = "System folder 3"
Sys1.Caption = "System folder 3"
Else
Option3.Caption = Text48.Text & " 3"
Sys1.Caption = Text48.Text & " 3"
End If
clear
End Sub
Public Sub clear()
Text1.Text = ""
text37.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text3.Text = ""
Text2.Text = ""
Text16.Text = ""
Text15.Text = ""
Text18.Text = ""
Text17.Text = ""
Text20.Text = ""
Text19.Text = ""
Text22.Text = ""
Text21.Text = ""

Command31.Caption = "Nothing"
Command32.Caption = "Nothing"
Command33.Caption = "Nothing"
Command34.Caption = "Nothing"
Command35.Caption = "Nothing"
Command36.Caption = "Nothing"
Command37.Caption = "Nothing"
Command38.Caption = "Nothing"
Command39.Caption = "Nothing"
Command40.Caption = "Nothing"

Tool1(0).Value = True
Tool1(1).Value = False
Tool1(2).Value = False
Tool1(3).Value = False
Tool1(4).Value = False
Tool1(5).Value = False
Tool1(6).Value = False

F1C1.Value = True
F1C2.Value = False
F1C3.Value = False

End Sub
Private Sub Option3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label28.Caption = "" & Option3.Caption & ""
End Sub

Private Sub Sys1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Text40.Text = "" Then
Label28.Caption = "Hold your mouse over a text and you will get help..."
Else
Label28.Caption = Text40.Text
End If
End Sub
Public Sub delc1()          'Delete folder 1 in classic mode
On Error Resume Next
DeleteKey HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Desktop\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\MyComputer\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA1}\"
End Sub
Public Sub delc2()          'Delete folder 2 in classic mode
On Error Resume Next
DeleteKey HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Desktop\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\MyComputer\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA2}\"
End Sub
Public Sub delc3()          'Delete folder 3 in classic mode
On Error Resume Next
DeleteKey HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Desktop\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\MyComputer\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64EA3}\"
End Sub

Private Sub checkff()
If Command31.Caption = "" Then Command31.Caption = "Nothing"
If Command32.Caption = "" Then Command32.Caption = "Nothing"
If Command33.Caption = "" Then Command33.Caption = "Nothing"
If Command34.Caption = "" Then Command34.Caption = "Nothing"
If Command35.Caption = "" Then Command35.Caption = "Nothing"
If Command36.Caption = "" Then Command36.Caption = "Nothing"
If Command37.Caption = "" Then Command37.Caption = "Nothing"
If Command38.Caption = "" Then Command38.Caption = "Nothing"
If Command39.Caption = "" Then Command39.Caption = "Nothing"
If Command40.Caption = "" Then Command40.Caption = "Nothing"

End Sub

Private Sub Text1_Change()
'(¯`'·.¸  NAME ¸.·'´¯)
'¯`·.¸¸.->NAME<-.¸¸.·´¯
'"¨¨°º©o.,NAME,.o©º°¨¨"
'°¤º°º¤.,¸¸,.¤º°º¤°¤º°º¤.,¸¸,.¤º°º¤¤º°º¤.,¸¸,.¤º°º¤° NAME °¤º°º¤.,¸¸,.¤º°º¤°¤º°º¤.,¸¸,.¤º°º¤¤º°º¤.,¸¸,.¤º°º¤°
'(¯`'·.¸(¯`'·.¸(¯`'·.¸  NAME ¸.·'´¯)¸.·'´¯)¸.·'´¯)
'>===========>>===========>>>===========>>>>  NAME  <<<<=============<<<===========<<===========<


If Tool1(0).Value = True Then Exit Sub

If Tool1(1).Value = True Then text37 = "(¯`'·.¸ " & Text1 & " ¸.·'´¯)"
If Tool1(2).Value = True Then text37 = "¯`·.¸¸.-> " & Text1 & " <-.¸¸.·´¯"
If Tool1(3).Value = True Then text37 = "¨¨°º©o., " & Text1 & " ,.o©º°¨¨"
If Tool1(4).Value = True Then text37 = "°º¤¤º°º¤.,¸¸,.¤º°º¤° " & Text1 & " °¤º°º¤.,¸¸,.¤º°º¤°¤º°"
If Tool1(5).Value = True Then text37 = "(¯`'·.¸(¯`'·.¸(¯`'·.¸" & Text1 & "¸.·'´¯)¸.·'´¯)¸.·'´¯)"
If Tool1(6).Value = True Then text37 = ">===>>===>>>===>>>>   " & Text1 & "   <<<<===<<<===<<===<"

If Option1.Value = True Then Option1.Caption = Text1.Text
If Option2.Value = True Then Option2.Caption = Text1.Text
If Option3.Value = True Then Option3.Caption = Text1.Text
End Sub

Private Sub Text38_Change()
If finishload = False Then
finishload = True
Exit Sub
End If

currentshape = "250"
Shape1.Visible = False
End Sub

Private Sub Timer3_Timer()
If Text40.Text = "" Then
Label28.Caption = "Hold your mouse over a text and you will get help..."
Else
Label28.Caption = Text40.Text
End If
Timer3.Enabled = False
End Sub

Private Sub loadpic()                       'Load icon from resource
On Error Resume Next                        'If an error accor
Dim i As Integer                            'Used for the image index
i = "0"                                     'Set i to "0"
Do Until i = 48                             'There is 48 icon to load therefor do this 48 times
Image1(i) = LoadResPicture(i + 100 + 1, 1)  'Becouse of the number the icon is saved as in the resource
If Val(Image1(i).Picture) = 0 Then Image1(i).Visible = False    'If there is no icon then it will set the image.visible to false
i = i + 1                                   'Add 1 to i
Loop                                        'Offcorce
End Sub                                     'Daahhh

Public Sub checkshape()
On Error Resume Next
Dim shapepos

If Option1.Value = True Then shapepos = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "ShapePosition", "")

If shapepos = "250" Then Command18_Click: Exit Sub
If shapepos = "" Then Command18_Click: Exit Sub

If Text1.Text = "" Then Exit Sub

Shape1.Visible = True
Shape1.Top = Image1(shapepos).Top
Shape1.Left = Image1(shapepos).Left

finishload = False

Set shapepos = Nothing
End Sub

Private Sub checkready()            'Check that everything is ready to make the folder
'\/ Check that folder have a name
If Text1 = "" Then MsgBox "You must enter a name of your folder...": readymake = False: Exit Sub
'-----------------------------------------------

'\/ Check that at least one item is someting
If Text5 = "" And _
Text7 = "" And _
Text9 = "" And _
Text11 = "" And _
Text13 = "" And _
Text3 = "" And _
Text16 = "" And _
Text18 = "" And _
Text20 = "" And _
Text22 = "" Then _
MsgBox "You must make at least one item in your folder...": readymake = False: Exit Sub
'-----------------------------------------------

'\/ Check that if an item have a name it also have a path
If Text5 <> "" And Text6 = "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text7 <> "" And Text8 = "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text9 <> "" And Text10 = "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text11 <> "" And Text12 = "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text13 <> "" And Text14 = "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text3 <> "" And Text2 = "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text16 <> "" And Text15 = "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text18 <> "" And Text17 = "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text20 <> "" And Text19 = "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text22 <> "" And Text21 = "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
'-----------------------------------------------

'\/ Check that if an item have a path it also have a name
If Text5 = "" And Text6 <> "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text7 = "" And Text8 <> "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text9 = "" And Text10 <> "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text11 = "" And Text12 <> "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text13 = "" And Text14 <> "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text3 = "" And Text2 <> "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text16 = "" And Text15 <> "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text18 = "" And Text17 <> "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text20 = "" And Text19 <> "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
If Text22 = "" And Text21 <> "" Then MsgBox "You must have a name and a path in all your item's...": readymake = False: Exit Sub
'-----------------------------------------------

'\/ Check that folder name is different from all item
If Text1 = Text5 Then MsgBox "You must have different name on your folder name and your item.": readymake = False: Exit Sub
If Text1 = Text7 Then MsgBox "You must have different name on your folder name and your item.": readymake = False: Exit Sub
If Text1 = Text9 Then MsgBox "You must have different name on your folder name and your item.": readymake = False: Exit Sub
If Text1 = Text11 Then MsgBox "You must have different name on your folder name and your item.": readymake = False: Exit Sub
If Text1 = Text13 Then MsgBox "You must have different name on your folder name and your item.": readymake = False: Exit Sub
If Text1 = Text3 Then MsgBox "You must have different name on your folder name and your item.": readymake = False: Exit Sub
If Text1 = Text16 Then MsgBox "You must have different name on your folder name and your item.": readymake = False: Exit Sub
If Text1 = Text18 Then MsgBox "You must have different name on your folder name and your item.": readymake = False: Exit Sub
If Text1 = Text20 Then MsgBox "You must have different name on your folder name and your item.": readymake = False: Exit Sub
If Text1 = Text22 Then MsgBox "You must have different name on your folder name and your item.": readymake = False: Exit Sub
'-----------------------------------------------

'\/ Check that all item name is different from all other item name

If Text5 = "" Then GoTo te7
If Text5 = Text7 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text5 = Text9 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text5 = Text11 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text5 = Text13 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text5 = Text3 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text5 = Text16 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text5 = Text18 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text5 = Text20 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text5 = Text22 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub

te7:
If Text7 = "" Then GoTo te9
If Text7 = Text9 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text7 = Text11 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text7 = Text13 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text7 = Text3 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text7 = Text16 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text7 = Text18 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text7 = Text20 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text7 = Text22 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub

te9:
If Text9 = "" Then GoTo te11
If Text9 = Text11 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text9 = Text13 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text9 = Text3 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text9 = Text16 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text9 = Text18 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text9 = Text20 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text9 = Text22 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub

te11:
If Text11 = "" Then GoTo te13
If Text11 = Text13 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text11 = Text3 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text11 = Text16 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text11 = Text18 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text11 = Text20 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text11 = Text22 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub

te13:
If Text13 = "" Then GoTo te3
If Text13 = Text3 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text13 = Text16 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text13 = Text18 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text13 = Text20 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text13 = Text22 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub

te3:
If Text3 = "" Then GoTo te16
If Text3 = Text16 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text3 = Text18 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text3 = Text20 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text3 = Text22 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub

te16:
If Text16 = "" Then GoTo te18
If Text16 = Text18 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text16 = Text20 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text16 = Text22 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub

te18:
If Text18 = "" Then GoTo te20
If Text18 = Text20 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub
If Text18 = Text22 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub

te20:
If Text20 = "" Then GoTo tend
If Text20 = Text22 Then MsgBox "You must have different names on all you item.": readymake = False: Exit Sub

tend:
'-----------------------------------------------

readymake = True
End Sub
Sub loadlist()
On Error Resume Next
Dim tempfolder As String
FolderNumber = "0"

Do Until FolderNumber = "99"

If Val(FolderNumber) <= "9" Then FolderNumber = "0" & FolderNumber
If FolderNumber = "0" Then FolderNumber = "00"

tempfolder = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "Name", "")

If tempfolder <> "" Then
listfolder.AddItem tempfolder
Else
listfolder.AddItem "System folder " & FolderNumber
End If
FolderNumber = FolderNumber + 1
Loop

FolderNumber = "99"
tempfolder = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "Name", "")
If tempfolder <> "" Then
listfolder.AddItem tempfolder
Else
listfolder.AddItem "System folder " & FolderNumber
End If

FolderNumber = ""
End Sub

Sub findfolder()            'Find folder and but the output in the textboxes (Advanced mode)
Dim TipStyle
TipStyle = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "ToolTipStyle", "")

If TipStyle = "0" Then Tool1(0).Value = True
If TipStyle = "1" Then Tool1(1).Value = True
If TipStyle = "2" Then Tool1(2).Value = True
If TipStyle = "3" Then Tool1(3).Value = True
If TipStyle = "4" Then Tool1(4).Value = True
If TipStyle = "5" Then Tool1(5).Value = True
If TipStyle = "6" Then Tool1(6).Value = True

Set TipStyle = Nothing

Text1.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "Name", "")
text37.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "Tip", "")
Text4.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "Click", "")
Text5.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P1", "")
Text6.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N1", "")
Text7.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P2", "")
Text8.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N2", "")
Text9.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P3", "")
Text10.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N3", "")
Text11.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P4", "")
Text12.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N4", "")
Text13.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P5", "")
Text14.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N5", "")

Text3 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P6", "")
Text2 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N6", "")
Text16 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P7", "")
Text15 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N7", "")
Text18 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P8", "")
Text17 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N8", "")
Text20 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P9", "")
Text19 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N9", "")
Text22 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P10", "")
Text21 = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N10", "")

text38.Text = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "Icon", "")

Command31.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T1", "")
Command32.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T2", "")
Command33.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T3", "")
Command34.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T4", "")
Command35.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T5", "")
Command36.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T6", "")
Command37.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T7", "")
Command38.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T8", "")
Command39.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T9", "")
Command40.Caption = GetSettingString(HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T10", "")

checkshape
checkff

If Text48.Text = "" Then GoTo Nor
If Text48.Text <> "" Then GoTo Nor1
GoTo Nor:

Nor1:
If Text1.Text = "" Then
    Option1.Caption = Text48 & " 1"
    Else
    Option1.Caption = Text1.Text
End If
Sys1.Caption = Option1.Caption
GoTo Nor

Nor:
Command2.BackColor = &HC4C28F

End Sub
Private Sub makefolder()        'Make folder (Advanced mode)
'FolderNumber = 00 to 99

If Val(FolderNumber) <= "9" Then FolderNumber = "0" & FolderNumber
If FolderNumber = "0" Then FolderNumber = "00"

Dim byTemp(3) As Byte
byTemp(0) = CByte(txtByte(0))
byTemp(1) = CByte(txtByte(1))
byTemp(2) = CByte(txtByte(2))
byTemp(3) = CByte(txtByte(3))

If Text1.Text = "" Or Text4.Text = "" Then GoTo err1

        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\", "", Text1.Text
        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\", "InfoTip", text37.Text
        
        SaveSettingByte HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\ShellFolder\", "Attributes", byTemp
        
        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\ShellEx\PropertySheetHandler\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\", "", ""
        SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\DefaultIcon\", "", text38.Text
        
SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text1.Text & "\Command", "", "c:\windows\explorer /n /root," & Text4.Text & ""
        
If Command31.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text5.Text & "\Command", "", "c:\windows\explorer /n /root," & Text6.Text & ""
If Command32.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text7.Text & "\Command", "", "c:\windows\explorer /n /root," & Text8.Text & ""
If Command33.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text9.Text & "\Command", "", "c:\windows\explorer /n /root," & Text10.Text & ""
If Command34.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text11.Text & "\Command", "", "c:\windows\explorer /n /root," & Text12.Text & ""
If Command35.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text13.Text & "\Command", "", "c:\windows\explorer /n /root," & Text14.Text & ""
If Command36.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text3.Text & "\Command", "", "c:\windows\explorer /n /root," & Text2.Text & ""
If Command37.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text16.Text & "\Command", "", "c:\windows\explorer /n /root," & Text15.Text & ""
If Command38.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text18.Text & "\Command", "", "c:\windows\explorer /n /root," & Text17.Text & ""
If Command39.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text20.Text & "\Command", "", "c:\windows\explorer /n /root," & Text19.Text & ""
If Command40.Caption = "Folder" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text22.Text & "\Command", "", "c:\windows\explorer /n /root," & Text21.Text & ""

If Command31.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text5.Text & "\Command", "", Text6.Text
If Command32.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text7.Text & "\Command", "", Text8.Text
If Command33.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text9.Text & "\Command", "", Text10.Text
If Command34.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text11.Text & "\Command", "", Text12.Text
If Command35.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text13.Text & "\Command", "", Text14.Text
If Command36.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text3.Text & "\Command", "", Text2.Text
If Command37.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text16.Text & "\Command", "", Text15.Text
If Command38.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text18.Text & "\Command", "", Text17.Text
If Command39.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text20.Text & "\Command", "", Text19.Text
If Command40.Caption = "File" Then SaveSettingString HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\Shell\" & Text22.Text & "\Command", "", Text21.Text



If F1C1.Value = True Then 'desktop
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Desktop\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\", "" & Text1.Text & "", ""
GoTo here
End If
If F1C2.Value = True Then 'My computer
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\MyComputer\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\", "" & Text1.Text & "", ""
GoTo here
End If
If F1C3.Value = True Then 'Both
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Desktop\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\", "" & Text1.Text & "", ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\MyComputer\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\", "" & Text1.Text & "", ""
GoTo here
End If
here:

SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "Name", "" & Text1 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "Tip", "" & text37 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "Click", "" & Text4 & ""

SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P1", "" & Text5 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N1", "" & Text6 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P2", "" & Text7 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N2", "" & Text8 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P3", "" & Text9 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N3", "" & Text10 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P4", "" & Text11 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N4", "" & Text12 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P5", "" & Text13 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N5", "" & Text14 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "Icon", "" & text38 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P6", "" & Text3 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P7", "" & Text16 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P8", "" & Text18 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P9", "" & Text20 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "P10", "" & Text22 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N6", "" & Text2 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N7", "" & Text15 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N8", "" & Text17 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N9", "" & Text19 & ""
SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "N10", "" & Text21 & ""

If Command31.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T1", "File"
If Command32.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T2", "File"
If Command33.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T3", "File"
If Command34.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T4", "File"
If Command35.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T5", "File"
If Command36.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T6", "File"
If Command37.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T7", "File"
If Command38.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T8", "File"
If Command39.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T9", "File"
If Command40.Caption = "File" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T10", "File"

If Command31.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T1", "Folder"
If Command32.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T2", "Folder"
If Command33.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T3", "Folder"
If Command34.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T4", "Folder"
If Command35.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T5", "Folder"
If Command36.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T6", "Folder"
If Command37.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T7", "Folder"
If Command38.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T8", "Folder"
If Command39.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T9", "Folder"
If Command40.Caption = "Folder" Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "T10", "Folder"

If Tool1(0).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "ToolTipStyle", "0"
If Tool1(1).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "ToolTipStyle", "1"
If Tool1(2).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "ToolTipStyle", "2"
If Tool1(3).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "ToolTipStyle", "3"
If Tool1(4).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "ToolTipStyle", "4"
If Tool1(5).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "ToolTipStyle", "5"
If Tool1(6).Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "ToolTipStyle", "6"

SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\", "ShapePosition", "" & currentshape & ""

If Text48.Text = "" Then
Label28.Caption = "Folder " & Text1.Text & " has been createt..."
Else
Label28.Caption = Text49.Text & " " & Text1.Text & " " & Text50.Text
End If
Timer3.Enabled = True

GoTo endd
err1:
If Text51.Text = "" Then
Label28.Caption = "You must fill in the name and where to go when you dobble click."
Else
Label28.Caption = Text51.Text
End If
endd:
End Sub
Public Sub deletefolder()       'Delete folder (Advanced mode)
On Error Resume Next
If Text49.Text = "" Or Text52.Text = "" Then
Label28.Caption = "Folder " & Text1.Text & " has been deleted..."
Else
Label28.Caption = Text49.Text & " " & Text52.Text
End If
Timer3.Enabled = True

DeleteKey HKEY_CLASSES_ROOT, "CLSID\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\long\" & FolderNumber & "\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\Desktop\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\"
DeleteKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\MyComputer\NameSpace\{AA0C4D12-ABAB-11D2-2F98-00D04FB64E" & FolderNumber & "}\"
If Text48.Text = "" Then
Sys1.Caption = "System folder " & FolderNumber
End If
clear
End Sub

Private Sub Tool1_Click(Index As Integer)

If Tool1(0).Value = True Then text37 = Text1
If Tool1(1).Value = True Then text37 = "(¯`'·.¸ " & Text1 & " ¸.·'´¯)"
If Tool1(2).Value = True Then text37 = "¯`·.¸¸.-> " & Text1 & " <-.¸¸.·´¯"
If Tool1(3).Value = True Then text37 = "¨¨°º©o., " & Text1 & " ,.o©º°¨¨"
If Tool1(4).Value = True Then text37 = "°º¤¤º°º¤.,¸¸,.¤º°º¤° " & Text1 & " °¤º°º¤.,¸¸,.¤º°º¤°¤º°"
If Tool1(5).Value = True Then text37 = "(¯`'·.¸(¯`'·.¸(¯`'·.¸" & Text1 & "¸.·'´¯)¸.·'´¯)¸.·'´¯)"
If Tool1(6).Value = True Then text37 = ">===>>===>>>===>>>>   " & Text1 & "   <<<<===<<<===<<===<"

End Sub

Private Sub Tool1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Index
Case 0: Label28.Caption = "If this is selected you must write your own Tool Tip."
Case 1: Label28.Caption = "Check this if you like this Tool Tip Style"
Case 2: Label28.Caption = "Check this if you like this Tool Tip Style"
Case 3: Label28.Caption = "Check this if you like this Tool Tip Style"
Case 4: Label28.Caption = "Click on this and hold your mouse over it and look at the ToolTip"
Case 5: Label28.Caption = "Click on this and hold your mouse over it and look at the ToolTip"
Case 6: Label28.Caption = "Click on this and hold your mouse over it and look at the ToolTip"
End Select
End Sub

Private Sub keyallow()          'Set key allowed in the different text box's
    ExTextbox1.BindControl Text5
    ExTextbox1.ValidKeys = "abcdefghijklmnopqrstuvwxyzæøåABCDEFGHIJKLMNOPQRSTUVWXYZÆØÅ0123456789-_! "
    
    ExTextbox2.BindControl Text7
    ExTextbox2.ValidKeys = "abcdefghijklmnopqrstuvwxyzæøåABCDEFGHIJKLMNOPQRSTUVWXYZÆØÅ0123456789-_! "

    ExTextbox3.BindControl Text9
    ExTextbox3.ValidKeys = "abcdefghijklmnopqrstuvwxyzæøåABCDEFGHIJKLMNOPQRSTUVWXYZÆØÅ0123456789-_! "
    
    ExTextbox4.BindControl Text11
    ExTextbox4.ValidKeys = "abcdefghijklmnopqrstuvwxyzæøåABCDEFGHIJKLMNOPQRSTUVWXYZÆØÅ0123456789-_! "

    ExTextbox5.BindControl Text13
    ExTextbox5.ValidKeys = "abcdefghijklmnopqrstuvwxyzæøåABCDEFGHIJKLMNOPQRSTUVWXYZÆØÅ0123456789-_! "
    
    ExTextbox6.BindControl Text3
    ExTextbox6.ValidKeys = "abcdefghijklmnopqrstuvwxyzæøåABCDEFGHIJKLMNOPQRSTUVWXYZÆØÅ0123456789-_! "

    ExTextbox7.BindControl Text16
    ExTextbox7.ValidKeys = "abcdefghijklmnopqrstuvwxyzæøåABCDEFGHIJKLMNOPQRSTUVWXYZÆØÅ0123456789-_! "
    
    ExTextbox8.BindControl Text18
    ExTextbox8.ValidKeys = "abcdefghijklmnopqrstuvwxyzæøåABCDEFGHIJKLMNOPQRSTUVWXYZÆØÅ0123456789-_! "

    ExTextbox9.BindControl Text20
    ExTextbox9.ValidKeys = "abcdefghijklmnopqrstuvwxyzæøåABCDEFGHIJKLMNOPQRSTUVWXYZÆØÅ0123456789-_! "
    
    ExTextbox10.BindControl Text22
    ExTextbox10.ValidKeys = "abcdefghijklmnopqrstuvwxyzæøåABCDEFGHIJKLMNOPQRSTUVWXYZÆØÅ0123456789-_! "
End Sub
