VERSION 5.00
Begin VB.Form frmWiz 
   BackColor       =   &H00843201&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Folder Wizard"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13530
   Icon            =   "frmWiz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   13530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00843201&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   9120
      TabIndex        =   55
      Top             =   120
      Width           =   4455
      Begin VB.Image Image1 
         Height          =   495
         Index           =   0
         Left            =   120
         Top             =   120
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   1
         Left            =   120
         Top             =   720
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   2
         Left            =   120
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   3
         Left            =   120
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   4
         Left            =   120
         Top             =   2520
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   5
         Left            =   120
         Top             =   3120
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   6
         Left            =   120
         Top             =   3720
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   7
         Left            =   720
         Top             =   120
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   8
         Left            =   720
         Top             =   720
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   9
         Left            =   720
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   10
         Left            =   720
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   11
         Left            =   720
         Top             =   2520
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   12
         Left            =   720
         Top             =   3120
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   13
         Left            =   720
         Top             =   3720
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   14
         Left            =   1320
         Top             =   120
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   15
         Left            =   1320
         Top             =   720
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   16
         Left            =   1320
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   17
         Left            =   1320
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   18
         Left            =   1320
         Top             =   2520
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   19
         Left            =   1320
         Top             =   3120
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   20
         Left            =   1320
         Top             =   3720
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   21
         Left            =   1920
         Top             =   120
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   22
         Left            =   1920
         Top             =   720
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   23
         Left            =   1920
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   24
         Left            =   1920
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   25
         Left            =   1920
         Top             =   2520
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   26
         Left            =   1920
         Top             =   3120
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   27
         Left            =   1920
         Top             =   3720
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   28
         Left            =   2520
         Top             =   120
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   29
         Left            =   2520
         Top             =   720
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   30
         Left            =   2520
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   31
         Left            =   2520
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   32
         Left            =   2520
         Top             =   2520
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   33
         Left            =   2520
         Top             =   3120
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   34
         Left            =   2520
         Top             =   3720
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   35
         Left            =   3120
         Top             =   120
         Width           =   495
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   495
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   36
         Left            =   3120
         Top             =   720
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   37
         Left            =   3120
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   38
         Left            =   3120
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   39
         Left            =   3120
         Top             =   2520
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   40
         Left            =   3120
         Top             =   3120
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   41
         Left            =   3120
         Top             =   3720
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   42
         Left            =   3720
         Top             =   120
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   43
         Left            =   3720
         Top             =   720
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   44
         Left            =   3720
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   45
         Left            =   3720
         Top             =   1920
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   46
         Left            =   3720
         Top             =   2520
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Index           =   47
         Left            =   3720
         Top             =   3120
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdFinish 
      BackColor       =   &H00C4C28F&
      Caption         =   "&Finish"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4C28F&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C4C28F&
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C4C28F&
      Caption         =   "&Next >"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox picWiz 
      BackColor       =   &H00843201&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3975
      ScaleWidth      =   2175
      TabIndex        =   25
      Top             =   0
      Width           =   2175
      Begin VB.Label lblSteps 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finish"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   30
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label lblSteps 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   29
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label lblSteps 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   28
         Top             =   960
         Width           =   465
      End
      Begin VB.Label lblSteps 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   27
         Top             =   480
         Width           =   465
      End
      Begin VB.Label lblSteps 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   26
         Top             =   120
         Width           =   420
      End
      Begin VB.Shape shpSteps 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   120
         Shape           =   1  'Square
         Top             =   1800
         Width           =   255
      End
      Begin VB.Shape shpSteps 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   360
         Shape           =   1  'Square
         Top             =   1440
         Width           =   255
      End
      Begin VB.Shape shpSteps 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   360
         Shape           =   1  'Square
         Top             =   960
         Width           =   255
      End
      Begin VB.Shape shpSteps 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   360
         Shape           =   1  'Square
         Top             =   480
         Width           =   255
      End
      Begin VB.Shape shpSteps 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   120
         Shape           =   1  'Square
         Top             =   120
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   1695
         Left            =   240
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame fraSteps 
      BackColor       =   &H00843201&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   0
      Left            =   2280
      TabIndex        =   31
      Top             =   120
      Width           =   6495
      Begin VB.Image Image2 
         Height          =   2895
         Left            =   3960
         Picture         =   "frmWiz.frx":030A
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWiz.frx":5DD1
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   3495
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblHead 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Create your own System folder (Wizard)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   6375
      End
   End
   Begin VB.Frame fraSteps 
      BackColor       =   &H00843201&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   1
      Left            =   2280
      TabIndex        =   32
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox text37 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "This is my own SystemFolder"
         Top             =   3360
         Width           =   3855
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "c:\"
         Top             =   2400
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Text            =   "Folder Name"
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   240
         X2              =   6240
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   240
         X2              =   6240
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWiz.frx":5E59
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   45
         Top             =   2880
         Width           =   5415
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWiz.frx":5EE0
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   44
         Top             =   1800
         Width           =   5415
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   240
         X2              =   6240
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   240
         X2              =   6240
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWiz.frx":5FA8
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   43
         Top             =   600
         Width           =   5415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1 (The first stuff to do)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   38
         Top             =   120
         Width           =   2790
      End
   End
   Begin VB.Frame fraSteps 
      BackColor       =   &H00843201&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   2
      Left            =   2280
      TabIndex        =   33
      Top             =   120
      Width           =   6495
      Begin VB.TextBox text38 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   3120
         Width           =   4335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C4C28F&
         Cancel          =   -1  'True
         Caption         =   "Choose Icon >"
         Height          =   285
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3120
         Width           =   1215
      End
      Begin VB.OptionButton F1C3 
         BackColor       =   &H00843201&
         Caption         =   "Both"
         ForeColor       =   &H00FFC0C0&
         Height          =   215
         Left            =   2520
         TabIndex        =   6
         Top             =   1680
         Width           =   735
      End
      Begin VB.OptionButton F1C1 
         BackColor       =   &H00843201&
         Caption         =   "Desktop"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton F1C2 
         BackColor       =   &H00843201&
         Caption         =   "My computer"
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   240
         X2              =   6240
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWiz.frx":6052
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   240
         TabIndex        =   47
         Top             =   2160
         Width           =   5415
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   240
         X2              =   6240
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   240
         X2              =   6240
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWiz.frx":617A
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   240
         TabIndex        =   46
         Top             =   600
         Width           =   5415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2 (Some useful options)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   39
         Top             =   120
         Width           =   2985
      End
   End
   Begin VB.Frame fraSteps 
      BackColor       =   &H00843201&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   4
      Left            =   2160
      TabIndex        =   35
      Top             =   120
      Width           =   6735
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Feel free to change or add something to your system folder before you click the red button after you press finsish..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   1080
         TabIndex        =   57
         Top             =   2280
         Width           =   4695
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Thank you for using this wizard!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1800
         TabIndex        =   56
         Top             =   3240
         Width           =   3360
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Thats it!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   42
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWiz.frx":62D1
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   1080
         TabIndex        =   41
         Top             =   1080
         Width           =   4695
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraSteps 
      BackColor       =   &H00843201&
      BorderStyle     =   0  'None
      Height          =   3855
      Index           =   3
      Left            =   2280
      TabIndex        =   34
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3165
         Width           =   255
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00C4C28F&
         Caption         =   "..."
         Height          =   255
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3450
         Width           =   255
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2880
         TabIndex        =   10
         Text            =   "Full path"
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2880
         TabIndex        =   12
         Top             =   3165
         Width           =   2295
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   2880
         TabIndex        =   14
         Top             =   3450
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "Name"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   3165
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H80000007&
         ForeColor       =   &H00FFFF80&
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   3450
         Width           =   1695
      End
      Begin VB.CommandButton Command31 
         BackColor       =   &H00C4C28F&
         Caption         =   "Folder"
         Height          =   255
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton Command32 
         BackColor       =   &H00C4C28F&
         Caption         =   "Folder"
         Height          =   255
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3165
         Width           =   735
      End
      Begin VB.CommandButton Command33 
         BackColor       =   &H00C4C28F&
         Caption         =   "Folder"
         Height          =   255
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3450
         Width           =   735
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   1650
         Left            =   4480
         Picture         =   "frmWiz.frx":638B
         ScaleHeight     =   1590
         ScaleWidth      =   1875
         TabIndex        =   49
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label18 
         BackColor       =   &H00843201&
         Caption         =   "Number 1:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   2880
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label Label17 
         BackColor       =   &H00843201&
         Caption         =   "Number 2:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   3165
         Width           =   855
      End
      Begin VB.Label Label16 
         BackColor       =   &H00843201&
         Caption         =   "Number 3:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   3450
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   51
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Path"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   50
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWiz.frx":7160
         ForeColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   120
         TabIndex        =   48
         Top             =   600
         Width           =   4335
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00C0C000&
         BorderWidth     =   2
         X1              =   240
         X2              =   6240
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 3 (The best and last part)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   40
         Top             =   120
         Width           =   3150
      End
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   2160
      Top             =   3840
      Width           =   6375
   End
End
Attribute VB_Name = "frmWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Index As Integer
Dim currentshape As Byte

Private Sub MoveIndex(PrevIndex As Integer, NextIndex As Integer)
    'Are we at the last step?
    If PrevIndex = shpSteps.Count - 1 Then
        'Yes
        shpSteps(PrevIndex).FillColor = vbRed
    Else
        'No
        shpSteps(PrevIndex).FillColor = &H808080
    End If
    'Set bold off
    lblSteps(PrevIndex).FontBold = False
    
    'Set next step color and font bold
    shpSteps(NextIndex).FillColor = vbGreen
    lblSteps(NextIndex).FontBold = True
End Sub

Private Sub cmdBack_Click()
    Index = Index - 1
    If Index <= 0 Then
        Index = 0
        cmdBack.Enabled = False
    End If
    
    MoveIndex Index + 1, Index

    'Set the frames
    fraSteps(Index).ZOrder 0
    
    'Set command buttons
    cmdNext.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    frmmain.Show
    Unload Me
End Sub

Private Sub cmdFinish_Click()
    'Put some code here to check if all data captured on the previous screens
    'was correct and then proceed
    PerformSomething
    frmmain.Show
    frmmain.Command2.BackColor = vbRed
    Unload Me
End Sub

Private Sub PerformSomething()
With frmmain
If F1C1.Enabled = True Then .F1C1.Enabled = True    'Place to put SysFolder
If F1C2.Enabled = True Then .F1C2.Enabled = True    'Place to put SysFolder
If F1C3.Enabled = True Then .F1C3.Enabled = True    'Place to put SysFolder

.text38 = text38        'Icon path

.Text5 = Text5          'Item name/path
.Text6 = Text6          'Item name/path
.Text7 = Text7          'Item name/path
.Text8 = Text8          'Item name/path
.Text9 = Text9          'Item name/path
.Text10 = Text10        'Item name/path

.Command31.Caption = Command31.Caption      'File or Folder
.Command32.Caption = Command32.Caption      'File or Folder
.Command33.Caption = Command33.Caption      'File or Folder

.Text1 = Text1          'Folder Name
.Text4 = Text4          'Dobble click path
.text37 = text37        'Tool Tip

If .Option1.Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\1\", "ShapePosition", "" & currentshape & ""
If .Option2.Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\2\", "ShapePosition", "" & currentshape & ""
If .Option3.Value = True Then SaveSettingString HKEY_LOCAL_MACHINE, "Software\Multima\SysFCreator\sysf\3\", "ShapePosition", "" & currentshape & ""

.checkshape

End With

End Sub

Private Sub cmdNext_Click()
    Index = Index + 1
    If Index >= shpSteps.Count - 1 Then
        Index = shpSteps.Count - 1
        cmdNext.Enabled = False
    End If
    
    MoveIndex Index - 1, Index
    
    'Set the frames
    fraSteps(Index).ZOrder 0
    
    'Set command buttons
    cmdBack.Enabled = True
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Choose Icon >" Then
Me.Width = "13620"
Command1.Caption = "Choose Icon <"
Else
Me.Width = "8970"
Command1.Caption = "Choose Icon >"
End If

    Me.Top = (Screen.Height * 0.85) / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
End Sub

Private Sub Command10_Click()
frmDir.Show
frmDir.Wizz2
Me.Hide
End Sub

Private Sub Command11_Click()
frmDir.Show
frmDir.Wizz3
Me.Hide
End Sub

Private Sub Command12_Click()
frmDir.Show
frmDir.Wizz4
Me.Hide
End Sub

Private Sub Command2_Click()
frmDir.Show
frmDir.Wizz1
Me.Hide
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

Private Sub Form_Load()
Me.Width = "8970"
    Index = 0
    currentshape = "0"
    loadpic
End Sub
Private Sub loadpic()
On Error Resume Next

Dim i As Integer
i = "0"

Do Until i = 48
Image1(i) = LoadResPicture(i + 100 + 1, 1)
If Val(Image1(i).Picture) = 0 Then Image1(i).Visible = False
i = i + 1
Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmmain.Show
End Sub

Private Sub Image1_Click(Index As Integer)
text38.Text = App.Path & "\" & App.EXEName & ".exe," & Index + 1
currentshape = Index
Shape3.Visible = True
Shape3.Top = Image1(Index).Top
Shape3.Left = Image1(Index).Left
End Sub

Private Sub Image1_DblClick(Index As Integer)
Command1_Click
End Sub
