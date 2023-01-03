VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form calciform 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mathemagix SciCalc"
   ClientHeight    =   4095
   ClientLeft      =   1845
   ClientTop       =   1605
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4095
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton MemClear 
      Caption         =   "MC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5880
      TabIndex        =   50
      Top             =   3480
      Width           =   540
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   120
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.CommandButton MemoryRecall 
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5880
      TabIndex        =   45
      Top             =   2280
      Width           =   540
   End
   Begin VB.CommandButton MemoryMinus 
      Caption         =   "M-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5880
      TabIndex        =   44
      Top             =   3080
      Width           =   540
   End
   Begin VB.CommandButton MemoryPlus 
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5880
      TabIndex        =   43
      Top             =   2680
      Width           =   540
   End
   Begin VB.CommandButton CLRButton 
      Caption         =   "CLR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7320
      TabIndex        =   42
      Top             =   120
      Width           =   540
   End
   Begin VB.CommandButton BSButton 
      Caption         =   "BS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6600
      TabIndex        =   41
      Top             =   120
      Width           =   540
   End
   Begin VB.CommandButton DelButton 
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5880
      TabIndex        =   40
      Top             =   120
      Width           =   540
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   810
      HideSelection   =   0   'False
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   105
      Width           =   4140
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Calculate"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   135
      TabIndex        =   2
      Top             =   1590
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Height          =   1920
      Left            =   90
      TabIndex        =   3
      Top             =   2040
      Width           =   1785
      Begin VB.CommandButton NumPeriod 
         BackColor       =   &H00C0C0C0&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1200
         TabIndex        =   28
         Top             =   1440
         Width           =   460
      End
      Begin VB.CommandButton Num 
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   0
         Left            =   660
         TabIndex        =   27
         Top             =   1440
         Width           =   460
      End
      Begin VB.CommandButton Num 
         BackColor       =   &H00C0C0C0&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   1200
         TabIndex        =   26
         Top             =   255
         WhatsThisHelpID =   1
         Width           =   465
      End
      Begin VB.CommandButton Num 
         BackColor       =   &H00C0C0C0&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   8
         Left            =   660
         TabIndex        =   25
         Top             =   255
         Width           =   460
      End
      Begin VB.CommandButton Num 
         BackColor       =   &H00C0C0C0&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   135
         TabIndex        =   24
         Top             =   255
         Width           =   465
      End
      Begin VB.CommandButton Num 
         BackColor       =   &H00C0C0C0&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   6
         Left            =   1200
         TabIndex        =   23
         Top             =   645
         Width           =   460
      End
      Begin VB.CommandButton Num 
         BackColor       =   &H00C0C0C0&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   5
         Left            =   660
         TabIndex        =   22
         Top             =   645
         Width           =   460
      End
      Begin VB.CommandButton Num 
         BackColor       =   &H00C0C0C0&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   645
         Width           =   460
      End
      Begin VB.CommandButton Num 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   3
         Left            =   1200
         TabIndex        =   20
         Top             =   1050
         Width           =   460
      End
      Begin VB.CommandButton Num 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   2
         Left            =   660
         TabIndex        =   19
         Top             =   1050
         Width           =   460
      End
      Begin VB.CommandButton Num 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   1
         Left            =   135
         TabIndex        =   18
         Top             =   1050
         Width           =   460
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      Height          =   1920
      Left            =   1995
      TabIndex        =   4
      Top             =   2040
      Width           =   1785
      Begin VB.CommandButton Operator 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   225
         Width           =   465
      End
      Begin VB.CommandButton Operator 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   1
         Left            =   645
         TabIndex        =   47
         Top             =   225
         Width           =   460
      End
      Begin VB.CommandButton Operator 
         Caption         =   "("
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   2
         Left            =   1185
         TabIndex        =   46
         Top             =   225
         Width           =   460
      End
      Begin VB.CommandButton Operator 
         Caption         =   "^"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   8
         Left            =   1185
         TabIndex        =   17
         Top             =   1065
         Width           =   460
      End
      Begin VB.CommandButton Operator 
         Caption         =   ")"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   5
         Left            =   1185
         TabIndex        =   16
         Top             =   645
         Width           =   460
      End
      Begin VB.CommandButton Operator 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   9
         Left            =   1185
         TabIndex        =   15
         Top             =   1485
         Width           =   460
      End
      Begin VB.CommandButton Operator 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   7
         Left            =   645
         TabIndex        =   14
         Top             =   1065
         Width           =   460
      End
      Begin VB.CommandButton Operator 
         Caption         =   "\"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   1065
         Width           =   460
      End
      Begin VB.CommandButton Operator 
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   4
         Left            =   645
         TabIndex        =   12
         Top             =   645
         Width           =   460
      End
      Begin VB.CommandButton Operator 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   645
         Width           =   460
      End
      Begin VB.CommandButton CalculateButton 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1485
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00008000&
      Height          =   1920
      Left            =   3900
      TabIndex        =   5
      Top             =   2040
      Width           =   1725
      Begin VB.CommandButton FunctionButton 
         Caption         =   "int"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   1155
         TabIndex        =   39
         Top             =   653
         Width           =   465
      End
      Begin VB.CommandButton FunctionButton 
         Caption         =   "tan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   11
         Left            =   1155
         TabIndex        =   38
         Top             =   1485
         Width           =   460
      End
      Begin VB.CommandButton FunctionButton 
         Caption         =   "sin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   1155
         TabIndex        =   37
         Top             =   1066
         Width           =   465
      End
      Begin VB.CommandButton FunctionButton 
         Caption         =   "cos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   2
         Left            =   1155
         TabIndex        =   36
         Top             =   240
         Width           =   460
      End
      Begin VB.CommandButton FunctionButton 
         Caption         =   "rnd"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   7
         Left            =   645
         TabIndex        =   35
         Top             =   1066
         Width           =   460
      End
      Begin VB.CommandButton FunctionButton 
         Caption         =   "sqr"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   10
         Left            =   645
         TabIndex        =   34
         Top             =   1485
         Width           =   460
      End
      Begin VB.CommandButton FunctionButton 
         Caption         =   "fix"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   4
         Left            =   645
         TabIndex        =   33
         Top             =   653
         Width           =   460
      End
      Begin VB.CommandButton FunctionButton 
         Caption         =   "atn"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   645
         TabIndex        =   32
         Top             =   240
         Width           =   460
      End
      Begin VB.CommandButton FunctionButton 
         Caption         =   "sgn"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   9
         Left            =   135
         TabIndex        =   9
         Top             =   1485
         Width           =   460
      End
      Begin VB.CommandButton FunctionButton 
         Caption         =   "exp"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   135
         TabIndex        =   8
         Top             =   653
         Width           =   460
      End
      Begin VB.CommandButton FunctionButton 
         Caption         =   "log"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   6
         Left            =   135
         TabIndex        =   7
         Top             =   1066
         Width           =   460
      End
      Begin VB.CommandButton FunctionButton 
         Caption         =   "abs"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   240
         Width           =   460
      End
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2280
      TabIndex        =   1
      Text            =   "0.0"
      Top             =   1095
      Width           =   3390
   End
   Begin VB.Label Memindicator 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5880
      TabIndex        =   49
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "X  Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   255
      TabIndex        =   31
      Top             =   1140
      Width           =   1125
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "Expression"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   210
      TabIndex        =   30
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label Result 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   2280
      TabIndex        =   29
      Top             =   1605
      Width           =   3405
   End
End
Attribute VB_Name = "calciform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Developed by Prashant Ganesh
'As a part of Mathemagix Application


Option Explicit
Private MemValue As Double
Private X As Double
Private SValue As Double
Private OpenLog As Integer

Private Sub CLRButton_Click()
    Text1.Text = ""
    Text1.SetFocus
End Sub

Private Sub Command1_Click()
Dim Value As Double
On Error GoTo EvalError
    X = Val(Text3.Text)
    ScriptControl1.ExecuteStatement "X=" & X
    Value = ScriptControl1.Eval(Trim(Text1.Text))
    Result.Caption = Value
    'UpdateLog
    Exit Sub
EvalError:
    Result.Caption = " Invalid Expression!"
    'MsgBox ScriptControl1.Error.Description
End Sub

'Private Sub Command2_Click()
 '   CopyrightForm.Show vbModal
'End Sub

Private Sub DelButton_Click()
    Dim sstart As Integer
    If Len(Text1.Text) = 0 Then
        Text1.SetFocus
        Exit Sub
    End If
    If Len(Text1.SelText) > 1 Then
        'present cursor position character is deleted
        Text1.SelText = ""
        Text1.SetFocus
        Exit Sub
    End If
    If Text1.SelStart = Len(Text1.Text) Then
        Text1.SetFocus
        Exit Sub
    End If
    'if the length of selected string is 1 handle it as a
    'special case
    sstart = Text1.SelStart
    Text1.Text = Left$(Text1.Text, Text1.SelStart) + Right$(Text1.Text, Len(Text1.Text) - Text1.SelStart - 1)
    Text1.SetFocus
    Text1.SelStart = sstart
End Sub

Private Sub BSButton_Click()
    Dim sstart As Integer
    sstart = Text1.SelStart
    If Text1.SelStart = 0 Then
        Text1.SetFocus
        Text1.SelStart = sstart
        Exit Sub
    Else
        If Text1.SelLength > 1 Then
            Text1.SelText = ""
            Text1.SetFocus
            Text1.SelStart = sstart
        Else
            Text1.Text = Left$(Text1.Text, Text1.SelStart - 1) + Right(Text1.Text, Len(Text1.Text) - Text1.SelStart)
            Text1.SetFocus
            Text1.SelStart = sstart - 1
        End If
    End If
End Sub

'Private Sub FileHideShow_Click()
    'If FileHideShow.Caption = "Hide Buttons" Then
       ' Form1.Height = 2790
       ' FileHideShow.Caption = "Show Buttons"
   ' Else
    '    Form1.Height = 4935
        'FileHideShow.Caption = "Hide Buttons"
   ' End If
'End Sub
'Private Sub LogBttn_Click()
 '   If LogBttn.Caption = "Show  Log" Then
  '      OpenLog = True
   '     LogBttn.Caption = "Hide  Log"
    '    LogForm.Show
    'Else
     '   OpenLog = False
      '  LogBttn.Caption = "Show  Log"
       ' LogForm.List1.Clear
        'LogForm.Hide
    'End If
'End Sub

'Private Sub UpdateLog()
 '   LogForm.List1.AddItem Trim(calciform.Text1.Text)
  '  LogForm.List1.AddItem "For X = " + calciform.Text3.Text
   ' LogForm.List1.AddItem ">        " + calciform.Result.Caption
'End Sub

Private Sub FunctionButton_Click(Index As Integer)
    Text1.SelText = ActiveControl.Caption + "("
    Text1.SetFocus
End Sub


Private Sub MemClear_Click()
MemValue = 0
Memindicator.Caption = ""
End Sub

Private Sub MemoryMinus_Click()
On Error Resume Next
    MemValue = MemValue - Result.Caption
    If MemValue = 0 Then
    Memindicator.Caption = ""
    Else
        Memindicator.Caption = " M"
    End If
End Sub

Private Sub MemoryPlus_Click()
On Error Resume Next
    MemValue = MemValue + Result.Caption
    If MemValue = 0 Then
    Memindicator.Caption = ""
    Else
        Memindicator.Caption = " M"
    End If
End Sub

Private Sub MemoryRecall_Click()
On Error Resume Next
    Result.Caption = MemValue
End Sub

Private Sub Num_Click(Index As Integer)
    Text1.SelText = ActiveControl.Caption
    Text1.SetFocus
End Sub

Private Sub NumPeriod_Click()
    Text1.SelText = ActiveControl.Caption
    Text1.SetFocus
End Sub

Private Sub CalculateButton_Click()
    Call Command1_Click
End Sub

Private Sub Operator_Click(Index As Integer)
    If ActiveControl.Caption = "%" Then
            Text1.SelText = " mod "
        ElseIf ActiveControl.Caption = "x" Then
            Text1.SelText = "*"
    Else
        Text1.SelText = ActiveControl.Caption
    End If
    Text1.SetFocus
End Sub






Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 61 Then
    'don't display anything in expression box
    KeyAscii = 0
    'evaluate expression
    Command1_Click
End If
End Sub
