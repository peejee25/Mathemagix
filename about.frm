VERSION 5.00
Begin VB.Form about 
   Caption         =   "About MatheMagix"
   ClientHeight    =   3600
   ClientLeft      =   3675
   ClientTop       =   2355
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3600
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1913
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label5 
      Caption         =   "Developed by: Prashant Ganesh"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Distribution:   Freeware"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Application Type:   Graph Plotter && Numerical Analyzer"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Version:   1.0"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "MatheMagix"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1373
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub
