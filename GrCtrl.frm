VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form ChartOptions 
   Caption         =   "Mathemagix - Chart Options"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Advanced >>"
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   2490
      Width           =   1455
   End
   Begin VB.OptionButton ChartType 
      Caption         =   "2D X-Y"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   16
      Left            =   2205
      TabIndex        =   20
      Top             =   2445
      Width           =   1440
   End
   Begin VB.OptionButton ChartType 
      Caption         =   "2D Pie"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   14
      Left            =   2205
      TabIndex        =   19
      Top             =   2085
      Width           =   1440
   End
   Begin VB.OptionButton ChartType 
      Caption         =   "2D Combination"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   9
      Left            =   2205
      TabIndex        =   18
      Top             =   1725
      Width           =   1830
   End
   Begin VB.OptionButton ChartType 
      Caption         =   "3D Combination"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   8
      Left            =   2205
      TabIndex        =   17
      Top             =   1380
      Width           =   1815
   End
   Begin VB.OptionButton ChartType 
      Caption         =   "2D Step"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   7
      Left            =   2205
      TabIndex        =   16
      Top             =   1020
      Width           =   1440
   End
   Begin VB.OptionButton ChartType 
      Caption         =   "3D Step"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   6
      Left            =   2205
      TabIndex        =   15
      Top             =   660
      Width           =   1440
   End
   Begin VB.OptionButton ChartType 
      Caption         =   "2D Area"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   480
      TabIndex        =   14
      Top             =   2445
      Width           =   1440
   End
   Begin VB.OptionButton ChartType 
      Caption         =   "3D Area"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   480
      TabIndex        =   13
      Top             =   2085
      Width           =   1440
   End
   Begin VB.OptionButton ChartType 
      Caption         =   "2D Line"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   480
      TabIndex        =   12
      Top             =   1725
      Width           =   1440
   End
   Begin VB.OptionButton ChartType 
      Caption         =   "3D Line"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   480
      TabIndex        =   11
      Top             =   1380
      Width           =   1440
   End
   Begin VB.OptionButton ChartType 
      Caption         =   "2D Bar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   495
      TabIndex        =   10
      Top             =   1020
      Value           =   -1  'True
      Width           =   1440
   End
   Begin VB.OptionButton ChartType 
      Caption         =   "3D Bar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   660
      Width           =   1440
   End
   Begin VB.Frame Frame3 
      Caption         =   "3D Chart Appearance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   6195
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Add Light"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4800
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         LargeChange     =   10
         Left            =   2340
         Max             =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   2235
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   10
         Left            =   2340
         Max             =   360
         TabIndex        =   3
         Top             =   630
         Width           =   2235
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Hide Visible Edges"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   2085
      End
      Begin VB.Label Label2 
         Caption         =   "Rotate Vertically"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2340
         TabIndex        =   8
         Top             =   1080
         Width           =   1845
      End
      Begin VB.Label Label1 
         Caption         =   "Rotate Horizontally"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2340
         TabIndex        =   7
         Top             =   390
         Width           =   1845
      End
      Begin VB.Label Label3 
         Caption         =   "Ambient Light 0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   6
         Top             =   555
         Width           =   1650
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chart Type "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   225
      TabIndex        =   0
      Top             =   240
      Width           =   3900
   End
End
Attribute VB_Name = "ChartOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ChartType_Click(Index As Integer)
'   Change the chart's type and enable/disable certain controls on the Form
'   according to whether the chart is 2D or 3D (can't rotate 2 2D chart, for example)
    If Left(ChartType(Index).Caption, 2) = "2D" Then
        Label3.Enabled = False
        UpDown1.Enabled = False
        Check1.Enabled = False
        Command4.Enabled = False
    Else
        Label3.Enabled = True
        UpDown1.Enabled = True
        Check1.Enabled = True
        Command4.Enabled = True
    End If
    Form1.MSChart1.ChartType = Index

End Sub

Private Sub Check1_Click()
'   Make edges visible
    Form1.MSChart1.Plot.Light.EdgeIntensity = Check1.Value
End Sub


Private Sub Command1_Click()
frmOptions.Show
End Sub

Private Sub Command4_Click()
    Form1.MSChart1.Plot.Light.LightSources(1).Set Rnd() * 10, Rnd() * 10, Rnd() * 10, Rnd()
End Sub


Private Sub HScroll1_Change()
'   Adjust the viewing angle of 3D charts
    Form1.MSChart1.Plot.View3d.Set HScroll1.Value, HScroll2.Value
End Sub

Private Sub HScroll2_Change()
'   Adjust the viewing angle of 3D charts
    Form1.MSChart1.Plot.View3d.Set HScroll1.Value, HScroll2.Value
End Sub

Private Sub MSChart1_SeriesSelected(Series As Integer, MouseFlags As Integer, Cancel As Integer)
'    redClr = MSChart1.Plot.SeriesCollection(Series).Pen.VtColor.Red
'    greenClr = MSChart1.Plot.SeriesCollection(Series).Pen.VtColor.Green
'    blueClr = MSChart1.Plot.SeriesCollection(Series).Pen.VtColor.Blue
'    Debug.Print "The selected series color is (" & redClr & ", " & greenClr & ", " & blueClr & ")"
End Sub

Private Sub UpDown1_Change()
'   Adjust intensity of ambient light
    Form1.MSChart1.Plot.Light.EdgeVisible = True
    Form1.MSChart1.Plot.Light.AmbientIntensity = UpDown1.Value / 10
    If UpDown1.Value = 0 Then Form1.MSChart1.Plot.Light.EdgeVisible = False
    Label3.Caption = "Ambient Light " & UpDown1.Value
End Sub

