VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   Caption         =   "PGrapher"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   6990
   Icon            =   "mathemagic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   6990
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   9615
      Left            =   0
      ScaleHeight     =   637
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   789
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   7215
         Left            =   120
         OleObjectBlob   =   "mathemagic.frx":0442
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   11895
      End
   End
   Begin VB.Menu menufile 
      Caption         =   "&File"
      Begin VB.Menu filenew 
         Caption         =   "&New"
      End
      Begin VB.Menu fileopen 
         Caption         =   "&Open"
      End
      Begin VB.Menu filesave 
         Caption         =   "&Save"
      End
      Begin VB.Menu filesaveas 
         Caption         =   "Save &As.."
      End
      Begin VB.Menu filesavebitmap 
         Caption         =   "Save as &Bitmap"
      End
      Begin VB.Menu fileseperator 
         Caption         =   "-"
      End
      Begin VB.Menu fileexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menutools 
      Caption         =   "&Tools"
      Begin VB.Menu toolscalci 
         Caption         =   "Scientific &Calculator"
      End
      Begin VB.Menu toolsnumal 
         Caption         =   "&Numerical Analyzer"
      End
      Begin VB.Menu toolsplotgraph 
         Caption         =   "&Plot Graph"
         Begin VB.Menu toolsplotfunc 
            Caption         =   "&Function"
            Begin VB.Menu func1 
               Caption         =   "1 Variable"
            End
            Begin VB.Menu func2 
               Caption         =   "2 Variables"
            End
         End
         Begin VB.Menu toolsplotTable 
            Caption         =   "&Tabular points"
         End
      End
      Begin VB.Menu Toolschartoptions 
         Caption         =   "Chart Options..."
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu HelpAbout 
         Caption         =   "About Mathemagix..."
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public fstring As String






Private Sub Form_Unload(Cancel As Integer)
    Unload calciform
    Unload tableplot
    Unload optiondata
    Unload ChartOptions
    Unload about
    Unload Form2
    Unload frmOptions
    Unload func2_form
End Sub

Private Sub func1_Click()
plottype = "Graph"

    
 Form2.Show

End Sub

Private Sub func2_Click()
MSChart1.Visible = False
func2_form.Show
End Sub

Private Sub HelpAbout_Click()
 about.Show
End Sub






Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
xangle = xangle - 2
yangle = yangle + 2

ElseIf KeyCode = 39 Then
xangle = xangle + 2
yangle = yangle - 2
End If

func2_form.draw_curve
End Sub

Private Sub toolscalci_Click()
    calciform.Show
End Sub

Private Sub Toolschartoptions_Click()
    ChartOptions.Show
End Sub


Private Sub toolsplotTable_Click()
    plottype = "chart"
    tableplot.Show
    Toolschartoptions.Enabled = True
    
End Sub
