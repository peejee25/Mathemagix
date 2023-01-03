VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4920
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6150
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   240
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1320
         TabIndex        =   26
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1320
         TabIndex        =   24
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1320
         TabIndex        =   19
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Footer Text"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   2295
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Y-Axis Label"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1695
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "X-Axis Label"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Chart Title"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   495
         Width           =   975
      End
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4080
      TabIndex        =   18
      Text            =   "Text4"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   0
      Top             =   4455
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   0
      Left            =   240
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3840
         TabIndex        =   15
         Text            =   "Text3"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   1140
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1980
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3840
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmOptions.frx":000C
         Left            =   240
         List            =   "frmOptions.frx":000E
         TabIndex        =   4
         Text            =   "X-Axis"
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Minor Divisions: "
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   3135
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Major Divisions: "
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   2655
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "AutoScale"
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Show Grid"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Max: "
         Height          =   255
         Left            =   3360
         TabIndex        =   8
         Top             =   2175
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Min: "
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   1680
         Width           =   375
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Axis"
            Key             =   "Group1"
            Object.ToolTipText     =   "Set Options for Group 1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Chart Labels"
            Key             =   "Group2"
            Object.ToolTipText     =   "Set Options for Group 2"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X(6) As Integer
Dim y(6) As Integer

Private Sub Check1_Click()

If Combo1.Text = "X-Axis" Then
X(2) = Check1.Value
Else
y(2) = Check1.Value
End If

End Sub

Private Sub Check2_Click()

If Check2.Value = 1 Then
If Combo1.Text = "X-Axis" Then
X(3) = 1
Else
y(3) = 1
End If
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False

Else
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True

If Combo1.Text = "X-Axis" Then
X(3) = 0
Else
y(3) = 0
End If

End If

End Sub

Private Sub cmdApply_Click()

Form1.MSChart1.Title.Text = Text5.Text
Form1.MSChart1.Footnote.Text = Text8.Text
Form1.MSChart1.Plot.Axis(VtChAxisIdX).AxisTitle.Text = Text6.Text
Form1.MSChart1.Plot.Axis(VtChAxisIdY).AxisTitle.Text = Text7.Text


If X(3) = 0 Then
Form1.MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
Else
Form1.MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Auto = True
End If

If y(3) = 0 Then
Form1.MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
Else
Form1.MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = True
End If



     
     Form1.MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Minimum = X(0)
     Form1.MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Maximum = X(1)
     Form1.MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = X(4)
     Form1.MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = X(5)
     Form1.MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = y(0)
     Form1.MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = y(1)
     Form1.MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = y(4)
     Form1.MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = y(5)
     
    If y(2) = 0 Then
     Form1.MSChart1.Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleNull
     End If
    If X(2) = 0 Then
    Form1.MSChart1.Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull
    End If
    
    If y(2) = 1 Then
    Form1.MSChart1.Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleDotted
    End If
    If X(2) = 1 Then
    Form1.MSChart1.Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleDotted
    End If
    'MsgBox "Place code here to set options w/o closing dialog!"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'MsgBox "Place code here to set options and close dialog!"
    Unload Me
End Sub


Private Sub Combo1_Click()
'Dim itemindex As Integer
'itemindex = Combo1.ItemData(Combo1.ListIndex)

'MsgBox itemindex
If Combo1.Text = "X-Axis" Then
Text1.Text = X(0)
Text2.Text = X(1)
Text3.Text = X(4)
Text4.Text = X(5)


Check1.Value = X(2)
Check2.Value = X(3)
Else
Text1.Text = y(0)
Text2.Text = y(1)
Text3.Text = y(4)
Text4.Text = y(5)

Check1.Value = y(2)
Check2.Value = y(3)


End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub Form_Load()
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    picOptions(1).Left = -20000
    picOptions(1).Enabled = False
    Check2.Value = 1
    Text1.Enabled = False
    Text2.Enabled = False
    Combo1.AddItem "X-Axis"
    Combo1.ItemData(Combo1.NewIndex) = 0
    Combo1.AddItem "Y-Axis"
    Combo1.ItemData(Combo1.NewIndex) = 1
    Combo1.ListIndex = 0
    Text1.Text = Form1.MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Minimum
    Text2.Text = Form1.MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Maximum
    Text3.Text = Form1.MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision
    Text4.Text = Form1.MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision
    X(0) = Form1.MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Minimum
    X(1) = Form1.MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.Maximum
    X(2) = 0
    Form1.MSChart1.Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen.Style = VtPenStyleNull
    X(3) = 1 'indicates autoscale is on
    X(4) = Form1.MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision
    X(5) = Form1.MSChart1.Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision
    
    y(0) = Form1.MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Minimum
    y(1) = Form1.MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum
    y(2) = 0
    Form1.MSChart1.Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen.Style = VtPenStyleNull
    y(3) = 1 'indicates auto scale is on
    y(4) = Form1.MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision
    y(5) = Form1.MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision
    
End Sub






Private Sub tbsOptions_Click()
    
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    
End Sub

Private Sub Text1_Change()
If Not IsNumeric(Text1.Text) Then
Exit Sub
End If

If Combo1.Text = "X-Axis" Then
X(0) = Text1.Text
Else
y(0) = Text1.Text
End If
End Sub

Private Sub Text2_Change()
If Not IsNumeric(Text2.Text) Then
Exit Sub
End If
If Combo1.Text = "X-Axis" Then

X(1) = Text2.Text
Else
y(1) = Text2.Text
End If


End Sub

Private Sub Text3_Change()
If Not IsNumeric(Text3.Text) Then
Exit Sub
End If
If Combo1.Text = "X-Axis" Then

X(4) = Text3.Text
Else
y(4) = Text3.Text
End If


End Sub

Private Sub Text4_Change()
If Not IsNumeric(Text4.Text) Then
Exit Sub
End If
If Combo1.Text = "X-Axis" Then

X(5) = Text4.Text
Else
y(5) = Text4.Text
End If


End Sub

