VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form tableplot 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   5325
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7320
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4913
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Plot 
      Caption         =   "Plot Graph"
      Height          =   375
      Left            =   1433
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5953
      _Version        =   393216
      Rows            =   10
      Cols            =   3
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the Data in following table or load from a file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   5535
   End
   Begin VB.Menu menufile 
      Caption         =   "File"
      Begin VB.Menu fileload 
         Caption         =   "Load data from file"
      End
      Begin VB.Menu filesave 
         Caption         =   "Save data to file"
      End
      Begin VB.Menu filrexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menuedit 
      Caption         =   "Edit"
      Begin VB.Menu EditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu EditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu EditPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu EditSelect 
         Caption         =   "Select All"
      End
      Begin VB.Menu EditClear 
         Caption         =   "Clear Selected"
      End
   End
   Begin VB.Menu menuoptions 
      Caption         =   "Options"
      Begin VB.Menu editdataoptions 
         Caption         =   "Data options..."
      End
      Begin VB.Menu editchartoptions 
         Caption         =   "Chart Options..."
      End
   End
End
Attribute VB_Name = "tableplot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rowcount As Integer
Public colcount As Integer



Private Sub btnCancel_Click()
Unload tableplot
End Sub

Private Sub editchartoptions_Click()
ChartOptions.Show
End Sub

Private Sub editdataoptions_Click()
    optiondata.Show
End Sub

Private Sub Form_Load()
rowcount = 10
colcount = 3
MSFlexGrid1.FormatString = "<     |^         X           |^          Y          "
MSFlexGrid1.FormatString = ";               "
End Sub


Private Sub MSFlexGrid1_Click()
' prepare another grid cell for editing
    
    SetTextBox
    
End Sub

Private Sub MSFlexGrid1_DblClick()
    
    
    
    
    If MSFlexGrid1.MouseRow = 0 Then
    MSFlexGrid1.Row = 0
    SetTextBox
    ElseIf MSFlexGrid1.MouseCol = 0 Then
    MSFlexGrid1.Col = 0
    SetTextBox
    End If
    
    
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Text1.Visible = False Then
    SetTextBox
    End If
End Sub

Private Sub MSFlexGrid1_Scroll()
' This handler ends the editing of the current cell
' when the grid control is scrolled
' Remove this line to see what happens when
' the FlexGrid control is scrolled
 Call Text1_keydown(40, 0)
End Sub

Private Sub Plot_Click()
Form1.MSChart1.Visible = True
On Error GoTo invaliddatavalue
 Dim irow As Integer
 Dim icol As Integer
 Dim rowcnt As Integer
 Dim dataerror As Boolean
 Dim i As Integer
 For i = 1 To rowcount - 2
 If IsNumeric(MSFlexGrid1.TextMatrix(i, 1)) Then
 rowcnt = rowcnt + 1
 
 End If
 If rowcnt = 0 Then
  GoTo invaliddatavalue
 End If
 Next
 Form1.MSChart1.rowcount = rowcnt
 Form1.MSChart1.ColumnCount = colcount - 1
 For irow = 1 To rowcnt
 Form1.MSChart1.Row = irow
 MSFlexGrid1.Col = 0
 MSFlexGrid1.Row = irow
 Form1.MSChart1.RowLabel = MSFlexGrid1.Text
 
 
 Next
 
 For icol = 1 To colcount - 1
 For irow = 1 To rowcnt
 Form1.MSChart1.Row = irow
 Form1.MSChart1.Column = icol
 If IsNumeric(MSFlexGrid1.TextMatrix(irow, icol)) Then
 Form1.MSChart1.Data = MSFlexGrid1.TextMatrix(irow, icol)
 Else
 Form1.MSChart1.Data = 0
 dataerror = True
 End If
 Next
 Next
 If dataerror Then
 GoTo invaliddatavalue
 Else
 Exit Sub
 End If
invaliddatavalue:
 MsgBox "Some Data values in the table are invalid or missing." & vbCrLf & "Please revise the entries", , "Chart Error"
End Sub

Private Sub Text1_Change()
' comit changes made to the TextBox control
' by copying its value to the active cell
    MSFlexGrid1.Text = Text1.Text
    If MSFlexGrid1.TextMatrix(rowcount - 1, 1) <> "" Then
    MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    rowcount = rowcount + 1
    End If
    
    If MSFlexGrid1.TextMatrix(rowcount - 2, 1) = "" Then
    If rowcount > 10 Then
    MSFlexGrid1.Rows = MSFlexGrid1.Rows - 1
    rowcount = rowcount - 1
    End If
    
    End If

End Sub
Private Sub Text1_keydown(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = 13 And Text1.Visible = True Then
       Text1.Visible = False
    ElseIf KeyCode = 40 And Text1.Visible = True Then
       KeyCode = 0
       If MSFlexGrid1.Row < (rowcount - 1) Then
       MSFlexGrid1.Row = MSFlexGrid1.Row + 1
       SetTextBox
       End If
    ElseIf KeyCode = 38 And Text1.Visible = True Then
       KeyCode = 0
       If MSFlexGrid1.Row > 1 Then
       MSFlexGrid1.Row = MSFlexGrid1.Row - 1
       SetTextBox
       End If
       
    ElseIf KeyCode = 37 And Text1.Visible = True And Text1.SelStart = 0 Then
       KeyCode = 0
       If MSFlexGrid1.Col > 1 Then
       MSFlexGrid1.Col = MSFlexGrid1.Col - 1
       SetTextBox
       End If
    ElseIf (KeyCode = 39 Or KeyCode = 9) _
    And Text1.Visible = True And _
    ((Text1.SelStart = 0 And Text1.SelLength = Len(Text1.Text)) _
    Or Text1.SelStart = Len(Text1.Text)) Then
       KeyCode = 0
       If MSFlexGrid1.Col < (colcount - 1) Then
       MSFlexGrid1.Col = MSFlexGrid1.Col + 1
       SetTextBox
       End If
    
       
    End If
  
End Sub
Private Sub Text1_LostFocus()
' When the TextBox loses the focus, the editing of the current cell ends
    Text1.Visible = False
    
 
End Sub

Sub SetTextBox()
Text1.BackColor = &HFFC0C0
If MSFlexGrid1.Col = 0 Or MSFlexGrid1.Row = 0 Then
Text1.BackColor = RGB(200, 200, 200)
Text1.ForeColor = RGB(0, 0, 0)
End If
' This subroutine resizes the TextBox according to the size
' of the current cell and places it on top of it, so that
' the user thinks he's editing a cell on the grid, not the TextBox
    ' Set the TextBox control's coordinates
    Text1.Left = MSFlexGrid1.Left + MSFlexGrid1.CellLeft
    Text1.Top = MSFlexGrid1.Top + MSFlexGrid1.CellTop
    ' Set the TextBox control's size
    Text1.Width = MSFlexGrid1.CellWidth
    Text1.Height = MSFlexGrid1.CellHeight
    ' copy the cell's contents to the TextBox and select it
    Text1.Text = MSFlexGrid1.Text
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    ' make the TextBox visible
    Text1.Visible = True
    ' and finally move the focus to the TextBox control
    Text1.SetFocus
    
    
End Sub


Private Sub EditClear_Click()
Dim irow As Integer, icol As Integer
Dim rowinc As Integer
Dim colinc As Integer
    For irow = MSFlexGrid1.Row To MSFlexGrid1.RowSel
        For icol = MSFlexGrid1.Col To MSFlexGrid1.ColSel
            MSFlexGrid1.TextMatrix(irow, icol) = ""
        Next
    Next
    
    rowinc = 0
    For irow = MSFlexGrid1.RowSel + 1 To MSFlexGrid1.Rows
        rowinc = rowinc + 1
        colinc = 0
        For icol = MSFlexGrid1.ColSel + 1 To MSFlexGrid1.Cols
            colinc = colinc + 1
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row + rowinc, MSFlexGrid1.Col + colinc) = MSFlexGrid1.TextMatrix(irow, icol)
        Next
    Next
    MSFlexGrid1.Rows = MSFlexGrid1.Row + MSFlexGrid1.Rows - MSFlexGrid1.RowSel
    
    'If MSFlexGrid1.RowSel >= (rowcount - 2) Then
    
    'If MSFlexGrid1.Row > 10 Then
    'MSFlexGrid1.Rows = MSFlexGrid1.Row
    
   ' Else
    'MSFlexGrid1.Rows = 10
    'End If
    'rowcount = MSFlexGrid1.Rows
    'MSFlexGrid1.Row = MSFlexGrid1.Row
    'SetTextBox
    'End If
    
End Sub

Private Sub EditCopy_Click()
Dim tmpText As String

    tmpText = MSFlexGrid1.Clip
    Clipboard.Clear
    Clipboard.SetText tmpText
End Sub

Private Sub EditCut_Click()
Dim tmpText As String

    tmpText = MSFlexGrid1.Clip
    Clipboard.Clear
    Clipboard.SetText tmpText
    EditClear_Click
End Sub

Private Sub EditPaste_Click()
Dim tmpText As String

    tmpText = Clipboard.GetText
    MSFlexGrid1.Clip = tmpText
End Sub

Private Sub EditSelect_Click()
    MSFlexGrid1.Row = 1
    MSFlexGrid1.Col = 1
    MSFlexGrid1.RowSel = MSFlexGrid1.Rows - 1
    MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1
    
End Sub

