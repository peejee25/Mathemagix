VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form2 
   Caption         =   "Function Input"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   LinkTopic       =   "Form2"
   ScaleHeight     =   5370
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   4200
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3398
      TabIndex        =   9
      Text            =   "10"
      Top             =   3960
      Width           =   780
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   998
      TabIndex        =   8
      Text            =   "0"
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PLOT"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   2580
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1980
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Function Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      Begin VB.OptionButton Option2 
         Caption         =   "Parametric Form  X = X(t)  ,  Y = Y(t)"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Explicit Form: Y= F(x)"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Scale for Independent variable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Max = "
      Height          =   255
      Left            =   2805
      TabIndex        =   11
      Top             =   4020
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Min = "
      Height          =   255
      Left            =   525
      TabIndex        =   10
      Top             =   4020
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Y(t) = "
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "F(x) = "
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim function_form As String




Private Sub Command1_Click()
'Start uncommenting here

'On Error Resume Next
'Dim min As Integer
'Dim max As Integer
'Dim inc As Double
'Dim t(101) As Double
'Dim X(101) As Double
'Dim y(101) As Double
'Dim i As Integer
'Dim j As Integer
'If Text3.Text = "" Then Text3.Text = "0"
'min = Text3.Text
'If Text4.Text = "" Then Text4.Text = "0"
'max = Text4.Text

'Form1.MSChart1.chartType = VtChChartType2dXY
'inc = (max - min) / 100
'If function_form = "explicit" Then
'For i = 0 To 100
'X(i) = min + i * inc
'Next

'For i = 0 To 100
'y(i) = evaluate_f(X(i), Text1.Text)
'Next

'Else
'For i = 0 To 100
't(i) = min + i * inc
'Next


'For i = 0 To 100
'X(i) = evaluate_f(t(i), Text1.Text)
'Next

'For i = 0 To 100
'y(i) = evaluate_f(t(i), Text2.Text)
'Next


'End If

'Form1.MSChart1.rowcount = 101
'Form1.MSChart1.ColumnCount = 2
'For j = 1 To 2
'For i = 1 To 101


'Form1.MSChart1.Row = i
'Form1.MSChart1.Column = j
'If j = 1 Then
'Form1.MSChart1.Data = X(i - 1)
'Else
'Form1.MSChart1.Data = y(i - 1)
'End If
'Next
'Next

'Form1.MSChart1.Visible = True


'Uncomment upto this to get plotting via chart
'--------------------------------------------
'--------------------------------------------
'This is manual plot

On Error GoTo funcerror
Form1.MSChart1.Visible = False
Dim min As Double, max As Double
Dim inc As Double
Dim t(101) As Double
Dim i As Integer
Dim j As Integer

If Text3.Text = "" Then Text3.Text = "0"
min = Text3.Text
If Text4.Text = "" Then Text4.Text = "0"
max = Text4.Text


inc = (max - min) / 100
If function_form = "explicit" Then
For i = 0 To 100
X(i) = min + i * inc
Next
For i = 0 To 100
y(i) = evaluate_f(X(i), Text1.Text)
Next

Else
For i = 0 To 100
t(i) = min + i * inc
Next

For i = 0 To 100
X(i) = evaluate_f(t(i), Text1.Text)
Next

For i = 0 To 100
y(i) = evaluate_f(t(i), Text2.Text)
Next


End If

plot2dcurve
Exit Sub
funcerror:
MsgBox "Error evaluating the function. Please Revise its syntax"
End Sub

Private Sub Form_Load()
function_form = "explicit"
Label2.Visible = False
Text2.Visible = False
End Sub

Private Sub Option1_Click()
    Label1.Caption = "F(x)= "
    Label2.Visible = False
    Text2.Visible = False
    function_form = "explicit"
End Sub

Private Sub Option2_Click()
    Label1.Caption = "X(t) = "
    Label2.Visible = True
    Text2.Visible = True
    
    function_form = "parametric"
End Sub

Private Function evaluate_f(var As Double, func As String) As Double
     
     If function_form = "explicit" Then
    ScriptControl1.AddCode ("x=" & var)
    Else
    ScriptControl1.AddCode ("t=" & var)
    End If
    'return the result
    evaluate_f = ScriptControl1.Eval(func)
    
    
End Function

Sub plot2dcurve(Optional autoscale As Boolean = True)
Dim scalegapx As Double, scalegapy As Double
Dim unitx As Double, unity As Double
Dim i As Integer, j As Integer
Dim temp As Double
Dim xmin As Double, xmax As Double
Dim ymin As Double, ymax As Double
Dim xscalemin As Double, xscalemax As Double
Dim yscalemin As Double, yscalemax As Double
Dim negox As Integer, negoy As Integer
Dim originxvalue As Double, originyvalue As Double
Form1.Picture1.Cls
Form1.Picture1.ForeColor = RGB(0, 0, 0)
On Error GoTo funcerror


If autoscale Then

'sort x-values and auto arrange y-values
'For i = 0 To 99
'For j = 0 To 99 - 1 - i
'If X(j) > X(j + 1) Then
'temp = X(j)
'X(j) = X(j + 1)
'X(j + 1) = temp

'temp = y(j)
'y(j) = y(j + 1)
'y(j + 1) = temp

'End If
'Next
'Next


'Find maximum and minimum x-values
xmax = X(0)
xmin = X(0)
For i = 1 To 100
If xmax < X(i) Then
xmax = X(i)
End If
If xmin > X(i) Then
xmin = X(i)
End If
Next


'Find maximum and minimum y-values
ymax = y(0)
ymin = y(0)
For i = 1 To 100
If ymax < y(i) Then
ymax = y(i)
End If
If ymin > y(i) Then
ymin = y(i)
End If
Next


'Compute maximum and minimum x and y scale values


xscalemin = CInt(xmin - 0.5)
xscalemax = CInt(xmax + 0.5)
yscalemin = CInt(ymin - 0.5)
yscalemax = CInt(ymax + 0.5)


'Find no. of negative x values in the scale
negox = 0
For i = 0 To 10
If (xscalemin + ((xscalemax - xscalemin) / 10) * i) < 0 Then
negox = negox + 1
End If
Next


'Find no. of negative y values in the scale

For i = 0 To 10
If (yscalemin + ((yscalemax - yscalemin) / 10) * i) < 0 Then
negoy = negoy + 1
End If
Next


'fix the coordinate system
minxcoord = 150
maxxcoord = 550
minycoord = 50
maxycoord = 450
centerx = (minxcoord + maxxcoord) / 2
centery = (minycoord + maxycoord) / 2


scalegapx = (maxxcoord - minxcoord) / 10
scalegapy = (maxycoord - minycoord) / 10


'unit distance in the coordinate system
unitx = (scalegapx) / ((xscalemax - xscalemin) / 10)
unity = (scalegapy) / ((yscalemax - yscalemin) / 10)

Dim originx As Integer, originy As Integer
originx = minxcoord + negox * scalegapx
originy = maxycoord - negoy * scalegapy
originxvalue = ((originx - minxcoord) / scalegapx) * ((xscalemax - xscalemin) / 10) + xscalemin
originyvalue = ((maxycoord - originy) / scalegapy) * ((yscalemax - yscalemin) / 10) + yscalemin


'draw x-axis
If 0 > yscalemin And 0 < yscalemax Then
MoveToEx Form1.Picture1.hdc, minxcoord, maxycoord - negoy * scalegapy + (originyvalue * unity), Module1.point
LineTo Form1.Picture1.hdc, maxxcoord, maxycoord - negoy * scalegapy + (originyvalue * unity)
Else
MoveToEx Form1.Picture1.hdc, minxcoord, maxycoord - negoy * scalegapy, Module1.point
LineTo Form1.Picture1.hdc, maxxcoord, maxycoord - negoy * scalegapy
End If
'draw y-axis
If 0 > xscalemin And 0 < xscalemax Then
MoveToEx Form1.Picture1.hdc, minxcoord + negox * scalegapx - (originxvalue * unitx), minycoord, Module1.point
LineTo Form1.Picture1.hdc, minxcoord + negox * scalegapx - (originxvalue * unitx), maxycoord
Else
MoveToEx Form1.Picture1.hdc, minxcoord + negox * scalegapx, minycoord, Module1.point
LineTo Form1.Picture1.hdc, minxcoord + negox * scalegapx, maxycoord
End If

'Draw graph

Form1.Picture1.ForeColor = RGB(0, 0, 255)


MoveToEx Form1.Picture1.hdc, originx + (X(0) - originxvalue) * (unitx), originy - ((y(0) - originyvalue) * unity), Module1.point


For i = 1 To 100
LineTo Form1.Picture1.hdc, originx + (X(i) - originxvalue) * (unitx), originy - ((y(i) - originyvalue) * unity)
Next

'draw x-axis ,yaxis labels

Form1.Picture1.ForeColor = RGB(0, 0, 0)
For i = 0 To 10
'x-axis label
If 0 > yscalemin And 0 < yscalemax Then
MoveToEx Form1.Picture1.hdc, minxcoord + scalegapx * i, maxycoord - negoy * scalegapy - 2 + (originyvalue * unity), Module1.point
LineTo Form1.Picture1.hdc, minxcoord + scalegapx * i, maxycoord - negoy * scalegapy + 4 + (originyvalue * unity)

Form1.Picture1.CurrentX = minxcoord + scalegapx * i - 6
Form1.Picture1.CurrentY = maxycoord - negoy * scalegapy + 5 + (originyvalue * unity)
Form1.Picture1.Print xscalemin + ((xscalemax - xscalemin) / 10) * i
Else
MoveToEx Form1.Picture1.hdc, minxcoord + scalegapx * i, maxycoord - negoy * scalegapy - 2, Module1.point
LineTo Form1.Picture1.hdc, minxcoord + scalegapx * i, maxycoord - negoy * scalegapy + 4

Form1.Picture1.CurrentX = minxcoord + scalegapx * i - 6
Form1.Picture1.CurrentY = maxycoord - negoy * scalegapy + 5
Form1.Picture1.Print xscalemin + ((xscalemax - xscalemin) / 10) * i
End If
'y-axis label
If 0 > xscalemin And 0 < xscalemax Then
MoveToEx Form1.Picture1.hdc, minxcoord + negox * scalegapx - 2 - (originxvalue * unitx), minycoord + scalegapy * i, Module1.point
LineTo Form1.Picture1.hdc, minxcoord + negox * scalegapx + 4 - (originxvalue * unitx), minycoord + scalegapy * i

Form1.Picture1.CurrentX = minxcoord + negox * scalegapx + 4 - (originxvalue * unitx)
Form1.Picture1.CurrentY = minycoord + scalegapy * i - 4
Form1.Picture1.Print yscalemin + ((yscalemax - yscalemin) / 10) * (10 - i)
Else
MoveToEx Form1.Picture1.hdc, minxcoord + negox * scalegapx - 2, minycoord + scalegapy * i, Module1.point
LineTo Form1.Picture1.hdc, minxcoord + negox * scalegapx + 4, minycoord + scalegapy * i

Form1.Picture1.CurrentX = minxcoord + negox * scalegapx + 4
Form1.Picture1.CurrentY = minycoord + scalegapy * i - 4
Form1.Picture1.Print yscalemin + ((yscalemax - yscalemin) / 10) * (10 - i)


End If

Next


End If

Exit Sub
funcerror:
MsgBox "Error evaluating the function"
End Sub

Private Sub ScriptControl1_Error()
    Debug.Print ScriptControl1.Error.Number
    Debug.Print ScriptControl1.Error.Text
End Sub

