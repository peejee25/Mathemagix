VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form func2_form 
   Caption         =   "Form3"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   LinkTopic       =   "Form3"
   ScaleHeight     =   3645
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   2760
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PLOT"
      Height          =   375
      Left            =   1515
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   225
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Scale for parameter t"
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
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "< t <"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Z(t) = "
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1455
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Y(t) = "
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   855
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "X(t) = "
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "func2_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Command1_Click()
On Error Resume Next


If Text3.Text = "" Then Text3.Text = "0"
min = Text3.Text
If Text4.Text = "" Then Text4.Text = "0"
max = Text4.Text


inc = (max - min) / 50


For i = 0 To 50
t(i) = min + i * inc
Next


For i = 0 To 50
X(i) = evaluate_f(t(i), Text1.Text)
Next

For i = 0 To 50
y(i) = evaluate_f(t(i), Text2.Text)
Next

For i = 0 To 50
z(i) = evaluate_f(t(i), Text3.Text)
Next

max = X(0)
For i = 1 To 50
If X(i) > max Then
max = X(i)
End If
Next

mulx = 200 / max


max = y(0)
For i = 1 To 50
If y(i) > max Then
max = y(i)
End If
Next

muly = 200 / max

max = z(0)
For i = 1 To 50
If z(i) > max Then
max = z(i)
Debug.Print z(i)
End If
Next

mulz = 300 / max
draw_curve



End Sub

Private Function evaluate_f(var As Double, func As String) As Double
On Error Resume Next
      
    ScriptControl1.AddCode ("t=" & var)
       
    evaluate_f = ScriptControl1.Eval(func)
    
    
End Function



Private Sub Form_Load()
wdt = 789
ht = 637
centerx = wdt / 2
centery = ht / 2
xangle = 15
yangle = 15
End Sub

Public Function func3dto2d(x3 As Double, y3 As Double, z3 As Double, ByRef x2 As Integer, ByRef y2 As Integer)
Debug.Print "x3: " & x3 & " y3: " & y3 & " z3: " & z3
x2 = CInt(centerx - x3 * Cos(degtorad(xangle)) + y3 * Cos(degtorad(yangle)))
y2 = CInt(centery + x3 * Sin(degtorad(xangle)) + y3 * Sin(degtorad(yangle)) - z3)

End Function

Public Function degtorad(ByVal deg As Double) As Double
degtorad = (deg * 3.141592654) / 180
End Function

Public Sub draw_curve()
On Error Resume Next
Form1.Picture1.Cls

MoveToEx Form1.Picture1.hdc, centerx, centery, Module1.point
LineTo Form1.Picture1.hdc, centerx + 200, centery + 200 * Tan(degtorad(xangle))
MoveToEx Form1.Picture1.hdc, centerx, centery, Module1.point
LineTo Form1.Picture1.hdc, centerx - 200, centery + 200 * Tan(degtorad(yangle))
MoveToEx Form1.Picture1.hdc, centerx, centery, Module1.point
LineTo Form1.Picture1.hdc, centerx, centery - 300
Form1.Picture1.Refresh

func3dto2d (X(0) * mulx), (y(0) * muly), (z(0) * mulz), x2, y2

MoveToEx Form1.Picture1.hdc, x2, y2, Module1.point

For i = 1 To 50
x3 = X(i) * mulx
y3 = y(i) * muly
z3 = z(i) * mulz
func3dto2d x3, y3, z3, x2, y2
LineTo Form1.Picture1.hdc, x2, y2
MoveToEx Form1.Picture1.hdc, x2, y2, Module1.point
Next
Form1.Picture1.Refresh

End Sub

