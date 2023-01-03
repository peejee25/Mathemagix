Attribute VB_Name = "Module1"
Option Explicit
Public plottype As String

Public Type POINTAPI
        X As Long
        y As Long
End Type
Public point As POINTAPI

Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long

Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public i As Integer
Public j As Integer
Public t(101) As Double
Public X(101) As Double
Public y(101) As Double
Public z(101) As Double
Public mulx As Single, muly As Single, mulz As Single
Public min As Integer
Public max As Integer
Public inc As Double

Public x2 As Integer, y2 As Integer
Public x3 As Double, y3 As Double, z3 As Double
Public wdt As Integer
Public ht As Integer
Public centerx As Double, centery As Double
Public maxxcoord As Double, minxcoord As Double
Public maxycoord As Double, minycoord As Double
Public xangle As Single, yangle As Single

