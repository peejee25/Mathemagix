VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form optiondata 
   Caption         =   "Data Options"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2895
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1620
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   375
      Left            =   540
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Columns"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2295
      Begin ComCtl2.UpDown UpDown1 
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   873
         _Version        =   327681
         Value           =   2
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "No of Columns"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "optiondata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    tableplot.MSFlexGrid1.Cols = UpDown1.Value + 1
    tableplot.colcount = UpDown1.Value + 1
End Sub


Private Sub Command2_Click()
Unload optiondata
End Sub

Private Sub Form_Load()
UpDown1.Value = tableplot.colcount - 1
Label2.Caption = UpDown1.Value
End Sub

Private Sub UpDown1_Change()
Label2.Caption = UpDown1.Value
End Sub
