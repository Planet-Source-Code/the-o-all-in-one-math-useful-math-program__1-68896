VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "All In One Math"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstmath 
      Height          =   1620
      ItemData        =   "frmmain.frx":0BC2
      Left            =   600
      List            =   "frmmain.frx":0BD5
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Double click which math application you would like to use...."
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lstmath_DblClick()
If lstmath.Text = "Three Variable System Of Equations Solver" Then
frmTVSOEPS.Show
End If
If lstmath.Text = "Two Variable System Of Equations Solver" Then
frmtwovariable.Show
End If
If lstmath.Text = "Quadratic Equation Solver" Then
frmquad.Show
End If
If lstmath.Text = "Graph" Then
frmgraph.Show
End If
If lstmath.Text = "Simple Calculator" Then
frmcalc.Show
End If
End Sub
