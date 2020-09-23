VERSION 5.00
Begin VB.Form frmTVSOEPS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Three Variable System Of Equations Solver"
   ClientHeight    =   6660
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7470
   Icon            =   "TVEPS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtXequals 
      Height          =   285
      Left            =   1680
      TabIndex        =   34
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3720
      TabIndex        =   33
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtX3 
      Height          =   285
      Left            =   1320
      TabIndex        =   27
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox txtz3 
      Height          =   285
      Left            =   1320
      TabIndex        =   26
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox txty3 
      Height          =   285
      Left            =   1320
      TabIndex        =   25
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox txtconstant3 
      Height          =   285
      Left            =   1320
      TabIndex        =   24
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox txtX2 
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtZ2 
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txtY2 
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtConstant2 
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtYequals 
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox txtZequals 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Solve"
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox txtconstant 
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtZ 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblequation 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3480
      TabIndex        =   40
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label Label14 
      Caption         =   "Equation 3:"
      Height          =   255
      Left            =   4800
      TabIndex        =   39
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Equation 2:"
      Height          =   255
      Left            =   4800
      TabIndex        =   38
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Equation1:"
      Height          =   255
      Left            =   4800
      TabIndex        =   37
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblequation3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3480
      TabIndex        =   36
      Top             =   3000
      Width           =   3855
   End
   Begin VB.Label lblequation2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3480
      TabIndex        =   35
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label11 
      Caption         =   "Equation 3"
      Height          =   255
      Left            =   1440
      TabIndex        =   32
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "X Coefficient"
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Y Coefficient"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Z Coefficient"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Constant"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Equation 2"
      Height          =   255
      Left            =   1440
      TabIndex        =   23
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "X Coefficient"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Y Coefficient"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Z Coefficient"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Constant"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Equation 1"
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblZequals 
      Caption         =   "Z = "
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label lblyequals 
      Caption         =   "Y = "
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label lblxequals 
      Caption         =   "X = "
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label lblconstant 
      Caption         =   "Constant"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblZ 
      Caption         =   "Z Coefficient"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblY 
      Caption         =   "Y Coefficient"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblx 
      Caption         =   "X Coefficient"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Menu mnmore 
      Caption         =   "More"
      Begin VB.Menu mnuquad 
         Caption         =   "Solve Quadratic Equations"
      End
      Begin VB.Menu mnutwo 
         Caption         =   "Two Variable Problem Solver"
      End
      Begin VB.Menu mnucalc 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnugraph 
         Caption         =   "Graph"
      End
   End
   Begin VB.Menu func 
      Caption         =   "Functions"
      Begin VB.Menu mnusolve 
         Caption         =   "Solve Equation"
      End
      Begin VB.Menu mnclear 
         Caption         =   "Clear All Entries"
      End
      Begin VB.Menu mnuequations 
         Caption         =   "Only Show Equations"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu mnuhelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmTVSOEPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
A = Val(txtX.Text)
a2 = Val(txtX2.Text)
a3 = Val(txtX3.Text)
b = Val(txtY.Text)
b2 = Val(txtY2.Text)
b3 = Val(txty3.Text)
c = Val(txtZ.Text)
c2 = Val(txtZ2.Text)
c3 = Val(txtz3.Text)
d = Val(txtconstant.Text)
d2 = Val(txtConstant2.Text)
d3 = Val(txtconstant3.Text)

txtXequals.Text = ((d * b2 * c3) + (b * c2 * d3) + (c * d2 * b3) - (d3 * b2 * c) - (b3 * c2 * d) - (c3 * d2 * b)) / ((A * b2 * c3) + (b * c2 * a3) + (c * a2 * b3) - (a3 * b2 * c) - (b3 * c2 * A) - (c3 * a2 * b))

txtYequals.Text = ((A * d2 * c3) + (d * c2 * a3) + (c * a2 * d3) - (a3 * d2 * c) - (d3 * c2 * A) - (c3 * a2 * d)) / ((A * b2 * c3) + (b * c2 * a3) + (c * a2 * b3) - (a3 * b2 * c) - (b3 * c2 * A) - (c3 * a2 * b))

txtZequals.Text = ((A * b2 * d3) + (b * d2 * a3) + (d * a2 * b3) - (a3 * b2 * d) - (b3 * d2 * A) - (d3 * a2 * b)) / ((A * b2 * c3) + (b * c2 * a3) + (c * a2 * b3) - (a3 * b2 * c) - (b3 * c2 * A) - (c3 * a2 * b))

lblequation.Caption = txtX.Text + "X" + " + " + txtY.Text + "Y" + " + " + txtZ.Text + "Z" + " " + "=" + " " + txtconstant.Text
lblequation2.Caption = txtX2.Text + "X" + " + " + txtY2.Text + "Y" + " + " + txtZ2.Text + "Z" + " " + "=" + " " + txtConstant2.Text
lblequation3.Caption = txtX3.Text + "X" + " + " + txty3.Text + "Y" + " + " + txtz3.Text + "Z" + " " + "=" + " " + txtconstant3.Text


End Sub

Private Sub Command2_Click()
txtX.Text = " "
txtX2.Text = " "
txtX3.Text = " "
txtY.Text = " "
txtY2.Text = " "
txty3.Text = " "
txtZ.Text = " "
txtZ2.Text = " "
txtz3.Text = " "
txtconstant.Text = " "
txtConstant2.Text = " "
txtconstant3.Text = " "
txtXequals.Text = " "
txtYequals.Text = " "
txtZequals.Text = " "
lblequation.Caption = " "
lblequation2.Caption = " "
lblequation3.Caption = " "
End Sub

Private Sub mnclear_Click()
txtX.Text = " "
txtX2.Text = " "
txtX3.Text = " "
txtY.Text = " "
txtY2.Text = " "
txty3.Text = " "
txtZ.Text = " "
txtZ2.Text = " "
txtz3.Text = " "
txtconstant.Text = " "
txtConstant2.Text = " "
txtconstant3.Text = " "
txtXequals.Text = " "
txtYequals.Text = " "
txtZequals.Text = " "
lblequation.Caption = " "
lblequation2.Caption = " "
lblequation3.Caption = " "
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnucalc_Click()
frmcalc.Show
End Sub

Private Sub mnuequations_Click()
lblequation.Caption = txtX.Text + "X" + " + " + txtY.Text + "Y" + " + " + txtZ.Text + "Z" + " " + "=" + " " + txtconstant.Text
lblequation2.Caption = txtX2.Text + "X" + " + " + txtY2.Text + "Y" + " + " + txtZ2.Text + "Z" + " " + "=" + " " + txtConstant2.Text
lblequation3.Caption = txtX3.Text + "X" + " + " + txty3.Text + "Y" + " + " + txtz3.Text + "Z" + " " + "=" + " " + txtconstant3.Text

End Sub

Private Sub mnugraph_Click()
frmgraph.Show
End Sub

Private Sub mnuquad_Click()
frmquad.Show
End Sub

Private Sub mnusolve_Click()
On Error Resume Next
A = Val(txtX.Text)
a2 = Val(txtX2.Text)
a3 = Val(txtX3.Text)
b = Val(txtY.Text)
b2 = Val(txtY2.Text)
b3 = Val(txty3.Text)
c = Val(txtZ.Text)
c2 = Val(txtZ2.Text)
c3 = Val(txtz3.Text)
d = Val(txtconstant.Text)
d2 = Val(txtConstant2.Text)
d3 = Val(txtconstant3.Text)

txtXequals.Text = ((d * b2 * c3) + (b * c2 * d3) + (c * d2 * b3) - (d3 * b2 * c) - (b3 * c2 * d) - (c3 * d2 * b)) / ((A * b2 * c3) + (b * c2 * a3) + (c * a2 * b3) - (a3 * b2 * c) - (b3 * c2 * A) - (c3 * a2 * b))

txtYequals.Text = ((A * d2 * c3) + (d * c2 * a3) + (c * a2 * d3) - (a3 * d2 * c) - (d3 * c2 * A) - (c3 * a2 * d)) / ((A * b2 * c3) + (b * c2 * a3) + (c * a2 * b3) - (a3 * b2 * c) - (b3 * c2 * A) - (c3 * a2 * b))

txtZequals.Text = ((A * b2 * d3) + (b * d2 * a3) + (d * a2 * b3) - (a3 * b2 * d) - (b3 * d2 * A) - (d3 * a2 * b)) / ((A * b2 * c3) + (b * c2 * a3) + (c * a2 * b3) - (a3 * b2 * c) - (b3 * c2 * A) - (c3 * a2 * b))

lblequation.Caption = txtX.Text + "X" + " + " + txtY.Text + "Y" + " + " + txtZ.Text + "Z" + " " + "=" + " " + txtconstant.Text
lblequation2.Caption = txtX2.Text + "X" + " + " + txtY2.Text + "Y" + " + " + txtZ2.Text + "Z" + " " + "=" + " " + txtConstant2.Text
lblequation3.Caption = txtX3.Text + "X" + " + " + txty3.Text + "Y" + " + " + txtz3.Text + "Z" + " " + "=" + " " + txtconstant3.Text

End Sub

Private Sub mnutwo_Click()
frmtwovariable.Show
End Sub
