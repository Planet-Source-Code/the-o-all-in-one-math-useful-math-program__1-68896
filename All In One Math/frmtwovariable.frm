VERSION 5.00
Begin VB.Form frmtwovariable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Two Variable System Of Equations Solver"
   ClientHeight    =   4815
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5865
   Icon            =   "frmtwovariable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtc2 
      Height          =   285
      Left            =   1080
      TabIndex        =   22
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtc 
      Height          =   285
      Left            =   1080
      TabIndex        =   21
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtYequals 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtXequals 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdsolve 
      Caption         =   "Solve"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtX2 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtY2 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Constant"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Constant"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblequation2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label lblequation 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label10 
      Caption         =   "Equation 2"
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Equation 1"
      Height          =   255
      Left            =   3720
      TabIndex        =   16
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Y ="
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "X ="
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Y Coefficient"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "X Coefficient"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Y Coefficient"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "X Coefficient"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Equation 2"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Equation 1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Menu mnufunctions 
      Caption         =   "Functions"
      Begin VB.Menu mnusolve 
         Caption         =   "Solve"
      End
      Begin VB.Menu mnuclear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnushow 
         Caption         =   "Show Equations Only"
      End
   End
End
Attribute VB_Name = "frmtwovariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclear_Click()
On Error Resume Next
txtX.Text = ""
txtX2.Text = ""
txtY.Text = ""
txtY2.Text = ""
txtc.Text = ""
txtc2.Text = ""
txtXequals.Text = ""
txtYequals.Text = ""
lblequation.Caption = ""
 lblequation2.Caption = ""
 
End Sub

Private Sub cmdgraph_Click()

End Sub

Private Sub cmdsolve_Click()
On Error Resume Next
A = Val(txtX.Text)
a2 = Val(txtX2.Text)
b = Val(txtY.Text)
b2 = Val(txtY2.Text)
c = Val(txtc.Text)
c2 = Val(txtc2.Text)

txtXequals.Text = ((c * b2) - (c2 * b)) / ((A * b2) - (a2 * b))

txtYequals.Text = ((A * c2) - (a2 * c)) / ((A * b2) - (a2 * b))
lblequation.Caption = txtX.Text + "X" + " + " + txtY.Text + "Y " + "= " + txtc.Text
lblequation2.Caption = txtX2.Text + "X" + " + " + txtY2.Text + "Y " + "= " + txtc2.Text

End Sub

Private Sub Command1_Click()

End Sub

Private Sub mnuclear_Click()
On Error Resume Next
txtX.Text = ""
txtX2.Text = ""
txtY.Text = ""
txtY2.Text = ""
txtc.Text = ""
txtc2.Text = ""
txtXequals.Text = ""
txtYequals.Text = ""
lblequation.Caption = ""
 lblequation2.Caption = ""
End Sub

Private Sub mnushow_Click()
On Error Resume Next
lblequation.Caption = txtX.Text + "X" + " + " + txtY.Text + "Y " + "= " + txtc.Text
lblequation2.Caption = txtX2.Text + "X" + " + " + txtY2.Text + "Y " + "= " + txtc2.Text
End Sub

Private Sub mnusolve_Click()
On Error Resume Next
A = Val(txtX.Text)
a2 = Val(txtX2.Text)
b = Val(txtY.Text)
b2 = Val(txtY2.Text)
c = Val(txtc.Text)
c2 = Val(txtc2.Text)

txtXequals.Text = ((c * b2) - (c2 * b)) / ((A * b2) - (a2 * b))

txtYequals.Text = ((A * c2) - (a2 * c)) / ((A * b2) - (a2 * b))
lblequation.Caption = txtX.Text + "X" + " + " + txtY.Text + "Y " + "= " + txtc.Text
lblequation2.Caption = txtX2.Text + "X" + " + " + txtY2.Text + "Y " + "= " + txtc2.Text
End Sub
