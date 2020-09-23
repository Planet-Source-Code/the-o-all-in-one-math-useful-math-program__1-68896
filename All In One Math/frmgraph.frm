VERSION 5.00
Begin VB.Form frmgraph 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graphing"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   Icon            =   "frmgraph.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Graph Midpoint"
      Height          =   495
      Left            =   8880
      TabIndex        =   29
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtmidy 
      Height          =   285
      Left            =   8400
      TabIndex        =   27
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txtmidx 
      Height          =   285
      Left            =   7920
      TabIndex        =   26
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txtdistance 
      Height          =   285
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtslope 
      Height          =   285
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox txtscaley2 
      Height          =   285
      Left            =   1800
      TabIndex        =   16
      ToolTipText     =   "Bottom Y value"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtscalex2 
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      ToolTipText     =   "Right X value"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtscaley1 
      Height          =   285
      Left            =   480
      TabIndex        =   14
      ToolTipText     =   "Top Y value"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtscalex1 
      Height          =   285
      Left            =   0
      TabIndex        =   13
      ToolTipText     =   "Left X Value"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw line"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtY2 
      Height          =   285
      Left            =   4920
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtX2 
      Height          =   285
      Left            =   4200
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtY1 
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtX1 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.PictureBox picoutput 
      Height          =   6615
      Left            =   0
      ScaleHeight     =   6555
      ScaleWidth      =   9795
      TabIndex        =   0
      Top             =   1200
      Width           =   9855
   End
   Begin VB.Label Label15 
      Caption         =   ","
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   28
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label14 
      Caption         =   "Midpoint"
      Height          =   255
      Left            =   7080
      TabIndex        =   25
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "Distance:"
      Height          =   255
      Left            =   7680
      TabIndex        =   23
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "Slope:"
      Height          =   255
      Left            =   7920
      TabIndex        =   21
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "Y2"
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "X2"
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "Y1"
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "X1"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "Scale"
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Y"
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "X"
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "Y"
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "X"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "Second Ordered Pair"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "First Ordered Pair"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmgraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
X1 = Val(txtX1.Text)
X2 = Val(txtX2.Text)
Y1 = Val(txtY1.Text)
Y2 = Val(txtY2.Text)
A = Val(txtscalex1.Text)
a2 = Val(txtscalex2.Text)
b = Val(txtscaley1.Text)
b2 = Val(txtscaley2.Text)
picoutput.Cls
picoutput.Scale (-A, b)-(a2, -b2)
picoutput.Line (-100, 0)-(100, 0)
picoutput.Line (0, -100)-(0, 100)
picoutput.Line (X1, Y1)-(X2, Y2)
picoutput.Circle (X1, Y1), 0.1
picoutput.Circle (X2, Y2), 0.1
If X2 - X1 = 0 Then
txtslope.Text = "Undefined"
Else
txtslope.Text = (Y2 - Y1) / (X2 - X1)
End If
txtdistance.Text = Sqr(((X2 - X1) ^ 2) + ((Y2 - Y1) ^ 2))
txtmidx.Text = ((X1 + X2) / 2)
txtmidy.Text = ((Y1 + Y2) / 2)
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
x = Val(txtmidx.Text)
y = Val(txtmidy.Text)
A = Val(txtscalex1.Text)
a2 = Val(txtscalex2.Text)
b = Val(txtscaley1.Text)
b2 = Val(txtscaley2.Text)
picoutput.Circle (x, y), 0.1
Command2.Enabled = False
End Sub
