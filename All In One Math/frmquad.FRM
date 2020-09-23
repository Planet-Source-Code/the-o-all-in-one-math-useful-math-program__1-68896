VERSION 5.00
Begin VB.Form frmquad 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quadratic Equation Solver"
   ClientHeight    =   3105
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   4260
   FillColor       =   &H80000000&
   Icon            =   "frmquad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtValues 
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtValues 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtValues 
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdSolve 
      Caption         =   "&Solve!"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      ToolTipText     =   "Click to Solve your Equation"
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "= 0"
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblc1 
      Caption         =   "c"
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "x +"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblb1 
      Caption         =   "b"
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblX 
      Caption         =   "x2 + "
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Type in your constants"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Solution"
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblSolution 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   360
      TabIndex        =   9
      Top             =   1560
      Width           =   3705
   End
   Begin VB.Label lblEquation 
      AutoSize        =   -1  'True
      Caption         =   "Equation: "
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   2760
      Width           =   720
   End
   Begin VB.Label lblA 
      Caption         =   "&a:"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   720
      Width           =   255
   End
   Begin VB.Label lblB 
      Caption         =   "&b:"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   720
      Width           =   255
   End
   Begin VB.Label lblC 
      Caption         =   "&c:"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.Label lbla1 
      Caption         =   "a"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   2760
      Width           =   255
   End
End
Attribute VB_Name = "frmquad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim A As Single
Dim b As Single
Dim c As Single
Dim X1 As Single
Dim X2 As Single
Dim Discriminant As Single
Dim Answer As String

Private Sub Form_Load()
lblX.Caption = "x " + Chr(178)
Call initZero
End Sub








Private Sub initZero()
X1 = 0#
X2 = 0#
Discriminant = 0#
A = 0#
b = 0#
c = 0#
End Sub
Private Sub cmdsolve_Click()
lblSolution.Caption = Answer & vbCrLf & "Discriminant = " & Discriminant
Dim i As Integer
i = 0
For i = 0 To 2
txtValues(i).Text = ""
Next i
lbla1.Caption = A
lblb1.Caption = b
lblc1.Caption = c
End Sub

Private Sub Label3_Click()

End Sub

Private Sub txtValues_Change(Index As Integer)
If Trim(txtValues(0).Text) = "" Or Trim(txtValues(1).Text) = "" Or Trim(txtValues(2).Text) = "" Then
    cmdSolve.Enabled = False
Else
    If Trim(txtValues(0).Text) = "0" Then
        cmdSolve.Enabled = False
        lblSolution.Caption = "The a value cannot equal 0"
    Else
        If (IsNumeric(txtValues(0)) = True And IsNumeric(txtValues(1)) = True And IsNumeric(txtValues(2)) = True) Then
            A = txtValues(0)
            b = txtValues(1)
            c = txtValues(2)
            cmdSolve.Enabled = True
            Call Calculate
        End If
    End If
End If
End Sub

Private Sub Calculate()
Discriminant = b * b - 4 * A * c

If (Discriminant < 0) Then
    Answer = CStr(-b) & " + " & CStr(Sqr(Abs(Discriminant))) & " i" & " / " & CStr(2 * A) & vbCrLf & CStr(-b) & " - " & CStr(Sqr(Abs(Discriminant))) & " i" & " / " & CStr(2 * A)
Else
    X1 = (-b + Sqr(Discriminant)) / (2 * A)
    X2 = (-b - Sqr(Discriminant)) / (2 * A)
    Answer = "X1 = " & CStr(X1) & vbCrLf & "X2 = " & CStr(X2)
End If
End Sub
