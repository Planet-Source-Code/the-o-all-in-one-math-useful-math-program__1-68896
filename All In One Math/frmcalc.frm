VERSION 5.00
Begin VB.Form frmcalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Calculator"
   ClientHeight    =   5265
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   5520
   Icon            =   "frmcalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmcalc.frx":0BC2
   ScaleHeight     =   5265
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3720
      TabIndex        =   7
      ToolTipText     =   "Divide the Top Number by the Bottom Number"
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3720
      TabIndex        =   5
      ToolTipText     =   "Subtract the Bottom Number fromt the Top Number"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H8000000E&
      Caption         =   "Tan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   27
      ToolTipText     =   "Calculates the "
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command21 
      Caption         =   "x^y"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   26
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Sin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   21
      ToolTipText     =   "Calculates the Sine of the top number"
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   20
      ToolTipText     =   "Calculates the logarithms of the top number"
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Cos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   19
      ToolTipText     =   "Calculates the cosine of the top number"
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Bottom +/-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   16
      ToolTipText     =   "Turns the bottom number into either a negative or positive number"
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Top +/-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   15
      ToolTipText     =   "Turns the top number into either a negative or positive number"
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Insert Pi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      ToolTipText     =   "Inserts pi (3.14) into the top number"
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear Answer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      ToolTipText     =   "Clears the answer"
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "<>?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   13
      ToolTipText     =   "Find out if the top number is greater than or less than the bottom number"
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Clear All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      ToolTipText     =   "Clears both the top and the bottom numbers"
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Clear Bottom Textbox"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      ToolTipText     =   "clears the bottom number"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Clear Top Textbox"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   9
      ToolTipText     =   "Clears the top number"
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "/ with R"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4320
      TabIndex        =   14
      ToolTipText     =   "Divide the Top Number by the Bottom Number with remainders"
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3120
      TabIndex        =   6
      ToolTipText     =   "The Top Number Multiplied by the Bottom Number"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "This is the bottom number"
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3120
      TabIndex        =   1
      ToolTipText     =   "Add the Top Number and the Bottom Number Together"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "This is the top number"
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton Command20 
      Caption         =   "x^3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   25
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command19 
      Caption         =   "x^2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   24
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   18
      ToolTipText     =   "Find the 1 over X of the Top Textbox"
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Find Square Root"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   17
      ToolTipText     =   "Finds the Square Root of the Top Textbox"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer"
      Height          =   255
      Left            =   2040
      TabIndex        =   28
      Top             =   1200
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   5520
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   5520
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Simple Functions"
      Height          =   255
      Left            =   3360
      TabIndex        =   23
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced Functions"
      Height          =   255
      Left            =   1560
      TabIndex        =   22
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "+ - / * + - * / + - * / + - * / + - *"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.Menu Funt 
      Caption         =   "Functions"
      Begin VB.Menu SimpFunct 
         Caption         =   "Simple Functions"
         Begin VB.Menu Add 
            Caption         =   "Addition"
         End
         Begin VB.Menu Sub 
            Caption         =   "Subtraction"
         End
         Begin VB.Menu Div 
            Caption         =   "Division"
         End
         Begin VB.Menu Mulitiply 
            Caption         =   "Multiplication"
         End
         Begin VB.Menu Divwithr 
            Caption         =   "Division with Remainders"
         End
         Begin VB.Menu Lessormore 
            Caption         =   "<>?"
         End
      End
      Begin VB.Menu AdvFunct 
         Caption         =   "Advance Functions"
         Begin VB.Menu SquareRoot 
            Caption         =   "Square Root"
         End
         Begin VB.Menu oneoverx 
            Caption         =   "1/x"
         End
         Begin VB.Menu xsquared 
            Caption         =   "x^2"
         End
         Begin VB.Menu xcubed 
            Caption         =   "x^3 "
         End
         Begin VB.Menu xtoy 
            Caption         =   "x^y"
         End
         Begin VB.Menu Coss 
            Caption         =   "Cos"
         End
         Begin VB.Menu Logg 
            Caption         =   "Log"
         End
         Begin VB.Menu Sinn 
            Caption         =   "Sin"
         End
         Begin VB.Menu Tann 
            Caption         =   "Tan"
         End
      End
      Begin VB.Menu otherfunct 
         Caption         =   "Other Functions"
         Begin VB.Menu Clear 
            Caption         =   "Clear"
            Begin VB.Menu ClearTop 
               Caption         =   "Clear Top Textbox"
            End
            Begin VB.Menu ClearBottom 
               Caption         =   "Clear Bottom Textbox"
            End
            Begin VB.Menu ClearAnsw 
               Caption         =   "Clear Answer"
            End
            Begin VB.Menu ClearAll 
               Caption         =   "Clear All"
            End
         End
         Begin VB.Menu negorpos 
            Caption         =   "+/-"
            Begin VB.Menu Topnegorpos 
               Caption         =   "Top +/-"
            End
            Begin VB.Menu Bottomnegorpos 
               Caption         =   "Bottom +/-"
            End
         End
         Begin VB.Menu insertpi 
            Caption         =   "Inser Pi"
         End
      End
   End
End
Attribute VB_Name = "frmcalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Add_Click()
Command1_Click
End Sub

Private Sub Command1_Click()

y = Val(Text1.Text)
x = Val(Text2.Text)
Label2.Caption = y + x
End Sub


Private Sub Command10_Click()
y = Val(Text1.Text)
x = Val(Text2.Text)
If y > x Then
Label2.Caption = frmcalc.Text1.Text + " > " + frmcalc.Text2.Text
End If
If y < x Then
Label2.Caption = frmcalc.Text1.Text + " < " + frmcalc.Text2.Text
End If
If y = x Then
Label2.Caption = frmcalc.Text1.Text + " = " + frmcalc.Text2.Text
End If
End Sub

Private Sub Command11_Click()
On Error Resume Next
Dim divisor As Single, dividend As Single
Dim quotient As Single, remainder As Single
divisor = Val(Text1.Text)
dividend = Val(Text2.Text)
If dividend = 0 Then
MsgBox "You can't divide by zero!"
Label2.Caption = "You can't divide by zero!"
Else
quotient = Int(dividend / divisor)
remainder = dividend - quotient * divisor
Label2.Caption = quotient & " R " & remainder
End If
End Sub

Private Sub Command12_Click()
    Text1.Text = -Val(Text1.Text)
End Sub

Private Sub Command13_Click()
   Text2.Text = -Val(Text2.Text)
End Sub





Private Sub Command14_Click()
   If Val(Text1.Text) <> 0 Then Label2.Caption = 1 / Val(Text1.Text)
    End Sub

Private Sub Command15_Click()
On Error Resume Next
Label2.Caption = Cos(Val(Text1.Text))
End Sub

Private Sub Command16_Click()
On Error Resume Next
If Text1.Text < 0 Then
Label2.Caption = "You can't find the square root of a negative number"
Else
Label2.Caption = Sqr(Val(Text1.Text))
End If
End Sub

Private Sub Command17_Click()
On Error Resume Next
Label2.Caption = Log(Val(Text1.Text))
End Sub

Private Sub Command18_Click()
On Error Resume Next
Label2.Caption = Sin(Val(Text1.Text))
End Sub

Private Sub Command19_Click()
y = Val(Text1.Text)
Label2.Caption = y ^ 2
End Sub

Private Sub Command2_Click()
y = Val(Text1.Text)
x = Val(Text2.Text)
Label2.Caption = y - x
End Sub

Private Sub Command20_Click()
y = Val(Text1.Text)
Label2.Caption = y ^ 3
End Sub

Private Sub Command21_Click()
y = Val(Text1.Text)
x = Val(Text2.Text)
Label2.Caption = y ^ x
End Sub

Private Sub Command22_Click()
On Error Resume Next
Label2.Caption = Tan(Val(Text1.Text))
End Sub

Private Sub Command3_Click()
y = Val(Text1.Text)
x = Val(Text2.Text)
Label2.Caption = y * x
End Sub

Private Sub Command4_Click()
On Error Resume Next
y = Val(Text1.Text)
x = Val(Text2.Text)
If x = 0 Then
MsgBox "You can't divide by zero"
Label2.Caption = "You can't divide by zero!"
End If
Label2.Caption = y / x
End Sub

Private Sub Command5_Click()
Label2.Caption = ""
End Sub

Private Sub Command6_Click()
Text1.Text = ""
End Sub

Private Sub Command7_Click()
Text2.Text = ""
End Sub

Private Sub Command8_Click()
Text1.Text = ""
Label2.Caption = ""
Text2.Text = ""
End Sub

Private Sub Command9_Click()
Text1.Text = "3.1415927"
End Sub


Private Sub Coss_Click()
Command15_Click
End Sub

Private Sub Div_Click()
Command4_Click
End Sub

Private Sub Divwithr_Click()
Command11_Click
End Sub

Private Sub Form_Load()
Command19.Caption = "x" + Chr(178)
Command20.Caption = "x" + Chr(179)
End Sub

Private Sub Lessormore_Click()
Command10_Click
End Sub

Private Sub Logg_Click()
Command17_Click
End Sub

Private Sub Mulitiply_Click()
Command3_Click
End Sub

Private Sub oneoverx_Click()
Command14_Click
End Sub

Private Sub Sinn_Click()
Command18_Click
End Sub

Private Sub SquareRoot_Click()
Command16_Click
End Sub

Private Sub Sub_Click()
Command2_Click
End Sub

Private Sub Tann_Click()
Command22_Click
End Sub

Private Sub xcubed_Click()
Command20_Click
End Sub

Private Sub xsquared_Click()
Command19_Click
End Sub

Private Sub xtoy_Click()
Command21_Click
End Sub
