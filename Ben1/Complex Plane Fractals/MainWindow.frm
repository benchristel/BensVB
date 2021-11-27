VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTolerance 
      Height          =   375
      Left            =   9360
      TabIndex        =   19
      Text            =   ".00000001"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X-Large"
      Height          =   375
      Left            =   9480
      TabIndex        =   18
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdLarge 
      Caption         =   "Large"
      Height          =   375
      Left            =   9480
      TabIndex        =   17
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdMed 
      Caption         =   "Medium"
      Height          =   375
      Left            =   9480
      TabIndex        =   16
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSmall 
      Caption         =   "Small"
      Height          =   375
      Left            =   9480
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdNewton 
      Caption         =   "Go!"
      Height          =   495
      Left            =   9360
      TabIndex        =   14
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox txtScale 
      Height          =   375
      Left            =   9360
      TabIndex        =   13
      Text            =   ".01"
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Index           =   9
      Left            =   6840
      TabIndex        =   10
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Index           =   8
      Left            =   6120
      TabIndex        =   9
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Index           =   7
      Left            =   5400
      TabIndex        =   8
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Index           =   6
      Left            =   4680
      TabIndex        =   7
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Index           =   5
      Left            =   3960
      TabIndex        =   6
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Index           =   4
      Left            =   3240
      TabIndex        =   5
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Index           =   3
      Left            =   2520
      TabIndex        =   4
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Text            =   "0"
      Top             =   6000
      Width           =   615
   End
   Begin VB.PictureBox picOutput 
      AutoRedraw      =   -1  'True
      Height          =   2655
      Left            =   120
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Tolerance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   21
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Scale:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   20
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label lblFPrime 
      Caption         =   "f'(z) = 0"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   6960
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.Label lblF 
      Caption         =   "f(z) = 0"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   6480
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PixScale As Double, XMin As Single, YMin As Single

Private Sub Text1_Change()
Dim output As Complex
Dim foo As Complex
If IsNumeric(Text1.Text) = True And IsNumeric(Text2.Text) = True Then
foo.real = Text1.Text
foo.imag = Text2.Text
output = ComplexPwr(foo, 2)
Label1.Caption = output.real & ", " & output.imag
End If
End Sub

Private Sub cmdLarge_Click()
picOutput.Width = 4575
picOutput.Height = 4575
XMin = -PixScale * picOutput.ScaleWidth / 2
YMin = -PixScale * picOutput.ScaleHeight / 2
End Sub

Private Sub cmdMed_Click()
picOutput.Width = 2655
picOutput.Height = 2655
XMin = -PixScale * picOutput.ScaleWidth / 2
YMin = -PixScale * picOutput.ScaleHeight / 2
End Sub

Private Sub cmdNewton_Click()
Dim x As Integer, y As Integer, j As Integer, oldA As Complex, newA As Complex, shade As Integer, slope As Complex
Dim Red As Single, Blue As Single, Green As Single
For x = 1 To picOutput.ScaleWidth
For y = 1 To picOutput.ScaleHeight
    oldA.real = x * PixScale + XMin
    oldA.imag = y * PixScale + YMin
    For j = 1 To 1000
    slope = ComplexDiv(FVal(oldA), FDeriv(oldA))
        newA.real = oldA.real - slope.real
        newA.imag = oldA.imag - slope.imag
        If complexdist(oldA, newA) < Tolerance Then
        'shade = 64 * Log(j)
        shade = 255 / (1 + Log(j) / 3)
        'shade = 255 * (1 - 0.9 / (1 + Exp(-j / 4 + 5)))
        'shade = 255 - 10 * j
        If shade > 255 Then shade = 255
        Red = Abs(Cos(Atn(newA.imag / newA.real)))
        Green = Abs(Cos(Atn(newA.imag / newA.real) + 2 / 3 * PI))
        Blue = Abs(Cos(Atn(newA.imag / newA.real) + 4 / 3 * PI))
        picOutput.PSet (x, y), RGB(Red * shade, Blue * shade, Green * shade)
        'Refresh
        Exit For
        End If
        oldA = newA
    Next j
Next y
Refresh
Next x
End Sub

Private Sub cmdSmall_Click()
picOutput.Width = 1335
picOutput.Height = 1335
XMin = -PixScale * picOutput.ScaleWidth / 2
YMin = -PixScale * picOutput.ScaleHeight / 2
End Sub

Private Sub Command1_Click()
picOutput.Width = 5775
picOutput.Height = 5775
XMin = -PixScale * picOutput.ScaleWidth / 2
YMin = -PixScale * picOutput.ScaleHeight / 2
End Sub

Private Sub Form_Load()
PixScale = 0.01
Tolerance = 0.00000001
XMin = -PixScale * picOutput.ScaleWidth / 2
YMin = -PixScale * picOutput.ScaleHeight / 2
End Sub

Private Sub txtInput_Change(Index As Integer)
Dim j As Integer
If IsNumeric(txtInput(Index).Text) = True Then
Coeff(Index) = txtInput(Index).Text
End If
lblF.Caption = "f(z) = "
For j = 0 To 9
If Coeff(j) <> 0 Then
    If j > 0 And lblF.Caption <> "f(z) = " Then lblF.Caption = lblF.Caption & " + "
    lblF.Caption = lblF.Caption & Coeff(j) & "z^" & j
End If
Next j
If lblF.Caption = "f(z) = " Then lblF.Caption = "f(z) = 0"
End Sub

Private Sub txtScale_Change()
If IsNumeric(txtScale.Text) = True Then
If txtScale.Text > 0 And txtScale.Text < 10 Then
PixScale = txtScale.Text
XMin = -PixScale * picOutput.ScaleWidth / 2
YMin = -PixScale * picOutput.ScaleHeight / 2
End If
End If
End Sub

Private Sub txtTolerance_Change()
If IsNumeric(txtTolerance.Text) = True Then
    If txtTolerance.Text > 0 And txtTolerance.Text < 1 Then Tolerance = txtTolerance.Text
End If
End Sub
