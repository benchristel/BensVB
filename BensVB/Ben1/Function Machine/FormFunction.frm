VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optLevel2 
      Caption         =   "Level 2"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdNewPlayer 
      Caption         =   "New Player"
      Height          =   285
      Left            =   2040
      TabIndex        =   19
      Top             =   4980
      Width           =   975
   End
   Begin VB.CommandButton cmdPlayAgain 
      Caption         =   "Play Again"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   15
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000B&
      Caption         =   "Quit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      MaskColor       =   &H8000000B&
      TabIndex        =   14
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton optLevel3 
      Caption         =   "Level 3"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   1215
   End
   Begin VB.OptionButton optLevel1 
      Caption         =   "Level 1"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Timer tmrCardMove 
      Enabled         =   0   'False
      Interval        =   90
      Left            =   480
      Top             =   0
   End
   Begin VB.CommandButton cmdEnterOK 
      Caption         =   "OK"
      Height          =   615
      Left            =   4320
      TabIndex        =   6
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdNumGuess 
      Caption         =   "1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   1
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdOpGuess 
      Caption         =   "+"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   0
      Top             =   4080
      Width           =   735
   End
   Begin VB.Timer tmrTime 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Guess What It's Doing!"
      Height          =   1455
      Left            =   5520
      TabIndex        =   2
      Top             =   3840
      Width           =   1815
      Begin VB.CommandButton cmdGuessOk 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.TextBox txtEnter 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   675
      Left            =   4320
      TabIndex        =   4
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Choose a level"
      Height          =   1335
      Left            =   0
      TabIndex        =   10
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Timer tmrLights 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Left            =   960
      Top             =   0
   End
   Begin VB.Label lblEntered 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1920
      TabIndex        =   24
      Top             =   4260
      Width           =   1215
   End
   Begin VB.Label lblEnteredtxt 
      Alignment       =   2  'Center
      Caption         =   "You entered:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1860
      TabIndex        =   23
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblEntertxt 
      BackColor       =   &H8000000B&
      Caption         =   "Enter A Number"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblScoreDetails 
      Caption         =   "This Turn         Total"
      Height          =   255
      Left            =   2040
      TabIndex        =   21
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblScoretxt 
      Alignment       =   2  'Center
      Caption         =   "Scoring"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   20
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2760
      TabIndex        =   18
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblAddScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2160
      TabIndex        =   17
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblExample2 
      BackColor       =   &H000080FF&
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape shpLightRed 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4680
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblExample1 
      Caption         =   "Label1"
      ForeColor       =   &H80000003&
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblCorrect 
      Caption         =   "CORRECT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   735
      Left            =   1680
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblIncorrect 
      Alignment       =   1  'Right Justify
      Caption         =   "IN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   735
      Left            =   1080
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000000&
      FillColor       =   &H80000000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   4080
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   4200
      X2              =   5520
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label lblCard 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   5
      Top             =   3960
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Level, Entered, RNumber, ROp, Op, Score, Chances, AddScore

Private Sub cmdEnterOK_Click()
'Level = 1
If Not IsNumeric(txtEnter.Text) Then
    MsgBox "Enter a number between -50 and 50."
    Exit Sub
End If
If txtEnter.Text >= 51 Or txtEnter.Text <= -51 Then
    MsgBox "Enter a number between -50 and 50."
    Exit Sub
End If
lblEntertxt.Visible = False
txtEnter.Visible = False
cmdEnterOK.Enabled = False
tmrCardMove.Enabled = True
lblCard.Caption = txtEnter.Text
lblEntered.Caption = txtEnter.Text
End Sub

Private Sub cmdGuessOk_Click()
Frame1.ForeColor = QBColor(0)
cmdQuit.BackColor = &H80FF&
cmdPlayAgain.BackColor = &H80FF&
cmdGuessOk.Enabled = False
cmdOpGuess.Enabled = False
cmdNumGuess.Enabled = False
If cmdNumGuess.Caption = RNumber And cmdOpGuess.Caption = Op Then
lblCorrect.Visible = True
lblCorrect.ForeColor = QBColor(10)
lblIncorrect.ForeColor = lblExample1.ForeColor
tmrLights.Enabled = True
Score = Score + AddScore
lblScore.Caption = Format(Score, "00#")
AddScore = AddScore + 1
lblAddScore.Caption = AddScore
Select Case Level
Case Is = 1
Chances = 5
Case Is = 2
Chances = 3
Case Else
Chances = 2
End Select
Else
lblIncorrect.Visible = True
lblCorrect.Visible = True
lblIncorrect.ForeColor = QBColor(12)
lblCorrect.ForeColor = QBColor(12)
Chances = Chances - 1
AddScore = AddScore - 1
lblAddScore.Caption = AddScore
Select Case Chances
Case Is = 0
MsgBox "Oops! You missed all your chances!" & vbCrLf & " You will score 0 points for this turn.", 48, "No points!"
Case Else
End Select
cmdQuit.Enabled = True
cmdPlayAgain.Enabled = True
Exit Sub
End If
cmdQuit.Enabled = True
cmdPlayAgain.Enabled = True
End Sub

'Private Sub cmdNumGuess_Click()
'cmdNumGuess.Caption = cmdNemGuess + 1
'If cmdNumGuess.Caption = 10 Then cmdNumGuess = 0
'Exit Sub
'If cmdNumGuess.Caption = "1" Then
'cmdNumGuess.Caption = "2"
'Exit Sub
'End If
'If cmdNumGuess.Caption = "2" Then
'cmdNumGuess.Caption = "3"
'Exit Sub
'End If
'If cmdNumGuess.Caption = "3" Then
'cmdNumGuess.Caption = "4"
'Exit Sub
'End If
'If cmdNumGuess.Caption = "4" Then
'cmdNumGuess.Caption = "5"
'Exit Sub
'End If
'If cmdNumGuess.Caption = "5" Then
'cmdNumGuess.Caption = "6"
'Exit Sub
'End If
'If cmdNumGuess.Caption = "6" Then
'cmdNumGuess.Caption = "7"
'Exit Sub
'End If
'If cmdNumGuess.Caption = "7" Then
'cmdNumGuess.Caption = "8"
'Exit Sub
'End If
'If cmdNumGuess.Caption = "8" Then
'cmdNumGuess.Caption = "9"
'Exit Sub
'End If
'If cmdNumGuess.Caption = "9" Then
'cmdNumGuess.Caption = "0"
'Exit Sub
'End If
'If cmdNumGuess.Caption = "0" Then
'cmdNumGuess.Caption = "1"
'Exit Sub
'End If
'End Sub


Private Sub cmdNumGuess_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then cmdNumGuess.Caption = cmdNumGuess.Caption + 1
If Button = 2 Then cmdNumGuess.Caption = cmdNumGuess.Caption - 1
If cmdNumGuess.Caption = 10 Then cmdNumGuess.Caption = 0
If cmdNumGuess.Caption = -1 Then cmdNumGuess.Caption = 9
End Sub

Private Sub cmdOpGuess_Click()
If cmdOpGuess.Caption = "+" Then
cmdOpGuess.Caption = "-"
Exit Sub
End If
If cmdOpGuess.Caption = "-" Then
cmdOpGuess.Caption = "*"
Exit Sub
End If
If cmdOpGuess.Caption = "*" Then
cmdOpGuess.Caption = "/"
Exit Sub
End If
If cmdOpGuess.Caption = "/" Then
cmdOpGuess.Caption = "+"
Exit Sub
End If
End Sub

Private Sub cmdPlayAgain_Click()
lblCorrect.Visible = False
lblIncorrect.Visible = False
lblEntertxt.Visible = True
cmdPlayAgain.Enabled = False
cmdQuit.Enabled = False
txtEnter.Visible = True
txtEnter.SetFocus
txtEnter.Text = ""
tmrLights.Enabled = False
shpLightRed.FillColor = Form1.BackColor
cmdEnterOK.Enabled = True
lblCorrect.ForeColor = lblExample1.ForeColor
lblIncorrect.ForeColor = lblExample1.ForeColor
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
Randomize
Level = 1
Form1.Caption = "Function Machine <Level " & Level & "> " & Now
shpLightRed.FillColor = Form1.BackColor
Score = 0
Chances = 5
AddScore = 5
End Sub

Private Sub optLevel1_Click()
Chances = 5
AddScore = 5
lblAddScore.Caption = AddScore
Level = 1
optLevel1.ForeColor = QBColor(12)
optLevel2.ForeColor = QBColor(0)
optLevel3.ForeColor = QBColor(0)
End Sub

Private Sub optLevel2_Click()
Chances = 3
AddScore = 5
lblAddScore.Caption = AddScore
Level = 2
optLevel1.ForeColor = QBColor(0)
optLevel2.ForeColor = QBColor(12)
optLevel3.ForeColor = QBColor(0)

End Sub

Private Sub optLevel3_Click()
Chances = 2
AddScore = 5
lblAddScore.Caption = AddScore
Level = 3
optLevel1.ForeColor = QBColor(0)
optLevel2.ForeColor = QBColor(0)
optLevel3.ForeColor = QBColor(12)
End Sub

Private Sub tmrCardMove_Timer()
lblCard.Top = lblCard.Top - 100
If lblCard.Top <= 1400 Then
tmrCardMove.Enabled = False
'cmdOpGuess.Enabled = True
'cmdNumGuess.Enabled = True
'cmdGuessOk.Enabled = True
    If Level = 1 Then
    ROp = Int(Rnd * 3) + 1
    RNumber = Int(Rnd * 5) + 1
        If ROp = 1 Then
        lblCard.Caption = lblCard.Caption - RNumber
        Op = "-"
        Else
        lblCard.Caption = lblCard.Caption + RNumber
        Op = "+"
        End If
    End If  ' End Level = 1 code
    If Level = 2 Then
    ROp = Int(Rnd * 3) + 1
    RNumber = Int(Rnd * 6) + 1
        If ROp = 1 Then
        lblCard.Caption = lblCard.Caption * RNumber
        Op = "*"
        Else
        lblCard.Caption = lblCard.Caption - RNumber
        Op = "*"
        End If
   End If
   If Level = 3 Then
'   ROp = 1
    ROp = Int(Rnd * 3) + 1
    RNumber = Int(Rnd * 4) + 1
        If ROp = 1 Then
        lblCard.Caption = lblCard.Caption / RNumber
        lblCard.Caption = Format(lblCard.Caption, "#.##")
        Op = "/"
        Else
        lblCard.Caption = lblCard.Caption * RNumber
        Op = "*"
        End If
   End If
lblCard.Top = 3960
Frame1.ForeColor = &HFF0000
cmdOpGuess.Enabled = True
cmdNumGuess.Enabled = True
cmdGuessOk.Enabled = True
End If
End Sub


Private Sub tmrCardStop_Timer()

End Sub

Private Sub tmrLights_Timer()
If shpLightRed.FillColor = Form1.BackColor Then
shpLightRed.FillColor = QBColor(12)
Exit Sub
End If
If shpLightRed.FillColor = QBColor(12) Then
shpLightRed.FillColor = lblExample2.BackColor
Exit Sub
End If
If shpLightRed.FillColor = lblExample2.BackColor Then
shpLightRed.FillColor = QBColor(14)
Exit Sub
End If
If shpLightRed.FillColor = QBColor(14) Then
shpLightRed.FillColor = QBColor(10)
Exit Sub
End If
If shpLightRed.FillColor = QBColor(10) Then
shpLightRed.FillColor = QBColor(9)
Exit Sub
End If
If shpLightRed.FillColor = QBColor(9) Then
shpLightRed.FillColor = QBColor(5)
Exit Sub
End If
If shpLightRed.FillColor = QBColor(5) Then
shpLightRed.FillColor = QBColor(12)
Exit Sub
End If
End Sub

Private Sub tmrTime_Timer()
Form1.Caption = "Function Machine <Level " & Level & "> " & Now
End Sub

Private Sub txtEnter_Change()
If Len(txtEnter.Text) > 3 Then
txtEnter.Text = Entered
Exit Sub
End If
Entered = txtEnter.Text
lblCard.Caption = txtEnter.Text
End Sub
