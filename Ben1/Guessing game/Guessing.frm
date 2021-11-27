VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Guessing Game"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdPlayAgain 
      BackColor       =   &H000000FF&
      Caption         =   "Play Again"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H8000000E&
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start New Game"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000A&
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Guess a number from 1 to 100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Guessing Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -360
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ans, GuessCt, userGuess
Private Sub cmdOK_Click()
Text1.SetFocus
userGuess = Val(Text1.Text)
Text1.Text = ""
If userGuess < 1 Or userGuess > 100 Then
MsgBox "Please give a number between 1 and 100", 16, "Silly Guess"

Exit Sub
End If
If userGuess < Ans Then
Label1.Caption = userGuess & " is too small"
GuessCt = GuessCt + 1
Label2.Caption = "Guess #" & GuessCt
End If
If userGuess > Ans Then
Label1.Caption = userGuess & " is too big"
GuessCt = GuessCt + 1
Label2.Caption = "Guess #" & GuessCt
End If
If userGuess = Ans Then
Label1.Caption = userGuess & " is correct"
GuessCt = GuessCt + 1
Label1.BackColor = QBColor(12)
Label2.BackColor = QBColor(12)
Text1.BackColor = QBColor(12)
cmdOK.Visible = False
Form1.BackColor = QBColor(14)
cmdQuit.Visible = True
CmdPlayAgain.Visible = True
Label2.Caption = "You won in " & GuessCt & " Guesses"

End If
If GuessCt = 13 Then
Label1.Caption = "You Lose"
Label2.Caption = "The number was " & Ans
Form1.BackColor = QBColor(0)
Label1.BackColor = QBColor(0)
Label1.ForeColor = QBColor(15)
Label2.BackColor = QBColor(0)
Label2.ForeColor = QBColor(15)
Text1.BackColor = QBColor(0)
cmdQuit.BackColor = QBColor(15)
CmdPlayAgain.BackColor = QBColor(15)
CmdPlayAgain.Visible = True
cmdQuit.Visible = True
cmdOK.Visible = False
End If
End Sub

Private Sub CmdPlayAgain_Click()
Text1.SetFocus
Randomize
Label1.BackColor = QBColor(7)
Label1.Caption = "Type your guess then click on OK"
Label2.BackColor = QBColor(7)
Text1.BackColor = QBColor(7)
cmdOK.Visible = True
Form1.BackColor = QBColor(7)
cmdQuit.Visible = False
CmdPlayAgain.Visible = False
Call cmdStart_Click
CmdPlayAgain.BackColor = QBColor(12)
cmdQuit.BackColor = QBColor(12)
Label1.ForeColor = QBColor(0)
Label2.ForeColor = QBColor(0)
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdStart_Click()
Ans = Int(Rnd * 100) + 1
GuessCt = 1
cmdStart.Visible = False
Text1.Visible = True
cmdOK.Visible = True
Text1.Text = ""
Label1.Caption = "  Type your guess then click on OK"
Label2.Caption = "Guess #1"
End Sub

Private Sub Form_Load()
Randomize
End Sub

Private Sub Text1_Change()
cmdOK.Enabled = True
If Text1.Text = "" Then
cmdOK.Enabled = False
End If
End Sub
