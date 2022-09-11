VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Dice"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   700
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   700
   End
   Begin VB.Label lblAnswer 
      Caption         =   "0"
      Height          =   705
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim D()
Dim Throws, Score

Private Sub cmdGo_Click()
ReDim D(6)
Throws = 0
Do
Call ThrowDice
Loop Until Throws = 600

End Sub

Private Sub Form_Load()
Randomize
End Sub

Private Sub ThrowDice()
Throws = Throws + 1
Score = Int(Rnd * 6) + 1
D(Score) = D(Score) + 1

End Sub
