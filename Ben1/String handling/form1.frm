VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
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
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblAnswer 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label lblInfo 
      Caption         =   "Using LEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub cmdGo_Click()
Dim Usertext
Dim ans
Dim letter
Dim Ascii
Dim place
Usertext = txtUser.Text
ans = ""
For place = 1 To Len(Usertext)
letter = Mid(Usertext, place, 1)
Ascii = Asc(letter)
Ascii = Ascii + 1
ans = ans & Chr(Ascii)
Next
lblAnswer.Caption = ans
txtUser.SetFocus
txtUser.SelStart = 0
'txtUser.Text = ""
End Sub


Private Sub Form_Load()
'x = 1
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdGo_Click
End Sub
