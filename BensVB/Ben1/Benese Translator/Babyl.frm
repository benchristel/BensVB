VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtGenerateNum 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Text            =   "1"
      ToolTipText     =   "# of words to generate"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdFrontback 
      Caption         =   "Front"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtOutput 
      Height          =   14535
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   18015
   End
   Begin VB.TextBox txtSyllables 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "1"
      ToolTipText     =   "letter count"
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frontback As Integer

Private Sub cmdFrontback_Click()
Select Case frontback
Case Is = 0
    frontback = 1
    cmdFrontback.Caption = "Back"
Case Else
    frontback = 0
    cmdFrontback.Caption = "Front"
End Select
End Sub

Private Sub cmdGo_Click()
Dim word As String, i, k
For k = 1 To Int(Val(txtGenerateNum.Text))
word = ""
Select Case frontback
Case Is = 0
If Int(Val(txtSyllables.Text)) Mod 2 = 1 Then
    Select Case Int(Rnd * 5 + 1)
        Case Is = 1
    word = "a"
        Case Is = 2
    word = "i"
        Case Is = 3
    word = "o"
        Case Is = 4
    word = "ai"
        Case Is = 5
    word = "ae"
    End Select
End If
On Error Resume Next
For i = 1 To Int(Val(txtSyllables.Text) / 2)
Select Case Int(Rnd * 5 + 1)
Case Is = 1
word = word & GenerateConsonant & "a"
Case Is = 2
word = word & GenerateConsonant & "i"
Case Is = 3
word = word & GenerateConsonant & "o"
Case Is = 4
word = word & GenerateConsonant & "ai"
Case Is = 5
word = word & GenerateConsonant & "ae"
End Select
Next i
txtOutput.Text = txtOutput.Text & " " & word
Case Else
'
'back vowels
'
If Int(Val(txtSyllables.Text)) Mod 2 = 1 Then
    Select Case Int(Rnd * 5 + 1)
        Case Is = 1
    word = "aa"
        Case Is = 2
    word = "e"
        Case Is = 3
    word = "o"
        Case Is = 4
    word = "u"
        Case Is = 5
    word = "ae"
    End Select
End If
On Error Resume Next
For i = 1 To Int(Val(txtSyllables.Text) / 2)
Select Case Int(Rnd * 5 + 1)
Case Is = 1
word = word & GenerateConsonant & "aa"
Case Is = 2
word = word & GenerateConsonant & "e"
Case Is = 3
word = word & GenerateConsonant & "o"
Case Is = 4
word = word & GenerateConsonant & "u"
Case Is = 5
word = word & GenerateConsonant & "ae"
End Select
Next i
txtOutput.Text = txtOutput.Text & " " & word
End Select
Next k
End Sub

Public Function GenerateConsonant() As String
Select Case Int(Rnd * 70 + 1)
Case Is <= 8
    GenerateConsonant = "t"
Case Is <= 14
    GenerateConsonant = "k"
Case Is <= 16
    GenerateConsonant = "yl"
Case Is <= 20
    GenerateConsonant = "y"
Case Is <= 23
    GenerateConsonant = "nt"
Case Is <= 26
    GenerateConsonant = "st"
Case Is <= 32
    GenerateConsonant = "n"
Case Is <= 36
    GenerateConsonant = "m"
Case Is <= 38
    GenerateConsonant = "ky"
Case Is <= 44
    GenerateConsonant = "l"
Case Is <= 50
    GenerateConsonant = "r"
Case Is <= 54
    GenerateConsonant = "s"
Case Is <= 56
    GenerateConsonant = "sh"
Case Is <= 58
    GenerateConsonant = "d"
Case Is <= 62
    GenerateConsonant = "v"
Case Is <= 65
    GenerateConsonant = "ly"
Case Is <= 68
    GenerateConsonant = "ry"
Case Is <= 70
    GenerateConsonant = "ny"
End Select
End Function
