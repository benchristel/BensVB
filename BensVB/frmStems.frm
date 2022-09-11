VERSION 5.00
Begin VB.Form frmStems 
   Caption         =   "Form5"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9675
   LinkTopic       =   "Form5"
   ScaleHeight     =   4875
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtOutput 
      Height          =   2085
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmStems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGenerate_Click()
Dim template As Integer, vowel As String, Initial As String, Final As String
'template #1 = [L, R, N, M, H, W, I]VC #2 = CV[L, R, N, S]
template = Int(Rnd * 2 + 1)
Select Case template
Case Is = 1
Initial = GenerateConsInitR
Final = GenerateConsFinal
vowel = GenerateVowel
Case Is = 2
Initial = GenerateConsInit
Final = GenerateConsFinalR
vowel = GenerateVowel
End Select
txtOutput.Text = txtOutput.Text & " " & Initial & vowel & Final
End Sub


Public Function GenerateVowel() As String
Dim index As Integer
index = Int(Rnd * 17 + 1)
Select Case index
    Case Is = 1
GenerateVowel = "a"
    Case Is = 2
GenerateVowel = "e"
    Case Is = 3
GenerateVowel = "i"
    Case Is = 4
GenerateVowel = "o"
    Case Is = 5
GenerateVowel = "u"
    Case Is = 6
GenerateVowel = "y"
    Case Is = 7
GenerateVowel = "a^"
    Case Is = 8
GenerateVowel = "e^"
    Case Is = 9
GenerateVowel = "i^"
    Case Is = 10
GenerateVowel = "o^"
    Case Is = 11
GenerateVowel = "u^"
    Case Is = 12
GenerateVowel = "y^"
    Case Is = 13
GenerateVowel = "ae"
    Case Is = 14
GenerateVowel = "ai"
    Case Is = 15
GenerateVowel = "au"
    Case Is = 16
GenerateVowel = "ei"
    Case Is = 17
GenerateVowel = "ui"
End Select
End Function


Public Function GenerateConsInitR() As String
Dim index As Integer
index = Int(Rnd * 17 + 1)
Select Case index
    Case Is = 1
GenerateConxInitR = ""
    Case Is = 2
GenerateConxInitR = "i"
    Case Is = 3
GenerateConxInitR = "r"
    Case Is = 4
GenerateConxInitR = "l"
    Case Is = 5
GenerateConxInitR = "w"
    Case Is = 6
GenerateConxInitR = "n"
    Case Is = 7
GenerateConxInitR = "m"
    Case Is = 8
GenerateConxInitR = "h"
    Case Is = 9
GenerateConxInitR = "lh"
    Case Is = 10
GenerateConxInitR = "rh"
    Case Is = 11
GenerateConxInitR = "br"
    Case Is = 12
GenerateConxInitR = "dr"
    Case Is = 13
GenerateConxInitR = "gl"
    Case Is = 14
GenerateConxInitR = "gr"
    Case Is = 15
GenerateConxInitR = "gw"
    Case Is = 16
GenerateConxInitR = "pr"
    Case Is = 17
GenerateConxInitR = "tr"
End Select
End Function

Public Function GenerateConsInit() As String
Dim index As Integer
index = Int(Rnd * 24 + 1)
Select Case index
    Case Is = 1
GenerateConsInit = "b"
    Case Is = 2
GenerateConsInit = "br"
    Case Is = 3
GenerateConsInit = "c"
    Case Is = 4
GenerateConsInit = "d"
    Case Is = 5
GenerateConsInit = "dr"
    Case Is = 6
GenerateConsInit = "f"
    Case Is = 7
GenerateConsInit = "g"
    Case Is = 8
GenerateConsInit = "gl"
    Case Is = 9
GenerateConsInit = "gr"
    Case Is = 10
GenerateConsInit = "gw"
    Case Is = 11
GenerateConsInit = "h"
    Case Is = 12
GenerateConsInit = "l"
    Case Is = 13
GenerateConsInit = "lh"
    Case Is = 14
GenerateConsInit = "m"
    Case Is = 15
GenerateConsInit = "n"
    Case Is = 16
GenerateConsInit = "p"
    Case Is = 17
GenerateConsInit = "pr"
    Case Is = 18
GenerateConsInit = "r"
    Case Is = 19
GenerateConsInit = "rh"
    Case Is = 20
GenerateConsInit = "s"
    Case Is = 21
GenerateConsInit = "t"
    Case Is = 22
GenerateConsInit = "th"
    Case Is = 23
GenerateConsInit = "tr"
    Case Is = 24
GenerateConsInit = "w"
End Select
End Function

Public Function GenerateConsFinalR() As String
Dim index As Integer
index = Int(Rnd * 18 + 1)
Select Case index
    Case Is = 1
GenerateConsFinalR = "lf"
    Case Is = 2
GenerateConsFinalR = "rm"
    Case Is = 3
GenerateConsFinalR = "n"
    Case Is = 4
GenerateConsFinalR = "l"
    Case Is = 5
GenerateConsFinalR = "r"
    Case Is = 6
GenerateConsFinalR = "ld"
    Case Is = 7
GenerateConsFinalR = "rd"
    Case Is = 8
GenerateConsFinalR = "nd"
    Case Is = 9
GenerateConsFinalR = "lt"
    Case Is = 10
GenerateConsFinalR = "rt"
    Case Is = 11
GenerateConsFinalR = "nt"
    Case Is = 12
GenerateConsFinalR = "rth"
    Case Is = 13
GenerateConsFinalR = "lth"
    Case Is = 14
GenerateConsFinalR = "nth"
    Case Is = 15
GenerateConsFinalR = "rdh"
    Case Is = 16
GenerateConsFinalR = "ldh"
    Case Is = 17
GenerateConsFinalR = "ndh"
    Case Is = 18
GenerateConsFinalR = "ng"
End Select
End Function


Public Function GenerateConsFinal() As String
Dim index As Integer
index = Int(Rnd * 25 + 1)
Select Case index
    Case Is = 1
GenerateConsFinal = "lf"
    Case Is = 2
GenerateConsFinal = "rm"
    Case Is = 3
GenerateConsFinal = "n"
    Case Is = 4
GenerateConsFinal = "l"
    Case Is = 5
GenerateConsFinal = "r"
    Case Is = 6
GenerateConsFinal = "ld"
    Case Is = 7
GenerateConsFinal = "rd"
    Case Is = 8
GenerateConsFinal = "nd"
    Case Is = 9
GenerateConsFinal = "lt"
    Case Is = 10
GenerateConsFinal = "rt"
    Case Is = 11
GenerateConsFinal = "nt"
    Case Is = 12
GenerateConsFinal = "rth"
    Case Is = 13
GenerateConsFinal = "lth"
    Case Is = 14
GenerateConsFinal = "nth"
    Case Is = 15
GenerateConsFinal = "rdh"
    Case Is = 16
GenerateConsFinal = "ldh"
    Case Is = 17
GenerateConsFinal = "ndh"
    Case Is = 18
GenerateConsFinal = "ng"
    Case Is = 19
GenerateConsFinal = "d"
    Case Is = 20
GenerateConsFinal = "g"
    Case Is = 21
GenerateConsFinal = "b"
    Case Is = 22
GenerateConsFinal = "s"
    Case Is = 23
GenerateConsFinal = "dh"
    Case Is = 24
GenerateConsFinal = "th"
    Case Is = 25
GenerateConsFinal = "v"
End Select
End Function

Private Sub Form_Load()
Randomize
End Sub

