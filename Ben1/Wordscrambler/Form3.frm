VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEntered 
      Height          =   2595
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   3435
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin VB.Label lblDisplay 
      Height          =   3075
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3555
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WordCount, InputWord(), OutputWord(), EnteredWord(), Words
Private Sub cmdGenerate_Click()
Dim i, x, NextLetter, wordlength, ReadPos, WriteIndex
Randomize
lblDisplay.Caption = ""
WordCount = 0
ReDim EnteredWord(1 To 1)
txtEntered.Text = " " & Trim(txtEntered.Text)
Do Until ReadPos = Len(txtEntered.Text)
ReadPos = ReadPos + 1
If Mid(txtEntered.Text, ReadPos, 1) = " " Then
WriteIndex = WriteIndex + 1
WordCount = WordCount + 1
ReDim Preserve EnteredWord(1 To WriteIndex)
ReadPos = ReadPos + 1
End If
If ReadPos <= Len(txtEntered.Text) Then EnteredWord(WriteIndex) = EnteredWord(WriteIndex) & Mid(txtEntered.Text, ReadPos, 1)
Loop
For i = 1 To WordCount
For x = 1 To Words
If InputWord(x) = EnteredWord(i) Then GoTo 1
Next x
Words = Words + 1
ReDim Preserve InputWord(1 To Words)
ReDim Preserve OutputWord(1 To Words)
InputWord(Words) = EnteredWord(i)
OutputWord(Words) = GenerateWord
1:
Next i
For i = 1 To WordCount
For x = 1 To Words
If InputWord(x) = EnteredWord(i) Then lblDisplay.Caption = lblDisplay.Caption & " " & OutputWord(x)
Next x
Next i
For i = 1 To WordCount
EnteredWord(WordCount) = ""
Next i
WriteIndex = 0
End Sub

Private Function GenerateSyllable()
Dim Form
'Form = Int(Rnd * 7 + 1)
'Select Case Form
'Case Is = 1
'GenerateSyllable = GenerateVowel
'Case Is = 2, 3
'GenerateSyllable = GenerateVowel & GenerateConsonant
'Case Is = 4, 5
GenerateSyllable = GenerateConsonant & GenerateVowel
'Case Is = 6, 7
'GenerateSyllable = GenerateConsonant & GenerateVowel & GenerateConsonant
'End Select
End Function

Private Function GenerateVowel()
Dim vowel
vowel = Int(Rnd * 5 + 1)
Select Case vowel
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
End Select
End Function

Private Function GenerateConsonant()
Dim cons
cons = Int(Rnd * 70 + 1)
Select Case cons
Case Is = 1, 2, 3
GenerateConsonant = "b"
Case Is = 4, 5, 6
GenerateConsonant = "c"
Case Is = 7, 8, 9, 10
GenerateConsonant = "d"
Case Is = 11, 12, 13
GenerateConsonant = "f"
Case Is = 14, 15, 16
GenerateConsonant = "g"
Case Is = 17
GenerateConsonant = "h"
Case Is = 18, 19, 20, 21, 22, 23, 24
GenerateConsonant = "l"
Case Is = 25, 26, 27, 28, 29, 30
GenerateConsonant = "m"
Case Is = 31, 32, 33, 34, 35, 36, 37
GenerateConsonant = "n"
Case Is = 38, 39, 40
GenerateConsonant = "p"
Case Is = 41, 42, 43
GenerateConsonant = "qu"
Case Is = 44, 45, 46, 47, 48
GenerateConsonant = "r"
Case Is = 49, 50, 51, 52, 53
GenerateConsonant = "s"
Case Is = 54, 55, 56, 57, 58
GenerateConsonant = "t"
Case Is = 59
GenerateConsonant = "v"
Case Is = 60
GenerateConsonant = "w"
Case Is = 61
GenerateConsonant = "x"
Case Is = 62
GenerateConsonant = "y"
Case Is = 63, 64, 65
GenerateConsonant = "th"
Case Is = 66, 67, 68
GenerateConsonant = "sh"
Case Is = 69
GenerateConsonant = "ch"
Case Is = 70
GenerateConsonant = "ng"
End Select
End Function

Private Function GenerateWord()
Dim wordlength, i, x
wordlength = Int(Rnd * 4 + 1)
For x = 1 To wordlength
GenerateWord = GenerateWord & GenerateSyllable
Next x

End Function

Private Sub Form_Load()
ReDim Preserve InputWord(1 To 1)
ReDim Preserve OutputWord(1 To 1)
Words = 1
End Sub

