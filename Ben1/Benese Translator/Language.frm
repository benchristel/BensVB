VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Language"
   ClientHeight    =   1095
   ClientLeft      =   8160
   ClientTop       =   7305
   ClientWidth     =   2655
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   2655
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtOutput 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "this is where words come out."
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Consonant(1 To 15) As String, Vowel(1 To 7) As String, NasalLiquid(1 To 18) As String, ClusterConsonant(1 To 17) As String
Dim Output As String
Private Sub cmdGenerate_Click()
Select Case Int(Rnd * 6) + 1
Case Is = 1
    Call GenerateCVCe
Case Is = 2
    Call GenerateCN
Case Is = 3
    Call GenerateCNYea
Case Is = 4
    Call GenerateCVCea
Case Is = 5
    Call GenerateNVCa
Case Is = 6
    Call GenerateNYa
End Select
End Sub

Private Sub Form_Load()
Consonant(1) = "mb"
Consonant(2) = "c"
Consonant(3) = "nd"
Consonant(4) = "f"
Consonant(5) = "ng"
Consonant(6) = "h"
Consonant(7) = "l"
Consonant(8) = "m"
Consonant(9) = "n"
Consonant(10) = "p"
Consonant(11) = "qu"
Consonant(12) = "r"
Consonant(13) = "s"
Consonant(14) = "t"
Consonant(15) = "v"
'''
Vowel(1) = "a"
Vowel(2) = "e"
Vowel(3) = "i"
Vowel(4) = "o"
Vowel(5) = "ai"
Vowel(6) = "ui"
Vowel(7) = "au"
'''
ClusterConsonant(1) = "b"
ClusterConsonant(2) = "c"
ClusterConsonant(3) = "d"
ClusterConsonant(4) = "f"
ClusterConsonant(5) = "g"
ClusterConsonant(6) = "h"
ClusterConsonant(7) = "l"
ClusterConsonant(8) = "m"
ClusterConsonant(9) = "n"
ClusterConsonant(10) = "p"
ClusterConsonant(11) = "qu"
ClusterConsonant(12) = "r"
ClusterConsonant(13) = "s"
ClusterConsonant(14) = "t"
ClusterConsonant(15) = "v"
ClusterConsonant(16) = "w"
ClusterConsonant(17) = "y"
'''
NasalLiquid(1) = "an"
NasalLiquid(2) = "en"
NasalLiquid(3) = "in"
NasalLiquid(4) = "on"
NasalLiquid(5) = "un"
NasalLiquid(6) = "al"
NasalLiquid(7) = "el"
NasalLiquid(8) = "il"
NasalLiquid(9) = "ol"
NasalLiquid(10) = "ul"
NasalLiquid(11) = "ar"
NasalLiquid(12) = "er"
NasalLiquid(13) = "ir"
NasalLiquid(14) = "or"
NasalLiquid(15) = "ur"
NasalLiquid(16) = "ien"
NasalLiquid(17) = "ier"
NasalLiquid(18) = "iel"

Randomize
End Sub

Public Sub GenerateCVCe()
Output = Consonant(Int(Rnd * 15 + 1)) & Vowel(Int(Rnd * 7 + 1)) & Consonant(Int(Rnd * 15 + 1)) & "e"
Output = ApplyChanges(Output)
txtOutput.Text = Output
End Sub

Public Sub GenerateCN()
Output = Consonant(Int(Rnd * 15 + 1)) & NasalLiquid(Int(Rnd * 15 + 1))
Output = ApplyChanges(Output)
txtOutput.Text = Output
End Sub

Public Sub GenerateCNYea()
Dim final
Select Case Int(Rnd * 2)
Case Is = 0
final = "e"
Case Is = 1
final = "a"
End Select
Output = Consonant(Int(Rnd * 15 + 1)) & NasalLiquid(Int(Rnd * 18 + 1)) & ClusterConsonant(Int(Rnd * 17 + 1)) & final
Output = ApplyChanges(Output)
txtOutput.Text = Output
End Sub

Public Sub GenerateCVCea()
Output = Consonant(Int(Rnd * 15 + 1)) & Vowel(Int(Rnd * 7 + 1)) & Consonant(Int(Rnd * 15 + 1)) & Vowel(Int(Rnd * 2 + 1))
Output = ApplyChanges(Output)
txtOutput.Text = Output
End Sub
Public Sub GenerateNVCa()
Output = NasalLiquid(Int(Rnd * 15 + 1)) & Vowel(Int(Rnd * 7 + 1)) & Consonant(Int(Rnd * 15 + 1)) & "a"
Output = ApplyChanges(Output)
txtOutput.Text = Output
End Sub
Public Sub GenerateNYa()
Output = NasalLiquid(Int(Rnd * 15 + 1)) & ClusterConsonant(Int(Rnd * 17 + 1)) & "a"
Output = ApplyChanges(Output)
txtOutput.Text = Output
End Sub

Public Function ApplyChanges(Text As String) As String
Text = "#" & Text & "#" 'mark beginning and end of string with # character
Text = Replace(Text, "#mb", "#m")
Text = Replace(Text, "#nd", "#n")
Text = Replace(Text, "#ng", "#ty")
Text = Replace(Text, "f", "lf")
Text = Replace(Text, "#lf", "f")
Text = Replace(Text, "ae", "e")
Text = Replace(Text, "ae", "ai")
Text = Replace(Text, "ao", "au")
'Text = Replace(Text, "ee", "i")
'Text = Replace(Text, "ei", "ie")
'Text = Replace(Text, "eo", "e")
'Text = Replace(Text, "aue", "ai")
'Text = Replace(Text, "yau", "ya")
'Text = Replace(Text, "oe", "oi")
'Text = Replace(Text, "oa", "o")
'Text = Replace(Text, "aie", "e")

Text = Replace(Text, "quu", "qua")
Text = Replace(Text, "quau", "que")

Text = Replace(Text, "nlf", "mv")
Text = Replace(Text, "llf", "lf")
Text = Replace(Text, "rlf", "rm")
Text = Replace(Text, "nr", "mr")
Text = Replace(Text, "np", "mp")
Text = Replace(Text, "mt", "nt")
Text = Replace(Text, "mg", "ng")

Text = Replace(Text, "aind", "ind")
Text = Replace(Text, "aund", "and")
Text = Replace(Text, "uind", "und")
Text = Replace(Text, "aimb", "imb")
Text = Replace(Text, "aumb", "amb")
Text = Replace(Text, "uimb", "umb")
Text = Replace(Text, "aing", "ing")
Text = Replace(Text, "aung", "ang")
Text = Replace(Text, "uing", "ung")
Text = Replace(Text, "auqu", "alqu")

Text = Replace(Text, "oo", "u")
Text = Replace(Text, "uu", "u")
Text = Replace(Text, "yi", "i")
Text = Replace(Text, "ii", "i")
Text = Replace(Text, "aa", "a")
Text = Replace(Text, "ee", "i")
Text = Replace(Text, "ou", "u")

Text = Replace(Text, "quu", "que")
Text = Replace(Text, "#", "")

ApplyChanges = Text
End Function
Private Function Replace(SearchText As String, FindText As String, ReplaceText As String)
Dim i
Replace = SearchText
For i = 1 To Len(SearchText)
If Mid(SearchText, i, Len(FindText)) = FindText Then
SearchText = Mid(SearchText, 1, i - 1) & ReplaceText & Right(SearchText, Len(SearchText) - i - Len(FindText) + 1)
Replace = SearchText
i = Len(SearchText) - Len(Right(SearchText, Len(SearchText) - i - Len(ReplaceText) + 1))
End If
Next i
End Function

