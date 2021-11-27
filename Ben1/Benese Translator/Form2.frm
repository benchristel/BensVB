VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   LinkTopic       =   "Form2"
   ScaleHeight     =   6210
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTranslate3 
      Caption         =   "Translate!"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdTranslate2 
      Caption         =   "Translate!"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdTranslate 
      Caption         =   "Translate!"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox txtOutput 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2880
      Width           =   10095
   End
   Begin VB.TextBox txtInput 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Output As String
Private Sub cmdTranslate_Click()
Output = LCase(" " & txtInput.Text & " ")
'put spaces before all puctuation
'Output = Replace(Output, "!", " !")
'Output = Replace(Output, ".", " .")
'Output = Replace(Output, "?", " ?")
'Output = Replace(Output, ",", " ,")
'Output = Replace(Output, ";", " ;")
'Output = Replace(Output, ":", " :")
'vowel adjustment
'Output = Replace(Output, "eal", "enel")
'Output = Replace(Output, "lea", "lene")
'Output = Replace(Output, " ea", " ëa")
'Output = Replace(Output, "ea", "a")
'Output = Replace(Output, "ee", "i")
'Output = Replace(Output, "u", "ui")
'Output = Replace(Output, "oo", "u")
'fix double letters and illegal clusters
'Output = Replace(Output, "bb", "mb")
'Output = Replace(Output, "br", "bar")
'output = Replace(Output, "cc", "c")
'Output = Replace(Output, "ck ", "qua ")
'Output = Replace(Output, "ck", "qu")
'Output = Replace(Output, "dd", "nd")
'Output = Replace(Output, "ff", "lf")
'Output = Replace(Output, "gg", "g")
'Output = Replace(Output, "j", "ss")
'Output = Replace(Output, "k", "c")
'Output = Replace(Output, "wi", "ui")
'Output = Replace(Output, "whi", "ui")
'Output = Replace(Output, "wh", "hw")
'derivational affixes
'Output = Replace(Output, "ed ", "ane ")
'Output = Replace(Output, "ing ", "ëa ")
'irregular words
'Output = Replace(Output, "the", "i")
'Output = Replace(Output, "ever", "uir")
'Output = Replace(Output, "over", "anur")
'Output = Replace(Output, " re", " en")
'other stuff
'Output = Replace(Output, "th", "st")
'Output = Replace(Output, " st", " ess")
'Output = Replace(Output, "b", "mb")
'Output = Replace(Output, "d", "nd")
'Output = Replace(Output, "g", "ng")
'Output = Replace(Output, "mmb", "mb")
'Output = Replace(Output, "nnd", "nd")
'Output = Replace(Output, "nng", "ng")
'Output = Replace(Output, "lmb", "lb")
'Output = Replace(Output, "lnd", "ld")
'Output = Replace(Output, "lng", "lg")
'Output = Replace(Output, "rmb", "rb")
'Output = Replace(Output, "rnd", "rd")
'Output = Replace(Output, "rng", "rg")
'Output = Replace(Output, " mb", " m")
'Output = Replace(Output, " nd", " n")
'Output = Replace(Output, " ng", " c")
'Output = Replace(Output, " lf", " f")
'Output = Replace(Output, " ch", " l")
'Output = Replace(Output, "ch ", "lye ")
'Output = Replace(Output, "ch", "ly")
'Output = Replace(Output, "z", "ss")
'Output = Replace(Output, " ss", " s")
'adjust final things
'Output = Replace(Output, "mb ", "mbe ")
'Output = Replace(Output, "c ", "ca ")
'Output = Replace(Output, "nd ", "nda ")
'Output = Replace(Output, "lf ", "lda ")
'Output = Replace(Output, "ng ", "nga ")
'Output = Replace(Output, "al ", "alya ")
'Output = Replace(Output, "m ", "ma ")
'Output = Replace(Output, "en ", "enya ")
'Output = Replace(Output, "p ", "pa ")
'Output = Replace(Output, "ar ", "arya ")
'Output = Replace(Output, "t ", "ta ")
'Output = Replace(Output, "v ", "va ")
'plurals
'Output = Replace(Output, "ses ", "tari ")
'Output = Replace(Output, "mbs ", "mbi ")
'Output = Replace(Output, "cs ", "car ")
'Output = Replace(Output, "nds ", "ndar ")
'Output = Replace(Output, "lfs ", "ldar ")
'Output = Replace(Output, "ngs ", "ngar ")
'Output = Replace(Output, "ls ", "lli ")
'Output = Replace(Output, "ms ", "mpi ")
'Output = Replace(Output, "ns ", "nti ")
'Output = Replace(Output, "ps ", "par ")
'Output = Replace(Output, "qus ", "quar ")
'Output = Replace(Output, "rs ", "ri ")
'Output = Replace(Output, "ts ", "tar ")
'Output = Replace(Output, "vs ", "var ")
'Output = Replace(Output, "ws ", "lyar ")
'genitive
'Output = Replace(Output, "'s ", "o ")
Output = Replace(Output, "ea", "ae")
Output = Replace(Output, "u", "ui")
Output = Replace(Output, "oo", "u")
Output = Replace(Output, "b", "m")
Output = Replace(Output, "g", "'")
Output = Replace(Output, "c", "g")
Output = Replace(Output, "h", "ch")
Output = Replace(Output, "gh", "ch")
Output = Replace(Output, " wch", " h")

Output = Replace(Output, "d", "dh")
Output = Replace(Output, "f", "lf")
Output = Replace(Output, " j", " rh")
Output = Replace(Output, "k", "c")
Output = Replace(Output, "p", "b")
Output = Replace(Output, "mb", "lf")
Output = Replace(Output, " lf", " f")
Output = Replace(Output, " la", " lha")
Output = Replace(Output, " qu", "p")
Output = Replace(Output, "qu", "gw")
Output = Replace(Output, "j", "ss")
Output = Replace(Output, "sch", "st")
Output = Replace(Output, "cch", "s")
Output = Replace(Output, " st", " est")
Output = Replace(Output, "t", "d")
Output = Replace(Output, "dch", "th")
Output = Replace(Output, "nd", "th")
Output = Replace(Output, " v", " m")
Output = Replace(Output, " x", " rh")
Output = Replace(Output, "x", "r")
Output = Replace(Output, " y", " i")
Output = Replace(Output, "ou", "y")
Output = Replace(Output, "z", "ss")
Output = Replace(Output, " dh", " d")
Output = Replace(Output, " ch", " h")

txtOutput.Text = UCase(Output)
End Sub

Private Function Replace(SearchText As String, FindText As String, ReplaceText As String)
Dim i
Replace = SearchText
i = 1
Do While i < Len(SearchText)
If Mid(SearchText, i, Len(FindText)) = FindText Then
SearchText = Mid(SearchText, 1, i - 1) & ReplaceText & Right(SearchText, Len(SearchText) - i - Len(FindText) + 1)
Replace = SearchText
i = Len(SearchText) - Len(Right(SearchText, Len(SearchText) - i - Len(ReplaceText) + 1))
End If
i = i + 1
Loop
End Function

Private Sub cmdTranslate2_Click()
Output = LCase(" " & txtInput.Text & " ")
Output = Replace(Output, "zyr", "0")
Output = Replace(Output, "zor", "00")
Output = Replace(Output, "in", "9")
Output = Replace(Output, "aed", "8")
Output = Replace(Output, "sve", "7")
Output = Replace(Output, "a", "5")
Output = Replace(Output, "fyr", "4")
Output = Replace(Output, "tyr", "3")
Output = Replace(Output, "tre", "33")
Output = Replace(Output, "du", "2")
Output = Replace(Output, "le", "1")
Output = Replace(Output, "9", "g")
Output = Replace(Output, "0", "o")
Output = Replace(Output, "8", "b")
Output = Replace(Output, "4", "a")
Output = Replace(Output, "7", "t")
Output = Replace(Output, "5", "s")
Output = Replace(Output, "2", "z")
Output = Replace(Output, "1", "l")
Output = Replace(Output, "3", "e")
Output = Replace(Output, "teh", "the")
Output = Replace(Output, "'", "")
txtOutput.Text = UCase(Output)
End Sub

Private Sub cmdTranslate3_Click()
Output = LCase(" " & txtInput.Text & " ")
Output = Replace(Output, "'", "")
Output = Replace(Output, "the", "teh")
Output = Replace(Output, "e", "3")
Output = Replace(Output, "l", "1")
Output = Replace(Output, "z", "2")
Output = Replace(Output, "s", "5")
Output = Replace(Output, "t", "7")
Output = Replace(Output, "a", "4")
Output = Replace(Output, "b", "8")
Output = Replace(Output, "g", "9")
Output = Replace(Output, "o", "0")
Output = Replace(Output, "1", "le")
Output = Replace(Output, "2", "du")
Output = Replace(Output, "33", "tre")
Output = Replace(Output, "3", "tyr")
Output = Replace(Output, "4", "fyr")
Output = Replace(Output, "5", "a")
Output = Replace(Output, "7", "sve")
Output = Replace(Output, "8", "aed")
Output = Replace(Output, "9", "in")
Output = Replace(Output, "00", "zor")
Output = Replace(Output, "0", "zyr")

txtOutput.Text = (Output)
End Sub
