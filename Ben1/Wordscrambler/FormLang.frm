VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   11130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10425
   LinkTopic       =   "Form4"
   ScaleHeight     =   11130
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO"
      Height          =   615
      Left            =   8280
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtOutput 
      Height          =   10815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Type your essay topic (noun phrase) in the box below.  Then click GO"
      Height          =   735
      Left            =   8280
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Noun(1 To 10) As Noun, Verb(1 To 10) As Verb, Adjective(1 To 10) As Adjective


Public Function GenNounPhrase(Number As Integer)
Dim RandNoun
RandNoun = Int(Rnd * 10 + 1)
If Noun(RandNoun).Proper = True Then
GenNounPhrase = Noun(RandNoun).Singular
Exit Function
End If
If Noun(RandNoun).Nebulous = True Then GoTo PluralOnly
If Number = 1 Then
    GenNounPhrase = Noun(RandNoun).Singular
    Select Case Int(Rnd * 2)
        Case Is = 0
        GenNounPhrase = "the " & GenNounPhrase
        Case Is = 1
        If Mid(GenNounPhrase, 1) = a Or Mid(GenNounPhrase, 1) = e Or Mid(GenNounPhrase, 1) = i Or Mid(GenNounPhrase, 1) = o Or Mid(GenNounPhrase, 1) = u Then
            GenNounPhrase = "an " & GenNounPhrase
        Else
            GenNounPhrase = "a " & GenNounPhrase
        End If
    End Select
Else
PluralOnly:
    GenNounPhrase = Noun(RandNoun).Plural
    Select Case Int(Rnd * 3)
        Case Is = 0
        GenNounPhrase = "the " & GenNounPhrase
        Case Is = 1
        GenNounPhrase = "some " & GenNounPhrase
    End Select
End If
End Function

Public Sub DefineWords()
With Noun(1)
    .Singular = "bear"
    .Plural = "bears"
    .Proper = False
    .Nebulous = False
End With
With Noun(2)
    .Singular = "Godzilla"
    .Plural = "Godzillai"
    .Proper = True
    .Nebulous = False
End With
With Noun(3)
    .Singular = "flame"
    .Plural = "flames"
    .Proper = True
    .Nebulous = False
End With
With Noun(4)
    .Singular = "mud"
    .Plural = "mud"
    .Proper = False
    .Nebulous = True
End With
With Noun(5)
    .Singular = "clown"
    .Plural = "clowns"
    .Proper = False
    .Nebulous = False
End With
With Noun(6)
    .Singular = "earwax"
    .Plural = "earwax"
    .Proper = False
    .Nebulous = True
End With
With Noun(7)
    .Singular = "ocean"
    .Plural = "oceans"
    .Proper = False
    .Nebulous = False
End With
With Noun(8)
    .Singular = "Bob"
    .Plural = "Bobs"
    .Proper = True
    .Nebulous = False
End With
With Noun(9)
    .Singular = "brain"
    .Plural = "brains"
    .Proper = False
    .Nebulous = False
End With
With Noun(10)
    .Singular = "chicken"
    .Plural = "chickens"
    .Proper = False
    .Nebulous = False
End With
End Sub

Private Sub cmdGo_Click()
txtOutput.Text = GenNounPhrase(Int(Rnd * 2))
End Sub

Private Sub Form_Load()
Randomize
DefineWords
End Sub

Public Sub GenVerbPhrase()
Dim Tense
Tense = Int(Rnd * 3) + 1
GenVerbPhrase = GenNounPhrase & GenVerb(Tense)
End Sub

Public Sub GenVerb(Tense As Integer) '1 = present 2= past 3 = future
Dim random
Select Case Tense
Case Is = 1
    random = Int(Rnd * 20 + 1)
    Select Case random
    Case Is = 1
        GenVerb = "eat " & GenNounPhrase
    Case Is = 2
        GenVerb = "walk to" & GenPlace
    Case Is = 3
        GenVerb = "eat " & GenNounPhrase
End Sub
