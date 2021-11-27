VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wordscrambler!"
   ClientHeight    =   4185
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6840
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdStory 
      BackColor       =   &H0000FFFF&
      Caption         =   "Story"
      Height          =   855
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   1635
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0000FFFF&
      Caption         =   "Clear"
      Height          =   855
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   1635
   End
   Begin VB.Timer tmrAuto 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   60
      Top             =   2760
   End
   Begin VB.CommandButton cmdAuto 
      BackColor       =   &H0000FFFF&
      Caption         =   "Autotimer"
      Height          =   855
      Left            =   1740
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   1635
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H0000FFFF&
      Caption         =   "Go"
      Height          =   855
      Left            =   60
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   1635
   End
   Begin VB.Label lblResult 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   9735
      Left            =   60
      TabIndex        =   0
      Top             =   1020
      Width           =   15075
   End
   Begin VB.Menu mnuMode 
      Caption         =   "Mode"
      Begin VB.Menu mnuFractured 
         Caption         =   "Fractured Sentences"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuStories 
         Caption         =   "Crazy Stories"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Adj, Conjunction, Article, ArticleB, NounB, VerbB, VerbPres, VerbSpeech, Occupation
Dim Noun, Verb, Adv, Sentence, SentStru, LabelFont, Number, Time, VerbIng, Mode, Page, Action, Prep

Private Sub cmdAuto_Click()
If tmrAuto.Enabled = False Then
tmrAuto.Enabled = True
cmdGo.Enabled = False
Call GenerateSentence
Else
tmrAuto.Enabled = False
cmdGo.Enabled = True
End If
End Sub

Private Sub cmdClear_Click()
lblResult.Caption = ""
End Sub

Private Sub cmdGo_Click()
Call GenerateSentence
End Sub

Private Sub cmdStory_Click()
Dim Story, Word(1 To 20)
lblResult.Caption = "Once upon a time there was a/an "
Word(1) = Int(Rnd * 10 + 1)
    Select Case Word(1)
    Case Is = 1
    Word(1) = "itty-bitty"
    Case Is = 2
    Word(1) = "violent"
    Case Is = 3
    Word(1) = "rotund"
    Case Is = 4
    Word(1) = "emaciated"
    Case Is = 5
    Word(1) = "chewy"
    Case Is = 6
    Word(1) = "cracked"
    Case Is = 7
    Word(1) = "insane"
    Case Is = 8
    Word(1) = "evil"
    Case Is = 9
    Word(1) = "drunk"
    Case Is = 10
    Word(1) = "cunning"
End Select
    lblResult.Caption = lblResult.Caption & Word(1) & " "

Word(2) = Int(Rnd * 10 + 1)
    Select Case Word(2)
    Case Is = 1
    Word(2) = "ant"
    Case Is = 2
    Word(2) = "camel"
    Case Is = 3
    Word(2) = "sorceress"
    Case Is = 4
    Word(2) = "aardvark"
    Case Is = 5
    Word(2) = "princess"
    Case Is = 6
    Word(2) = "snake charmer"
    Case Is = 7
    Word(2) = "idiot"
    Case Is = 8
    Word(2) = "dragon"
    Case Is = 9
    Word(2) = "piff"
    Case Is = 10
    Word(2) = "beagle"
    End Select
    lblResult.Caption = lblResult.Caption & Word(2) & " named little "
Word(3) = Int(Rnd * 10 + 1)
    Select Case Word(3)
    Case Is = 1
    Word(3) = "mauve"
    Case Is = 2
    Word(3) = "beige"
    Case Is = 3
    Word(3) = "ultramarine"
    Case Is = 4
    Word(3) = "lavender"
    Case Is = 5
    Word(3) = "turquoise"
    Case Is = 6
    Word(3) = "navy"
    Case Is = 7
    Word(3) = "red, white, and blue"
    Case Is = 8
    Word(3) = "gold and black"
    Case Is = 9
    Word(3) = "muddy green"
    Case Is = 10
    Word(3) = "orange"
    End Select
    lblResult.Caption = lblResult.Caption & Word(3) & " riding hood who lived in a "
    Word(4) = Int(Rnd * 10 + 1)
    Select Case Word(4)
    Case Is = 1
    Word(4) = "barrel"
    Case Is = 2
    Word(4) = "lunchbox"
    Case Is = 3
    Word(4) = "pineapple"
    Case Is = 4
    Word(4) = "bureau drawer"
    Case Is = 5
    Word(4) = "castle"
    Case Is = 6
    Word(4) = "tent"
    Case Is = 7
    Word(4) = "bladder"
    Case Is = 8
    Word(4) = "cave"
    Case Is = 9
    Word(4) = "asylum"
    Case Is = 10
    Word(4) = "teepee"
    End Select
    lblResult.Caption = lblResult.Caption & Word(4) & " by a spooky "
    Word(5) = Int(Rnd * 10 + 1)
    Select Case Word(5)
    Case Is = 1
    Word(5) = "desert"
    Case Is = 2
    Word(5) = "lake"
    Case Is = 3
    Word(5) = "mountain"
    Case Is = 4
    Word(5) = "city"
    Case Is = 5
    Word(5) = "plain"
    Case Is = 6
    Word(5) = "island"
    Case Is = 7
    Word(5) = "ocean"
    Case Is = 8
    Word(5) = "peninsula"
    Case Is = 9
    Word(5) = "volcano"
    Case Is = 10
    Word(5) = "boy scout camp"
    End Select
    lblResult.Caption = lblResult.Caption & Word(5) & " with her mother, who was a professional "
End Sub

Private Sub Form_Load()
Randomize
Mode = "Sentence"
End Sub





Private Sub RandomArticleAn()
ArticleB = Int(Rnd * 2 + 1)
If ArticleB = 1 Then
ArticleB = "an"
Else
ArticleB = "the"
End If
End Sub

Private Sub RandomAdjective()
Adj = Int(Rnd * 30 + 1)
        Select Case Adj
        Case Is = 1
        Adj = "exuberant"
        Article = "An"
        Case Is = 2
        Adj = "thoughtful"
        Article = "A"
        Case Is = 3
        Adj = "disheveled"
        Article = "A"
        Case Is = 4
        Adj = "incapacitated"
        Article = "An"
        Case Is = 5
        Adj = "explosive"
        Article = "An"
        Case Is = 6
        Adj = "bubbly"
        Article = "A"
        Case Is = 7
        Adj = "ungrateful"
        Article = "An"
        Case Is = 8
        Adj = "awkward"
        Article = "An"
        Case Is = 9
        Adj = "microscopic"
        Article = "A"
        Case Is = 10
        Adj = "gigantic"
        Article = "A"
        Case Is = 11
        Adj = "crunchy"
        Article = "A"
        Case Is = 12
        Adj = "bulbous"
        Article = "A"
        Case Is = 13
        Adj = "disembodied"
        Article = "A"
        Case Is = 14
        Adj = "crumbly"
        Article = "A"
        Case Is = 15
        Adj = "effervescent"
        Article = "An"
        Case Is = 16
        Adj = "reinforced"
        Article = "A"
        Case Is = 17
        Adj = "turbocharged"
        Article = "A"
        Case Is = 18
        Adj = "disgruntled"
        Article = "A"
        Case Is = 19
        Adj = "stupified"
        Article = "A"
        Case Is = 20
        Adj = "paranoid"
        Article = "A"
        Case Is = 21
        Adj = "petrified"
        Article = "A"
        Case Is = 22
        Adj = "carved"
        Article = "A"
        Case Is = 23
        Adj = "captivated"
        Article = "A"
        Case Is = 24
        Adj = "slimy"
        Article = "A"
        Case Is = 25
        Adj = "massive"
        Article = "A"
        Case Is = 26
        Adj = "rotund"
        Article = "A"
        Case Is = 27
        Adj = "flat"
        Article = "A"
        Case Is = 28
        Adj = "feathered"
        Article = "A"
        Case Is = 29
        Adj = "dead"
        Article = "A"
        Case Is = 30
        Adj = "odiferous"
        Article = "An"
        End Select
End Sub

Private Sub RandomNoun()
    Noun = Int(Rnd * 52 + 1)
        Select Case Noun
        Case Is = 1
        Noun = "anvil"
        Case Is = 2
        Noun = "tuba"
        Case Is = 3
        Noun = "penguin"
        Case Is = 4
        Noun = "porcupine"
        Case Is = 5
        Noun = "pumpkin"
        Case Is = 6
        Noun = "piff"
        Case Is = 7
        Noun = "aardvark"
        Case Is = 8
        Noun = "slug"
        Case Is = 9
        Noun = "hutt"
        Case Is = 10
        Noun = "ewok"
        Case Is = 11
        Noun = "duckling"
        Case Is = 12
        Noun = "mattress"
        Case Is = 13
        Noun = "violin"
        Case Is = 14
        Noun = "crouton"
        Case Is = 15
        Noun = "igloo"
        Case Is = 16
        Noun = "moccasin"
        Case Is = 17
        Noun = "blubber bag"
        Case Is = 18
        Noun = "bladder"
        Case Is = 19
        Noun = "noodle"
        Case Is = 20
        Noun = "tongue"
        Case Is = 21
        Noun = "bathtub"
        Case Is = 22
        Noun = "lawnmower"
        Case Is = 23
        Noun = "locomotive"
        Case Is = 24
        Noun = "dragon"
        Case Is = 25
        Noun = "coal shovel"
        Case Is = 26
        Noun = "rhomboid"
        Case Is = 27
        Noun = "deity"
        Case Is = 28
        Noun = "thug"
        Case Is = 29
        Noun = "hooligan"
        Case Is = 30
        Noun = "tooth"
        Case Is = 31
        Noun = "spleen"
        Case Is = 32
        Noun = "post-it"
        Case Is = 33
        Noun = "shrapnel fragment"
        Case Is = 34
        Noun = "ghost"
        Case Is = 35
        Noun = "pork chop"
        Case Is = 36
        Noun = "axe-head"
        Case Is = 37
        Noun = "fugal horn"
        Case Is = 38
        Noun = "snork"
        Case Is = 39
        Noun = "tarantula"
        Case Is = 40
        Noun = "swillbomb"
        Case Is = 41
        Noun = "enchilada"
        Case Is = 42
        Noun = "kokomoko"
        Case Is = 43
        Noun = "molay"
        Case Is = 44
        Noun = "arbol"
        Case Is = 45
        Noun = "mono"
        Case Is = 46
        Noun = "Nina"
        Case Is = 47
        Noun = "Pinta"
        Case Is = 48
        Noun = "Santa Maria"
        Case Is = 49
        Noun = "spouse"
        Case Is = 50
        Noun = "braces"
        Case Is = 51
        Noun = "taus"
        Case Is = 52
        Noun = "obeirpfapfnenhuoppfhen"
        End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Form1 = Nothing
End Sub

Private Sub mnuFractured_Click()
Mode = "Fractured"
End Sub

Private Sub tmrAuto_Timer()
cmdGo.Enabled = False
Call GenerateSentence
End Sub

Private Sub GenerateSentence()
Select Case Mode
Case Is = "Sentence"
    SentStru = Int(Rnd * 7 + 1)
        Select Case SentStru
        Case Is = 1
        Call RandomWords
        Call RandomAdverb
        Sentence = Article & " " & Adj & " " & Noun & " " & Verb & " " & Adv & "."
        Case Is = 2
        Conjunction = Int(Rnd * 2 + 1)
            Select Case Conjunction
            Case Is = 1
            Conjunction = "Although"
            Case Is = 2
            Conjunction = "Since"
            End Select
        Call RandomNounB
        Call RandomVerbPast
        Call RandomWords
        Call RandomAdverb
        Sentence = Conjunction & " " & ArticleB & " " & NounB & " " & VerbB & "," & "the" & " " & Adj & " " & Noun & " " & Verb & " " & Adv & "."
        Case Is = 3
        VerbPres = Int(Rnd * 6 + 1)
            Select Case VerbPres
            Case Is = 1
            VerbPres = "squish"
            Case Is = 2
            VerbPres = "expose"
            Case Is = 3
            VerbPres = "eliminate"
            Case Is = 4
            VerbPres = "extrude"
            Case Is = 5
            VerbPres = "compress"
            Case Is = 6
            VerbPres = "kick"
            End Select
        Call RandomAdjective
        Call RandomNoun
        Call RandomVerbSpeech
            Call RandomOccupation
        Sentence = Chr$(34) & "We must " & VerbPres & " this " & Adj & " " & Noun & ", " & Chr$(34) & " " & VerbSpeech & " the " & Occupation & "."
        Case Is = 4
        Call RandomNumber
        Call RandomTime
        Call RandomAdjective
        Call RandomVerbIng
        Call RandomNounB
        Call RandomVerbPast
        Sentence = "After " & Number & " " & Time & " of " & Adj & " " & VerbIng & ", " & ArticleB & " " & NounB & " " & VerbB & "."
        Case Is = 5
        Call RandomAdjective
        Call RandomAction
        Call RandomVerbPast
        Call RandomTime
        Call RandomVerbSpeech
        Call RandomOccupation
        Call RandomNoun
        Sentence = Chr$(34) & "This is the most " & Adj & " " & Noun & " I've " & Action & " in " & Time & "!" & Chr$(34) & " " & VerbSpeech & " the " & Occupation & "."
        Case Is = 6
        Call RandomNounB
        Call RandomAction
        Call RandomNoun
        Call RandomAdverb
        Call RandomAction
        Call RandomPrep
        Sentence = ArticleB & " " & NounB & "'s pet " & Noun & " was " & Adv & " " & Action & " " & Prep & "."
        Case Is = 7
        Call RandomNounB
        Call RandomAction
        Call RandomNoun
        Call RandomAdjective
        Call RandomPrep
        Sentence = "Has " & NounB & " " & Action & " the " & Adj & " " & Noun & " " & Prep & "?"
        End Select

    lblResult.Caption = lblResult.Caption & "   " & Sentence
'    Case Is = "Fractured"
'    Page = 1
'        Do
'        Select Case Page
'        Case Is = 1
'        Call RandomAdjective
'        Call RandomOccupation
'        Call RandomNoun
'        lblResult.Caption = "Once upon a time there was a girl named Red Riding " & Noun & " Who lived with her mother"
  End Select
End Sub



Private Sub RandomArticleA()
ArticleB = Int(Rnd * 2 + 1)
If ArticleB = 1 Then
ArticleB = "a"
Else
ArticleB = "the"
End If
End Sub

Private Sub RandomWords()
Call RandomAdjective
Call RandomNoun

    Verb = Int(Rnd * 30 + 1)
        Select Case Verb
        Case Is = 1
        Verb = "shriveled"
        Case Is = 2
        Verb = "belched"
        Case Is = 3
        Verb = "plummeted"
        Case Is = 4
        Verb = "sizzled"
        Case Is = 5
        Verb = "burst"
        Case Is = 6
        Verb = "disembarked"
        Case Is = 7
        Verb = "sniffed"
        Case Is = 8
        Verb = "smiled"
        Case Is = 9
        Verb = "writhed"
        Case Is = 10
        Verb = "cheered"
        Case Is = 11
        Verb = "masticated"
        Case Is = 12
        Verb = "evacuated"
        Case Is = 13
        Verb = "evaporated"
        Case Is = 14
        Verb = "materialized"
        Case Is = 15
        Verb = "multiplied"
        Case Is = 16
        Verb = "scribbled"
        Case Is = 17
        Verb = "gnawed"
        Case Is = 18
        Verb = "blushed"
        Case Is = 19
        Verb = "jiggled"
        Case Is = 20
        Verb = "retreated"
        Case Is = 21
        Verb = "died"
        Case Is = 22
        Verb = "surrenderred"
        Case Is = 23
        Verb = "fluffed"
        Case Is = 24
        Verb = "sucked"
        Case Is = 25
        Verb = "slurped"
        Case Is = 26
        Verb = "sank"
        Case Is = 27
        Verb = "bled"
        Case Is = 28
        Verb = "tumbled"
        Case Is = 29
        Verb = "whistled"
        Case Is = 30
        Verb = "chugged"
        End Select
End Sub

Private Sub RandomNumber()
Number = Int(Rnd * 9 + 2)
End Sub

Private Sub RandomTime()
Time = Int(Rnd * 10 + 1)
    Select Case Time
    Case Is = 1
    Time = "years"
    Case Is = 2
    Time = "days"
    Case Is = 3
    Time = "months"
    Case Is = 4
    Time = "weeks"
    Case Is = 5
    Time = "hours"
    Case Is = 6
    Time = "minutes"
    Case Is = 7
    Time = "seconds"
    Case Is = 8
    Time = "decades"
    Case Is = 9
    Time = "centuries"
    Case Is = 10
    Time = "millenia"
    End Select
End Sub

Private Sub RandomAdverb()
Adv = Int(Rnd * 20 + 1)
        Select Case Adv
        Case Is = 1
        Adv = "inwardly"
        Case Is = 2
        Adv = "rudely"
        Case Is = 3
        Adv = "dangerously"
        Case Is = 4
        Adv = "cautiously"
        Case Is = 5
        Adv = "minutely"
        Case Is = 6
        Adv = "slowly"
        Case Is = 7
        Adv = "directly"
        Case Is = 8
        Adv = "painfully"
        Case Is = 9
        Adv = "furiously"
        Case Is = 10
        Adv = "messily"
        Case Is = 11
        Adv = "weakly"
        Case Is = 12
        Adv = "forcefully"
        Case Is = 13
        Adv = "loudly"
        Case Is = 14
        Adv = "crisply"
        Case Is = 15
        Adv = "sharply"
        Case Is = 16
        Adv = "hungrily"
        Case Is = 17
        Adv = "actively"
        Case Is = 18
        Adv = "horrifically"
        Case Is = 19
        Adv = "broadly"
        Case Is = 20
        Adv = "heavily"
        End Select
End Sub

Private Sub RandomVerbIng()
VerbIng = Int(Rnd * 10 + 1)
    Select Case VerbIng
    Case Is = 1
    VerbIng = "frowning"
    Case Is = 2
    VerbIng = "screeching"
    Case Is = 3
    VerbIng = "sputtering"
    Case Is = 4
    VerbIng = "coughing"
    Case Is = 5
    VerbIng = "crawling"
    Case Is = 6
    VerbIng = "sneezing"
    Case Is = 7
    VerbIng = "flopping"
    Case Is = 8
    VerbIng = "overheating"
    Case Is = 9
    VerbIng = "degrading"
    Case Is = 10
    VerbIng = "climbing"
    End Select
End Sub

Public Sub RandomNounB()
    NounB = Int(Rnd * 23 + 1)
        Select Case NounB
        Case Is = 1
        NounB = "Bob"
        Case Is = 2
        NounB = "Quetzalcoatl"
        Case Is = 3
        NounB = "Huitzilopochtli"
        Case Is = 4
        NounB = "Xochicalco"
        Case Is = 5
        NounB = "Tlaloc"
        Case Is = 6
        NounB = "Zeus"
        Case Is = 7
        NounB = "Ares"
        Case Is = 8
        NounB = "Kokopelli"
        Case Is = 9
        NounB = "Feather"
        Case Is = 10
        NounB = "Pwt"
        Case Is = 11
        NounB = "Yoda"
        ArticleB = ""
        Case Is = 12
        NounB = "Obi-Wan Kenobi"
        ArticleB = ""
        Case Is = 13
        NounB = "Darth Vader"
        ArticleB = ""
        Case Is = 14
        NounB = "Putt-Putt"
        ArticleB = ""
        Case Is = 15
        NounB = "Manny Calavera"
        ArticleB = ""
        Case Is = 16
        NounB = "Ms. Frizzle"
        ArticleB = ""
        Case Is = 17
        NounB = "Ramses II"
        ArticleB = ""
        Case Is = 18
        NounB = "Nebuchadnezzar"
        ArticleB = ""
        Case Is = 19
        NounB = "Tutankhamun"
        ArticleB = ""
        Case Is = 20
        NounB = "Moses"
        ArticleB = ""
        Case Is = 21
        NounB = "Dr. Seuss"
        ArticleB = ""
        Case Is = 22
        NounB = "George W. Bush"
        ArticleB = ""
        Case Is = 23
        NounB = "Saddam Hussein"
        ArticleB = ""
        End Select
End Sub

Public Sub RandomVerbPast()
 VerbB = Int(Rnd * 20 + 1)
        Select Case VerbB
        Case Is = 1
        VerbB = "was toppled"
        Case Is = 2
        VerbB = "was vanquished"
        Case Is = 3
        VerbB = "was eaten"
        Case Is = 4
        VerbB = "was broken"
        Case Is = 5
        VerbB = "was eletrocuted"
        Case Is = 6
        VerbB = "collapsed"
        Case Is = 7
        VerbB = "inflated"
        Case Is = 8
        VerbB = "detonated"
        Case Is = 9
        VerbB = "imploded"
        Case Is = 10
        VerbB = "snorted"
        Case Is = 11
        VerbB = "vomited"
        Case Is = 12
        VerbB = "wobbled"
        Case Is = 13
        VerbB = "went mad"
        Case Is = 14
        VerbB = "volunteered"
        Case Is = 15
        VerbB = "was impaled"
        Case Is = 16
        VerbB = "jabbered incessantly"
        Case Is = 17
        VerbB = "was undermined"
        Case Is = 18
        VerbB = "recovered"
        Case Is = 19
        VerbB = "was excavated"
        Case Is = 20
        VerbB = "sneezed"
        End Select
End Sub

Private Sub RandomOccupation()
        Occupation = Int(Rnd * 20 + 1)
            Select Case Occupation
            Case Is = 1
            Occupation = "dentist"
            Case Is = 2
            Occupation = "cook"
            Case Is = 3
            Occupation = "Indian chief"
            Case Is = 4
            Occupation = "lawyer"
            Case Is = 5
            Occupation = "Spartan warrior"
            Case Is = 6
            Occupation = "grandmother"
            Case Is = 7
            Occupation = "babysitter"
            Case Is = 8
            Occupation = "sentenced criminal"
            Case Is = 9
            Occupation = "rhinologist"
            Case Is = 10
            Occupation = "undersecretary"
            Case Is = 11
            Occupation = "engineer"
            Case Is = 12
            Occupation = "zookeeper"
            Case Is = 13
            Occupation = "burglar"
            Case Is = 14
            Occupation = "mason"
            Case Is = 15
            Occupation = "surgeon"
            Case Is = 16
            Occupation = "cowboy"
            Case Is = 17
            Occupation = "astronaut"
            Case Is = 18
            Occupation = "teacher"
            Case Is = 19
            Occupation = "principal"
            Case Is = 20
            Occupation = "king"
            End Select
End Sub

Public Sub RandomVerbSpeech()
        VerbSpeech = Int(Rnd * 11 + 1)
            Select Case VerbSpeech
            Case Is = 1
            VerbSpeech = "stated"
            Case Is = 2
            VerbSpeech = "exploded"
            Case Is = 3
            VerbSpeech = "spat"
            Case Is = 4
            VerbSpeech = "wheezed"
            Case Is = 5
            VerbSpeech = "cried"
            Case Is = 6
            VerbSpeech = "whispered"
            Case Is = 7
            VerbSpeech = "moaned"
            Case Is = 8
            VerbSpeech = "proclaimed"
            Case Is = 9
            VerbSpeech = "concluded"
            Case Is = 10
            VerbSpeech = "mumbled"
            Case Is = 11
            VerbSpeech = "drooled"
            End Select
End Sub

Public Sub RandomAction()
        Action = Int(Rnd * 10 + 1)
            Select Case Action
            Case Is = 1
            Action = "pureed"
            Case Is = 2
            Action = "misted"
            Case Is = 3
            Action = "kicked"
            Case Is = 4
            Action = "eaten"
            Case Is = 5
            Action = "cremated"
            Case Is = 6
            Action = "pasted"
            Case Is = 7
            Action = "shot"
            Case Is = 8
            Action = "coughed"
            Case Is = 9
            Action = "strangled"
            Case Is = 10
            Action = "refrigerated"
            End Select
End Sub

Public Sub RandomPrep()
Prep = Int(Rnd * 10 + 1)
    Select Case Prep
    Case Is = 1
    Prep = "up the chimney"
    Case Is = 2
    Prep = "down the toilet"
    Case Is = 3
    Prep = "through the window"
    Case Is = 4
    Prep = "into the sky"
    Case Is = 5
    Prep = "into a brick wall"
    Case Is = 6
    Prep = "into the ocean"
    Case Is = 7
    Prep = "down the tubes"
    Case Is = 8
    Prep = "into the whale's mouth"
    Case Is = 9
    Prep = "over the deadly desert"
    Case Is = 10
    Prep = "through the secret tunnel"
    End Select
End Sub
