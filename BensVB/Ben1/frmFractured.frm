VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "frmFractured.frx":0000
      Left            =   5520
      List            =   "frmFractured.frx":0019
      TabIndex        =   4
      Top             =   1080
      Width           =   1755
   End
   Begin VB.ListBox lstVerbsOI 
      Height          =   1035
      ItemData        =   "frmFractured.frx":004B
      Left            =   9120
      List            =   "frmFractured.frx":0061
      TabIndex        =   3
      Top             =   0
      Width           =   1755
   End
   Begin VB.ListBox lstPlNouns 
      Height          =   1035
      ItemData        =   "frmFractured.frx":0091
      Left            =   7320
      List            =   "frmFractured.frx":00A4
      TabIndex        =   2
      Top             =   0
      Width           =   1755
   End
   Begin VB.ListBox lstNouns 
      Height          =   1035
      ItemData        =   "frmFractured.frx":00DA
      Left            =   5520
      List            =   "frmFractured.frx":00F3
      TabIndex        =   1
      Top             =   0
      Width           =   1755
   End
   Begin VB.Label lblDisplay 
      Caption         =   "Label1"
      Height          =   5475
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5355
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "Load"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add..."
         Begin VB.Menu mnuAddN 
            Caption         =   "Noun"
         End
         Begin VB.Menu mnuAddPN 
            Caption         =   "Plural Noun"
         End
         Begin VB.Menu mnuAddOVI 
            Caption         =   "Objective Verb Infinitive"
         End
         Begin VB.Menu mnuAddOVPs 
            Caption         =   "Objective Verb Past"
         End
         Begin VB.Menu mnuAddOVPr 
            Caption         =   "Objective Verb Present"
         End
         Begin VB.Menu mnuAddOVF 
            Caption         =   "Objective Verb Future"
         End
         Begin VB.Menu mnuAddSVI 
            Caption         =   "Subjective Verb Infinitive"
         End
         Begin VB.Menu mnuAddSVPs 
            Caption         =   "Subjective Verb Past"
         End
         Begin VB.Menu mnuAddSVPr 
            Caption         =   "Subjective Verb Present"
         End
         Begin VB.Menu mnuAddSVF 
            Caption         =   "Subjective Verb Future"
         End
         Begin VB.Menu mnuAddAdj 
            Caption         =   "Adjective"
         End
         Begin VB.Menu mnuAddadv 
            Caption         =   "Adverb"
         End
         Begin VB.Menu mnuAddExc 
            Caption         =   "Exclamation!"
         End
         Begin VB.Menu mnuAddPro 
            Caption         =   "Profession"
         End
         Begin VB.Menu mnuAddFood 
            Caption         =   "Food"
         End
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "Select..."
         Begin VB.Menu mnuSelectCav 
            Caption         =   "Cavemen"
         End
         Begin VB.Menu mnuSelectRSM 
            Caption         =   "Review of a Spooky Movie"
         End
         Begin VB.Menu mnuSelectAJE 
            Caption         =   "Advice to Jungle Explorers"
         End
         Begin VB.Menu mnuSelectTeach 
            Caption         =   "The Teacher"
         End
      End
   End
   Begin VB.Menu mnuMakeYourOwn 
      Caption         =   "Make Your Own!"
      Begin VB.Menu mnuNewHome 
         Caption         =   "New"
      End
      Begin VB.Menu mnuSaveHome 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuLoadHome 
         Caption         =   "Load"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim VerbI1()
'Dim VerbI2()
'Dim VerbI3()
'Dim VerbPs()
'Dim VerbPsP()
'Dim VerbPr()
'Dim VerbF()
'Dim Noun()
'Dim PlNoun()
'Dim Adj()
'Dim Adv()
'Dim Selected()
Dim Word(1 To 50), Response, MadLib
Private Sub Lbldisplay_Click()
Select Case MadLib
Case Is = "CAVEMEN"
Call Cavemen
Case Is = "SPOOKY MOVIE"
Call SpookyMovie
End Select
End Sub

Private Sub Form_Load()
'ReDim VerbI1(1 To 10)
'ReDim VerbI3(1 To 10)
'ReDim VerbPs(1 To 10)
'ReDim VerbPsP(1 To 10)
'ReDim VerbPr(1 To 10)
'ReDim VerbF(1 To 10)
'ReDim Noun(1 To 10)
'ReDim PlNoun(1 To 10)
'ReDim Adj(1 To 10)
'ReDim Adv(1 To 10)
'VerbI1(1) = "laugh"
'VerbI1(2) = "sizzle"
'VerbI1(3) = "jump"
'VerbI1(4) = "throw up"
'VerbI1(5) = "pick"
'VerbI1(6) = "eat"
'VerbI1(7) = "quack"
'VerbI1(8) = "snicker"
'VerbI1(9) = "roll"
'VerbI1(10) = "disembowel"
'VerbI2(1) = "think"
'VerbI2(2) = "chew"
'VerbI2(3) = "staple"
'VerbI2(4) = "flick"
'VerbI2(5) = "snort"
'VerbI2(6) = "gulp"
'VerbI2(7) = "chain"
'VerbI2(8) = "think"
'VerbI2(9) = "think"
'VerbI2(10) = "think"
MadLib = "CAVEMEN"
End Sub

Private Sub mnuSelectAJE_Click()
Dim Slot(1 To 5)
Call RandAdj
Slot(1) = Adj(Selected)
Call RandNoun
Slot(2) = Noun(Selected)
Call RandNoun
Slot(3) = Noun(Selected)
Call RandPlNoun
Slot(4) = PlNoun(Selected)
Call RandVerbSI
Slot(5) = VerbSI(Selected)
End Sub

Private Sub RandWord(Word1, Word2, Word3, Word4, Word5, Word6, Word7, Word8, Word9, Word10)
Randomize
Select Case Int(Rnd * 10 + 1)
Case Is = 1
Response = Word1
Case Is = 2
Response = Word2
Case Is = 3
Response = Word3
Case Is = 4
Response = Word4
Case Is = 5
Response = Word5
Case Is = 6
Response = Word6
Case Is = 7
Response = Word7
Case Is = 8
Response = Word8
Case Is = 9
Response = Word9
Case Is = 10
Response = Word10
End Select
End Sub

Private Sub Cavemen()
Call RandWord("DUCKS", "TREES", "JUICEBOXES", "PRESIDENTS", "ANVILS", "TOILETS", "PENCILS", "STAPLERS", "GRUMPS", "CHEESEMONKEYS")
Word(1) = Response
Call RandWord("JELL-O", "TIN", "LATEX", "NYLON", "KRYPTONITE", "CRISCO", "WAX", "LINT", "ASPIRIN", "WOOD")
Word(2) = Response
Call RandWord("BOXES", "CARS", "TANKS", "BLADDERS", "BUCKETS", "TEACUPS", "APARTMENTS", "HOLES", "LAKES", "BAGGIES")
Word(3) = Response
Call RandWord("FLEXIBLE", "RIDICULOUS", "FLABBY", "ORANGE", "STICKY", "MULTIPLE", "HOT", "RUBBER", "GRUMPY", "DISGUSTING")
Word(4) = Response
Call RandWord("THUG", "TEDDY BEAR", "BOULDER", "BRONTOSAURUS", "GERBIL", "TENNIS BALL", "PIRHANA", "GECKO", "PARAKEET", "POLITICIAN")
Word(5) = Response
Call RandWord("THIN", "RAGGED", "MEATY", "ENERGETIC", "SQUASHED", "GLASS", "FRAGMENTED", "REALISTIC", "CRAZY", "GOOFY")
Word(6) = Response
Call RandWord("SCISSORS", "PENCILS", "PAPERCLIPS", "CARROTS", "SKEWERS", "MARKERS", "ANTENNAS", "RAZORS", "AXES", "EX-ACTO KNIVES")
Word(7) = Response
Call RandWord("EARWAX", "EGGPLANTS", "UNDERWEAR", "SOFAS", "PAPYRUS", "ANTIQUES", "FISH HEADS", "NOSEPLUGS", "HAIR", "SLOTHS")
Word(8) = Response
lblDisplay.Caption = "CAVEMEN WERE PRIMITIVE " & Word(1) & " WHO LIVED DURING THE " & Word(2) & " AGE.  THEY LIVED IN " & Word(3) & _
" FOR PROTECTION FROM THEIR ENEMIES, THE " & Word(4) & "-TOOTHED TIGER AND THE WOOLY " & Word(5) & ".  THEY SPEARED THESE " & Word(6) & _
" ANIMALS ON LONG " & Word(7) & " AND ROASTED THEM OVER A CAMPFIRE MADE OF " & Word(8) & "."
End Sub

Private Sub SpookyMovie()
Call RandWord("DIPPY", "DESPOTIC", "ODIFEROUS", "JUICY", "LAZY", "FRESH", "CRISPY", "SHARP", "TASTELESS", "CHEESY")
Word(1) = Response
Call RandWord("BARNEY", "BATMAN", "BILL NYE", "ARNOLD SJHUWURRZENEGGER", "DR. SEUSS", "DUMBLEDORE", "BOB", "THE INVISIBLE MAN", "THE COOK", "A CHEESE")
Word(2) = Response
Call RandWord("SLUG", "BULLETPROOF VEST", "SHOE", "BRUSSELS SPROUT", "ENCHILADA", "CHRISTMAS TREE", "BEARD", "DUST BUNNY", "SNEEZE", "GRANDPAPA")
Word(3) = Response
Call RandWord("NOSE", "TONGUE", "STOMACH", "SCALP", "NECK", "FINGER", "BUTT", "HEEL", "FACE", "KNEES")
Word(4) = Response
Call RandWord("TOMATO", "TRAIN", "TELEPHONE", "ONION", "PAPERWEIGHT", "COOKIE", "PIRHANA", "GECKO", "PARAKEET", "POLITICIAN")
Word(5) = Response
Call RandWord("FISH", "BLOOD", "ELVES", "ELVIS", "MOSES", "LITTLE THINGS WITH LEGS", "BUSES", "HALLOWEEN", "DENTISTS", "BARF")
Word(6) = Response
Call RandWord("LITTLE PIGS", "BEARS", "MUSKETEERS", "WISHES", "FAIRIES", "OUT OF 7 DWARVES", "KINGS OF ORIENTAR", "MAGI", "GHOSTS", "TELEMARKETERS")
Word(7) = Response
Call RandWord("SANTA CLAUS", "MR. HYDE", "CAPTAIN UNDERPANTS", "AN INANIMATE OBJECT", "HUITZILOPOCHTLI", "MARK TWAIN", "GOD", "THE GUY DOWN THE STREET", "MY DOG", "A SLOTH")
Word(8) = Response
Call RandWord("BEAN", "SPHERE", "ROCK", "PUNCH", "KISS", "BOWL", "LIZARD", "BEZOAR", "TRICHYGLYPHLOPOD", "ARROW")
Word(9) = Response

lblDisplay.Caption = "LAST WEEK I WENT TO SEE A VERY " & Word(1) & " MOVIE.  IT STARRED " & Word(2) & _
    " AS THE HERO WHO IS TRANSFORMED INTO A GIANT " & Word(3) & " WHEN HE WAS BITTEN ON THE " & Word(4) & _
    " BY A VICIOUS " & Word(5) & ".  AFTER THIS HAPPENS, HE STARTS HAVING DREAMS ABOUT " & Word(6) & _
    " AND WHEN HE WALKS IN HIS SLEEP HE KILLS THREE " & Word(7) & ".  AFTER HE REALIZES WHAT'S HAPPENING, HE GOES TO A DOCTOR, PLAYED BY " _
    & Word(8) & " WHO GIVES HIM A MAGIC " & Word(9) & " WHICH CURES 3/5 OF HIM.  UNFORTUNATELY, THE REST OF HIM IS STILL A " & Word(3) & _
    ".  I'M LOOKING FORWARD TO THE SEQUEL."

End Sub

Private Sub mnuSelectCav_Click()
MadLib = "CAVEMEN"
End Sub

Private Sub mnuSelectRSM_Click()
MadLib = "SPOOKY MOVIE"
End Sub
