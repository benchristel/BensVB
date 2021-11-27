VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAmalim 
      Caption         =   "Ask Dalboz"
      Height          =   555
      Left            =   1620
      TabIndex        =   3
      Top             =   2580
      Width           =   1515
   End
   Begin VB.CommandButton cmdSwami 
      Caption         =   "Ask the Swami"
      Height          =   555
      Left            =   60
      TabIndex        =   2
      Top             =   2580
      Width           =   1515
   End
   Begin VB.TextBox txtEnter 
      Height          =   915
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   60
      Width           =   4575
   End
   Begin VB.Label lblDisplay 
      Caption         =   "The swami knows all.  Type your qvestion."
      Height          =   1995
      Left            =   60
      TabIndex        =   0
      Top             =   1140
      Width           =   4575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Random
Private Sub Text1_Change()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdAmalim_Click()
If LCase(Mid(txtEnter.Text, 1, 11)) = "who are you" Then
lblDisplay.Caption = "I'm Dalboz of Gurth."
Exit Sub
End If
'
'WHO QUESTIONS
'
If LCase(Mid(txtEnter.Text, 1, 3)) = "who" Then
If LCase(Mid(txtEnter.Text, 4, 4)) = " are" Or LCase(Mid(txtEnter.Text, 4, 4)) = " were" Then
Random = Int(Rnd * 15 + 1)
Select Case Random
    Case Is = 0
    lblDisplay.Caption = "Squirrels."
    Case Is = 1
    lblDisplay.Caption = "I'd say the dwarves, but they've been defeated."
    Case Is = 2
    lblDisplay.Caption = "Mike and his pants."
    Case Is = 3
    lblDisplay.Caption = "The Accardi chapter of the Guild of Enchanters."
    Case Is = 4
    lblDisplay.Caption = "The King and Queen of Vestigia."
    Case Is = 5
    lblDisplay.Caption = "The Huns."
    Case Is = 6
    lblDisplay.Caption = "Rat-ants."
    Case Is = 7
    lblDisplay.Caption = "A large amount of carpenters."
    Case Is = 8
    lblDisplay.Caption = "One more like that and you're totemized, buddy."
    Case Is = 9
    lblDisplay.Caption = "I must decline to state."
    Case Is = 10
    lblDisplay.Caption = "The army of Tuulius Pompus."
    Case Is = 11
    lblDisplay.Caption = "Will someone bring in the swami?"
    Case Is = 12
    lblDisplay.Caption = "Lord Dimwit the Fifth and your auntie Grazelda."
    Case Is = 13
    lblDisplay.Caption = "Mr. and Mrs. Subothrumbus"
    Case Is = 14
    lblDisplay.Caption = "The Dunedain."
    Case Is = 15
    lblDisplay.Caption = "Hobbits."
End Select
Else
Random = Int(Rnd * 15 + 1)
Select Case Random
    Case Is = 0
    lblDisplay.Caption = "I dunno... Maybe Obi-Wan Kenobi."
    Case Is = 1
    lblDisplay.Caption = "Yo momma!"
    Case Is = 2
    lblDisplay.Caption = "Geeorge W. Boosh."
    Case Is = 3
    lblDisplay.Caption = "Dunno.  Some celebrity or other."
    Case Is = 4
    lblDisplay.Caption = "I believe that would be the Grand Inquisitor."
    Case Is = 5
    lblDisplay.Caption = "Me."
    Case Is = 6
    lblDisplay.Caption = "No, no, no, Who's on first."
    Case Is = 7
    lblDisplay.Caption = "Your mother's father's sister's nephew's cousin's son's former roommate's best friend's gardener."
    Case Is = 8
    lblDisplay.Caption = "One more like that and you're totemized, buddy."
    Case Is = 9
    lblDisplay.Caption = "I must decline to state."
    Case Is = 10
    lblDisplay.Caption = "Tuulius Pompus."
    Case Is = 11
    lblDisplay.Caption = "Will someone bring in the swami?"
    Case Is = 12
    lblDisplay.Caption = "Bizboz...or else his granny."
    Case Is = 13
    lblDisplay.Caption = "Benjamin Franklin."
    Case Is = 14
    lblDisplay.Caption = "Elessar, Elessar."
    Case Is = 15
    lblDisplay.Caption = "Sauron, Lord of Mordor, Forger of the Ruling Ring."
End Select
End If
Exit Sub
End If
'
'WHERE QUESTIONS
'
If LCase(Mid(txtEnter.Text, 1, 5)) = "where" Or LCase(Mid(txtEnter.Text, 1, 6)) = "whence" _
Or LCase(Mid(txtEnter.Text, 1, 7)) = "whither" Then
Random = Int(Rnd * 15 + 1)
Select Case Random
    Case Is = 0
    lblDisplay.Caption = "Madagascar."
    Case Is = 1
    lblDisplay.Caption = "Cape Cod."
    Case Is = 2
    lblDisplay.Caption = "Texas, Land of Geeorge W. Boosh."
    Case Is = 3
    lblDisplay.Caption = "Dunno.  Probably a volcanic island."
    Case Is = 4
    lblDisplay.Caption = "Port Foozle."
    Case Is = 5
    lblDisplay.Caption = "My house."
    Case Is = 6
    lblDisplay.Caption = "Panama, Where Everything Smells like Bananas."
    Case Is = 7
    lblDisplay.Caption = "A long time ago in a galaxy far, far away."
    Case Is = 8
    lblDisplay.Caption = "One more like that and you're totemized, buddy."
    Case Is = 9
    lblDisplay.Caption = "I must decline to state."
    Case Is = 10
    lblDisplay.Caption = "Rome, obviously."
    Case Is = 11
    lblDisplay.Caption = "Will someone bring in the swami?"
    Case Is = 12
    lblDisplay.Caption = "GUE TECH."
    Case Is = 13
    lblDisplay.Caption = "My garden."
    Case Is = 14
    lblDisplay.Caption = "Gondor, Gondor, Between the Mountains and the Sea!"
    Case Is = 15
    lblDisplay.Caption = "MORDOR!"
End Select
Exit Sub
End If
'
'HOW DO YOU QUESTIONS
'
If LCase(Mid(txtEnter.Text, 1, 10)) = "how do you" Or LCase(Mid(txtEnter.Text, 1, 12)) = "how does one" _
Or LCase(Mid(txtEnter.Text, 1, 8)) = "how do i" Or LCase(Mid(txtEnter.Text, 1, 9)) = "how can i" _
Or LCase(Mid(txtEnter.Text, 1, 11)) = "how can you" Or LCase(Mid(txtEnter.Text, 1, 11)) = "how can one" Or LCase(Mid(txtEnter.Text, 1, 3)) = "how" Then
Random = Int(Rnd * 15 + 1)
Select Case Random
    Case Is = 0
    lblDisplay.Caption = "Very sloppily, I can assure you."
    Case Is = 1
    lblDisplay.Caption = "Badly."
    Case Is = 2
    lblDisplay.Caption = "First, find three blind mice, add water, and stir."
    Case Is = 3
    lblDisplay.Caption = "Put the lime in the coconut and call me in the morning."
    Case Is = 4
    lblDisplay.Caption = "Don't."
    Case Is = 5
    lblDisplay.Caption = "Why would you want to?"
    Case Is = 6
    lblDisplay.Caption = "With an edible monkey."
    Case Is = 7
    lblDisplay.Caption = "With a great deal of cheese."
    Case Is = 8
    lblDisplay.Caption = "Take one bowl of granola."
    Case Is = 9
    lblDisplay.Caption = "I must decline to state."
    Case Is = 10
    lblDisplay.Caption = "Dangerously."
    Case Is = 11
    lblDisplay.Caption = "Will someone bring in the swami?"
    Case Is = 12
    lblDisplay.Caption = "First, become king of the world by beating the grand inquisitor in a burping contest."
    Case Is = 13
    lblDisplay.Caption = "Gain 86 experience points and find a healing potion."
    Case Is = 14
    lblDisplay.Caption = "Pretend you're cool just because you have a sloppily glued sword."
    Case Is = 15
    lblDisplay.Caption = "Once you've eaten your crackers, do not hesitate to swallow."
End Select
Exit Sub
End If
'
'WHY QUESTIONS
'
If LCase(Mid(txtEnter.Text, 1, 3)) = "why" Then
Random = Int(Rnd * 15 + 1)
Select Case Random
    Case Is = 0
    lblDisplay.Caption = "Because the sum of the squares of the legs of any right triangle is equal to the square of the hypotenuse."
    Case Is = 1
    lblDisplay.Caption = "Because you are a fool."
    Case Is = 2
    lblDisplay.Caption = "Because all dung tastes the same with your eyes closed."
    Case Is = 3
    lblDisplay.Caption = "Because you are fat."
    Case Is = 4
    lblDisplay.Caption = "Because I pulled Moses's beard."
    Case Is = 5
    lblDisplay.Caption = "Because strange little kids often strangle people"
    Case Is = 6
    lblDisplay.Caption = "Because monkeys are edible."
    Case Is = 7
    lblDisplay.Caption = "For the same reason that milk has an expiration date on it."
    Case Is = 8
    lblDisplay.Caption = "For the same reason that evil people never get executed."
    Case Is = 9
    lblDisplay.Caption = "For the same reason that the earth is flat."
    Case Is = 10
    lblDisplay.Caption = "For the same reason that if you eat a smashed fly, it will come back to haunt you one day."
    Case Is = 11
    lblDisplay.Caption = "Will someone bring in the swami?"
    Case Is = 12
    lblDisplay.Caption = "Because the grand inquisitor's breath stinks."
    Case Is = 13
    lblDisplay.Caption = "Because dwarves are immortal."
    Case Is = 14
    lblDisplay.Caption = "Because the men of Westernesse are sloppy smiths."
    Case Is = 15
    lblDisplay.Caption = "Because excessive swallowing will kill you."
End Select
Exit Sub
End If
'
'WHICH QUESTIONS
'
If LCase(Mid(txtEnter.Text, 1, 5)) = "which" Then
    lblDisplay.Caption = "The one right behind you"
Exit Sub
End If
'
'WHAT IS QUESTIONS
'
If LCase(Mid(txtEnter.Text, 1, 7)) = "what is" Or LCase(Mid(txtEnter.Text, 1, 7)) = "what am" Or LCase(Mid(txtEnter.Text, 1, 12)) = "what are you" Or LCase(Mid(txtEnter.Text, 1, 13)) = "what art thou" Then
Random = Int(Rnd * 8 + 1)
Select Case Random
    Case Is = 0
    lblDisplay.Caption = "A toasted chicken."
    Case Is = 1
    lblDisplay.Caption = "An egg salad sandwich."
    Case Is = 2
    lblDisplay.Caption = "The scientific term for a disembodied head."
    Case Is = 3
    lblDisplay.Caption = "An ancient chinese bowling pin."
    Case Is = 4
    lblDisplay.Caption = "A funny little man with odd habits."
    Case Is = 5
    lblDisplay.Caption = "One of the many slaves of Santa Claus."
    Case Is = 6
    lblDisplay.Caption = "A cheesemonkey."
    Case Is = 7
    lblDisplay.Caption = "One of the rat-catchers of Minas Tirith."
    Case Is = 8
    lblDisplay.Caption = "An exhumed rodent."
End Select
Exit Sub
End If
'
'YES/NO QUESTIONS:
'
Random = Int(Rnd * 15)
Select Case Random
    Case Is = 0
    lblDisplay.Caption = "Let me look in the enclicklopedia...yup, yup, yup."
    Case Is = 1
    lblDisplay.Caption = "Well, all I'll say is your luck ain't gonna hold out."
    Case Is = 2
    lblDisplay.Caption = "Ho, ho, ho...yes."
    Case Is = 3
    lblDisplay.Caption = "BUT OF COURSE...not."
    Case Is = 4
    lblDisplay.Caption = "If I said yes, would you slap me?"
    Case Is = 5
    lblDisplay.Caption = "Oh, yes, undoubtedly so."
    Case Is = 6
    lblDisplay.Caption = "DOES A HUNGUS HAVE STRIPES?  IS THE GRAND INQUISITOR CATHOLIC?   yes, yes, yes!"
    Case Is = 7
    lblDisplay.Caption = "Uh-huh, uh-huh, it really is true."
    Case Is = 8
    lblDisplay.Caption = "One more like that and you're totemized, buddy."
    Case Is = 9
    lblDisplay.Caption = "I must decline to state."
    Case Is = 10
    lblDisplay.Caption = "Hargh, hargh. yes, gngngngngngngngngn."
    Case Is = 11
    lblDisplay.Caption = "Will someone bring in the swami?"
    Case Is = 12
    lblDisplay.Caption = "If I'm not allowed to hazard a guess, I'll throck you.  I'd say yes."
    Case Is = 13
    lblDisplay.Caption = "Why, yes.  You must be an almanac."
    Case Is = 14
    lblDisplay.Caption = "Frankly, yes."
    Case Is = 15
    lblDisplay.Caption = "You ask too many questions.  Well all right.  No."
    Case Is = 16
    lblDisplay.Caption = "Yes, no, maybe so, yes, no, maybe so, yes, no, maybe so..."
End Select
End Sub




Private Sub cmdSwami_Click()
If LCase(Mid(txtEnter.Text, 1, 11)) = "who are you" Then
lblDisplay.Caption = "I'm the Swami."
Exit Sub
End If

'
'WHO QUESTIONS
'
If LCase(Mid(txtEnter.Text, 1, 3)) = "who" Then
If LCase(Mid(txtEnter.Text, 4, 4)) = " are" Or LCase(Mid(txtEnter.Text, 4, 4)) = " were" Then
Random = Int(Rnd * 15 + 1)
Select Case Random
    Case Is = 0
    lblDisplay.Caption = "Squirrels."
    Case Is = 1
    lblDisplay.Caption = "I'd say the dwarves, but they've been defeated."
    Case Is = 2
    lblDisplay.Caption = "Mike and his pants."
    Case Is = 3
    lblDisplay.Caption = "The Accardi chapter of the Guild of Enchanters."
    Case Is = 4
    lblDisplay.Caption = "The King and Queen of Vestigia."
    Case Is = 5
    lblDisplay.Caption = "The Huns."
    Case Is = 6
    lblDisplay.Caption = "Rat-ants."
    Case Is = 7
    lblDisplay.Caption = "A large amount of carpenters."
    Case Is = 8
    lblDisplay.Caption = "One more like that and you're totemized, buddy."
    Case Is = 9
    lblDisplay.Caption = "I must decline to state."
    Case Is = 10
    lblDisplay.Caption = "The army of Tuulius Pompus."
    Case Is = 11
    lblDisplay.Caption = "Will someone bring in the swami?"
    Case Is = 12
    lblDisplay.Caption = "Lord Dimwit the Fifth and your auntie Grazelda."
    Case Is = 13
    lblDisplay.Caption = "Mr. and Mrs. Subothrumbus"
    Case Is = 14
    lblDisplay.Caption = "The Dunedain."
    Case Is = 15
    lblDisplay.Caption = "Hobbits."
End Select
Else
Random = Int(Rnd * 15 + 1)
Select Case Random
    Case Is = 0
    lblDisplay.Caption = "I dunno... Maybe Obi-Wan Kenobi."
    Case Is = 1
    lblDisplay.Caption = "Yo daddy!"
    Case Is = 2
    lblDisplay.Caption = "Geeorge W. Boosh."
    Case Is = 3
    lblDisplay.Caption = "Dunno.  Some celebrity or other."
    Case Is = 4
    lblDisplay.Caption = "I believe that would be the Grand Mogul."
    Case Is = 5
    lblDisplay.Caption = "Me."
    Case Is = 6
    lblDisplay.Caption = "No, no, no, Who's on first."
    Case Is = 7
    lblDisplay.Caption = "Your mother's father's sister's nephew's cousin's son's former roommate's best friend's gardener."
    Case Is = 8
    lblDisplay.Caption = "One more like that and you're stir-fried, buddy."
    Case Is = 9
    lblDisplay.Caption = "I must decline to state."
    Case Is = 10
    lblDisplay.Caption = "Tuulius Pompus."
    Case Is = 11
    lblDisplay.Caption = "Will someone bring in the swami?"
    Case Is = 12
    lblDisplay.Caption = "Bizboz...or else his granny."
    Case Is = 13
    lblDisplay.Caption = "Benjamin Franklin."
    Case Is = 14
    lblDisplay.Caption = "Elessar, Elessar."
    Case Is = 15
    lblDisplay.Caption = "Sauron, Lord of Mordor, Forger of the Ruling Ring."
End Select
End If
Exit Sub
End If
'
'WHERE QUESTIONS
'
If LCase(Mid(txtEnter.Text, 1, 5)) = "where" Or LCase(Mid(txtEnter.Text, 1, 6)) = "whence" _
Or LCase(Mid(txtEnter.Text, 1, 7)) = "whither" Then
Random = Int(Rnd * 15 + 1)
Select Case Random
    Case Is = 0
    lblDisplay.Caption = "Sri Lanka."
    Case Is = 1
    lblDisplay.Caption = "The Cape of Good Hope."
    Case Is = 2
    lblDisplay.Caption = "Mass-of-chew-sits, Land of John Kerry."
    Case Is = 3
    lblDisplay.Caption = "Dunno.  Probably a volcanic island."
    Case Is = 4
    lblDisplay.Caption = "Port Foozle."
    Case Is = 5
    lblDisplay.Caption = "My house."
    Case Is = 6
    lblDisplay.Caption = "Franistan, Where Everything Smells like Latex."
    Case Is = 7
    lblDisplay.Caption = "A long time ago in farthest Asia."
    Case Is = 8
    lblDisplay.Caption = "One more like that and you're stir-fried, buddy."
    Case Is = 9
    lblDisplay.Caption = "I must decline to state."
    Case Is = 10
    lblDisplay.Caption = "Carthage, obviously."
    Case Is = 11
    lblDisplay.Caption = "Will someone bring in Dalboz?"
    Case Is = 12
    lblDisplay.Caption = "In the land of the kangaroo."
    Case Is = 13
    lblDisplay.Caption = "My rock garden."
    Case Is = 14
    lblDisplay.Caption = "Mongolia, Mongolia, Between the Mountains and the Sea!"
    Case Is = 15
    lblDisplay.Caption = "AFGHANISTAN!"
End Select
Exit Sub
End If
'
'HOW DO YOU QUESTIONS
'
If LCase(Mid(txtEnter.Text, 1, 10)) = "how do you" Or LCase(Mid(txtEnter.Text, 1, 12)) = "how does one" _
Or LCase(Mid(txtEnter.Text, 1, 8)) = "how do i" Or LCase(Mid(txtEnter.Text, 1, 9)) = "how can i" _
Or LCase(Mid(txtEnter.Text, 1, 11)) = "how can you" Or LCase(Mid(txtEnter.Text, 1, 11)) = "how can one" Or LCase(Mid(txtEnter.Text, 1, 3)) = "how" Then
Random = Int(Rnd * 15 + 1)
Select Case Random
    Case Is = 0
    lblDisplay.Caption = "Very skilfully, I can assure you."
    Case Is = 1
    lblDisplay.Caption = "Quite well."
    Case Is = 2
    lblDisplay.Caption = "First, find three dead ducks, add beansprouts, and stir-fry."
    Case Is = 3
    lblDisplay.Caption = "Put the eye of newt in a cellophane bag and call me in the morning."
    Case Is = 4
    lblDisplay.Caption = "Don't."
    Case Is = 5
    lblDisplay.Caption = "Why would you want to?"
    Case Is = 6
    lblDisplay.Caption = "With an edible vegetable."
    Case Is = 7
    lblDisplay.Caption = "With a great deal of grease."
    Case Is = 8
    lblDisplay.Caption = "Take one bowl of rice."
    Case Is = 9
    lblDisplay.Caption = "I must decline to state."
    Case Is = 10
    lblDisplay.Caption = "Proceed with caution."
    Case Is = 11
    lblDisplay.Caption = "Will someone bring in Dalboz?"
    Case Is = 12
    lblDisplay.Caption = "First, become king of the world by beating the grand mogul in a burping contest."
    Case Is = 13
    lblDisplay.Caption = "Go into the cave and bring me the lamp!"
    Case Is = 14
    lblDisplay.Caption = "Pretend you're cool just because you have a Damascus Blade."
    Case Is = 15
    lblDisplay.Caption = "Once you've taken your potion, do not hesitate to swallow."
End Select
Exit Sub
End If
'
'WHY QUESTIONS
'
If LCase(Mid(txtEnter.Text, 1, 3)) = "why" Then
Random = Int(Rnd * 15 + 1)
Select Case Random
    Case Is = 0
    lblDisplay.Caption = "Because the sum of the cubes of the legs of any table is equal to the volume of the table."
    Case Is = 1
    lblDisplay.Caption = "Because Confucius say so."
    Case Is = 2
    lblDisplay.Caption = "Because all cloves taste the same with your eyes closed."
    Case Is = 3
    lblDisplay.Caption = "Because you are fat."
    Case Is = 4
    lblDisplay.Caption = "Because I pulled Cleopatra's beard."
    Case Is = 5
    lblDisplay.Caption = "Because strange little kids often see dead people"
    Case Is = 6
    lblDisplay.Caption = "Because tofu is edible."
    Case Is = 7
    lblDisplay.Caption = "For the same reason that opium has an expiration date on it."
    Case Is = 8
    lblDisplay.Caption = "For the same reason that great greek gods almost never get executed."
    Case Is = 9
    lblDisplay.Caption = "For the same reason that the earth is round."
    Case Is = 10
    lblDisplay.Caption = "For the same reason that you have writer's cramp and reader's digest."
    Case Is = 11
    lblDisplay.Caption = "Will someone bring in Dalboz?"
    Case Is = 12
    lblDisplay.Caption = "Because the grand mogul's breath stinks."
    Case Is = 13
    lblDisplay.Caption = "Because genies are immortal."
    Case Is = 14
    lblDisplay.Caption = "Because the men of Damascus are skilled in their art."
    Case Is = 15
    lblDisplay.Caption = "Because excessive swallowing will kill you."
End Select
Exit Sub
End If
'
'WHICH QUESTIONS
'
If LCase(Mid(txtEnter.Text, 1, 5)) = "which" Then
    lblDisplay.Caption = "The one which is before you."
Exit Sub
End If
'
'WHAT IS QUESTIONS
'
If LCase(Mid(txtEnter.Text, 1, 7)) = "what is" Or LCase(Mid(txtEnter.Text, 1, 7)) = "what am" Or LCase(Mid(txtEnter.Text, 1, 12)) = "what are you" Or LCase(Mid(txtEnter.Text, 1, 13)) = "what art thou" Then
Random = Int(Rnd * 8 + 1)
Select Case Random
    Case Is = 0
    lblDisplay.Caption = "A toasted eggplant."
    Case Is = 1
    lblDisplay.Caption = "A cosmic egg."
    Case Is = 2
    lblDisplay.Caption = "The scientific term for a glowing orb."
    Case Is = 3
    lblDisplay.Caption = "An ancient chinese utensil."
    Case Is = 4
    lblDisplay.Caption = "A mighty wizard with a dark power."
    Case Is = 5
    lblDisplay.Caption = "A towering citadel."
    Case Is = 6
    lblDisplay.Caption = "A golden monkey from the ancient sacred crypts of Egypt."
    Case Is = 7
    lblDisplay.Caption = "The name of one of the thieves of Ali Baba."
    Case Is = 8
    lblDisplay.Caption = "One of the mysteries of farthest Asia."
End Select
Exit Sub
End If
'yes/no
    Random = Int(Rnd * 15 + 1)
Select Case Random
    Case Is = 0
    lblDisplay.Caption = "You crazy piff.  You're more stupid than than me to ask a question like that.  But yes."
    Case Is = 1
    lblDisplay.Caption = "No, and if you're Bobit, go live in a hole."
    Case Is = 2
    lblDisplay.Caption = "Hack, wheeze cough...yes."
    Case Is = 3
    lblDisplay.Caption = "BUT OF COURSE!!!"
    Case Is = 4
    lblDisplay.Caption = "It pains me to tell you this, but, um, no."
    Case Is = 5
    lblDisplay.Caption = "As a matter of fact, no."
    Case Is = 6
    lblDisplay.Caption = "Do I know?  DO I KNOW?  DOES A COW HAVE STRIPES?  IS THE POPE CATHOLIC?   yes, yes, yes!"
    Case Is = 7
    lblDisplay.Caption = "Yus, yus, yus.  Noo, noo, noo.  Yus, yus, yus.  Noo, noo, noo."
    Case Is = 8
    lblDisplay.Caption = "Certainly not!  Why would you think that?"
    Case Is = 9
    lblDisplay.Caption = "No comment."
    Case Is = 10
    lblDisplay.Caption = "Hargh, hargh. yes, gngngngngngngngngn."
    Case Is = 11
    lblDisplay.Caption = "Will someone bring in Afgncaap?"
    Case Is = 12
    lblDisplay.Caption = "If I am allowed to hazard a guess, I'd have to say no."
    Case Is = 13
    lblDisplay.Caption = "Why, yes.  How did you know?"
    Case Is = 14
    lblDisplay.Caption = "Frankly, no."
    Case Is = 15
    lblDisplay.Caption = "Nope, you're probably right, so I wouldn't have to decline to agree to say yes...but then again, maybe not."
End Select
End Sub




Private Sub Form_Load()
Randomize
End Sub
