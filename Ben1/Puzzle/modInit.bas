Attribute VB_Name = "Module1"
Option Explicit
Type Room
Links As Integer 'the number of links in a room
Link() As Integer 'determines what room or event to link a particular command to
LinkText() As String 'the text that the user will enter to link to the indexed command
LinkItem() As Integer 'the item that the user must have to link to the indexed command -- if a 0, no item is necessary
Index As Integer 'the number used in the link property of other rooms to identify the room or event
Item(1 To 10) As Boolean 'determines which items can be picked up in a particular room
Text As String  'the text that will be displayed
GetItem As Integer 'what item (if any) you will obtain from visiting the room
LoseItem As Boolean 'determines whether you lose the item required to access the room or event
End Type
Global Score, Scored(1 To 24) As Boolean
Global RoomData(1 To 24) As Room, ActiveRoom As Integer

'items:
'1: Gem of Power

Public Sub InitiateRooms()
'<<<Room 1: The Barrel>>>
RoomData(1).Text = "You find yourself in a barrel.  It appears to be empty (except for you, of course)."
RoomData(1).Links = 1
ReDim RoomData(1).Link(1 To RoomData(1).Links)
ReDim RoomData(1).LinkText(1 To RoomData(1).Links)
RoomData(1).Link(1) = 2
RoomData(1).LinkText(1) = "go up"
'<<<Room 2: The Dungeon>>>
RoomData(2).Text = "You are in what appears to be a dungeon.  In one wall is a high window.  A door leads through the north wall.  In a corner sleeps a troll."
RoomData(2).Links = 4
ReDim RoomData(2).Link(1 To RoomData(2).Links)
ReDim RoomData(2).LinkText(1 To RoomData(2).Links)
RoomData(2).Link(1) = 3
RoomData(2).LinkText(1) = "examine window"
RoomData(2).Link(2) = 4
RoomData(2).LinkText(2) = "examine troll"
RoomData(2).Link(3) = 5
RoomData(2).LinkText(3) = "go north"
RoomData(2).Link(4) = 6
RoomData(2).LinkText(4) = "wake up troll"
'<<<Room 3: The High Window>>>
RoomData(3).Text = "The window is unbarred, but too high for you to reach."
RoomData(3).Links = 2
ReDim RoomData(3).Link(1 To RoomData(3).Links)
ReDim RoomData(3).LinkText(1 To RoomData(3).Links)
RoomData(3).Link(1) = 2
RoomData(3).LinkText(1) = "go back"
RoomData(3).Link(2) = 7
RoomData(3).LinkText(2) = "use barrel"
'<<<Room 4: The Troll>>>
RoomData(4).Text = "It's big, green, and ugly.  Lucky for you, it's asleep."
RoomData(4).Links = 2
ReDim RoomData(4).Link(1 To RoomData(4).Links)
ReDim RoomData(4).LinkText(1 To RoomData(4).Links)
RoomData(4).Link(1) = 2
RoomData(4).LinkText(1) = "go back"
RoomData(4).Link(2) = 6
RoomData(4).LinkText(2) = "wake up troll"
'<<<Room 5: The Door>>>
RoomData(5).Text = "The door is locked."
RoomData(5).Links = 1
ReDim RoomData(5).Link(1 To RoomData(5).Links)
ReDim RoomData(5).LinkText(1 To RoomData(5).Links)
RoomData(5).Link(1) = 2
RoomData(5).LinkText(1) = "go back"
'<<<Room 6: Death by Troll>>>
RoomData(6).Text = "That was stupid.  IT'S A TROLL, FOR CRYING OUT LOUD!  What on earth were you thinking?" & vbLf & vbLf & "YOU HAVE DIED.  "
'<<<Room 7: Outside the Dungeon>>>
RoomData(7).Text = "  You are in a forest clearing.  To the east is the window of the dungeon.  On the fringes of the trees stand four altars of Earth, Water, Air, and Fire."
RoomData(7).Links = 8
ReDim RoomData(7).Link(1 To RoomData(7).Links)
ReDim RoomData(7).LinkText(1 To RoomData(7).Links)
RoomData(7).Link(1) = 8
RoomData(7).LinkText(1) = "examine earth altar"
RoomData(7).Link(2) = 9
RoomData(7).LinkText(2) = "examine water altar"
RoomData(7).Link(3) = 10
RoomData(7).LinkText(3) = "examine air altar"
RoomData(7).Link(4) = 11
RoomData(7).LinkText(4) = "examine fire altar"
RoomData(7).Link(5) = 2
RoomData(7).LinkText(5) = "go east"
RoomData(7).Link(6) = 12
RoomData(7).LinkText(6) = "go north"
RoomData(7).Link(7) = 13
RoomData(7).LinkText(7) = "go south"
RoomData(7).Link(8) = 14
RoomData(7).LinkText(8) = "go west"
'<<<Room 8: The Earth Altar>>>
RoomData(8).Text = "A cool, dark calm surrounds you.  You hear a voice, deep and slow.  It says, 'You must find the Staff of the Magi.  Do this, and worlds will open to you.'"
RoomData(8).Links = 1
ReDim RoomData(8).Link(1 To RoomData(8).Links)
ReDim RoomData(8).LinkText(1 To RoomData(8).Links)
RoomData(8).Link(1) = 7
RoomData(8).LinkText(1) = "go back"
'<<<Room 9: The Water Altar>>>
RoomData(9).Text = "You hear a rushing, as of wings or a mountain stream.  A voice speaks from a deep void, light and strong.  It says, 'You must find the Stone of Truth.  Do this, and from your mind will wisdom flow.'"
RoomData(9).Links = 1
ReDim RoomData(9).Link(1 To RoomData(9).Links)
ReDim RoomData(9).LinkText(1 To RoomData(9).Links)
RoomData(9).Link(1) = 7
RoomData(9).LinkText(1) = "go back"
'<<<Room 10: The Air Altar>>>
RoomData(10).Text = "You are plunged into a realm of cool green.  A voice shouts to you across a great distance.  It says, 'You must find the Orb of Light.  Do this, and to you will come the glory of kings.'"
RoomData(10).Links = 1
ReDim RoomData(10).Link(1 To RoomData(10).Links)
ReDim RoomData(10).LinkText(1 To RoomData(10).Links)
RoomData(10).Link(1) = 7
RoomData(10).LinkText(1) = "go back"
'<<<Room 11: The Fire Altar>>>
RoomData(11).Text = "You are plunged into a pit of intense heat and light.  A voice roars, flaming with power.  It says, 'You must find the Adamant Cube.  Do this, and from basest lead will gold be wrought.'"
RoomData(11).Links = 1
ReDim RoomData(11).Link(1 To RoomData(11).Links)
ReDim RoomData(11).LinkText(1 To RoomData(11).Links)
RoomData(11).Link(1) = 7
RoomData(11).LinkText(1) = "go back"
'<<<Room 12: The Cave Under the Stair>>>
RoomData(12).Text = "You come to the edge of the forest and find a great cliff looming up north of you.  Up the cliff winds a staircase cut into the stone.  Beneath the staircase is the entrance to a dark cave."
RoomData(12).Links = 4
ReDim RoomData(12).Link(1 To RoomData(12).Links)
ReDim RoomData(12).LinkText(1 To RoomData(12).Links)
RoomData(12).Link(1) = 7
RoomData(12).LinkText(1) = "go south"
RoomData(12).Link(2) = 19
RoomData(12).LinkText(2) = "go north"
RoomData(12).Link(3) = 7
RoomData(12).LinkText(3) = "where is the desert unblessed by breeze?"
RoomData(12).Link(4) = 20
RoomData(12).LinkText(4) = "smek juffle zuchinni"

'<<<Room 13: The Waterfall beyond the Lake>>>
RoomData(13).Text = "You hear the sound of rushing water, and soon approach a raging waterfall boiling over a cliff.  To the east stands a mighty portal, a huge burnished iron door set into the edifice of the citadel."
RoomData(13).Links = 3
ReDim RoomData(13).Link(1 To RoomData(13).Links)
ReDim RoomData(13).LinkText(1 To RoomData(13).Links)
RoomData(13).Link(1) = 7
RoomData(13).LinkText(1) = "go north"
RoomData(13).Link(2) = 7
RoomData(13).LinkText(2) = "what is the fire that does not burn?"
RoomData(13).Link(3) = 7
RoomData(13).LinkText(3) = "what is the water that will not freeze?"
'<<<Room 14: The Summoning Stone>>>
RoomData(14).Text = "You come to a forest glade.  In the center, resting upon a stump, is a slab of marble with strange designs carved into it.  You recognize it as a summoning stone."
RoomData(14).Links = 4
ReDim RoomData(14).Link(1 To RoomData(14).Links)
ReDim RoomData(14).LinkText(1 To RoomData(14).Links)
RoomData(14).Link(1) = 7
RoomData(14).LinkText(1) = "go east"
RoomData(14).Link(2) = 15
RoomData(14).LinkText(2) = "go west"
RoomData(14).Link(3) = 16
RoomData(14).LinkText(3) = "go north"
RoomData(14).Link(4) = 17
RoomData(14).LinkText(4) = "go south"
'<<<Room 15: The Wizard by the River>>>
RoomData(15).Text = "You peer through the bushes and vines and see a roaring waterfall pouring from a high cliff.  Along the riverbank walks a wizard with a crooked staff.  He appears to be trying to read a piece of weathered parchment he is holding, but can't maneage it in the dim light under the trees.  Finally, he says, 'SMEK JUFFLE ZUCHINNI'.  A ball of light appears in his hand."
RoomData(15).Links = 2
ReDim RoomData(15).Link(1 To RoomData(15).Links)
ReDim RoomData(15).LinkText(1 To RoomData(15).Links)
RoomData(15).Link(1) = 14
RoomData(15).LinkText(1) = "go east"
RoomData(15).Link(2) = 18
RoomData(15).LinkText(2) = "examine wizard"
'<<<Room 16: The Lake Under the Cliff>>>
RoomData(16).Text = "You almost slip and fall into a lake that happens to be located under a cliff."
RoomData(16).Links = 3
ReDim RoomData(16).Link(1 To RoomData(16).Links)
ReDim RoomData(16).LinkText(1 To RoomData(16).Links)
RoomData(16).Link(1) = 14
RoomData(16).LinkText(1) = "twisty me leg"
RoomData(16).Link(2) = 18
RoomData(16).LinkText(2) = "examine pool"
RoomData(16).Link(3) = 14
RoomData(16).LinkText(3) = "go south"
'<<<Room 17: The Well beneath the Tree>>>
RoomData(17).Text = "You encounter a giant tree that dwarfs the well that squats beneath it."
RoomData(17).Links = 3
ReDim RoomData(17).Link(1 To RoomData(17).Links)
ReDim RoomData(17).LinkText(1 To RoomData(17).Links)
RoomData(17).Link(1) = 14
RoomData(17).LinkText(1) = "twisty me leg"
RoomData(17).Link(2) = 18
RoomData(17).LinkText(2) = "what is the sphere that does not turn?"
RoomData(17).Link(3) = 14
RoomData(17).LinkText(3) = "go north"
'<<<Room 18: Death By Wizard>>>
RoomData(18).Text = "You start to examine the wizard, but that's as far as you get.  He lifts his staff and hits you with a blast of lightning.  You sink to your knees and mutter something idiotic before you expire." & vbLf & vbLf & "YOU HAVE DIED.  "
'<<<Room 19: Death By Rabid Cheesemonkey>>>
RoomData(19).Text = "You wander aimlessly in the darkness for some time, hopelessly lost.  However, you are spared the horrible fate of death by starvation by the local rabid cheesemonkey, who pops you into his microwave.  Unfortunately, he forgets to puncture you with a fork first so your head explodes." & vbLf & vbLf & "YOU HAVE DIED.  "
'<<<Room 20: Entering the Cave>>>
RoomData(20).Text = "As the light flares up in your hand, you step into the cave." & vbLf & vbLf & "PRESS ENTER TO CONTINUE."
RoomData(20).Links = 1
ReDim RoomData(20).Link(1 To RoomData(20).Links)
ReDim RoomData(20).LinkText(1 To RoomData(20).Links)
RoomData(20).Link(1) = 21
RoomData(20).LinkText(1) = ""
'<<<Room 21: Inside the Cave>>>
RoomData(21).Text = "The air in the cave is cool and damp.  Water drips from the ceiling, sculpting luminous stalactites and fantastic forms.  Passages branch off in every direction.  You are glad you brought your compass."
RoomData(21).Links = 4
ReDim RoomData(21).Link(1 To RoomData(21).Links)
ReDim RoomData(21).LinkText(1 To RoomData(21).Links)
RoomData(21).Link(1) = 12
RoomData(21).LinkText(1) = "go south"
RoomData(21).Link(2) = 22
RoomData(21).LinkText(2) = "go west"
RoomData(21).Link(3) = 23
RoomData(21).LinkText(3) = "go north"
RoomData(21).Link(4) = 24
RoomData(21).LinkText(4) = "go east"
'<<<Room 22: The Western Cavern>>>
RoomData(22).Text = "You come to a small cavern with a high ceiling.  "
RoomData(22).Links = 3
ReDim RoomData(22).Link(1 To RoomData(21).Links)
ReDim RoomData(22).LinkText(1 To RoomData(21).Links)
RoomData(22).Link(1) = 21
RoomData(22).LinkText(1) = "go east"
RoomData(22).Link(2) = 25
RoomData(22).LinkText(2) = "use bernathos"
RoomData(22).Link(3) = 23
RoomData(22).LinkText(3) = "examine ledge"
'<<<Final Initiation>>>
ActiveRoom = 1
Scored(1) = True
End Sub

Public Sub CalculateScore()
Score = 0
If Scored(1) = True Then Score = Score + 5
If Scored(2) = True Then Score = Score + 10
If Scored(7) = True Then Score = Score + 15
If Scored(12) = True Then Score = Score + 5
If Scored(13) = True Then Score = Score + 5
If Scored(16) = True Then Score = Score + 5
If Scored(17) = True Then Score = Score + 5
If Scored(21) = True Then Score = Score + 15
End Sub
