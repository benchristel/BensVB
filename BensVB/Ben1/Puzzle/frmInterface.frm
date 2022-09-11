VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEnter 
      Height          =   915
      Left            =   5880
      TabIndex        =   1
      Top             =   60
      Width           =   1995
   End
   Begin VB.Label lblInventory 
      Caption         =   "Key"
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   4
      Top             =   1620
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label lblInventory 
      Caption         =   "Crucible"
      Height          =   255
      Index           =   2
      Left            =   5880
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label lblInventory 
      Caption         =   "Gem of Power"
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   2
      Top             =   1020
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label lblDisplay 
      Height          =   5355
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5715
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Inventory(0 To 10) As Boolean 'determines whether or not a particular item is being carried by the player

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim i
If KeyAscii = 13 Then 'enter key
For i = 1 To RoomData(ActiveRoom).Links
If LCase(Trim(txtEnter.Text)) = RoomData(ActiveRoom).LinkText(i) Then
'If RoomData(ActiveRoom).LinkItem(i) = 0 Or Inventory(RoomData(ActiveRoom).LinkItem(i)) = True Then
'If RoomData(RoomData(ActiveRoom).Link(i)).LoseItem = True Then
'Inventory(RoomData(ActiveRoom).LinkItem(i)) = False
'lblInventory(RoomData(ActiveRoom).LinkItem(i)).Visible = False
'End If
ActiveRoom = RoomData(ActiveRoom).Link(i)
Scored(ActiveRoom) = True
Call CalculateScore
Exit For
End If
Next i
lblDisplay.Caption = RoomData(ActiveRoom).Text
txtEnter.Text = ""
If ActiveRoom = 6 Or ActiveRoom = 18 Or ActiveRoom = 19 Then lblDisplay.Caption = lblDisplay.Caption & "YOUR SCORE IS " & Score & " OUT OF 100."
End If
End Sub

Private Sub Form_Load()
Call InitiateRooms
'RoomData(1).Text = "  This sanctuary for the dead was abandoned long ago, but the shades still lurk among the vaulted chambers of the innermost tombs.  Only the dead have time to learn the intricate secrets of this place, for there are no teachers.  the Enigma says something different to everyone, and the interpretations would outsize the annals of all the history of Khazoum -- if any of the pupils survived." & vbLf & "  Now the door stands -- locked and unopened.  The magic that permeates this place has waited for a worthy adventurer for a thousand years.  It can wait a little more."
'RoomData(1).Links = 2
'ReDim RoomData(1).Link(1 To RoomData(1).Links)
'ReDim RoomData(1).LinkText(1 To RoomData(1).Links)
'ReDim RoomData(1).LinkItem(1 To RoomData(1).Links)
'RoomData(1).Link(1) = 2
'RoomData(1).LinkText(1) = "examine door"
'RoomData(1).LinkItem(1) = 0
'RoomData(1).Link(2) = 3
'RoomData(1).LinkText(2) = "go east"
'RoomData(1).LinkItem(2) = 0
'RoomData(1).GetItem = 0
'RoomData(1).LoseItem = False
''''<<<Room 2>>>'''
'RoomData(2).Text = "  The heavy bronzework of this ancient portal is weathered and scarred by the sands of time.  You can't quite make out the inscription above the door, but the runes have a sinister look.  You sense the ancient magic hovering around you.  When you step away from the door, you realize you hold the Gem of Power."
'RoomData(2).Links = 2
'ReDim RoomData(2).Link(1 To RoomData(2).Links)
'ReDim RoomData(2).LinkText(1 To RoomData(2).Links)
'ReDim RoomData(2).LinkItem(1 To RoomData(2).Links)
'RoomData(2).Link(1) = 1
'RoomData(2).LinkText(1) = "go back"
'RoomData(2).LinkItem(1) = 0
'RoomData(2).Link(2) = 8
'RoomData(2).LinkText(2) = "use key"
'RoomData(2).LinkItem(2) = 3
'RoomData(2).GetItem = 1
'RoomData(2).LoseItem = False
'lblDisplay.Caption = RoomData(1).Text
'ActiveRoom = 1
''''<<<Room 3>>>'''
'RoomData(3).Text = "  Among the skeletal trees lie a discarded sarcophagus and an elaborately carved stone altar.  Judging by the intricate designs tediously engraved into the stone, the altar predates the Tombs themselves.  This place may have been an alchemist's sanctuary in times long past."
'RoomData(3).Links = 3
'ReDim RoomData(3).Link(1 To RoomData(3).Links)
'ReDim RoomData(3).LinkText(1 To RoomData(3).Links)
'ReDim RoomData(3).LinkItem(1 To RoomData(3).Links)
'RoomData(3).Link(1) = 1
'RoomData(3).LinkText(1) = "go west"
'RoomData(3).LinkItem(1) = 0
'RoomData(3).Link(2) = 4
'RoomData(3).LinkText(2) = "examine sarcophagus"
'RoomData(3).LinkItem(2) = 0
'RoomData(3).Link(3) = 5
'RoomData(3).LinkText(3) = "examine altar"
'RoomData(3).LinkItem(3) = 0
'RoomData(3).GetItem = 0
'RoomData(3).LoseItem = False
''''<<<Room 4>>>'''
'RoomData(4).Text = "  You heave open the stone lid of the sarcophagus, but inside there is nothing but darkness.  You reach cautiously into the casket and find that, although your hand disappears into the thick blackness, you can feel the stone bottom.  But you can feel something else.  Slowly, you lift the ancient bronze crucible out of the void."
'RoomData(4).Links = 1
'ReDim RoomData(4).Link(1 To RoomData(4).Links)
'ReDim RoomData(4).LinkText(1 To RoomData(4).Links)
'ReDim RoomData(4).LinkItem(1 To RoomData(4).Links)
'RoomData(4).Link(1) = 3
'RoomData(4).LinkText(1) = "go back"
'RoomData(4).LinkItem(1) = 0
'RoomData(4).GetItem = 2
'RoomData(4).LoseItem = False
''''<<<Room 5>>>'''
'RoomData(5).Text = "  The surface of the altar is engraved with necromantic symbols whose purposes you can only guess at.  The only smooth part of the surface is the circular indentation in the center, marred by a crude carving of a key.  You don't know what this altar was used for in ancient times, but the magic emenating from it is strong.  The rites performed here were doubtless terrible and great."
'RoomData(5).Links = 2
'ReDim RoomData(5).Link(1 To RoomData(5).Links)
'ReDim RoomData(5).LinkText(1 To RoomData(5).Links)
'ReDim RoomData(5).LinkItem(1 To RoomData(5).Links)
'RoomData(5).Link(1) = 3
'RoomData(5).LinkText(1) = "go back"
'RoomData(5).LinkItem(1) = 0
'RoomData(5).Link(2) = 6
'RoomData(5).LinkText(2) = "use crucible"
'RoomData(5).LinkItem(2) = 2
'RoomData(5).GetItem = 0
'RoomData(5).LoseItem = False
''''<<<Room 6>>>'''
'RoomData(6).Text = "  The crucible fits neatly into the indentation in the altar.  It begins to glow steadily, and you realize it will evaporate if left here too long."
'RoomData(6).Links = 2
'ReDim RoomData(6).Link(1 To RoomData(6).Links)
'ReDim RoomData(6).LinkText(1 To RoomData(6).Links)
'ReDim RoomData(6).LinkItem(1 To RoomData(6).Links)
'RoomData(6).Link(1) = 3
'RoomData(6).LinkText(1) = "go back"
'RoomData(6).LinkItem(1) = 0
'RoomData(6).Link(2) = 7
'RoomData(6).LinkText(2) = "use gem of power"
'RoomData(6).LinkItem(2) = 1
'RoomData(6).GetItem = 0
'RoomData(6).LoseItem = True
''''<<<Room 7>>>'''
'RoomData(7).Text = "  As you drop the gem into the crucible, they both begin to vibrate and then melt, finally becoming a pool of molten metal and stone.  the molten bronze pools on top and vaporizes, and the melted gemstone oozes into the form of a key.  As soon as it's cool enough to touch, you thrust it into your pocket."
'RoomData(7).Links = 1
'ReDim RoomData(7).Link(1 To RoomData(7).Links)
'ReDim RoomData(7).LinkText(1 To RoomData(7).Links)
'ReDim RoomData(7).LinkItem(1 To RoomData(7).Links)
'RoomData(7).Link(1) = 3
'RoomData(7).LinkText(1) = "go back"
'RoomData(7).LinkItem(1) = 0
'RoomData(7).GetItem = 3
'RoomData(7).LoseItem = True
''''<<<Room 8>>>'''
'RoomData(8).Text = "  You look around the antechamber of the necropolis.  Nothing moves, and you cannot hear even the slightest sound.  To the west, a single high window lets the last rays of the dying sun in to illuminate part of the room.  The eastern half is shrouded in ghastly darkness.  But there is something else.  The beam of light from the setting sun illuminates a decrepit dais."
'RoomData(8).Links = 1
'ReDim RoomData(8).Link(1 To RoomData(8).Links)
'ReDim RoomData(8).LinkText(1 To RoomData(8).Links)
'ReDim RoomData(8).LinkItem(1 To RoomData(8).Links)
'RoomData(8).Link(1) = 9
'RoomData(8).LinkText(1) = "go south"
'RoomData(8).LinkItem(1) = 0
'RoomData(8).Link(1) = 10
'RoomData(8).LinkText(1) = "go east"
'RoomData(8).LinkItem(1) = 4
'RoomData(8).GetItem = 0
'RoomData(8).LoseItem = False
lblDisplay.Caption = RoomData(ActiveRoom).Text
End Sub

