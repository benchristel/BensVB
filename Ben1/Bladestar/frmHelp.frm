VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bladestar Help"
   ClientHeight    =   5895
   ClientLeft      =   7485
   ClientTop       =   8115
   ClientWidth     =   4335
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   4335
   Begin VB.CommandButton Command1 
      Caption         =   "Exit Help"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdControls 
      Caption         =   "Controls"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdGeneral 
      Caption         =   "General"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label lblHelpText 
      Caption         =   "Welcome to Bladestar Help.  Click on one of the categories below."
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdControls_Click()
lblHelpText = "Player 1:" & vbLf & vbLf & "Move Forward: I" & vbLf & "Move Backward: K" & vbLf & "Turn Left: J" & vbLf & "Turn Right: L" & vbLf & "Strafe Left: U" & vbLf & "Strafe Right: O" & vbLf & "Reload/Pick Up: P" & vbLf & "Fire/Club: spacebar" & vbLf & "Scope: C" & vbLf & vbLf & _
    "Player 2:" & vbLf & vbLf & "Move Forward: Numpad 8" & vbLf & "Move Backward: Numpad 5" & vbLf & "Turn Left: Numpad 4" & vbLf & "Turn Right: Numpad 6" & vbLf & "Strafe Left: Numpad 7" & vbLf & "Strafe Right: Numpad 9" & vbLf & "Reload/Pick Up: Numpad +" & vbLf & "Fire/Club: Numpad 0" & vbLf & "Scope: Numpad 1" & vbLf & _
    vbLf & "NOTE: Make sure that Num Lock is ON.  Otherwise, the controls for Player 2 will not work!"
    
End Sub

Private Sub cmdGeneral_Click()
lblHelpText = "-You can change the name of a player by double-clicking the player's current name and typing in a new name." & vbLf & _
    "-The highlighted player names on each side of the main screen indicate which players will be competing in the next round.  To change the active player on a side, click once on the player's name.  If a name is greyed-out and unavailable, it is because that player is selected on the other side." & vbLf & _
    "-The numbers to the right of the player names display wins/losses.  This is useful for tournaments where you can have up to ten players competing for the most wins." & vbLf & _
    "-You can change the number of points needed to win for each player.  Simply highlight the number labelled 'score to win', and type in the new number of points.  You cannot set this number below 1 or above 100." & vbLf & _
    "-There are currently three maps available for play in Bladestar: Oracle, Giza, and Quicksilver.  More will be added in later versions.  To select the map you want to play, click on its name.  The selected map will be highlighted." & vbLf & _
    "-There are four modes of play available: Normal, Grenades, Sabers, and Hardcore.  To select a variant, click on its name.  The selected variant will be highlighted."
End Sub

Private Sub Command1_Click()
Unload frmHelp
End Sub

