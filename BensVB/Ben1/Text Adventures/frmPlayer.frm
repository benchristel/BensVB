VERSION 5.00
Begin VB.Form frmAdventure 
   Caption         =   "Text Adventure Player!!!"
   ClientHeight    =   6525
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstInventory 
      Height          =   6300
      ItemData        =   "frmPlayer.frx":0000
      Left            =   8520
      List            =   "frmPlayer.frx":0002
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "ENTER"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtNotes 
      Height          =   5175
      Left            =   6120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmPlayer.frx":0004
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txtInput 
      Height          =   735
      Left            =   6120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmPlayer.frx":0018
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblDisplay 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmAdventure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdENTER_Click()
Room = EnterText(LCase(Trim(txtInput.Text)))
lblDisplay.Caption = RoomData(Room).Text
End Sub

Private Sub mnuNew_Click()
Dim i, j, temp, MapName
MapName = LCase(InputBox("Type the name of the scenario you want to open.", "Open Scenario"))
On Error GoTo 1
Open App.Path & "\Scenarios\" & MapName & ".dat" For Input As #1
Line Input #1, temp
    MapName = temp
Line Input #1, temp
    Rooms = temp
ReDim RoomData(1 To Rooms)
For i = 1 To Rooms
Line Input #1, temp
    RoomData(i).Name = temp
Line Input #1, temp
    RoomData(i).Text = temp
For j = 0 To 31
Line Input #1, temp
    RoomData(i).Link(j) = temp
Line Input #1, temp
    RoomData(i).LinkText(j) = temp
Line Input #1, temp
    RoomData(i).ItemName(j) = temp
Line Input #1, temp
    RoomData(i).ItemFunction(j) = temp
Next j
Next i
Close #1
Room = 1
lblDisplay.Caption = RoomData(Room).Text
cmdEnter.Enabled = True
Exit Sub
1:
MsgBox "Could not find requested scenario.", , "Error"
End Sub

