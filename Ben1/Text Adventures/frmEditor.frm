VERSION 5.00
Begin VB.Form frmEditor 
   Caption         =   "Emrui Adventure Editor v.1.0"
   ClientHeight    =   8145
   ClientLeft      =   5445
   ClientTop       =   3960
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   12000
   Begin VB.CommandButton cmdItemFunction 
      Caption         =   "None"
      Height          =   255
      Index           =   7
      Left            =   8280
      TabIndex        =   56
      Top             =   6000
      Width           =   1245
   End
   Begin VB.TextBox txtItemName 
      Height          =   285
      Index           =   7
      Left            =   8280
      TabIndex        =   55
      Top             =   6360
      Width           =   3615
   End
   Begin VB.CommandButton cmdItemFunction 
      Caption         =   "None"
      Height          =   255
      Index           =   6
      Left            =   8280
      TabIndex        =   54
      Top             =   5280
      Width           =   1245
   End
   Begin VB.TextBox txtItemName 
      Height          =   285
      Index           =   6
      Left            =   8280
      TabIndex        =   53
      Top             =   5640
      Width           =   3615
   End
   Begin VB.CommandButton cmdItemFunction 
      Caption         =   "None"
      Height          =   255
      Index           =   5
      Left            =   8280
      TabIndex        =   52
      Top             =   4560
      Width           =   1245
   End
   Begin VB.TextBox txtItemName 
      Height          =   285
      Index           =   5
      Left            =   8280
      TabIndex        =   51
      Top             =   4920
      Width           =   3615
   End
   Begin VB.CommandButton cmdItemFunction 
      Caption         =   "None"
      Height          =   255
      Index           =   4
      Left            =   8280
      TabIndex        =   50
      Top             =   3840
      Width           =   1245
   End
   Begin VB.TextBox txtItemName 
      Height          =   285
      Index           =   4
      Left            =   8280
      TabIndex        =   49
      Top             =   4200
      Width           =   3615
   End
   Begin VB.CommandButton cmdItemFunction 
      Caption         =   "None"
      Height          =   255
      Index           =   3
      Left            =   8280
      TabIndex        =   48
      Top             =   3120
      Width           =   1245
   End
   Begin VB.TextBox txtItemName 
      Height          =   285
      Index           =   3
      Left            =   8280
      TabIndex        =   47
      Top             =   3480
      Width           =   3615
   End
   Begin VB.CommandButton cmdItemFunction 
      Caption         =   "None"
      Height          =   255
      Index           =   2
      Left            =   8280
      TabIndex        =   46
      Top             =   2400
      Width           =   1245
   End
   Begin VB.TextBox txtItemName 
      Height          =   285
      Index           =   2
      Left            =   8280
      TabIndex        =   45
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CommandButton cmdItemFunction 
      Caption         =   "None"
      Height          =   255
      Index           =   1
      Left            =   8280
      TabIndex        =   44
      Top             =   1680
      Width           =   1245
   End
   Begin VB.TextBox txtItemName 
      Height          =   285
      Index           =   1
      Left            =   8280
      TabIndex        =   43
      Top             =   2040
      Width           =   3615
   End
   Begin VB.CommandButton cmdItemFunction 
      Caption         =   "None"
      Height          =   255
      Index           =   0
      Left            =   8280
      TabIndex        =   42
      Top             =   960
      Width           =   1245
   End
   Begin VB.TextBox txtItemName 
      Height          =   285
      Index           =   0
      Left            =   8280
      TabIndex        =   41
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CommandButton cmdRoomSelect2 
      Caption         =   "Location >"
      Height          =   375
      Left            =   3360
      TabIndex        =   40
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdRoomSelect1 
      Caption         =   "< Location"
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdAdvancedLink 
      Caption         =   "Advanced..."
      Height          =   255
      Index           =   7
      Left            =   6480
      TabIndex        =   37
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdvancedLink 
      Caption         =   "Advanced..."
      Height          =   255
      Index           =   6
      Left            =   6480
      TabIndex        =   36
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdvancedLink 
      Caption         =   "Advanced..."
      Height          =   255
      Index           =   5
      Left            =   6480
      TabIndex        =   35
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdvancedLink 
      Caption         =   "Advanced..."
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   34
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdvancedLink 
      Caption         =   "Advanced..."
      Height          =   255
      Index           =   3
      Left            =   6480
      TabIndex        =   33
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdvancedLink 
      Caption         =   "Advanced..."
      Height          =   255
      Index           =   2
      Left            =   6480
      TabIndex        =   32
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdvancedLink 
      Caption         =   "Advanced..."
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   31
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdvancedLink 
      Caption         =   "Advanced..."
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   30
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdPageDownLink 
      Caption         =   "Page Down"
      Height          =   375
      Left            =   7320
      TabIndex        =   18
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton cmdScrollDownLink 
      Caption         =   "Scroll Down"
      Height          =   375
      Left            =   7320
      TabIndex        =   29
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton cmdScrollLinkUp 
      Caption         =   "Scroll Up"
      Height          =   375
      Left            =   7320
      TabIndex        =   28
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddLink 
      Caption         =   "Add Link"
      Height          =   255
      Index           =   7
      Left            =   5280
      TabIndex        =   27
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddLink 
      Caption         =   "Add Link"
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   26
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddLink 
      Caption         =   "Add Link"
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   25
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddLink 
      Caption         =   "Add Link"
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   24
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddLink 
      Caption         =   "Add Link"
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   23
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddLink 
      Caption         =   "Add Link"
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   22
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddLink 
      Caption         =   "Add Link"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   21
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddLink 
      Caption         =   "Add Link"
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   20
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtRoomName 
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Text            =   "Name"
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdPageUpLink 
      Caption         =   "Page Up"
      Height          =   375
      Left            =   7320
      TabIndex        =   17
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtLinkIndex 
      Height          =   285
      Index           =   7
      Left            =   4440
      TabIndex        =   16
      Text            =   "0"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox txtLinkText 
      Height          =   285
      Index           =   7
      Left            =   4440
      TabIndex        =   15
      Top             =   6360
      Width           =   3615
   End
   Begin VB.TextBox txtLinkIndex 
      Height          =   285
      Index           =   6
      Left            =   4440
      TabIndex        =   14
      Text            =   "0"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox txtLinkText 
      Height          =   285
      Index           =   6
      Left            =   4440
      TabIndex        =   13
      Top             =   5640
      Width           =   3615
   End
   Begin VB.TextBox txtLinkIndex 
      Height          =   285
      Index           =   5
      Left            =   4440
      TabIndex        =   12
      Text            =   "0"
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtLinkText 
      Height          =   285
      Index           =   5
      Left            =   4440
      TabIndex        =   11
      Top             =   4920
      Width           =   3615
   End
   Begin VB.TextBox txtLinkIndex 
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   10
      Text            =   "0"
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox txtLinkText 
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   9
      Top             =   4200
      Width           =   3615
   End
   Begin VB.TextBox txtLinkIndex 
      Height          =   285
      Index           =   3
      Left            =   4440
      TabIndex        =   8
      Text            =   "0"
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtLinkText 
      Height          =   285
      Index           =   3
      Left            =   4440
      TabIndex        =   7
      Top             =   3480
      Width           =   3615
   End
   Begin VB.TextBox txtLinkIndex 
      Height          =   285
      Index           =   2
      Left            =   4440
      TabIndex        =   6
      Text            =   "0"
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtLinkText 
      Height          =   285
      Index           =   2
      Left            =   4440
      TabIndex        =   5
      Top             =   2760
      Width           =   3615
   End
   Begin VB.TextBox txtLinkIndex 
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   4
      Text            =   "0"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtLinkText 
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   3
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox txtLinkIndex 
      Height          =   285
      Index           =   0
      Left            =   4440
      TabIndex        =   2
      Text            =   "0"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtLinkText 
      Height          =   285
      Index           =   0
      Left            =   4440
      TabIndex        =   1
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox txtRoomText 
      Height          =   6975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmEditor.frx":0000
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label lblRoomIndex 
      Caption         =   "1"
      Height          =   375
      Left            =   3720
      TabIndex        =   38
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpenGame 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuLocation 
      Caption         =   "Location"
      Begin VB.Menu mnuNewRoom 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpenRoom 
         Caption         =   "Open"
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurrentRoom As Integer, Rooms As Integer, LinkView As Integer, Room() As Room
Dim MapName As String



Private Sub cmdItemFunction_Click(Index As Integer)
Room(CurrentRoom).ItemFunction(Index + LinkView) = Room(CurrentRoom).ItemFunction(Index + LinkView) + 1
If Room(CurrentRoom).ItemFunction(Index + LinkView) = 5 Then Room(CurrentRoom).ItemFunction(Index + LinkView) = 0
Select Case Room(CurrentRoom).ItemFunction(Index + LinkView)
Case Is = 0 'none
    cmdItemFunction(Index).Caption = "None"
Case Is = 1 'get item
    cmdItemFunction(Index).Caption = "Get"
Case Is = 2 'lose item
    cmdItemFunction(Index).Caption = "Lose"
Case Is = 3 'unlock link only if player holds this item
    cmdItemFunction(Index).Caption = "Unlock"
Case Is = 4 'unlock link and lose the item
    cmdItemFunction(Index).Caption = "Unlock + Lose"
End Select
End Sub

Private Sub cmdPageDownLink_Click()
LinkView = LinkView + 8
If LinkView > 24 Then LinkView = 24
Call RefreshData
End Sub

Private Sub cmdPageUpLink_Click()
LinkView = LinkView - 8
If LinkView < 0 Then LinkView = 0
Call RefreshData
End Sub

Private Sub cmdRoomSelect1_Click()
CurrentRoom = CurrentRoom - 1
If CurrentRoom = 0 Then CurrentRoom = Rooms
Call RefreshData
End Sub

Private Sub cmdRoomSelect2_Click()
CurrentRoom = CurrentRoom + 1
If CurrentRoom > Rooms Then CurrentRoom = 1
Call RefreshData
End Sub

Private Sub Form_Load()
ReDim Preserve Room(1 To 1)
Rooms = 1
CurrentRoom = 1
End Sub

Private Sub mnuNewRoom_Click()
Rooms = Rooms + 1
ReDim Preserve Room(1 To Rooms)
CurrentRoom = Rooms
Call RefreshData
End Sub

Private Sub mnuOpenGame_Click()
Dim i, j, temp
MapName = LCase(InputBox("Type the name of the scenario you want to open.", "Open Scenario"))
On Error GoTo 1
Open App.Path & "\Scenarios\" & MapName & ".dat" For Input As #1
Line Input #1, temp
    MapName = temp
Line Input #1, temp
    Rooms = temp
ReDim Room(1 To Rooms)
For i = 1 To Rooms
Line Input #1, temp
    Room(i).Name = temp
Line Input #1, temp
    Room(i).Text = temp
For j = 0 To 31
Line Input #1, temp
    Room(i).Link(j) = temp
Line Input #1, temp
    Room(i).LinkText(j) = temp
Line Input #1, temp
    Room(i).ItemName(j) = temp
Line Input #1, temp
    Room(i).ItemFunction(j) = temp
Next j
Next i
Close #1
CurrentRoom = 1
Call RefreshData
Exit Sub
1:
MsgBox "Could not find requested scenario.", , "Error"
End Sub

Private Sub mnuSave_Click()
Dim i, j
If MapName = "" Then
MsgBox "Could not save file.  Use Save As first.", , "Error"
Exit Sub
End If
Open App.Path & "\Scenarios\" & MapName & ".dat" For Output As #1
Print #1, MapName
Print #1, Rooms
For i = 1 To Rooms
Print #1, Room(i).Name
Print #1, Room(i).Text
For j = 0 To 31
Print #1, Room(i).Link(j)
Print #1, Room(i).LinkText(j)
Print #1, Room(i).ItemName(j)
Print #1, Room(i).ItemFunction(j)

Next j
Next i
Close #1
End Sub

Private Sub mnuSaveAs_Click()
Dim i, j
On Error GoTo 1
MapName = LCase(InputBox("What name do you want to save this scenario under?", "Save Scenario", "My Text Adventure"))
Open App.Path & "\Scenarios\" & MapName & ".dat" For Output As #1
Print #1, MapName
Print #1, Rooms
For i = 1 To Rooms
Print #1, Room(i).Name
Print #1, Room(i).Text
For j = 0 To 31
Print #1, Room(i).Link(j)
Print #1, Room(i).LinkText(j)
Print #1, Room(i).ItemName(j)
Print #1, Room(i).ItemFunction(j)
Next j
Next i
Close #1
Exit Sub
1:
MsgBox "Invalid File Name", , "Error"
End Sub

Private Sub txtItemName_Change(Index As Integer)
Room(CurrentRoom).ItemName(Index + LinkView) = txtItemName(Index).Text
End Sub

Private Sub txtLinkIndex_LostFocus(Index As Integer)
txtLinkIndex(Index).Text = Int(Val(txtLinkIndex(Index).Text))
Room(CurrentRoom).Link(Index + LinkView) = txtLinkIndex(Index).Text
End Sub

Private Sub txtLinkText_Change(Index As Integer)
Room(CurrentRoom).LinkText(Index + LinkView) = txtLinkText(Index).Text
End Sub

Private Sub txtRoomName_Change()
Room(CurrentRoom).Name = txtRoomName.Text
End Sub

Private Sub txtRoomText_Change()
Room(CurrentRoom).Text = txtRoomText.Text
End Sub

Private Sub RefreshData()
Dim i
For i = 0 To 7
txtLinkText(i).Text = Room(CurrentRoom).LinkText(i + LinkView)
txtLinkIndex(i).Text = Room(CurrentRoom).Link(i + LinkView)
txtItemName(i).Text = Room(CurrentRoom).ItemName(i + LinkView)
Select Case Room(CurrentRoom).ItemFunction(i + LinkView)
Case Is = 0 'none
    cmdItemFunction(i).Caption = "None"
Case Is = 1 'get item
    cmdItemFunction(i).Caption = "Get"
Case Is = 2 'lose item
    cmdItemFunction(i).Caption = "Lose"
Case Is = 3 'unlock link only if player holds this item
    cmdItemFunction(i).Caption = "Unlock"
Case Is = 4 'unlock link and lose the item
    cmdItemFunction(i).Caption = "Unlock + Lose"
End Select
Next i
txtRoomText.Text = Room(CurrentRoom).Text
txtRoomName.Text = Room(CurrentRoom).Name
lblRoomIndex.Caption = CurrentRoom
End Sub
