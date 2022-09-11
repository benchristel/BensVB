VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7890
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picInventory 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   960
      Index           =   9
      Left            =   9240
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   11
      Top             =   6540
      Width           =   960
   End
   Begin VB.PictureBox picInventory 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   960
      Index           =   8
      Left            =   8220
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   10
      Top             =   6540
      Width           =   960
   End
   Begin VB.PictureBox picInventory 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   960
      Index           =   7
      Left            =   7200
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   9
      Top             =   6540
      Width           =   960
   End
   Begin VB.PictureBox picInventory 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   960
      Index           =   6
      Left            =   6180
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   8
      Top             =   6540
      Width           =   960
   End
   Begin VB.PictureBox picInventory 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   960
      Index           =   5
      Left            =   5160
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   6540
      Width           =   960
   End
   Begin VB.PictureBox picInventory 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   960
      Index           =   4
      Left            =   4140
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   6
      Top             =   6540
      Width           =   960
   End
   Begin VB.PictureBox picInventory 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   960
      Index           =   3
      Left            =   3120
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   5
      Top             =   6540
      Width           =   960
   End
   Begin VB.PictureBox picInventory 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   960
      Index           =   2
      Left            =   2100
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   6540
      Width           =   960
   End
   Begin VB.PictureBox picInventory 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   960
      Index           =   1
      Left            =   1080
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   3
      Top             =   6540
      Width           =   960
   End
   Begin VB.PictureBox picInventory 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   960
      Index           =   0
      Left            =   60
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   2
      Top             =   6540
      Width           =   960
   End
   Begin VB.TextBox txtKeys 
      Height          =   435
      Left            =   8880
      TabIndex        =   1
      Top             =   -1000
      Width           =   255
   End
   Begin VB.PictureBox picField 
      Enabled         =   0   'False
      Height          =   6375
      Left            =   60
      ScaleHeight     =   421
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   421
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin VB.Timer tmrCycle 
      Interval        =   100
      Left            =   9180
      Top             =   60
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   9
      Left            =   9240
      TabIndex        =   21
      Top             =   7500
      Width           =   975
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   8
      Left            =   8220
      TabIndex        =   20
      Top             =   7500
      Width           =   975
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   7
      Left            =   7200
      TabIndex        =   19
      Top             =   7500
      Width           =   975
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   6
      Left            =   6180
      TabIndex        =   18
      Top             =   7500
      Width           =   975
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   5
      Left            =   5160
      TabIndex        =   17
      Top             =   7500
      Width           =   975
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   4
      Left            =   4140
      TabIndex        =   16
      Top             =   7500
      Width           =   975
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   3
      Left            =   3120
      TabIndex        =   15
      Top             =   7500
      Width           =   975
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   2100
      TabIndex        =   14
      Top             =   7500
      Width           =   975
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   13
      Top             =   7500
      Width           =   975
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   12
      Top             =   7500
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i
Randomize
For i = 1 To 23
ItemDC(i) = GenerateDC(App.Path & "\graphics\item" & i & ".bmp")
IMaskDC(i) = GenerateDC(App.Path & "\graphics\IMask" & i & ".bmp")
Next i
BackBuffDC = GenerateDC(App.Path & "\graphics\background.bmp")
BackgroundDC = GenerateDC(App.Path & "\graphics\background.bmp")
ItemMin = 1
Call InitializeTerrain
For i = 0 To 6
TerrainDC(i) = GenerateDC(App.Path & "\graphics\terrain" & i & ".bmp")
Next i
For i = 1 To 4
PlayerDC(i) = GenerateDC(App.Path & "\graphics\Player" & i & ".bmp")
PMaskDC(i) = GenerateDC(App.Path & "\graphics\PMask" & i & ".bmp")
Next i
For i = 1 To 4
TreeDC(i) = GenerateDC(App.Path & "\graphics\Tree" & i & ".bmp")
TMaskDC(i) = GenerateDC(App.Path & "\graphics\TMask" & i & ".bmp")
Next i
Player.x = 1
Player.y = 1
Player.Position = 1
Call GenerateItem("Flint", 1, 9, 18, True)
Call GenerateItem("Flint", 1, 41, 5, True)
End Sub
Private Sub Form_Terminate()
Dim i
Unload Me
For i = 1 To 20
DeleteGeneratedDC (ItemDC(i))
DeleteGeneratedDC (IMaskDC(i))
Next i
For i = 0 To 6
DeleteGeneratedDC (TerrainDC(i))
Next i
For i = 1 To 4
DeleteGeneratedDC (PlayerDC(i))
DeleteGeneratedDC (PMaskDC(i))
Next i
For i = 1 To 4
DeleteGeneratedDC (TreeDC(i))
DeleteGeneratedDC (TMaskDC(i))
Next i
End Sub

Private Sub lblInventory_Click(Index As Integer)
Call DropItem(Index)
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuOpen_Click()
Dim i, j, ReadItems, temp
On Error GoTo 1 'trap error where file has not been created
Open App.Path & "\Saves\SavedGame.dat" For Input As #1
    Line Input #1, temp
    Player.x = temp
    Line Input #1, temp
    Player.y = temp
For i = 0 To 9
    Line Input #1, temp
    Player.Item(i) = temp
    Line Input #1, temp
    Player.ItemNum(i) = temp
Next i
    Line Input #1, ReadItems
    ReDim Item(1 To ReadItems)
    ItemMin = 1
    Items = ReadItems
For i = 1 To ReadItems
    Line Input #1, temp
    Item(i).DC = temp
    Line Input #1, temp
    Item(i).Deleted = temp
    Line Input #1, temp
    Item(i).Name = temp
    Line Input #1, temp
    Item(i).PickUp = temp
    Line Input #1, temp
    Item(i).x = temp
    Line Input #1, temp
    Item(i).y = temp
Next i
    Line Input #1, temp
    TreeCount = temp
    ReDim Tree(1 To TreeCount)
For i = 1 To TreeCount
    Line Input #1, temp
    Tree(i).DC = temp
    Line Input #1, temp
    Tree(i).Deleted = temp
    Line Input #1, temp
    Tree(i).Fruit = temp
    Line Input #1, temp
    Tree(i).HP = temp
    Line Input #1, temp
    Tree(i).MaxHp = temp
    Line Input #1, temp
    Tree(i).Name = temp
    Line Input #1, temp
    Tree(i).RespawnTime = temp
    Line Input #1, temp
    Tree(i).TimeRemaining = temp
    Line Input #1, temp
    Tree(i).Tool = temp
    Line Input #1, temp
    Tree(i).Wood = temp
    Line Input #1, temp
    Tree(i).x = temp
    Line Input #1, temp
    Tree(i).y = temp
Next i
For i = 1 To 50
For j = 1 To 40
    Line Input #1, temp
    Terrain(i, j) = temp
Next j
Next i
    Line Input #1, temp
    RespawnCount = temp
    ReDim RespawnPoint(1 To RespawnCount)
For i = 1 To RespawnCount
    Line Input #1, temp
    RespawnPoint(i).DC = temp
    Line Input #1, temp
    RespawnPoint(i).Frequency = temp
    Line Input #1, temp
    RespawnPoint(i).TimeRemaining = temp
    Line Input #1, temp
    RespawnPoint(i).x = temp
    Line Input #1, temp
    RespawnPoint(i).y = temp
Next i
    Line Input #1, temp
    Fish = temp
Close #1
Call RefreshInventory
Exit Sub
1:
MsgBox "Could not find a saved game to open", , "Error"
Close #1
End Sub

Private Sub mnuSave_Click()
Dim i, j
Open App.Path & "\Saves\SavedGame.dat" For Output As #1
    Print #1, Player.x
    Print #1, Player.y
For i = 0 To 9
    Print #1, Player.Item(i)
    Print #1, Player.ItemNum(i)
Next i
    Print #1, Items - ItemMin + 1
For i = ItemMin To Items
    Print #1, Item(i).DC
    Print #1, Item(i).Deleted
    Print #1, Item(i).Name
    Print #1, Item(i).PickUp
    Print #1, Item(i).x
    Print #1, Item(i).y
Next i
    Print #1, TreeCount
For i = 1 To TreeCount
    Print #1, Tree(i).DC
    Print #1, Tree(i).Deleted
    Print #1, Tree(i).Fruit
    Print #1, Tree(i).HP
    Print #1, Tree(i).MaxHp
    Print #1, Tree(i).Name
    Print #1, Tree(i).RespawnTime
    Print #1, Tree(i).TimeRemaining
    Print #1, Tree(i).Tool
    Print #1, Tree(i).Wood
    Print #1, Tree(i).x
    Print #1, Tree(i).y
Next i
For i = 1 To 50
For j = 1 To 40
    Print #1, Terrain(i, j)
Next j
Next i
    Print #1, RespawnCount
For i = 1 To RespawnCount
    Print #1, RespawnPoint(i).DC
    Print #1, RespawnPoint(i).Frequency
    Print #1, RespawnPoint(i).TimeRemaining
    Print #1, RespawnPoint(i).x
    Print #1, RespawnPoint(i).y
Next i
    Print #1, Fish
Close #1
End Sub



Private Sub tmrCycle_Timer()
Dim i, x, DeleteItem As Boolean, BlitX, BlitY
Static Initialized As Boolean, FrameNo As Integer

If Initialized = False Then
txtKeys.SetFocus
Initialized = True
End If
If FrameNo = 3 Then
FrameNo = 1
Call MovePlayer
Else
FrameNo = FrameNo + 1
End If
BitBlt BackBuffDC, 0, 0, 420, 420, BackgroundDC, 0, 0, vbSrcCopy
If ItemMin <= Items Then
For i = ItemMin To Items
If Item(i).Deleted = True And DeleteItem = False Then
ItemMin = i + 1
Else
DeleteItem = True
End If
If Item(i).Deleted = False And Item(i).x >= Player.x - 3 And Item(i).x <= Player.x + 3 And Item(i).y >= Player.y - 3 And Item(i).y <= Player.y + 3 Then
BlitX = Item(i).x - Player.x + 3
BlitY = Item(i).y - Player.y + 3
End If
Next i
End If
For i = Player.x - 3 To Player.x + 3
For x = Player.y - 3 To Player.y + 3
'bitblit terrain to the backbuffer
If i < 1 Or i > 400 Or x < 1 Or x > 400 Then GoTo 1 'trap error where player is standing near edge of map
BitBlt BackBuffDC, (i - Player.x + 3) * 60, (x - Player.y + 3) * 60, 60, 60, TerrainDC(Terrain(i, x)), 0, 0, vbSrcCopy
1:
Next x
Next i
'blit items to backbuffer
For i = ItemMin To Items
If Item(i).x >= Player.x - 3 And Item(i).x <= Player.x + 3 And Item(i).y >= Player.y - 3 And Item(i).y <= Player.y + 3 And Item(i).Deleted = False Then
BlitX = Item(i).x - Player.x + 3
BlitY = Item(i).y - Player.y + 3
BitBlt BackBuffDC, BlitX * 60, BlitY * 60, 60, 60, IMaskDC(Item(i).DC), 0, 0, vbSrcAnd
BitBlt BackBuffDC, BlitX * 60, BlitY * 60, 60, 60, ItemDC(Item(i).DC), 0, 0, vbSrcPaint
End If
Next i
'blit trees to backbuffer
For i = 1 To TreeCount
If Tree(i).x >= Player.x - 3 And Tree(i).x <= Player.x + 3 And Tree(i).y >= Player.y - 3 And Tree(i).y <= Player.y + 3 And Tree(i).Deleted = False Then
BlitX = Tree(i).x - Player.x + 3
BlitY = Tree(i).y - Player.y + 3
BitBlt BackBuffDC, BlitX * 60, BlitY * 60, 60, 60, TMaskDC(Tree(i).DC), 0, 0, vbSrcAnd
BitBlt BackBuffDC, BlitX * 60, BlitY * 60, 60, 60, TreeDC(Tree(i).DC), 0, 0, vbSrcPaint
End If
Next i
'blit player graphics to backbuffer
BitBlt BackBuffDC, 180, 180, 60, 60, PMaskDC(Player.Position), 0, 0, vbSrcAnd
BitBlt BackBuffDC, 180, 180, 60, 60, PlayerDC(Player.Position), 0, 0, vbSrcPaint
'blit the backbuffer to the screen
BitBlt picField.hdc, 0, 0, 420, 420, BackBuffDC, 0, 0, vbSrcCopy
End Sub

Private Sub txtKeys_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i
If KeyCode <> 13 Then Player.ChopTree = 0
Select Case KeyCode
Case Is = 104 'keypad 8
    NMove = True
    EMove = False
    WMove = False
    Player.Position = 1 'north
Case Is = 100 'keypad 4
    WMove = True
    NMove = False
    SMove = False
    Player.Position = 4 'west
Case Is = 102 'keypad 6
    EMove = True
    NMove = False
    SMove = False
    Player.Position = 2 'east
Case Is = 98 'keypad 2
    SMove = True
    EMove = False
    WMove = False
    Player.Position = 3 'south
Case Is = 101 'keypad 5
Call PickUpItem(Player.x, Player.y)
Case Is = 13 'enter
    Select Case Player.Position
    Case Is = 1
    For i = 1 To TreeCount
    If Tree(i).y = Player.y - 1 And Tree(i).x = Player.x And Tree(i).Deleted = False Then Player.ChopTree = i
    Next i
    Case Is = 2
    For i = 1 To TreeCount
    If Tree(i).y = Player.y And Tree(i).x = Player.x + 1 And Tree(i).Deleted = False Then Player.ChopTree = i
    Next i
    Case Is = 3
    For i = 1 To TreeCount
    If Tree(i).y = Player.y + 1 And Tree(i).x = Player.x And Tree(i).Deleted = False Then Player.ChopTree = i
    Next i
    Case Is = 4
    For i = 1 To TreeCount
    If Tree(i).y = Player.y And Tree(i).x = Player.x - 1 And Tree(i).Deleted = False Then Player.ChopTree = i
    Next i
    End Select
Case Is = 96 'keypad 0
    frmMakeItem.Show 1
End Select
'MsgBox KeyCode
End Sub

Private Sub txtKeys_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = 104
NMove = False
Case Is = 100
WMove = False
Case Is = 102
EMove = False
Case Is = 98
SMove = False
End Select

End Sub

Private Sub txtKeys_LostFocus()
'txtKeys.SetFocus
End Sub

