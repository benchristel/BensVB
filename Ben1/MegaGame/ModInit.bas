Attribute VB_Name = "ModInit"
Option Explicit
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Constants for the GenerateDC function
'**LoadImage Constants**
Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
Const LR_DEFAULTSIZE As Long = &H40
'variable declarations
Public BackBuffDC, BackgroundDC
Public NMove As Boolean, NEMove As Boolean, EMove As Boolean, SEMove As Boolean, SMove As Boolean, SWMove As Boolean, WMove As Boolean, NWMove As Boolean
Public Items As Integer, Item() As Item, ItemMin As Integer, Player As Player, ItemDC(1 To 30) As Long, IMaskDC(1 To 30) As Long, Terrain(1 To 400, 1 To 400) As Integer
Public TerrainDC(0 To 6) As Long, PlayerDC(1 To 4) As Long, PMaskDC(1 To 4) As Long, TreeDC(1 To 5) As Long, Tree() As Tree, TreeCount As Integer
Public TMaskDC(1 To 5) As Long, RespawnPoint() As RespawnPoint, RespawnCount As Integer
Public Fish As Integer 'number of fish left in all fishing spots
Type Item
    Name As String ' the name of the object type
    DC As Integer ' the number which refers to this object dc and item type in code
    x As Integer ' the x-position of the object
    y As Integer ' the y-position of the object
    Deleted As Boolean 'whether or not the object still exists
    PickUp As Boolean 'whether the item can be picked up
End Type
Type Player
    Name As String
    x As Integer
    y As Integer
    Position As Integer '1 to 4 = N, S, E, W
    Item(0 To 9) As Integer '0 = no item
    ItemNum(0 To 9) As Integer 'the number of items of 1 type in slots 1 to 10
    ChopTree As Integer '0 = no tree is being chopped, > 0 = index of tree being chopped
    XP(1 To 10) '1-Attack, 2-Defense, 3-Strength, 4-Precision, 5-Crafting, 6-Smithing,
    '7-Mining, 8-Woodcutting, 9-Cooking, 10-Carpentry
    End Type
Type Tree
    Name As String
    DC As Integer 'a number which denotes the tree type and graphics dc
    x As Integer
    y As Integer
    Fruit As Integer 'type of fruit on the tree  0 = no fruit
    Wood As Integer
    Deleted As Boolean
    MaxHp As Integer
    HP As Integer
    Tool As Integer '1 = Axe, 2 = pick
    RespawnTime As Integer 'time until tree respawns after being chopped
    TimeRemaining As Integer
End Type
Type RespawnPoint
    DC As Integer
    x As Integer
    y As Integer
    Frequency As Integer
    TimeRemaining As Integer
End Type
'IN: FileName: The file name of the graphics
'OUT: The Generated DC
Public Function GenerateDC(FileName As String) As Long
Dim DC As Long
Dim hBitmap As Long

'Create a Device Context, compatible with the screen
DC = CreateCompatibleDC(0)

If DC < 1 Then
    GenerateDC = 0
    Exit Function
End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC DC
    GenerateDC = 0
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject DC, hBitmap

'Return the device context
GenerateDC = DC

'Delte the bitmap handle object
DeleteObject hBitmap

End Function
'Deletes a generated DC
Public Function DeleteGeneratedDC(DC As Long) As Long

If DC > 0 Then
    DeleteGeneratedDC = DeleteDC(DC)
Else
    DeleteGeneratedDC = 0
End If

End Function


Public Sub GenerateItem(Name As String, DC As Integer, x As Integer, y As Integer, PickUp As Boolean)
Items = Items + 1
ReDim Preserve Item(1 To Items)
With Item(Items)
    .Name = Name
    .DC = DC
    .x = x
    .y = y
    .PickUp = PickUp
End With
End Sub

Public Sub InitializeTerrain()
Terrain(1, 1) = 1
Terrain(4, 2) = 1
Call GenerateTree("Maple", 1, 4, 3, False, 2, 100, 600, 1)
Terrain(2, 6) = 1
Call GenerateTree("Maple", 1, 7, 6, False, 2, 100, 600, 1)
Terrain(3, 4) = 1
Call GenerateTree("Maple", 1, 1, 2, False, 2, 100, 600, 1)
Terrain(9, 18) = 1
Terrain(7, 21) = 1
Terrain(13, 19) = 1
Terrain(32, 5) = 1
Terrain(34, 3) = 1
Terrain(35, 2) = 1
Terrain(35, 5) = 1
Terrain(36, 3) = 1
Terrain(36, 5) = 1
Terrain(36, 6) = 1
Terrain(37, 4) = 1
Terrain(38, 4) = 1
Terrain(38, 5) = 1
Terrain(2, 21) = 4
Call GenerateTree("Fishing Spot", 5, 2, 21, False, 22, 5, 10, 3)
Terrain(40, 15) = 4
Call GenerateTree("Fishing Spot", 5, 40, 15, False, 21, 5, 10, 3)
Terrain(7, 23) = 5
Call GenerateTree("Maple", 1, 5, 11, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 5, 20, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 6, 16, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 6, 19, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 6, 23, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 8, 21, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 8, 23, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 9, 15, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 9, 17, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 10, 11, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 10, 18, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 11, 20, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 11, 25, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 12, 18, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 12, 25, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 13, 13, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 13, 24, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 14, 19, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 14, 20, False, 2, 100, 600, 1)
Call GenerateTree("Maple", 1, 15, 22, False, 2, 100, 600, 1)
Call GenerateTree("Copper", 2, 42, 6, False, 5, 110, 800, 2)
Call GenerateTree("Copper", 2, 41, 7, False, 5, 110, 800, 2)
Call GenerateTree("Clay", 3, 32, 15, False, 8, 80, 500, 2)
Call GenerateTree("Clay", 3, 33, 12, False, 8, 80, 500, 2)
Call GenerateTree("Clay", 3, 42, 11, False, 8, 80, 500, 2)
Call GenerateTree("Clay", 3, 43, 12, False, 8, 80, 500, 2)
Call GenerateTree("Clay", 3, 44, 15, False, 8, 80, 500, 2)
Call GenerateTree("Tin", 4, 42, 3, False, 14, 110, 800, 2)
Call GenerateTree("Tin", 4, 41, 4, False, 14, 110, 800, 2)
Call GenerateRespawn(20, 3, 21, 800)
Call GenerateRespawn(20, 6, 15, 800)
Call GenerateRespawn(20, 6, 20, 800)
ItemName(1) = "Flint"
ItemName(2) = "Maple Wood"
ItemName(3) = "Branch"
ItemName(4) = "Flint Axe"
ItemName(5) = "Copper Ore"
ItemName(6) = "Bucket"
ItemName(7) = "Bucket of Water"
ItemName(8) = "Clay"
ItemName(9) = "Tinder"
ItemName(10) = "Brick"
ItemName(11) = "Furnace"
ItemName(12) = "Copper Bar"
ItemName(13) = "Copper Hammer"
ItemName(14) = "Tin Ore"
ItemName(15) = "Tinderbox"
ItemName(16) = "Tin Bar"
ItemName(17) = "Bronze Bar"
ItemName(18) = "Bronze Axe"
ItemName(19) = "Fishing Net"
ItemName(20) = "Plant Fibers"
ItemName(21) = "Raw Anchovies"
ItemName(22) = "Raw Shrimps"
ItemName(23) = "Cooked Shrimps"
Fish = 50
End Sub


Public Sub MovePlayer()
Dim i
If NMove = True Then
    If Player.y = 1 Then Exit Sub
    Select Case Terrain(Player.x, Player.y - 1)
    Case Is = 2, 3
    'Nothing happens--terrain is impassable
    Case Else
    Player.y = Player.y - 1
    End Select
End If
If SMove = True Then
    If Player.y = 400 Then Exit Sub
    Select Case Terrain(Player.x, Player.y + 1)
    Case Is = 2, 3
    'Nothing happens--terrain is impassable
    Case Else
    Player.y = Player.y + 1
    End Select
End If
If EMove = True Then
    If Player.x = 400 Then Exit Sub
    Select Case Terrain(Player.x + 1, Player.y)
    Case Is = 2, 3
    'Nothing happens--terrain is impassable
    Case Else
    Player.x = Player.x + 1
    End Select
End If
If WMove = True Then
    If Player.x = 1 Then Exit Sub
    Select Case Terrain(Player.x - 1, Player.y)
    Case Is = 2, 3
    'Nothing happens--terrain is impassable
    Case Else
    Player.x = Player.x - 1
    End Select
End If
'Chop at trees
If Player.ChopTree > 0 Then
    For i = 0 To 9
    If Player.Item(i) = 18 And Tree(Player.ChopTree).Tool = 1 Then
    Tree(Player.ChopTree).HP = Tree(Player.ChopTree).HP - 4
    GoTo CheckTreeState
    End If
    Next i
    For i = 0 To 9
    If Player.Item(i) = 4 Then
    Tree(Player.ChopTree).HP = Tree(Player.ChopTree).HP - 3
    GoTo CheckTreeState
    End If
    Next i
    For i = 0 To 9
    If Player.Item(i) = 1 Then
    Tree(Player.ChopTree).HP = Tree(Player.ChopTree).HP - 2
    GoTo CheckTreeState
    End If
    Next i
    For i = 0 To 9
    If Player.Item(i) = 19 And Tree(Player.ChopTree).Tool = 3 Then
    Tree(Player.ChopTree).HP = Tree(Player.ChopTree).HP - 1
    GoTo CheckTreeState
    End If
    Next i
CheckTreeState:
    If Tree(Player.ChopTree).HP <= 0 Then
    With Tree(Player.ChopTree)
        .Deleted = True
        .TimeRemaining = Tree(Player.ChopTree).RespawnTime
        .HP = Tree(Player.ChopTree).MaxHp
    End With
    Call GenerateItem(ItemName(Tree(Player.ChopTree).Wood), Tree(Player.ChopTree).Wood, Player.x, Player.y, True)
    If Int(Rnd * 2) = 0 And Tree(Player.ChopTree).Tool = 1 Then Call GenerateItem("Branch", 3, Player.x, Player.y, True)
    If Int(Rnd * 2) = 0 And Tree(Player.ChopTree).Tool = 1 Then Call GenerateItem("Tinder", 9, Player.x, Player.y, True)
    If Int(Rnd * 2) = 0 And Tree(Player.ChopTree).Tool = 2 Then Call GenerateItem("Flint", 1, Player.x, Player.y, True)
    Select Case Tree(Player.ChopTree).DC
    Case Is = 1
    Player.XP(8) = Player.XP(8) + 7
    Case Is = 2
    Player.XP(7) = Player.XP(7) + 12
    Case Is = 3
    Player.XP(7) = Player.XP(7) + 7
    Case Is = 4
    Player.XP(7) = Player.XP(7) + 12
    End Select
    Player.ChopTree = 0
    End If
End If
'respawn trees + advance timers
For i = 1 To TreeCount
If Tree(i).Deleted = True Then
Tree(i).TimeRemaining = Tree(i).TimeRemaining - 1
If Tree(i).TimeRemaining = 0 Then Tree(i).Deleted = False
End If
Next i
'Check item respawn points
For i = 1 To RespawnCount
If RespawnPoint(i).TimeRemaining = 0 Then
Call GenerateItem(ItemName(RespawnPoint(i).DC), RespawnPoint(i).DC, RespawnPoint(i).x, RespawnPoint(i).y, True)
RespawnPoint(i).TimeRemaining = RespawnPoint(i).Frequency
Else
RespawnPoint(i).TimeRemaining = RespawnPoint(i).TimeRemaining - 1
End If
Next i
'add fish
If Int(Rnd * 6) = 0 Then Fish = Fish + 1
End Sub

Public Sub GenerateTree(Name As String, DC As Integer, x As Integer, y As Integer, Fruit As Boolean, Wood As Integer, HP As Integer, RespawnTime As Integer, Tool As Integer)
TreeCount = TreeCount + 1
ReDim Preserve Tree(1 To TreeCount)
With Tree(TreeCount)
    .Name = Name
    .DC = DC
    .x = x
    .y = y
    .Fruit = Fruit
    .Wood = Wood
    .HP = HP
    .MaxHp = HP
    .RespawnTime = RespawnTime
    .TimeRemaining = 0
    .Tool = Tool
End With
If Tool = 1 Then Terrain(x, y) = 2 ' if this is actually a rock, don't draw tree trunk
If Tool = 2 Then Terrain(x, y) = 3
End Sub

Public Sub PickUpItem(x, y)
Dim i, j
For i = ItemMin To Items
If Item(i).x = x And Item(i).y = y And Item(i).Deleted = False And Item(i).PickUp = True Then
For j = 0 To 9
If Player.Item(j) = Item(i).DC And Player.ItemNum(j) > 0 Then
Player.ItemNum(j) = Player.ItemNum(j) + 1
Item(i).Deleted = True
GoTo 1
End If
Next j
For j = 0 To 9
If Player.Item(j) = 0 Or Player.ItemNum(j) = 0 Then
Player.Item(j) = Item(i).DC
Player.ItemNum(j) = 1
Item(i).Deleted = True
Exit For
End If
Next j
1:
End If
Next i
Call RefreshInventory
End Sub

Public Sub RefreshInventory()
Dim i
For i = 0 To 9
If Player.Item(i) > 0 And Player.ItemNum(i) > 0 Then
BitBlt Form1.picInventory(i).hdc, 0, 0, 60, 60, ItemDC(Player.Item(i)), 0, 0, vbSrcCopy
Else
Form1.picInventory(i).Cls
End If
Form1.lblInventory(i).Caption = Player.ItemNum(i)
Next i
End Sub

Public Sub DropItem(Index As Integer)
If Player.ItemNum(Index) > 0 Then
Player.ItemNum(Index) = Player.ItemNum(Index) - 1
Call GenerateItem(ItemName(Player.Item(Index)), Player.Item(Index), Player.x, Player.y, True)
End If
If Player.ItemNum(Index) = 0 Then Player.Item(Index) = 0
Call RefreshInventory
End Sub

Public Sub GenerateRespawn(DC, x, y, Frequency)
RespawnCount = RespawnCount + 1
ReDim Preserve RespawnPoint(1 To RespawnCount)
With RespawnPoint(RespawnCount)
.DC = DC
.x = x
.y = y
.Frequency = Frequency
End With
End Sub
