Attribute VB_Name = "Module1"
Option Explicit
Public ReqItem(1 To 10, 1 To 10) As Integer, ReqNumber(1 To 10, 1 To 10) As Integer
'2d array consists of stacks of required items to make another item; 1 item type from each
'stack is used in making the item.  Generally items with alternate possibilities
'are tools and not deleted from the inventory when making an item, so any items that have
'a 1st dimension index value where the corresponding delete array item is set to true are
'deleted.
Public DeleteItem(1 To 10) As Boolean, ItemName(1 To 30) As String
'ITEMS
'=====
'1: Flint
'2: Maple Wood
'3: Branch
'4: Flint Axe
'5: Copper Ore
'6: Bucket
'7: Bucket of Water
'8: Clay
'9: Tinder
'10: Clay Brick
'11: Furnace
'12: Copper Bar
'13: Copper Hammer
'14: Tin Ore
'15: Tinderbox
'16: Tin Bar
'17: Bronze Bar
'18: Bronze Axe
'19: Fishing Net
'20: Plant Fibers
'21: Anchovies
'22: Shrimp
'23: Cooked Shrimp
'TERRAIN
'=======
'0: Grass
'1: Stones
'2: Tree Trunk
'3: Mining Rocks
'4: Spring
'5: Big Stone
'6: Anvil
Public Sub MakeItem(Name As String, DC As Integer, PickUp As Boolean)
'<<<all variables to be set prior to the calling of this sub>>>
Dim i, j, Alternate(1 To 10), ItemInSlot(1 To 10) As Integer
For i = 1 To 10 'cycle thru required items
If ReqItem(i, 1) > 0 Then
    For j = 0 To 9 'cycle thru inventory slots
        If Player.Item(j) = ReqItem(i, 1) Then
        Alternate(i) = 1
        GoTo 1 'program has found needed item
        End If
        If Player.Item(j) = ReqItem(i, 2) And ReqItem(i, 2) > 0 Then
        Alternate(i) = 2
        GoTo 1 'program has found needed item
        End If
    Next j
    Exit Sub
1:
    If Player.ItemNum(j) < ReqNumber(i, Alternate(i)) Then
    Exit Sub 'insufficient materials
    Else
    ItemInSlot(i) = j
    End If
End If
Next i
'If program has gotten this far all needed items have been found
For i = 1 To 10 'cycle thru needed items
'delete items that should be deleted
If DeleteItem(i) = True Then Player.ItemNum(ItemInSlot(i)) = Player.ItemNum(ItemInSlot(i)) - ReqNumber(i, Alternate(i))
Next i
Call GenerateItem(Name, DC, Player.x, Player.y, PickUp) 'create item on the ground
Call RefreshInventory
End Sub

Public Sub MakeFlintAxe()
ClearReqItems
ReqItem(1, 1) = 1 'flint
ReqItem(2, 1) = 3 'branch
ReqNumber(1, 1) = 1
ReqNumber(2, 1) = 1
DeleteItem(1) = True
DeleteItem(2) = True
Call MakeItem("Flint Axe", 4, True)
End Sub
Public Sub MakeBucket()
ClearReqItems
ReqItem(1, 1) = 1 'flint
ReqItem(1, 2) = 4 'flint axe
ReqItem(2, 1) = 2 'maple log
ReqNumber(1, 1) = 1
ReqNumber(1, 2) = 1
ReqNumber(2, 1) = 1
DeleteItem(1) = False
DeleteItem(2) = True
Call MakeItem("Bucket", 6, True)
End Sub

Public Sub ClearReqItems()
Dim i, j
For i = 1 To 10
For j = 1 To 10
ReqItem(i, j) = 0
ReqNumber(i, j) = 0
Next j
DeleteItem(i) = False
Next i
End Sub

Public Sub FillBucket()
ClearReqItems
ReqItem(1, 1) = 6 'bucket
ReqNumber(1, 1) = 1
DeleteItem(1) = True
If Terrain(Player.x, Player.y) = 4 Then Call MakeItem("Bucket of Water", 7, True)
End Sub

Public Sub MakeBrick()
ClearReqItems
ReqItem(1, 1) = 8 'clay
ReqItem(2, 1) = 7 'bucket of water
ReqNumber(1, 1) = 1
ReqNumber(2, 1) = 1
DeleteItem(1) = False
DeleteItem(2) = False
Call MakeItem("Brick", 10, True)
DeleteItem(1) = True
DeleteItem(2) = True
Call MakeItem("Bucket", 6, True)
End Sub

Public Sub BuildFurnace()
ClearReqItems
ReqItem(1, 1) = 10 'clay brick
ReqItem(2, 1) = 2 'maple logs
ReqItem(3, 1) = 9 'tinder
ReqItem(4, 1) = 1 'flint
DeleteItem(1) = True
DeleteItem(2) = True
DeleteItem(3) = True
DeleteItem(4) = True
ReqNumber(1, 1) = 5
ReqNumber(2, 1) = 2
ReqNumber(3, 1) = 2
ReqNumber(4, 1) = 1
Call MakeItem("Furnace", 11, False)
End Sub

Public Sub SmeltCopper()
Dim i
ClearReqItems
ReqItem(1, 1) = 5
DeleteItem(1) = True
ReqNumber(1, 1) = 1
For i = ItemMin To Items
If Item(i).x = Player.x And Item(i).y = Player.y And Item(i).DC = 11 Then Call MakeItem("Copper Bar", 12, True)
Next i
'if player is standing on a forge
End Sub

Public Sub MakeCopperHammer()
Dim i
ClearReqItems
ReqItem(1, 1) = 12
ReqItem(2, 1) = 3
DeleteItem(1) = True
DeleteItem(2) = True
ReqNumber(1, 1) = 1
ReqNumber(2, 1) = 1
Call MakeItem("Copper Hammer", 13, True)
End Sub


Public Sub MakeStoneAnvil()
ClearReqItems
ReqItem(1, 1) = 1
ReqItem(1, 2) = 4
DeleteItem(1) = False
ReqNumber(1, 1) = 1
ReqNumber(1, 2) = 1
If Terrain(Player.x, Player.y) = 5 Then
Call MakeItem("Flint", 1, False)
Item(Items).Deleted = True
Terrain(Player.x, Player.y) = 6
End If
End Sub

Public Sub SmeltTin()
Dim i
ClearReqItems
ReqItem(1, 1) = 14
DeleteItem(1) = True
ReqNumber(1, 1) = 1
For i = ItemMin To Items
If Item(i).x = Player.x And Item(i).y = Player.y And Item(i).DC = 11 Then Call MakeItem("Tin Bar", 16, True)
Next i
End Sub

Public Sub MakeTinderbox()
ClearReqItems
ReqItem(1, 1) = 1
ReqItem(1, 2) = 13
ReqItem(1, 3) = 4
ReqItem(2, 1) = 16
ReqItem(3, 1) = 9
ReqItem(4, 1) = 1
DeleteItem(1) = False
DeleteItem(2) = True
DeleteItem(3) = True
DeleteItem(4) = True
ReqNumber(1, 1) = 1
ReqNumber(2, 1) = 1
ReqNumber(1, 2) = 1
ReqNumber(1, 3) = 1
ReqNumber(3, 1) = 1
ReqNumber(4, 1) = 1
If Terrain(Player.x, Player.y) = 6 Then Call MakeItem("Tinderbox", 15, True)
End Sub

Public Sub SmeltBronze()
Dim i
ClearReqItems
ReqItem(1, 1) = 14
ReqItem(2, 1) = 5
DeleteItem(1) = True
DeleteItem(2) = True
ReqNumber(1, 1) = 1
ReqNumber(2, 1) = 1
For i = ItemMin To Items
If Item(i).x = Player.x And Item(i).y = Player.y And Item(i).DC = 11 Then Call MakeItem("Bronze Bar", 17, True)
Next i
End Sub

Public Sub MakeBronzeAxe()
ClearReqItems
ReqItem(1, 1) = 17
ReqItem(2, 1) = 3
ReqItem(3, 1) = 13
DeleteItem(1) = True
DeleteItem(2) = True
DeleteItem(3) = False
ReqNumber(1, 1) = 1
ReqNumber(2, 1) = 1
ReqNumber(3, 1) = 1
If Terrain(Player.x, Player.y) = 6 Then Call MakeItem("Bronze Axe", 18, True)
End Sub

Public Sub MakeFishingNet()
ClearReqItems
ReqItem(1, 1) = 20
ReqItem(2, 1) = 3
DeleteItem(1) = True
DeleteItem(2) = True
ReqNumber(1, 1) = 3
ReqNumber(2, 1) = 1
Call MakeItem("Fishing Net", 19, True)
End Sub

Public Sub FishForFish()
Dim i
    For i = 0 To 9
    If Player.Item(i) = 19 And Terrain(Player.x, Player.y) = 4 Then
    Select Case Int(Rnd * 4 + 1)
    Case Is = 1
    Call GenerateItem("Raw Anchovies", 21, Player.x, Player.y, True)
    Fish = Fish - 1
    Case Is = 2
    Call GenerateItem("Raw Shrimps", 22, Player.x, Player.y, True)
    Fish = Fish - 1
    Case Else
    'Call GenerateItem("Fishing Bait", 23, Player.x, Player.y, True)
    End Select
    Exit Sub
    End If
    Next i
End Sub

Public Sub CookShrimps()
ClearReqItems
ReqItem(1, 1) = 22
ReqItem(2, 1) = 2
ReqItem(3, 1) = 15
DeleteItem(1) = True
DeleteItem(2) = True
DeleteItem(3) = False
ReqNumber(1, 1) = 1
ReqNumber(2, 1) = 1
ReqNumber(3, 1) = 1
Call MakeItem("Cooked Shrimps", 23, True)
If Int(Rnd * 2) = 0 Then Item(Items).Deleted = True
End Sub
