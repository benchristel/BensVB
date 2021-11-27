Attribute VB_Name = "ModInit"
Public Room As Integer, Rooms As Integer, RoomData() As Room
Type Room
    Name As String 'name of room
    Text As String 'text displayed
    Link(0 To 31) As Integer 'link to room with this index
    LinkText(0 To 31) As String 'when this text is entered
    ItemName(0 To 31) As String
    ItemFunction(0 To 31) As Integer
End Type



Public Function EnterText(Text As String) As Integer
Dim i, k
For i = 0 To 31
If LCase(Trim(RoomData(Room).LinkText(i))) = LCase(Trim(Text)) And Trim(RoomData(Room).Link(i)) <> 0 Then
Select Case RoomData(Room).ItemFunction(i)
Case Is = 0 'none
    EnterText = RoomData(Room).Link(i)
    Exit For
Case Is = 1 'get
    frmAdventure.lstInventory.AddItem (RoomData(Room).ItemName(i))
    frmAdventure.lstInventory.Refresh
    EnterText = RoomData(Room).Link(i)
    Exit For
Case Is = 2 'lose
    For k = 0 To frmAdventure.lstInventory.ListCount
        If frmAdventure.lstInventory.Text = RoomData(Room).ItemName(i) Then
            frmAdventure.lstInventory.RemoveItem (k)
            frmAdventure.lstInventory.Refresh
            Exit For
        End If
    EnterText = RoomData(Room).Link(i)
    Exit For
    Next k
Case Is = 3 'unlock
    If frmAdventure.lstInventory.Text = RoomData(Room).ItemName(i) Then
    EnterText = RoomData(Room).Link(i)
    Exit For
    End If
Case Is = 4 'unlock and lose
    If frmAdventure.lstInventory.Text = RoomData(Room).ItemName(i) Then
        For k = 0 To frmAdventure.lstInventory.ListCount
        If frmAdventure.lstInventory.Text = RoomData(Room).ItemName(i) Then
            frmAdventure.lstInventory.RemoveItem (k)
            frmAdventure.lstInventory.Refresh
            Exit For
        End If
        Next k
    EnterText = RoomData(Room).Link(i)
    Exit For
    End If
End Select
End If
Next i
If EnterText = 0 Then 'if no link match found then
EnterText = Room 'keep room the same
End If
End Function
