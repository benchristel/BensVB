Attribute VB_Name = "Module1"
Type Room
    Name As String 'name of room
    Text As String 'text displayed
    Link(0 To 31) As Integer 'link to room with this index
    LinkText(0 To 31) As String 'when this text is entered
    ItemName(0 To 31) As String
    ItemFunction(0 To 31) As Integer
End Type
