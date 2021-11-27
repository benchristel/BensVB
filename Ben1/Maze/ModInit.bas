Attribute VB_Name = "Module1"
Option Explicit
Type Wall
    x As Integer
    y As Integer
    Length As Integer
    Position As Integer '1 = vertical, 2 = horizontal
End Type
Type Player
    Name As String * 20
    x As Single
    y As Single
    Health As Integer
    Position As Integer
    Ammo As Integer
    Bombs As Integer
    Movement As Integer '0 = standing still, 1 = forwards, -1 = move back
    Sidestep As Integer '0 = no lateral movement, 1 = right, -1 = move left
    Turning As Integer '0 = none, 1= right, -1= left
    Running As Integer '1 = normal speed, 2 = running
End Type
Type Bullet
    x As Single
    y As Single
    Position As Integer
    State As Integer
End Type
Type AmmoBox
    x As Integer
    y As Integer
    State As Integer '1  to 200 - each tick adds one - when at 200, box can be used again
End Type
Type Bomb
    x As Integer
    y As Integer
    State As Integer '0 = detonated, 1 = armed
End Type
