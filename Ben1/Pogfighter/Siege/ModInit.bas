Attribute VB_Name = "Module1"
Option Explicit
Type Enemy
GraphicsDC As Integer 'number used in index of graphics dc variables to determine which graphics
    'are used for a given object
x As Single 'x-coordinate of ship
y As Single 'y-coordinate of ship
AniFrames(1 To 3) As Integer  'the number of frames in the ship's animation, 1 = normal,
    '2 = firing, 3 = explode
FireRate As Integer 'how many frames will elapse between each time the ship fires. 0 = does
    'not fire
Damage As Integer 'the amount of damage the ship's attacks do
Health As Integer 'the amount of health the ship has
Speed As Integer 'the speed of the ship
YVar As Boolean 'whether the ship has elevational variation
CurFrame As Integer 'the current frame to blit to the screen.
Mode As Integer 'what the ship is doing(normal, firing, exploding, dead)
End Type
'''''
Type Player
GraphicsDC As Integer
x As Single
y As Single
Mode As Integer
MoveRate As Double
AniFrames(1 To 3) As Integer
CurFrame As Integer
End Type
'''''
Type Missile
x As Single
y As Single
Exploded As Boolean
YVar As Single
XVar As Single
YMinus As Single
Damage As Integer
GraphicsDC As Integer
End Type
