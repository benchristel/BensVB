Attribute VB_Name = "modDeclarations"
Type EnemyData
    Name As String
    Speed As Double
    Altitude As Double
    Damage As Integer
    Health As Integer
    GraphicsDC As Integer
    Points As Integer
    PowerDrain As Integer
End Type
Type Enemy
    Name As String
    Speed As Double
    XCoord As Double
    YCoord As Double
    Damage As Integer
    Points As Integer
    PowerDrain As Integer
    Health As Integer
    GraphicsDC As Integer
    Dead As Boolean
    DeadYVector As Double
    Deleted As Boolean
End Type
Type Shot
    XCoord As Double
    YCoord As Double
    XVector As Double
    YVector As Double
    Damage As Integer
    Deleted As Boolean
End Type
Type Player
    Health As Integer
    Ready As Boolean
    Damage As Integer
    Kills As Integer
    TargetX As Single
    TargetY As Single
    DrawX As Single
    DrawY As Single
    Points As Long
End Type
Type Slope
    Rise As Double
    Run As Double
End Type
Public Paused As Boolean, Terminated As Boolean, Player As Player
Public Enemy() As Enemy, EnemyData(1 To 8) As EnemyData, EnemyMin As Integer, EnemyCount As Integer
Public Shot() As Shot, ShotMin As Integer, ShotCount As Integer
Public SpawnThreshold As Double
Public BackgroundDC As Long, BackBuffDC As Long, ShotDC As Long, EnemyDC(1 To 8) As Long
Public ShotMaskDC As Long, EnemyMaskDC(1 To 8) As Long

Public Sub InitializeData()
With EnemyData(1)
.Altitude = 20
.Damage = 50
.GraphicsDC = 1
.Health = 5
.Name = "Orc"
.Points = 60
.PowerDrain = 0
.Speed = 0.5
End With
With EnemyData(2)
.Altitude = 150
.Damage = 150
.GraphicsDC = 2
.Health = 10
.Name = "Phoenix"
.Points = 200
.PowerDrain = 0
.Speed = 0.75
End With
With EnemyData(3)
.Altitude = 20
.Damage = 30
.GraphicsDC = 3
.Health = 4
.Name = "Dark Mage"
.Points = 100
.PowerDrain = 0
.Speed = 0.85
End With
With EnemyData(4)
.Altitude = 20
.Damage = -100
.GraphicsDC = 4
.Health = 5
.Name = "Dwarven Smith"
.Points = -100
.PowerDrain = 0
.Speed = 0.7
End With
With EnemyData(5)
.Altitude = 190
.Damage = 0
.GraphicsDC = 5
.Health = 5
.Name = "Wisp"
.Points = -100
.PowerDrain = -3
.Speed = 0.7
End With
With EnemyData(6)
.Altitude = 100
.Damage = 200
.GraphicsDC = 6
.Health = 16
.Name = "Genie"
.Points = 275
.PowerDrain = 2
.Speed = 0.8
End With
End Sub
