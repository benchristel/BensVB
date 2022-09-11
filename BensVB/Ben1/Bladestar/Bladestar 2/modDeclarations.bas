Attribute VB_Name = "modDeclarations"
Type Player
    XCoord As Double ' the x-coordinate of the player on the map
    YCoord As Double ' the y-coordinate of the player on the map
    TargetX As Double 'the coordinates of the cursor point
    TargetY As Double
    RadarBlipX As Double
    RadarBlipY As Double
    TurnLeft As Boolean
    TurnRight As Boolean
    MoveForward As Boolean
    MoveBack As Boolean
    StrafeLeft As Boolean
    StrafeRight As Boolean
    Theta As Double 'the angle of rotation, expressed in a percentage of 360 degrees.
    HP As Integer 'the player's hitpoints value
    Shield As Integer 'the player's shield hitpoints that recharge after every kill
    Weapon As Integer 'index of currently carried weapon
    Ammo As Integer
    ClipAmmo As Integer
    CoolDownTime As Integer
    Deaths As Integer
    Score As Integer
    Firing As Boolean ' whether or not the player is shooting.
    Reloading As Boolean
    PickUpItem As Boolean
    ReloadTime As Integer 'in frames.
    Dead As Boolean
    RespawnTime As Integer 'in seconds
    Sniping As Boolean
    SnipeToggle As Boolean
    RecordIndex As Integer 'tournament use to keep track of wins/losses
End Type
Type Weapon
    Name As String
    XCoord As Double
    YCoord As Double
    WeaponType As Integer 'index of weapon data
    ClipSize As Integer 'number of shots before reloading
    Ammo As Integer 'number of shots left
    ClipAmmo As Integer
    Deleted As Boolean 'whether or not this object has been deleted
    'SemiAuto As Boolean 'whether or not a separate keypress is required for each shot
    DespawnEnabled As Boolean   'whether this weapon will eventually despawn.  set to true for weapons that get dropped.
    TimeToDespawn As Integer 'in seconds.  Running timer of how long the weapon will wait before despawning.
End Type
Type WeaponData
    Name As String
    GraphicsIndex As Integer 'Index of graphics used ON SCREEN
    IconIndex As Integer 'Index of graphics used as the ICON
    ShotGraphicsIndex As Integer
    Damage As Integer
    ClipSize As Integer
    SpeedBonus As Double 'bonus to player's speed
    WallPiercing As Boolean 'whether the shot pierces walls
    StartAmmo As Integer 'COUNTING starting clip
    ShotRadius As Integer 'pixels
    ShotExplosive As Boolean 'whether shot explodes
    CoolDown As Integer 'in frames
    ReloadTime As Integer 'time in frames
    Melee As Boolean 'whether the weapon can be used in close combat
    MeleeDamage As Integer ' damage done in melee
    Bounce As Boolean 'whether or not shots bounce
    DamageMultiplier As Double 'totaldamage = damage + damagemultiplier * distance
    ShotSpeed As Double 'pixels per frame -- should not exceed 32.00
    ShotSpread As Integer 'max variation from mean in pixels at 100 px range
    ShotVisible As Boolean 'whether shot is visible
    ShotLifespan As Integer
    SemiAuto As Boolean
End Type
Type Shot
    Bounce As Boolean
    GraphicsIndex As Integer
    Alignment As Integer 'Player 1 or 2
    Damage As Integer
    XCoord As Double
    YCoord As Double
    XVector As Double
    YVector As Double
    Explosive As Boolean
    Radius As Integer 'pixels - radius of damage done
    Lifespan As Integer 'frames
    Distance As Double 'in pixels
    DamageMultiplier As Double
    Deleted As Boolean
    Speed As Double
    Visible As Boolean
    WallPiercing As Boolean
End Type
Type Object
    GraphicsIndex As Integer
    XCoord As Double
    YCoord As Double
    Age As Integer 'frames
    Deleted As Boolean
End Type
Type Wall
    XCoord As Double
    YCoord As Double
    Height As Double
    Width As Double
    Type As Integer '1 = blocks players + shots, '2 = chain link blocks only players, '3 = door blocks only shots
End Type
Type Slope
    Run As Double
    Rise As Double
End Type
Type PlayerSpawnPoint
    XCoord As Double
    YCoord As Double
End Type
Type WeaponSpawnPoint
    XCoord As Double
    YCoord As Double
    WeaponType As Integer
    Enabled As Integer 'whether this point is counting down currently
    LastSpawnedIndex As Integer ' the index of the last weapon spawned by this point.  When the weapon with this index is deleted, the point will start counting down to the next spawn.
    Frequency As Integer 'in seconds.  Coordinated with weapon despawn time to ensure that duplicate weapons are not spawned on one point.
    TimeLeft As Integer 'current time remaining until next spawn
End Type
Type Explosion
    XCoord As Double
    YCoord As Double
    Lifespan As Integer ' in frames
    XVector As Double
    YVector As Double
    GraphicsIndex As Integer
    Deleted As Boolean
End Type
Type PlayerRecord 'tournament use
    Name As String
    Wins As Integer
    Losses As Integer
End Type
Public ObjectGraphicsDC(1 To 1) As Long, PlayerGraphicsDC(1 To 2) As Long, WeaponGraphicsDC(1 To 9) As Long, ShotGraphicsDC(1 To 4) As Long
Public ExplosionGraphicsDC(1 To 1) As Long, ExplosionMaskDC(1 To 1) As Long, ExplosionCount As Integer, ExplosionMin As Integer, Explosion() As Explosion
Public PlayerMaskDC(1 To 2) As Long, BackBuffDC(1 To 2) As Long, BackgroundDC As Long, WeaponMaskDC(1 To 9) As Long, SightDC As Long, SightMaskDC As Long, ShotMaskDC(1 To 4) As Long
Public WallDC(1 To 2) As Long '1 = normal 2 = blocks players not shots
Public RadarBlipDC(1 To 2) As Long, RadarBlipMaskDC As Long
Public ShieldBarDC As Long, HealthBarDC As Long, GameOverDC As Long
Public Wall() As Wall, WallCount As Integer
Public Shot() As Shot, ShotCount As Integer, ShotMin As Integer
Public Weapon() As Weapon, WeaponCount As Integer, WeaponData() As WeaponData, WeaponMin As Integer
Public PlayerSpawn() As PlayerSpawnPoint, WeaponSpawn() As WeaponSpawnPoint
Public PlayerSpawnCount As Integer, WeaponSpawnCount As Integer
Public ScoreToWin(1 To 2) As Integer, GameOver As Boolean, UnloadCountdown As Integer, Terminated As Boolean
Public StartWeapon As Integer
'dimension object stat database
Public Player(1 To 2) As Player
Public VisualContact As Boolean
Public RuleVariant As Integer, VariantWeapon(1 To 10) As Integer, MapSelected As Integer
Public VariantDescription(0 To 2) As String, MapDescription(0 To 1) As String
Public PlayerRespawnTime As Integer
Public RandomWeapons As Boolean, WeaponDataCount As Integer
'=======================================
'Global Player Variables for Tournaments
'=======================================
Public PlayerRecord(0 To 9) As PlayerRecord
Public ScoreMethod As Integer 'wins/losses = 1, kills/deaths = 2


Public Sub InitializeWeaponsNormal()
ReDim WeaponData(1 To 10)
WeaponDataCount = 10
StartWeapon = 1
With WeaponData(1)
    .Bounce = False
    .CoolDown = 10
    .ClipSize = 80
    .Damage = 18
    .DamageMultiplier = 0
    .ShotExplosive = False
    .GraphicsIndex = 1
    .IconIndex = 1
    .Melee = True
    .MeleeDamage = 80
    .Name = "Laser Pistol"
    .ReloadTime = 100
    .SemiAuto = False
    .ShotGraphicsIndex = 1
    .ShotLifespan = 80
    .ShotRadius = 5
    .ShotSpeed = 10
    .ShotSpread = 0
    .ShotVisible = True
    .StartAmmo = 160
    .SpeedBonus = 0
    .WallPiercing = False
End With
With WeaponData(2)
    .Bounce = True
    .CoolDown = 80
    .ClipSize = 4
    .Damage = 160
    .DamageMultiplier = 0
    .ShotExplosive = True
    .GraphicsIndex = 2
    .IconIndex = 2
    .Melee = True
    .MeleeDamage = 30
    .Name = "Grenades"
    .ReloadTime = 0
    .SemiAuto = True
    .ShotGraphicsIndex = 2
    .ShotLifespan = 120
    .ShotRadius = 60
    .ShotSpeed = 6
    .ShotSpread = 0
    .ShotVisible = True
    .StartAmmo = 4
    .SpeedBonus = 0
    .WallPiercing = False
End With
With WeaponData(3)
    .Bounce = False
    .CoolDown = 5
    .ClipSize = 80
    .Damage = 10
    .DamageMultiplier = 0
    .ShotExplosive = False
    .GraphicsIndex = 4
    .IconIndex = 4
    .Melee = True
    .MeleeDamage = 50
    .Name = "SMG"
    .ReloadTime = 120
    .SemiAuto = False
    .ShotGraphicsIndex = 4
    .ShotLifespan = 110
    .ShotRadius = 14
    .ShotSpeed = 10
    .ShotSpread = 6
    .ShotVisible = True
    .StartAmmo = 240
    .SpeedBonus = 0
    .WallPiercing = False
End With
With WeaponData(4)
    .Bounce = False
    .CoolDown = 40
    .ClipSize = 6
    .Damage = 220
    .DamageMultiplier = -0.55
    .ShotExplosive = False
    .GraphicsIndex = 3
    .IconIndex = 3
    .Melee = False
    .MeleeDamage = 0
    .Name = "Shotgun"
    .ReloadTime = 120
    .SemiAuto = True
    .ShotGraphicsIndex = 1
    .ShotLifespan = 35
    .ShotRadius = 40
    .ShotSpeed = 10
    .ShotSpread = 0
    .ShotVisible = False
    .StartAmmo = 12
    .SpeedBonus = 0
    .WallPiercing = False
End With
With WeaponData(5)
    .Bounce = False
    .CoolDown = 25
    .ClipSize = 6
    .Damage = 1
    .DamageMultiplier = 0.1
    .ShotExplosive = False
    .GraphicsIndex = 5
    .IconIndex = 5
    .Melee = True
    .MeleeDamage = 130
    .Name = "Sniper Rifle"
    .ReloadTime = 60
    .SemiAuto = True
    .ShotGraphicsIndex = 4
    .ShotLifespan = 120
    .ShotRadius = 25
    .ShotSpeed = 30
    .ShotSpread = 0
    .ShotVisible = True
    .StartAmmo = 24
    .SpeedBonus = 0
    .WallPiercing = False
End With
With WeaponData(6)
    .Bounce = False
    .CoolDown = 25
    .ClipSize = 3
    .Damage = 150
    .DamageMultiplier = 0
    .ShotExplosive = True
    .GraphicsIndex = 6
    .IconIndex = 6
    .Melee = True
    .MeleeDamage = 30
    .Name = "Mines"
    .ReloadTime = 0
    .SemiAuto = True
    .ShotGraphicsIndex = 1
    .ShotLifespan = 2500
    .ShotRadius = 60
    .ShotSpeed = 0
    .ShotSpread = 0
    .ShotVisible = False
    .StartAmmo = 6
    .SpeedBonus = 0
    .WallPiercing = False
End With
With WeaponData(7)
    .Bounce = False
    .CoolDown = 100
    .ClipSize = 1000
    .Damage = 300
    .DamageMultiplier = 0
    .ShotExplosive = False
    .GraphicsIndex = 7
    .IconIndex = 7
    .Melee = False
    .MeleeDamage = 0
    .Name = "Saber"
    .ReloadTime = 0
    .SemiAuto = True
    .ShotGraphicsIndex = 1
    .ShotLifespan = 8
    .ShotRadius = 10
    .ShotSpeed = 15
    .ShotSpread = 0
    .ShotVisible = False
    .StartAmmo = 30
    .SpeedBonus = 1
    .WallPiercing = False
End With
With WeaponData(8)
    .Bounce = True
    .CoolDown = 4
    .ClipSize = 100
    .Damage = 8
    .DamageMultiplier = 0
    .ShotExplosive = False
    .GraphicsIndex = 8
    .IconIndex = 1
    .Melee = True
    .MeleeDamage = 70
    .Name = "Chain Gun"
    .ReloadTime = 150
    .SemiAuto = False
    .ShotGraphicsIndex = 3
    .ShotLifespan = 150
    .ShotRadius = 15
    .ShotSpeed = 7
    .ShotSpread = 12
    .ShotVisible = True
    .StartAmmo = 200
    .SpeedBonus = 0
    .WallPiercing = False
End With
With WeaponData(9)
    .Bounce = False
    .CoolDown = 0
    .ClipSize = 12
    .Damage = 40
    .DamageMultiplier = 0
    .ShotExplosive = False
    .GraphicsIndex = 9
    .IconIndex = 1
    .Melee = True
    .MeleeDamage = 80
    .Name = "Pistol"
    .ReloadTime = 50
    .SemiAuto = True
    .ShotGraphicsIndex = 4
    .ShotLifespan = 50
    .ShotRadius = 12
    .ShotSpeed = 14
    .ShotSpread = 1
    .ShotVisible = True
    .StartAmmo = 36
    .SpeedBonus = 0
    .WallPiercing = False
End With
With WeaponData(10)
    .Bounce = False
    .CoolDown = 0
    .ClipSize = 1
    .Damage = 250
    .DamageMultiplier = 0.1
    .ShotExplosive = True
    .GraphicsIndex = 8
    .IconIndex = 1
    .Melee = True
    .MeleeDamage = 80
    .Name = "Rocket Launcher"
    .ReloadTime = 300
    .SemiAuto = True
    .ShotGraphicsIndex = 1
    .ShotLifespan = 167
    .ShotRadius = 90
    .ShotSpeed = 12
    .ShotSpread = 0
    .ShotVisible = True
    .StartAmmo = 3
    .SpeedBonus = -0.3
    .WallPiercing = False
End With

VariantWeapon(1) = 1
VariantWeapon(2) = 2
VariantWeapon(3) = 3
VariantWeapon(4) = 4
VariantWeapon(5) = 5
VariantWeapon(6) = 6
VariantWeapon(7) = 7
VariantWeapon(8) = 8
VariantWeapon(9) = 9
VariantWeapon(10) = 10

End Sub
Public Sub InitializeWeaponsGrenades()
WeaponDataCount = 1
ReDim WeaponData(1 To 1)
StartWeapon = 1
With WeaponData(1)
    .Bounce = True
    .CoolDown = 80
    .ClipSize = 4
    .Damage = 170
    .DamageMultiplier = 0
    .ShotExplosive = True
    .GraphicsIndex = 2
    .IconIndex = 2
    .Melee = True
    .MeleeDamage = 155
    .Name = "Grenades"
    .ReloadTime = 0
    .SemiAuto = True
    .ShotGraphicsIndex = 2
    .ShotLifespan = 120
    .ShotRadius = 60
    .ShotSpeed = 6
    .ShotSpread = 0
    .ShotVisible = True
    .StartAmmo = 2
    .SpeedBonus = 0
    .WallPiercing = False
End With
VariantWeapon(1) = 1
VariantWeapon(2) = 1
VariantWeapon(3) = 1
VariantWeapon(4) = 1
VariantWeapon(5) = 1
VariantWeapon(6) = 1
VariantWeapon(7) = 1
VariantWeapon(8) = 1
VariantWeapon(9) = 1
VariantWeapon(10) = 1

End Sub
Public Sub InitializeWeaponsSwords()
WeaponDataCount = 2
ReDim WeaponData(1 To 2)
StartWeapon = 1
With WeaponData(1)
    .Bounce = False
    .CoolDown = 100
    .ClipSize = 1000
    .Damage = 300
    .DamageMultiplier = 0
    .ShotExplosive = False
    .GraphicsIndex = 7
    .IconIndex = 7
    .Melee = False
    .MeleeDamage = 0
    .Name = "Saber"
    .ReloadTime = 0
    .SemiAuto = True
    .ShotGraphicsIndex = 1
    .ShotLifespan = 8
    .ShotRadius = 10
    .ShotSpeed = 15
    .ShotSpread = 0
    .ShotVisible = False
    .StartAmmo = 20
    .SpeedBonus = 1
    .WallPiercing = False
End With
With WeaponData(2)
    .Bounce = False
    .CoolDown = 50
    .ClipSize = 4
    .Damage = 1
    .DamageMultiplier = 0.1
    .ShotExplosive = False
    .GraphicsIndex = 5
    .IconIndex = 5
    .Melee = True
    .MeleeDamage = 180
    .Name = "Sniper Rifle"
    .ReloadTime = 60
    .SemiAuto = True
    .ShotGraphicsIndex = 4
    .ShotLifespan = 120
    .ShotRadius = 25
    .ShotSpeed = 30
    .ShotSpread = 0
    .ShotVisible = True
    .StartAmmo = 16
    .SpeedBonus = 0
    .WallPiercing = False
End With
VariantWeapon(1) = 1
VariantWeapon(2) = 1
VariantWeapon(3) = 1
VariantWeapon(4) = 1
VariantWeapon(5) = 2
VariantWeapon(6) = 1
VariantWeapon(7) = 1
VariantWeapon(8) = 1
VariantWeapon(9) = 1
VariantWeapon(10) = 1
End Sub

Public Sub InitializeWeaponsHardcoreNormal()
WeaponDataCount = 10
ReDim WeaponData(1 To 10)
StartWeapon = 1
With WeaponData(1)
    .Bounce = False
    .CoolDown = 8
    .ClipSize = 80
    .Damage = 22
    .DamageMultiplier = 0
    .ShotExplosive = False
    .GraphicsIndex = 1
    .IconIndex = 1
    .Melee = True
    .MeleeDamage = 100
    .Name = "Laser Pistol"
    .ReloadTime = 50
    .SemiAuto = False
    .ShotGraphicsIndex = 1
    .ShotLifespan = 80
    .ShotRadius = 5
    .ShotSpeed = 10
    .ShotSpread = 0
    .ShotVisible = True
    .StartAmmo = 160
    .SpeedBonus = 1
    .WallPiercing = False
End With
With WeaponData(2)
    .Bounce = True
    .CoolDown = 40
    .ClipSize = 4
    .Damage = 180
    .DamageMultiplier = 0
    .ShotExplosive = True
    .GraphicsIndex = 2
    .IconIndex = 2
    .Melee = True
    .MeleeDamage = 40
    .Name = "Grenades"
    .ReloadTime = 0
    .SemiAuto = True
    .ShotGraphicsIndex = 2
    .ShotLifespan = 120
    .ShotRadius = 60
    .ShotSpeed = 6
    .ShotSpread = 0
    .ShotVisible = True
    .StartAmmo = 4
    .SpeedBonus = 1
    .WallPiercing = False
End With
With WeaponData(3)
    .Bounce = False
    .CoolDown = 4
    .ClipSize = 80
    .Damage = 12
    .DamageMultiplier = 0
    .ShotExplosive = False
    .GraphicsIndex = 4
    .IconIndex = 4
    .Melee = True
    .MeleeDamage = 45
    .Name = "SMG"
    .ReloadTime = 60
    .SemiAuto = False
    .ShotGraphicsIndex = 4
    .ShotLifespan = 110
    .ShotRadius = 14
    .ShotSpeed = 10
    .ShotSpread = 6
    .ShotVisible = True
    .StartAmmo = 240
    .SpeedBonus = 1
    .WallPiercing = False
End With
With WeaponData(4)
    .Bounce = False
    .CoolDown = 25
    .ClipSize = 6
    .Damage = 240
    .DamageMultiplier = -0.55
    .ShotExplosive = False
    .GraphicsIndex = 3
    .IconIndex = 3
    .Melee = False
    .MeleeDamage = 0
    .Name = "Shotgun"
    .ReloadTime = 60
    .SemiAuto = True
    .ShotGraphicsIndex = 1
    .ShotLifespan = 35
    .ShotRadius = 40
    .ShotSpeed = 10
    .ShotSpread = 0
    .ShotVisible = False
    .StartAmmo = 12
    .SpeedBonus = 1
    .WallPiercing = False
End With
With WeaponData(5)
    .Bounce = False
    .CoolDown = 10
    .ClipSize = 12
    .Damage = 25
    .DamageMultiplier = 0.05
    .ShotExplosive = False
    .GraphicsIndex = 5
    .IconIndex = 5
    .Melee = True
    .MeleeDamage = 120
    .Name = "Battle Rifle"
    .ReloadTime = 30
    .SemiAuto = True
    .ShotGraphicsIndex = 4
    .ShotLifespan = 120
    .ShotRadius = 25
    .ShotSpeed = 30
    .ShotSpread = 0
    .ShotVisible = True
    .StartAmmo = 36
    .SpeedBonus = 1
    .WallPiercing = False
End With
With WeaponData(6)
    .Bounce = False
    .CoolDown = 25
    .ClipSize = 3
    .Damage = 160
    .DamageMultiplier = 0
    .ShotExplosive = True
    .GraphicsIndex = 6
    .IconIndex = 6
    .Melee = True
    .MeleeDamage = 40
    .Name = "Mines"
    .ReloadTime = 0
    .SemiAuto = True
    .ShotGraphicsIndex = 1
    .ShotLifespan = 2000
    .ShotRadius = 60
    .ShotSpeed = 0
    .ShotSpread = 0
    .ShotVisible = False
    .StartAmmo = 6
    .SpeedBonus = 1
    .WallPiercing = False
End With
With WeaponData(7)
    .Bounce = False
    .CoolDown = 100
    .ClipSize = 1000
    .Damage = 300
    .DamageMultiplier = 0
    .ShotExplosive = False
    .GraphicsIndex = 7
    .IconIndex = 7
    .Melee = False
    .MeleeDamage = 0
    .Name = "Saber"
    .ReloadTime = 0
    .SemiAuto = True
    .ShotGraphicsIndex = 1
    .ShotLifespan = 8
    .ShotRadius = 20
    .ShotSpeed = 15
    .ShotSpread = 0
    .ShotVisible = False
    .StartAmmo = 30
    .SpeedBonus = 2
    .WallPiercing = False
End With
With WeaponData(8)
    .Bounce = True
    .CoolDown = 3
    .ClipSize = 120
    .Damage = 8
    .DamageMultiplier = 0
    .ShotExplosive = False
    .GraphicsIndex = 8
    .IconIndex = 1
    .Melee = False
    .MeleeDamage = 80
    .Name = "Chain Gun"
    .ReloadTime = 75
    .SemiAuto = False
    .ShotGraphicsIndex = 3
    .ShotLifespan = 105
    .ShotRadius = 15
    .ShotSpeed = 10
    .ShotSpread = 12
    .ShotVisible = True
    .StartAmmo = 240
    .SpeedBonus = 0.8
    .WallPiercing = False
End With
With WeaponData(9)
    .Bounce = False
    .CoolDown = 0
    .ClipSize = 12
    .Damage = 52
    .DamageMultiplier = 0
    .ShotExplosive = False
    .GraphicsIndex = 9
    .IconIndex = 1
    .Melee = True
    .MeleeDamage = 90
    .Name = "Pistol"
    .ReloadTime = 25
    .SemiAuto = True
    .ShotGraphicsIndex = 4
    .ShotLifespan = 50
    .ShotRadius = 12
    .ShotSpeed = 14
    .ShotSpread = 1
    .ShotVisible = True
    .StartAmmo = 36
    .SpeedBonus = 1
    .WallPiercing = False
End With
With WeaponData(10)
    .Bounce = False
    .CoolDown = 0
    .ClipSize = 1
    .Damage = 300
    .DamageMultiplier = 0
    .ShotExplosive = True
    .GraphicsIndex = 8
    .IconIndex = 1
    .Melee = True
    .MeleeDamage = 155
    .Name = "Rocket Launcher"
    .ReloadTime = 180
    .SemiAuto = True
    .ShotGraphicsIndex = 1
    .ShotLifespan = 133
    .ShotRadius = 90
    .ShotSpeed = 15
    .ShotSpread = 0
    .ShotVisible = True
    .StartAmmo = 5
    .SpeedBonus = 0.4
    .WallPiercing = False
End With

VariantWeapon(1) = 1
VariantWeapon(2) = 2
VariantWeapon(3) = 3
VariantWeapon(4) = 4
VariantWeapon(5) = 5
VariantWeapon(6) = 6
VariantWeapon(7) = 7
VariantWeapon(8) = 8
VariantWeapon(9) = 9
VariantWeapon(10) = 10

End Sub
