Attribute VB_Name = "modDeclarations"

'Enemy Graphix Codes
'3 = 2
'1 = 3
'4 = 4
'5 = 5
'6 = 2
'7 = 5
Type Player
    XCoord As Double ' the x-coordinate of the player on the map
    YCoord As Double ' the y-coordinate of the player on the map
    TargetX As Double 'the coordinates of the cursor point
    TargetY As Double
    LastTargetX As Double
    LastTargetY As Double
    MoveX As Double
    MoveY As Double
    TurnLeft As Boolean
    TurnRight As Boolean
    MoveForward As Boolean
    MoveBack As Boolean
    StrafeLeft As Boolean
    StrafeRight As Boolean
    Theta As Double 'the angle of rotation, expressed in a percentage of 360 degrees.
    HP As Integer 'the player's hitpoints value
    Shield As Integer 'the player's shield hitpoints that recharge after every kill
    Weapon(1 To 2) As Integer 'index of currently carried weapons (1 = active)
    Ammo(1 To 2) As Integer
    ClipAmmo(1 To 2) As Integer
    CoolDownTime As Integer
    Deaths As Integer
    Score As Integer
    Firing As Boolean ' whether or not the player is shooting.
    Reloading As Boolean
    PickUpItem As Boolean
    ReloadTime As Integer 'in frames.
    RechargeTime As Integer 'time for shields to recharge after a hit
    Dead As Boolean
    RespawnTime As Integer 'in seconds
    Sniping As Boolean
    SnipeToggle As Boolean
    Switching As Boolean
End Type
Type PlayerData
    Name As String
    Level As Integer
    Difficulty As Integer
    Weapon(1 To 2) As Integer
    Ammo(1 To 2) As Integer
    ClipAmmo(1 To 2) As Integer
    Layout As Integer '1 = basic 2 = classic 3 = benny's
    Inverted As Boolean
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
    AmmoName(1 To 2) As String '1 = singular, 2 = plural
    GraphicsIndex As Integer 'Index of graphics used ON SCREEN
    IconIndex As Integer 'Index of graphics used as the ICON
    ShotGraphicsIndex As Integer
    Damage As Integer
    Despawns As Boolean
    ClipSize As Integer
    MaxAmmo As Integer
    SpeedBonus As Double 'bonus to player's speed
    WallPiercing As Boolean 'whether the shot pierces walls
    StartAmmo As Integer 'COUNTING starting clip
    ShotRadius As Integer 'pixels
    ExplodeRadius As Integer
    CoolDown As Integer 'in frames
    Reloads As Boolean 'whether or not the weapon reloads
    ReloadTime As Integer 'time in frames
    MeleeDamage As Integer ' damage done in melee
    PickUpAmmo As Boolean
    Bounce As Boolean 'whether or not shots bounce
    DamageMultiplier As Double 'totaldamage = damage + damagemultiplier * distance
    ExplodeDamage As Integer 'damage done at center of explosion - decreases to 0 at edge of explosion radius.
    Arc As Double 'multiplied by distance to cursor to determine range
    ShotSpeed As Double 'pixels per frame -- should not exceed 32.00
    FrameSteps As Integer 'number of times to move shot in 1 frame
    ShotSpread As Integer 'max variation from mean in pixels at 100 px range
    ShotVisible As Boolean 'whether shot is visible
    ShotLifespan As Integer
    SemiAuto As Boolean
End Type
Type Shot
    Bounce As Boolean
    GraphicsIndex As Integer
    Alignment As Integer 'Player = 1, hostile = 2
    Damage As Integer
    XCoord As Double
    YCoord As Double
    XVector As Double
    YVector As Double
    ExplodeRadius As Integer
    ExplodeDamage As Integer
    FrameSteps As Integer
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
Type Holoscreen
    XCoord As Double
    YCoord As Double
    Text As String
    Goal As Boolean
    GraphicsDC As Integer
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
Type Enemy
    XCoord As Double
    YCoord As Double
    RadarBlipX As Double
    RadarBlipY As Double
    TargetX As Double 'the coordinates of the point at which this enemy most recently saw the player
    TargetY As Double
    MoveSpeed As Double
    MoveForward As Integer
    HP As Integer 'the enemy's hitpoints value
    Weapon As Integer 'index of currently carried weapon
    Ammo As Integer
    ClipAmmo As Integer
    CoolDownTime As Integer
    Reloading As Boolean
    ReloadTime As Integer 'in frames.
    Stance As Integer '1 = dormant defensive 2 = dormant aggressive 3 = active defensive 4 = active agressive
    Dead As Boolean
    VisualContact As Boolean 'whether enemy can see the player
    GraphicsDC As Long
End Type
Type Trigger
    XCoord As Double
    YCoord As Double
    Height As Double
    Width As Double
    Link As Integer
End Type
Public ObjectGraphicsDC(1 To 3) As Long, ObjectMaskDC(1 To 3) As Long, PlayerGraphicsDC As Long, WeaponGraphicsDC(1 To 12) As Long, ShotGraphicsDC(1 To 6) As Long
Public ExplosionGraphicsDC(0 To 1) As Long, ExplosionMaskDC(0 To 1) As Long, ExplosionCount As Integer, ExplosionMin As Integer, Explosion() As Explosion
Public EnemyMaskDC(1 To 10) As Long, EnemyDC(1 To 10) As Long
Public PlayerMaskDC As Long, BackBuffDC As Long, BackgroundDC As Long, WeaponMaskDC(1 To 12) As Long, SightDC As Long, SightMaskDC As Long, ShotMaskDC(1 To 6) As Long
Public WallDC(1 To 2) As Long '1 = normal 2 = blocks players not shots
Public RadarBlipDC(1 To 2) As Long, RadarBlipMaskDC As Long
Public ShieldBarDC As Long, HealthBarDC As Long, GameOverDC As Long
Public Wall() As Wall, WallCount As Integer
Public Shot() As Shot, ShotCount As Integer, ShotMin As Integer
Public Weapon() As Weapon, WeaponCount As Integer, WeaponData() As WeaponData, WeaponMin As Integer
Public ScoreToWin(1 To 2) As Integer, GameOver As Boolean, UnloadCountdown As Integer, Terminated As Boolean
Public StartWeapon(1 To 2) As Integer
Public Holoscreen() As Holoscreen, Holocount As Integer
'dimension object stat database
Public Player As Player
Public GameOverCountdown As Integer
Public Enemy() As Enemy, EnemyCount As Integer
Public PlayerRespawnTime As Integer
Public CurrentLevel As Integer, CurrentAccount As Integer
Public AccountData(0 To 9) As PlayerData
Public victory As Boolean
Public Const PlayerMaxHealth = 250, PlayerMaxShield = 125
Public PlayerBonus As Double, EnemyBonus As Double
Public PromptTime As Integer
Public Trigger() As Trigger, TriggerCount As Integer



Public Sub InitializeWeapons()
WeapondataCount = 11
ReDim WeaponData(1 To WeapondataCount)
With WeaponData(1)
    .Name = "Battle Rifle"
    .AmmoName(1) = "battle rifle shell"
    .AmmoName(2) = "battle rifle shells"
    .Arc = 0
    .Bounce = False
    .ClipSize = 12
    .CoolDown = 24
    .Damage = 15
    .DamageMultiplier = 0.01
    .Despawns = False
    .ExplodeDamage = 0
    .ExplodeRadius = 0
    .FrameSteps = 3
    .GraphicsIndex = 5
    .IconIndex = 5
    .MaxAmmo = 72
    .MeleeDamage = 14
    .PickUpAmmo = True
    .Reloads = True
    .ReloadTime = 100
    .SemiAuto = True
    .ShotGraphicsIndex = 4
    .ShotLifespan = 70
    .ShotRadius = 12
    .ShotSpeed = 30
    .ShotSpread = 0
    .ShotVisible = True
    .SpeedBonus = 0
    .StartAmmo = 60
End With
With WeaponData(2)
    .Name = "Grenades"
    .AmmoName(1) = "grenade"
    .AmmoName(2) = "grenades"
    .Arc = 1.2
    .Bounce = True
    .ClipSize = 0
    .CoolDown = 80
    .Damage = 0
    .DamageMultiplier = 0
    .Despawns = True
    .ExplodeDamage = 300
    .ExplodeRadius = 225
    .FrameSteps = 1
    .GraphicsIndex = 2
    .IconIndex = 2
    .MaxAmmo = 4
    .MeleeDamage = 7
    .PickUpAmmo = True
    .Reloads = False
    .ReloadTime = 0
    .SemiAuto = True
    .ShotGraphicsIndex = 2
    .ShotLifespan = 0
    .ShotRadius = 12
    .ShotSpeed = 8
    .ShotSpread = 0
    .ShotVisible = True
    .SpeedBonus = 0
    .StartAmmo = 2
End With
With WeaponData(3)
    .Name = "Laser Pistol"
    .AmmoName(1) = ""
    .AmmoName(2) = ""
    .Arc = 0
    .Bounce = False
    .ClipSize = 0
    .CoolDown = 10
    .Damage = 6
    .DamageMultiplier = 0
    .Despawns = False
    .ExplodeDamage = 0
    .ExplodeRadius = 0
    .FrameSteps = 1
    .GraphicsIndex = 1
    .IconIndex = 1
    .MaxAmmo = 0
    .MeleeDamage = 15
    .PickUpAmmo = False
    .Reloads = False
    .ReloadTime = 0
    .SemiAuto = False
    .ShotGraphicsIndex = 1
    .ShotLifespan = 42
    .ShotRadius = 7
    .ShotSpeed = 15
    .ShotSpread = 5
    .ShotVisible = True
    .SpeedBonus = 0
    .StartAmmo = 120
End With
With WeaponData(4)
    .Name = "Magnum"
    .AmmoName(1) = "round for magnum"
    .AmmoName(2) = "rounds for magnum"
    .Arc = 0
    .Bounce = False
    .ClipSize = 10
    .CoolDown = 2
    .Damage = 28
    .DamageMultiplier = 0
    .Despawns = False
    .ExplodeDamage = 0
    .ExplodeRadius = 0
    .FrameSteps = 2
    .GraphicsIndex = 9
    .IconIndex = 8
    .MaxAmmo = 50
    .MeleeDamage = 12
    .PickUpAmmo = True
    .Reloads = True
    .ReloadTime = 65
    .SemiAuto = True
    .ShotGraphicsIndex = 4
    .ShotLifespan = 26
    .ShotRadius = 3
    .ShotSpeed = 28
    .ShotSpread = 1
    .ShotVisible = True
    .SpeedBonus = 0
    .StartAmmo = 20
End With
With WeaponData(5)
    .Name = "SMG"
    .AmmoName(1) = "round for SMG"
    .AmmoName(2) = "rounds for SMG"
    .Arc = 0
    .Bounce = False
    .ClipSize = 90
    .CoolDown = 6
    .Damage = 12
    .DamageMultiplier = -0.02
    .Despawns = False
    .ExplodeDamage = 0
    .ExplodeRadius = 0
    .FrameSteps = 1
    .GraphicsIndex = 4
    .IconIndex = 4
    .MaxAmmo = 360
    .MeleeDamage = 20
    .PickUpAmmo = True
    .Reloads = True
    .ReloadTime = 100
    .SemiAuto = False
    .ShotGraphicsIndex = 4
    .ShotLifespan = 50
    .ShotRadius = 50
    .ShotSpeed = 15
    .ShotSpread = 12
    .ShotVisible = True
    .SpeedBonus = 0
    .StartAmmo = 180
End With
With WeaponData(6)
    .Name = "Laser Rifle"
    .AmmoName(1) = ""
    .AmmoName(2) = ""
    .Arc = 0
    .Bounce = False
    .ClipSize = 0
    .CoolDown = 15
    .Damage = 17
    .DamageMultiplier = 0
    .Despawns = False
    .ExplodeDamage = 0
    .ExplodeRadius = 0
    .FrameSteps = 1
    .GraphicsIndex = 12
    .IconIndex = 1
    .MaxAmmo = 0
    .MeleeDamage = 18
    .PickUpAmmo = False
    .Reloads = False
    .ReloadTime = 0
    .SemiAuto = False
    .ShotGraphicsIndex = 1
    .ShotLifespan = 70
    .ShotRadius = 7
    .ShotSpeed = 15
    .ShotSpread = 5
    .ShotVisible = True
    .SpeedBonus = 0
    .StartAmmo = 120
End With
With WeaponData(7)
    .Name = "Assault Rifle"
    .AmmoName(1) = "round for assault rife"
    .AmmoName(2) = "rounds for assault rifle"
    .Arc = 0
    .Bounce = False
    .ClipSize = 60
    .CoolDown = 8
    .Damage = 12
    .DamageMultiplier = 0.02
    .Despawns = False
    .ExplodeDamage = 0
    .ExplodeRadius = 0
    .FrameSteps = 1
    .GraphicsIndex = 10
    .IconIndex = 4
    .MaxAmmo = 240
    .MeleeDamage = 22
    .PickUpAmmo = True
    .Reloads = True
    .ReloadTime = 85
    .SemiAuto = False
    .ShotGraphicsIndex = 4
    .ShotLifespan = 60
    .ShotRadius = 7
    .ShotSpeed = 15
    .ShotSpread = 12
    .ShotVisible = True
    .SpeedBonus = 0
    .StartAmmo = 120
End With
With WeaponData(8)
    .Name = "Shock Rifle"
    .AmmoName(1) = "shock rifle cylinder"
    .AmmoName(2) = "shock rifle cylinders"
    .Arc = 0
    .Bounce = True
    .ClipSize = 8
    .CoolDown = 24
    .Damage = 50
    .DamageMultiplier = -0.02
    .Despawns = False
    .ExplodeDamage = 0
    .ExplodeRadius = 0
    .FrameSteps = 1
    .GraphicsIndex = 11
    .IconIndex = 11
    .MaxAmmo = 48
    .MeleeDamage = 14
    .PickUpAmmo = True
    .Reloads = True
    .ReloadTime = 100
    .SemiAuto = True
    .ShotGraphicsIndex = 4
    .ShotLifespan = 20
    .ShotRadius = 20
    .ShotSpeed = 30
    .ShotSpread = 0
    .ShotVisible = True
    .SpeedBonus = 0
    .StartAmmo = 32
End With
With WeaponData(9)
    .Name = "Shotgun"
    .AmmoName(1) = "shotgun shell"
    .AmmoName(2) = "shotgun shells"
    .Arc = 0
    .Bounce = True
    .ClipSize = 12
    .CoolDown = 24
    .Damage = 100
    .DamageMultiplier = -0.05
    .Despawns = False
    .ExplodeDamage = 0
    .ExplodeRadius = 0
    .FrameSteps = 1
    .GraphicsIndex = 3
    .IconIndex = 3
    .MaxAmmo = 36
    .MeleeDamage = 23
    .PickUpAmmo = True
    .Reloads = True
    .ReloadTime = 100
    .SemiAuto = True
    .ShotGraphicsIndex = 4
    .ShotLifespan = 10
    .ShotRadius = 12
    .ShotSpeed = 30
    .ShotSpread = 0
    .ShotVisible = False
    .SpeedBonus = 0
    .StartAmmo = 12
End With

End Sub

