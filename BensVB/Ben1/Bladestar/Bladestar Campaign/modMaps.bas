Attribute VB_Name = "modMaps"
Public Sub GenerateWall(x As Double, y As Double, Height As Double, Width As Double, WallType As Integer)
WallCount = WallCount + 1
ReDim Preserve Wall(1 To WallCount)
With Wall(WallCount)
    .Height = Height
    .Type = WallType
    .Width = Width
    .XCoord = x
    .YCoord = y
End With
End Sub
Public Sub LoadLevel1()

With Player
    .Ammo(1) = 0
    .Ammo(2) = 0
    .ClipAmmo(1) = 0
    .ClipAmmo(2) = 0
    .CoolDownTime = 0
    .Dead = False
    .Deaths = 0
    .Firing = False
    .HP = PlayerMaxHealth * PlayerBonus
    .MoveBack = False
    .MoveForward = False
    .PickUpItem = False
    .Reloading = False
    .ReloadTime = 0
    .RespawnTime = 0
    .Score = 0
    .Shield = 0
    .StrafeLeft = False
    .StrafeRight = False
    .Switching = False
    .TargetX = 0
    .TargetY = 0
    .Theta = PI
    .TurnLeft = False
    .TurnRight = False
    .Weapon(1) = 3 'laser pistol
    .Weapon(2) = 4 'magnum
    .XCoord = 600
    .YCoord = 3600
End With
Dim i
'=================
'load map graphics
'=================
BackgroundDC = GenerateDC(App.Path & "\Graphics\Background1.bmp")
BackBuffDC = GenerateDC(App.Path & "\Graphics\Background1.bmp")
For i = 1 To 2
WallDC(i) = GenerateDC(App.Path & "\Graphics\Wall" & i & "_1.bmp")
Next i
'1
Call GenerateWall(0, 0, 500, 3200, 1)
'2
Call GenerateWall(0, 500, 3700, 500, 1)
'3
Call GenerateWall(700, 800, 500, 200, 1)
'4
Call GenerateWall(1100, 700, 900, 700, 1)
'5
Call GenerateWall(2400, 500, 600, 800, 1)
'6
Call GenerateWall(2400, 1100, 300, 100, 1)
'7
Call GenerateWall(2700, 1100, 500, 500, 1)
'8
Call GenerateWall(900, 1600, 200, 2300, 1)
'9
Call GenerateWall(500, 1600, 400, 200, 1)
'10
Call GenerateWall(900, 1800, 400, 100, 1)
'11
Call GenerateWall(2700, 1800, 600, 500, 1)
'12
Call GenerateWall(1200, 2000, 200, 200, 1)
'13
Call GenerateWall(2000, 2000, 200, 200, 1)
'14
Call GenerateWall(700, 2200, 300, 300, 1)
'15
Call GenerateWall(2400, 2400, 300, 800, 1)
'16
Call GenerateWall(2700, 2700, 700, 600, 1)
'17
Call GenerateWall(2800, 3400, 900, 500, 1)
'18
Call GenerateWall(2400, 3800, 500, 400, 1)
'19
Call GenerateWall(1800, 3700, 700, 600, 1)
'20
Call GenerateWall(1200, 3900, 500, 600, 1)
'21
Call GenerateWall(500, 3700, 700, 700, 1)
'22
Call GenerateWall(700, 3200, 300, 500, 1)
'23
Call GenerateWall(500, 2800, 400, 700, 1)
'24
Call GenerateWall(1200, 2400, 900, 1000, 1)
'25
Call GenerateWall(2200, 2900, 500, 300, 1)
'26
Call GenerateWall(1800, 3300, 200, 600, 1)
WeaponCount = 10
ReDim Weapon(1 To WeaponCount)
With Weapon(1)
    .XCoord = 2500
    .YCoord = 3500
    .Ammo = 120
    .ClipAmmo = 0
    .Deleted = False
    .WeaponType = 3
End With
With Weapon(2)
    .XCoord = 2700
    .YCoord = 3500
    .Ammo = 48
    .ClipAmmo = 12
    .Deleted = False
    .WeaponType = 1
End With
With Weapon(3)
    .XCoord = 2500
    .YCoord = 3700
    .Ammo = 36
    .ClipAmmo = 12
    .Deleted = False
    .WeaponType = 1
End With
With Weapon(4)
    .XCoord = 2700
    .YCoord = 3700
    .Ammo = 20
    .ClipAmmo = 10
    .Deleted = False
    .WeaponType = 4
End With
With Weapon(5)
    .XCoord = 800
    .YCoord = 2100
    .Ammo = 4
    .ClipAmmo = 0
    .Deleted = False
    .WeaponType = 2
End With
With Weapon(6)
    .XCoord = 600
    .YCoord = 1500
    .Ammo = 4
    .ClipAmmo = 0
    .Deleted = False
    .WeaponType = 2
End With
With Weapon(7)
    .XCoord = 1000
    .YCoord = 1500
    .Ammo = 4
    .ClipAmmo = 0
    .Deleted = False
    .WeaponType = 2
End With
With Weapon(8)
    .XCoord = 800
    .YCoord = 600
    .Ammo = 20
    .ClipAmmo = 10
    .Deleted = False
    .WeaponType = 4
End With
With Weapon(9)
    .XCoord = 1400
    .YCoord = 600
    .Ammo = 36
    .ClipAmmo = 12
    .Deleted = False
    .WeaponType = 1
End With
With Weapon(10)
    .XCoord = 1500
    .YCoord = 600
    .Ammo = 120
    .ClipAmmo = 0
    .Deleted = False
    .WeaponType = 3
End With

EnemyCount = 14
ReDim Enemy(1 To EnemyCount)
With Enemy(1)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 2100
    .TargetY = 1900
    .XCoord = 2100
    .YCoord = 1900
End With
With Enemy(2)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 2600
    .TargetY = 1900
    .XCoord = 2600
    .YCoord = 1900
End With
With Enemy(3)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 2600
    .TargetY = 2300
    .XCoord = 2600
    .YCoord = 2300
End With
With Enemy(4)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 1800
    .TargetY = 2300
    .XCoord = 1800
    .YCoord = 2300
End With
With Enemy(5)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 1700
    .TargetY = 2100
    .XCoord = 1700
    .YCoord = 2100
End With
With Enemy(6)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 1100
    .TargetY = 2100
    .XCoord = 1100
    .YCoord = 2100
End With
With Enemy(7)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 70
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 600
    .TargetY = 2700
    .XCoord = 600
    .YCoord = 2700
End With
With Enemy(8)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 70
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 600
    .TargetY = 2400
    .XCoord = 600
    .YCoord = 2400
End With
With Enemy(9)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 120
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 600
    .TargetY = 1200
    .XCoord = 600
    .YCoord = 1200
End With
With Enemy(10)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 120
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 1000
    .TargetY = 1200
    .XCoord = 1000
    .YCoord = 1200
End With
With Enemy(11)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 120
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 800
    .TargetY = 600
    .XCoord = 800
    .YCoord = 600
End With
With Enemy(12)
    .Weapon = 3
    .Ammo = 120
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 2
    .HP = 80
    .MoveSpeed = 2.5
    .Stance = 3
    .TargetX = 1900
    .TargetY = 1500
    .XCoord = 1900
    .YCoord = 1500
End With
With Enemy(13)
    .Weapon = 3
    .Ammo = 120
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 2
    .HP = 80
    .MoveSpeed = 2.5
    .Stance = 3
    .TargetX = 2300
    .TargetY = 1500
    .XCoord = 2300
    .YCoord = 1500
End With
With Enemy(14)
    .Weapon = 3
    .Ammo = 120
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 2
    .HP = 80
    .MoveSpeed = 2.5
    .Stance = 3
    .TargetX = 2100
    .TargetY = 1300
    .XCoord = 2100
    .YCoord = 1300
End With

Holocount = 10
ReDim Holoscreen(1 To Holocount)
With Holoscreen(1)
    .Goal = False
Select Case AccountData(CurrentAccount).Layout
    Case Is = 1
    .Text = "Welcome to Bladestar.  If you already know how to play, you can skip this tutorial by using the green exit portal.  Move forward and backward using the W and S keys.  Move sideways by pressing A and D.  Go east to the next holoscreen to continue training."
    Case Is = 2
    .Text = "Welcome to Bladestar.  If you already know how to play, you can skip this tutorial by using the green exit portal.  Move forward and backward using the W and S keys.  Turn using A and D.  You can move sideways by pressing Q and E.  Go east to the next holoscreen to continue training."
    Case Is = 3
    .Text = "Welcome to Bladestar.  If you already know how to play, you can skip this tutorial by using the green exit portal.  Move forward and backward using the W and S keys.  Move sideways by pressing A and D.  Use the mouse to turn and aim.  Go east to the next holoscreen to continue training."
    End Select
    .XCoord = 600
    .YCoord = 3600
    .GraphicsDC = 2
End With
With Holoscreen(2)
    .Goal = False
    .Text = "To the east is a weapons storeroom.  Your active weapon is displayed at the top right of the screen, along with the number of shots in your ammo inventory.  Your second weapon is displayed underneath.  To switch weapons, press the shift key.  You can pick up a weapon on the ground by standing near it and right-clicking it.  Note that this causes you to drop whichever weapon you are using."
    .XCoord = 1500
    .YCoord = 3600
    .GraphicsDC = 2
End With
With Holoscreen(3)
    .Goal = True
    .Text = "Level complete."
    .XCoord = 600
    .YCoord = 3300
    .GraphicsDC = 3
End With
With Holoscreen(4)
    .Goal = False
    .Text = "When you move over a weapon of the same type as one you are holding, you can pick up any spare ammo.  Try it with the battle rifle - the laser pistol's energy is stored in a battery, so you can't pick up more ammo for it.  Once you have selected your weapons, proceed to the north.  The next stage of your training will be target practice using some deactivated droids."
    .XCoord = 2600
    .YCoord = 3600
    .GraphicsDC = 2
End With
With Holoscreen(5)
    .Goal = False
    .Text = "When your line of sight to an enemy is not blocked, you will see a light blue dot near you which will indicate the enemy's position and distance.  The dot turns red when the enemy is within range of your current weapon.  To shoot, left-click anywhere on the screen.  Some weapons, like the laser pistol, also allow you to hold down the mouse button for rapid-fire.  Click anywhere with the right mouse button to reload.  Note that the laser pistol doesn't need to reload."
    .XCoord = 2300
    .YCoord = 2800
    .GraphicsDC = 2
End With
With Holoscreen(6)
    .Goal = False
    .Text = "You can also destroy targets using melee combat.  While standing near an enemy, move your cursor quickly over them to deal damage in melee.  Practice on these droids."
    .XCoord = 1100
    .YCoord = 2400
    .GraphicsDC = 2
End With
With Holoscreen(7)
    .Goal = False
    .Text = "The next skill you will learn is how to use grenades.  Unlike other weapons, grenades do not have a fixed maximum range.  The farther away you click to throw them, the farther they will go.  You can also bounce them off walls.  They can hurt you if they explode nearby, so don't throw them too close.  Pick up some nades and practice on the targets in the next room."
    .XCoord = 600
    .YCoord = 2100
    .GraphicsDC = 2
End With
With Holoscreen(8)
    .Goal = False
    .Text = "Now you will fight against activated droids.  They will move and shoot at you, so be sure to dodge.  If you are hit, your green shield bar will decrease, and if it reaches zero, your health will start to take damage.  Your health only heals between levels, but your shields will recharge fully if you can avoid getting hit for a few seconds.  Happy hunting!"
    .XCoord = 1200
    .YCoord = 600
    .GraphicsDC = 2
End With
With Holoscreen(9)
    .Goal = False
    .Text = "Congrats!  You survived the preliminary training session...now head north to the green portal to start the actual game!  If you want to replay this level, click the exit button in the lower right corner of the screen."
    .XCoord = 2600
    .YCoord = 1500
    .GraphicsDC = 2
End With
With Holoscreen(10)
    .Goal = True
    .Text = "Level complete."
    .XCoord = 2600
    .YCoord = 1200
    .GraphicsDC = 3
End With
TriggerCount = 0
End Sub

Public Sub LoadLevel0()

With Player
    .Ammo(1) = 48
    .Ammo(2) = 2
    .ClipAmmo(1) = 12
    .ClipAmmo(2) = 0
    .CoolDownTime = 0
    .Dead = False
    .Deaths = 0
    .Firing = False
    .HP = PlayerMaxHealth * PlayerBonus
    .MoveBack = False
    .MoveForward = False
    .PickUpItem = False
    .Reloading = False
    .ReloadTime = 0
    .RespawnTime = 0
    .Score = 0
    .Shield = PlayerMaxShield * PlayerBonus
    .StrafeLeft = False
    .StrafeRight = False
    .Switching = False
    .TargetX = 0
    .TargetY = 0
    .Theta = 0
    .TurnLeft = False
    .TurnRight = False
    .Weapon(1) = 1 'battle rifle
    .Weapon(2) = 2 'grenades
    .XCoord = 680
    .YCoord = 540
End With
Dim i
'=================
'load map graphics
'=================
BackgroundDC = GenerateDC(App.Path & "\Graphics\Background1.bmp")
BackBuffDC = GenerateDC(App.Path & "\Graphics\Background1.bmp")
For i = 1 To 2
WallDC(i) = GenerateDC(App.Path & "\Graphics\Wall" & i & "_1.bmp")
Next i
'1
Call GenerateWall(320, 200, 40, 200, 2)
'2
Call GenerateWall(520, 200, 40, 400, 1)
'3
Call GenerateWall(920, 200, 40, 1120, 2)
'4
Call GenerateWall(320, 240, 600, 40, 2)
'5
Call GenerateWall(520, 400, 280, 40, 1)
'6
Call GenerateWall(880, 400, 40, 160, 1)
'7
Call GenerateWall(1200, 400, 40, 160, 1)
'8
Call GenerateWall(1360, 400, 40, 520, 2)
'9
Call GenerateWall(2040, 200, 1040, 40, 2)
'10
Call GenerateWall(560, 640, 40, 320, 1)
'11
Call GenerateWall(880, 440, 400, 40, 1)
'12
Call GenerateWall(1320, 440, 400, 40, 1)
'13
Call GenerateWall(1840, 440, 800, 40, 2)
'14
Call GenerateWall(320, 840, 40, 560, 2)
'15
Call GenerateWall(880, 840, 40, 160, 1)
'16
Call GenerateWall(1200, 840, 40, 160, 1)
'17
Call GenerateWall(1000, 880, 280, 40, 2)
'18
Call GenerateWall(1200, 880, 280, 40, 2)
'19
Call GenerateWall(760, 1160, 240, 40, 2)
'20
Call GenerateWall(800, 1160, 40, 240, 2)
'21
Call GenerateWall(1200, 1160, 40, 240, 2)
'22
Call GenerateWall(1440, 1160, 240, 40, 2)
'23
Call GenerateWall(80, 1240, 40, 480, 1)
'24
Call GenerateWall(1640, 1240, 40, 240, 1)
'25
Call GenerateWall(2040, 1240, 40, 400, 1)
'26
Call GenerateWall(40, 1240, 800, 40, 1)
'27
Call GenerateWall(160, 1480, 40, 120, 1)
'28
Call GenerateWall(280, 1400, 120, 40, 1)
'29
Call GenerateWall(560, 1240, 200, 40, 1)
'30
Call GenerateWall(600, 1400, 40, 200, 2)
'31
Call GenerateWall(1440, 1400, 40, 200, 2)
'32
Call GenerateWall(1640, 1280, 160, 40, 2)
'33
Call GenerateWall(1840, 1440, 160, 40, 1)
'34
Call GenerateWall(2200, 1280, 160, 40, 1)
'35
Call GenerateWall(2400, 1280, 520, 40, 1)
'36
Call GenerateWall(400, 1720, 40, 160, 1)
'37
Call GenerateWall(560, 1600, 440, 40, 1)
'38
Call GenerateWall(600, 1600, 40, 200, 2)
'39
Call GenerateWall(760, 1640, 240, 40, 2)
'40
Call GenerateWall(800, 1840, 40, 640, 2)
'41
Call GenerateWall(1440, 1640, 240, 40, 2)
'42
Call GenerateWall(1440, 1600, 40, 200, 2)
'43
Call GenerateWall(1640, 1600, 160, 40, 2)
'44
Call GenerateWall(1640, 1760, 40, 600, 1)
'45
Call GenerateWall(2200, 1600, 160, 40, 1)
'46
Call GenerateWall(200, 2040, 720, 40, 2)
'47
Call GenerateWall(400, 2040, 520, 40, 2)
'48
Call GenerateWall(920, 2200, 40, 120, 1)
'49
Call GenerateWall(1040, 2200, 40, 160, 2)
'50
Call GenerateWall(1200, 2200, 40, 120, 1)
'51
Call GenerateWall(2200, 1800, 760, 40, 2)
'52
Call GenerateWall(2400, 1800, 960, 40, 2)
'53
Call GenerateWall(80, 2000, 40, 160, 1)
'54
Call GenerateWall(400, 2000, 40, 160, 1)
'55
Call GenerateWall(920, 2240, 320, 40, 1)
'56
Call GenerateWall(1280, 2240, 320, 40, 1)
'57
Call GenerateWall(400, 2560, 40, 520, 2)
'58
Call GenerateWall(920, 2560, 40, 120, 1)
'59
Call GenerateWall(1200, 2560, 40, 120, 1)
'60
Call GenerateWall(1320, 2560, 40, 920, 2)
'61
Call GenerateWall(200, 2760, 40, 720, 2)
'62
Call GenerateWall(920, 2760, 40, 120, 1)
'63
Call GenerateWall(1200, 2760, 40, 120, 1)
'64
Call GenerateWall(1320, 2760, 40, 1120, 2)
'65
Call GenerateWall(920, 2800, 200, 40, 1)
'66
Call GenerateWall(1280, 2800, 200, 40, 1)
'67
Call GenerateWall(960, 2960, 40, 320, 1)
WeaponCount = 3
ReDim Weapon(1 To WeaponCount)
With Weapon(1)
    .XCoord = 800
    .YCoord = 720
    .Ammo = 360
    .ClipAmmo = 120
    .Deleted = False
    .WeaponType = 3
End With
With Weapon(2)
    .XCoord = 1120
    .YCoord = 1540
    .Ammo = 9
    .ClipAmmo = 1
    .Deleted = False
    .WeaponType = 4
End With
With Weapon(3)
    .XCoord = 1120
    .YCoord = 640
    .Ammo = 9
    .ClipAmmo = 1
    .Deleted = False
    .WeaponType = 5
End With

EnemyCount = 3
ReDim Enemy(1 To EnemyCount)
With Enemy(1)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(1).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(1).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 150
    .MoveSpeed = 3
    .TargetX = 240
    .TargetY = 1600
    .XCoord = 240
    .YCoord = 1600
End With
With Enemy(2)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(2).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(2).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 150
    .MoveSpeed = 3
    .TargetX = 320
    .TargetY = 1800
    .XCoord = 320
    .YCoord = 1800
End With
With Enemy(3)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(3).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(3).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 150
    .MoveSpeed = 3
    .TargetX = 160
    .TargetY = 1920
    .XCoord = 160
    .YCoord = 1920
End With
Holocount = 1
ReDim Holoscreen(1 To Holocount)
With Holoscreen(1)
    .Goal = True
    .Text = ""
    .XCoord = 1120
    .YCoord = 2400
    .GraphicsDC = 1
End With
TriggerCount = 0
End Sub

Public Sub LoadLevel2()
With Player
    .Ammo(1) = 12
    .Ammo(2) = 120
    .ClipAmmo(1) = 12
    .ClipAmmo(2) = 0
    .CoolDownTime = 0
    .Dead = False
    .Deaths = 0
    .Firing = False
    .HP = PlayerMaxHealth * PlayerBonus
    .MoveBack = False
    .MoveForward = False
    .PickUpItem = False
    .Reloading = False
    .ReloadTime = 0
    .RespawnTime = 0
    .Score = 0
    .Shield = PlayerMaxShield * PlayerBonus
    .StrafeLeft = False
    .StrafeRight = False
    .Switching = False
    .TargetX = 0
    .TargetY = 0
    .Theta = PI
    .TurnLeft = False
    .TurnRight = False
    .Weapon(1) = 1 'battle rifle
    .Weapon(2) = 3 'laser pistol
    .XCoord = 900
    .YCoord = 700
End With
Dim i
'=================
'load map graphics
'=================
BackgroundDC = GenerateDC(App.Path & "\Graphics\Background4.bmp")
BackBuffDC = GenerateDC(App.Path & "\Graphics\Background4.bmp")
For i = 1 To 2
WallDC(i) = GenerateDC(App.Path & "\Graphics\Wall" & i & "_4.bmp")
Next i
'1
Call GenerateWall(0, 0, 600, 3200, 1)
'2
Call GenerateWall(400, 600, 200, 100, 0)
'3
Call GenerateWall(700, 600, 900, 100, 2)
'4
Call GenerateWall(800, 800, 100, 1200, 2)
'5
Call GenerateWall(2000, 800, 300, 200, 1)
'6
Call GenerateWall(2600, 600, 2300, 200, 1)
'7
Call GenerateWall(0, 800, 2500, 500, 1)
'8
Call GenerateWall(1100, 1100, 400, 400, 1)
'9
Call GenerateWall(2000, 1300, 300, 200, 1)
'10
Call GenerateWall(2400, 1500, 100, 200, 1)
'11
Call GenerateWall(2000, 1600, 200, 50, 2)
'12
Call GenerateWall(2000, 1800, 100, 600, 2)
'13
Call GenerateWall(700, 1500, 200, 300, 1)
'14
Call GenerateWall(1200, 1500, 1300, 200, 1)
'15
Call GenerateWall(1700, 900, 800, 100, 2)
'16
Call GenerateWall(1600, 1700, 500, 200, 1)
'17
Call GenerateWall(1800, 2100, 100, 100, 2)
'18
Call GenerateWall(1900, 2100, 400, 400, 1)
'19
Call GenerateWall(800, 2100, 200, 100, 1)
'20
Call GenerateWall(500, 2800, 500, 2600, 1)
WeaponCount = 1
ReDim Weapon(1 To WeaponCount)
With Weapon(1)
    .XCoord = 2100
    .YCoord = 1200
    .Ammo = 180
    .ClipAmmo = 90
    .Deleted = False
    .WeaponType = 5
End With

Holocount = 2
ReDim Holoscreen(1 To Holocount)
With Holoscreen(1)
    .Goal = False
    .GraphicsDC = 2
    .Text = "Enemy droids have invaded the base.  You must destroy them to escape.  Watch out for snipers.  Remember to pick up new weapons and ammo.  The weapons you finish this level with will carry over to the next level, so pick good ones."
    .XCoord = 900
    .YCoord = 700
End With
With Holoscreen(2)
    .Goal = True
    .GraphicsDC = 3
    .Text = "Level complete."
    .XCoord = 600
    .YCoord = 700
End With
    
EnemyCount = 11
ReDim Enemy(1 To EnemyCount)
With Enemy(1)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(1).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(1).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 3
    .TargetX = 2500
    .TargetY = 1000
    .XCoord = 2500
    .YCoord = 1000
End With
With Enemy(2)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(2).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(2).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 3
    .TargetX = 2300
    .TargetY = 1400
    .XCoord = 2300
    .YCoord = 1400
End With
With Enemy(3)
    .Weapon = 1
    .Ammo = WeaponData(Enemy(3).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(3).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 3
    .HP = 75 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 1
    .TargetX = 2100
    .TargetY = 1700
    .XCoord = 2100
    .YCoord = 1700
End With
With Enemy(4)
    .Weapon = 1
    .Ammo = WeaponData(Enemy(4).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(4).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 3
    .HP = 75 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 1
    .TargetX = 2500
    .TargetY = 1700
    .XCoord = 2500
    .YCoord = 1700
End With
With Enemy(5)
    .Weapon = 1
    .Ammo = WeaponData(Enemy(5).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(5).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 3
    .HP = 90 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 3
    .TargetX = 1500
    .TargetY = 1600
    .XCoord = 1500
    .YCoord = 1600
End With
With Enemy(6)
    .Weapon = 4
    .Ammo = WeaponData(Enemy(6).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(6).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 4
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 3
    .TargetX = 1850
    .TargetY = 2250
    .XCoord = 1850
    .YCoord = 2250
End With
With Enemy(7)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(7).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(7).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2250
    .TargetY = 2550
    .XCoord = 2250
    .YCoord = 2550
End With
With Enemy(8)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(8).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(8).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2000
    .TargetY = 2650
    .XCoord = 2000
    .YCoord = 2650
End With
With Enemy(9)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(9).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(9).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 750
    .TargetY = 1750
    .XCoord = 750
    .YCoord = 1750
End With
With Enemy(10)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(10).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(10).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 3
    .TargetX = 900
    .TargetY = 2700
    .XCoord = 900
    .YCoord = 2700
End With
With Enemy(11)
    .Weapon = 5
    .Ammo = WeaponData(Enemy(11).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(11).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 5
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 800
    .TargetY = 2400
    .XCoord = 800
    .YCoord = 2400
End With
TriggerCount = 0
End Sub
Public Sub LoadLevel3()
'all weapons carry over from last level
With Player
    .CoolDownTime = 0
    .Dead = False
    .Deaths = 0
    .Firing = False
    .HP = PlayerMaxHealth * PlayerBonus
    .MoveBack = False
    .MoveForward = False
    .PickUpItem = False
    .Reloading = False
    .ReloadTime = 0
    .RespawnTime = 0
    .Score = 0
    .Shield = PlayerMaxShield * PlayerBonus
    .StrafeLeft = False
    .StrafeRight = False
    .Switching = False
    .TargetX = 0
    .TargetY = 0
    .Theta = PI
    .TurnLeft = False
    .TurnRight = False
    .XCoord = 3000
    .YCoord = 3900
End With
Dim i
'=================
'load map graphics
'=================
BackgroundDC = GenerateDC(App.Path & "\Graphics\Background3.bmp")
BackBuffDC = GenerateDC(App.Path & "\Graphics\Background3.bmp")
For i = 1 To 2
WallDC(i) = GenerateDC(App.Path & "\Graphics\Wall" & i & "_3.bmp")
Next i
'1
Call GenerateWall(0, 0, 500, 3800, 1)
'2
Call GenerateWall(1900, 500, 400, 400, 1)
'3
Call GenerateWall(0, 500, 3500, 100, 2)
'4
Call GenerateWall(500, 1200, 400, 400, 1)
'5
Call GenerateWall(1300, 1000, 400, 400, 1)
'6
Call GenerateWall(2100, 1300, 200, 200, 1)
'7
Call GenerateWall(2600, 700, 1200, 400, 1)
'8
Call GenerateWall(3200, 500, 1800, 600, 1)
'9
Call GenerateWall(600, 2200, 200, 200, 1)
'10
Call GenerateWall(1300, 2200, 300, 300, 1)
'11
Call GenerateWall(2600, 2000, 200, 200, 1)
'12
Call GenerateWall(500, 2900, 200, 500, 1)
'13
Call GenerateWall(1600, 2600, 300, 800, 1)
'14
Call GenerateWall(2400, 2700, 100, 200, 2)
'15
Call GenerateWall(2600, 2300, 1500, 1200, 1)
'16
Call GenerateWall(600, 3500, 200, 200, 1)
'17
Call GenerateWall(1300, 3400, 400, 200, 1)
'18
Call GenerateWall(1800, 3300, 200, 500, 1)
'19
Call GenerateWall(3150, 3800, 200, 50, 0)
'20
Call GenerateWall(0, 4000, 600, 3800, 1)
'21
Call GenerateWall(100, 1800, 100, 2500, 1)

Holocount = 2
ReDim Holoscreen(1 To Holocount)
With Holoscreen(1)
    .Goal = False
    .GraphicsDC = 2
    .Text = "Probe droids have infiltrated this storage facility and stolen vital weapons technology.  Dispatch them and recover whatever you can.  Your weapons will carry over to the next level."
    .XCoord = 3000
    .YCoord = 3900
End With
With Holoscreen(2)
    .Goal = True
    .GraphicsDC = 3
    .Text = "Level complete."
    .XCoord = 200
    .YCoord = 1100
End With
    
EnemyCount = 14
ReDim Enemy(1 To EnemyCount)
With Enemy(1)
    .Weapon = 1
    .Ammo = WeaponData(Enemy(1).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(1).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 3
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 3
    .TargetX = 2500
    .TargetY = 2600
    .XCoord = 2500
    .YCoord = 2600
End With
With Enemy(2)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(2).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(2).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 2
    .Stance = 4
    .TargetX = 1900
    .TargetY = 3200
    .XCoord = 1900
    .YCoord = 3200
End With
With Enemy(3)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(3).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(3).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2100
    .TargetY = 3000
    .XCoord = 2100
    .YCoord = 3000
End With
With Enemy(4)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(4).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(4).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1200
    .TargetY = 3500
    .XCoord = 1200
    .YCoord = 3500
End With
With Enemy(5)
    .Weapon = 4
    .Ammo = WeaponData(Enemy(5).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(5).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 4
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 550
    .TargetY = 2850
    .XCoord = 550
    .YCoord = 2850
End With
With Enemy(6)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(6).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(6).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 700
    .TargetY = 2100
    .XCoord = 700
    .YCoord = 2100
End With
With Enemy(7)
    .Weapon = 5
    .Ammo = WeaponData(Enemy(7).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(7).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 5
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1400
    .TargetY = 2000
    .XCoord = 1400
    .YCoord = 2000
End With
With Enemy(8)
    .Weapon = 4
    .Ammo = WeaponData(Enemy(8).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(8).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 4
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 3100
    .TargetY = 2100
    .XCoord = 3100
    .YCoord = 2100
End With
With Enemy(9)
    .Weapon = 1
    .Ammo = WeaponData(Enemy(9).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(9).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 3
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 3050
    .TargetY = 850
    .XCoord = 3050
    .YCoord = 850
End With
With Enemy(10)
    .Weapon = 1
    .Ammo = WeaponData(Enemy(10).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(10).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 3
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 3150
    .TargetY = 550
    .XCoord = 3150
    .YCoord = 550
End With
With Enemy(11)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(11).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(11).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1850
    .TargetY = 950
    .XCoord = 1850
    .YCoord = 950
End With
With Enemy(12)
    .Weapon = 5
    .Ammo = WeaponData(Enemy(12).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(12).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 5
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2050
    .TargetY = 1550
    .XCoord = 2050
    .YCoord = 1550
End With
With Enemy(13)
    .Weapon = 5
    .Ammo = WeaponData(Enemy(13).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(13).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 5
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1250
    .TargetY = 1150
    .XCoord = 1250
    .YCoord = 1150
End With
With Enemy(14)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(14).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(14).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 950
    .TargetY = 1550
    .XCoord = 950
    .YCoord = 1550
End With
TriggerCount = 0
End Sub
Public Sub LoadLevel4()
'all weapons carry over from last level
With Player
    .CoolDownTime = 0
    .Dead = False
    .Deaths = 0
    .Firing = False
    .HP = PlayerMaxHealth * PlayerBonus
    .MoveBack = False
    .MoveForward = False
    .PickUpItem = False
    .Reloading = False
    .ReloadTime = 0
    .RespawnTime = 0
    .Score = 0
    .Shield = PlayerMaxShield * PlayerBonus
    .StrafeLeft = False
    .StrafeRight = False
    .Switching = False
    .TargetX = 0
    .TargetY = 0
    .Theta = PI
    .TurnLeft = False
    .TurnRight = False
    .XCoord = 2600
    .YCoord = 3900
End With
Dim i
'=================
'load map graphics
'=================
BackgroundDC = GenerateDC(App.Path & "\Graphics\Background3.bmp")
BackBuffDC = GenerateDC(App.Path & "\Graphics\Background3.bmp")
For i = 1 To 2
WallDC(i) = GenerateDC(App.Path & "\Graphics\Wall" & i & "_3.bmp")
Next i
'1
Call GenerateWall(0, 0, 4000, 500, 1)
'2
Call GenerateWall(500, 0, 600, 900, 1)
'3
Call GenerateWall(1400, 450, 50, 1200, 2)
'4
Call GenerateWall(2600, 0, 2300, 1000, 1)
'5
Call GenerateWall(500, 800, 700, 500, 1)
'6
Call GenerateWall(1200, 1200, 300, 500, 1)
'7
Call GenerateWall(2300, 1500, 200, 300, 1)
'8
Call GenerateWall(500, 1700, 600, 500, 1)
'9
Call GenerateWall(1200, 1700, 600, 500, 1)
'10
Call GenerateWall(1600, 2500, 200, 200, 1)
'11
Call GenerateWall(1000, 2800, 200, 200, 1)
'12
Call GenerateWall(2400, 2800, 400, 400, 1)
'13
Call GenerateWall(3100, 2300, 1200, 500, 1)
'14
Call GenerateWall(1400, 3200, 200, 500, 1)
'15
Call GenerateWall(500, 3500, 1000, 2000, 1)
'16
Call GenerateWall(2500, 4000, 50, 200, 0)
'17
Call GenerateWall(2700, 3500, 1000, 900, 1)


Holocount = 2
ReDim Holoscreen(1 To Holocount)
With Holoscreen(1)
    .Goal = False
    .GraphicsDC = 2
    .Text = "The enemy has taken over this city, and are currently using it as their base of operations for an attack on one of our bases further south.  We're not exactly sure what they did with the technology they managed to steal from us, but you should be on your guard."
    .XCoord = 2600
    .YCoord = 3900
End With
With Holoscreen(2)
    .Goal = True
    .GraphicsDC = 3
    .Text = "Level complete."
    .XCoord = 600
    .YCoord = 700
End With
WeaponCount = 0
EnemyCount = 14
ReDim Enemy(1 To EnemyCount)
With Enemy(1)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(1).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(1).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2900
    .TargetY = 2700
    .XCoord = 2900
    .YCoord = 2700
End With
With Enemy(2)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(2).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(2).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2700
    .TargetY = 2500
    .XCoord = 2700
    .YCoord = 2500
End With
With Enemy(3)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(3).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(3).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1100
    .TargetY = 3300
    .XCoord = 1100
    .YCoord = 3300
End With
With Enemy(4)
    .Weapon = 7
    .Ammo = WeaponData(Enemy(4).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(4).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1100
    .TargetY = 2700
    .XCoord = 1100
    .YCoord = 2700
End With
With Enemy(5)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(5).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(5).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1100
    .TargetY = 1600
    .XCoord = 1100
    .YCoord = 1600
End With
With Enemy(6)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(6).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(6).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1350
    .TargetY = 1650
    .XCoord = 1350
    .YCoord = 1650
End With
With Enemy(7)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(7).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(7).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 850
    .TargetY = 1550
    .XCoord = 850
    .YCoord = 1550
End With
With Enemy(8)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(8).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(8).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 650
    .TargetY = 1600
    .XCoord = 650
    .YCoord = 1600
End With
With Enemy(9)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(9).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(9).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1150
    .TargetY = 1850
    .XCoord = 1150
    .YCoord = 1850
End With
With Enemy(10)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(10).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(10).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1050
    .TargetY = 1350
    .XCoord = 1050
    .YCoord = 1350
End With
With Enemy(11)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(11).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(11).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1100
    .TargetY = 1100
    .XCoord = 1100
    .YCoord = 1100
End With
With Enemy(12)
    .Weapon = 7
    .Ammo = WeaponData(Enemy(12).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(12).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2450
    .TargetY = 1450
    .XCoord = 2450
    .YCoord = 1450
End With
With Enemy(13)
    .Weapon = 2
    .Ammo = WeaponData(Enemy(13).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(13).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1800
    .TargetY = 800
    .XCoord = 1800
    .YCoord = 800
End With
With Enemy(14)
    .Weapon = 2
    .Ammo = WeaponData(Enemy(14).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(14).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2000
    .TargetY = 600
    .XCoord = 2000
    .YCoord = 600
End With
TriggerCount = 0
End Sub

Public Sub LoadLevel5()
'all weapons carry over from last level
With Player
    .CoolDownTime = 0
    .Dead = False
    .Deaths = 0
    .Firing = False
    .HP = PlayerMaxHealth * PlayerBonus
    .MoveBack = False
    .MoveForward = False
    .PickUpItem = False
    .Reloading = False
    .ReloadTime = 0
    .RespawnTime = 0
    .Score = 0
    .Shield = PlayerMaxShield * PlayerBonus
    .StrafeLeft = False
    .StrafeRight = False
    .Switching = False
    .TargetX = 0
    .TargetY = 0
    .Theta = PI
    .TurnLeft = False
    .TurnRight = False
    .XCoord = 2200
    .YCoord = 1900
End With
Dim i
'=================
'load map graphics
'=================
BackgroundDC = GenerateDC(App.Path & "\Graphics\Background3.bmp")
BackBuffDC = GenerateDC(App.Path & "\Graphics\Background3.bmp")
For i = 1 To 2
WallDC(i) = GenerateDC(App.Path & "\Graphics\Wall" & i & "_3.bmp")
Next i
'1
Call GenerateWall(0, -500, 500, 3700, 1)
'2
Call GenerateWall(1200, 0, 100, 1100, 2)
'3
Call GenerateWall(2600, 0, 300, 600, 1)
'4
Call GenerateWall(3200, 0, 3800, 500, 1)
'5
Call GenerateWall(1200, 300, 100, 1100, 2)
'6
Call GenerateWall(2700, 300, 200, 50, 0)
'7
Call GenerateWall(1200, 600, 100, 1100, 2)
'8
Call GenerateWall(2600, 500, 200, 600, 1)
'9
Call GenerateWall(1200, 700, 200, 2000, 1)
'10
Call GenerateWall(700, 1000, 200, 300, 1)
'11
Call GenerateWall(1200, 900, 1100, 100, 1)
'12
Call GenerateWall(1900, 1000, 100, 100, 1)
'13
Call GenerateWall(1900, 1200, 100, 100, 1)
'14
Call GenerateWall(1900, 1400, 400, 500, 1)
'15
Call GenerateWall(1600, 1750, 50, 300, 2)
'16
Call GenerateWall(2700, 1700, 200, 200, 1)
'17
Call GenerateWall(2300, 1800, 200, 100, 1)
'18
Call GenerateWall(1000, 2000, 400, 1400, 1)
'19
Call GenerateWall(800, 2000, 100, 200, 1)
'20
Call GenerateWall(500, 2000, 100, 300, 1)
'21
Call GenerateWall(0, 0, 3800, 500, 1)
'22
Call GenerateWall(700, 2300, 100, 300, 1)
'23
Call GenerateWall(1600, 2400, 200, 200, 1)
'24
Call GenerateWall(2700, 2600, 200, 200, 1)
'25
Call GenerateWall(1300, 2800, 400, 200, 1)
'26
Call GenerateWall(1800, 3000, 300, 600, 1)
'27
Call GenerateWall(500, 3300, 500, 2700, 1)



Holocount = 2
ReDim Holoscreen(1 To Holocount)
With Holoscreen(1)
    .Goal = False
    .GraphicsDC = 2
    .Text = "This subway terminal should connect to a tunnel which leads to the next city, controlled by our forces.  Unfortunately, it doesn't appear that the enemy patrols are any less vigilant underground."
    .XCoord = 2600
    .YCoord = 3900
End With
With Holoscreen(2)
    .Goal = True
    .GraphicsDC = 3
    .Text = "Level complete."
    .XCoord = 2600
    .YCoord = 400
End With
    
EnemyCount = 18
ReDim Enemy(1 To EnemyCount)
With Enemy(1)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(1).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(1).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1600
    .TargetY = 1100
    .XCoord = 1600
    .YCoord = 1100
End With
With Enemy(2)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(2).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(2).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1500
    .TargetY = 1300
    .XCoord = 1500
    .YCoord = 1300
End With
With Enemy(3)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(3).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(3).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1800
    .TargetY = 1500
    .XCoord = 1800
    .YCoord = 1500
End With
With Enemy(4)
    .Weapon = 7
    .Ammo = WeaponData(Enemy(4).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(4).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 4
    .Stance = 4
    .TargetX = 2800
    .TargetY = 1300
    .XCoord = 2800
    .YCoord = 1300
End With
With Enemy(5)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(5).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(5).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2600
    .TargetY = 1900
    .XCoord = 2600
    .YCoord = 1900
End With
With Enemy(6)
    .Weapon = 5
    .Ammo = WeaponData(Enemy(6).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(6).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 5
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 3000
    .TargetY = 1900
    .XCoord = 3000
    .YCoord = 1900
End With
With Enemy(7)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(7).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(7).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2400
    .TargetY = 2800
    .XCoord = 2400
    .YCoord = 2800
End With
With Enemy(8)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(8).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(8).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 600
    .TargetY = 1600
    .XCoord = 600
    .YCoord = 1600
End With
With Enemy(9)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(9).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(9).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1500
    .TargetY = 2600
    .XCoord = 1500
    .YCoord = 2600
End With
With Enemy(10)
    .Weapon = 1
    .Ammo = WeaponData(Enemy(10).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(10).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 3
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1100
    .TargetY = 2700
    .XCoord = 1100
    .YCoord = 2700
End With
With Enemy(11)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(11).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(11).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1200
    .TargetY = 3200
    .XCoord = 1200
    .YCoord = 3200
End With
With Enemy(12)
    .Weapon = 1
    .Ammo = WeaponData(Enemy(12).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(12).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 3
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 600
    .TargetY = 2200
    .XCoord = 600
    .YCoord = 2200
End With
With Enemy(13)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(13).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(13).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1100
    .TargetY = 1100
    .XCoord = 1100
    .YCoord = 1100
End With
With Enemy(14)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(14).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(14).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 600
    .TargetY = 1100
    .XCoord = 600
    .YCoord = 1100
End With
With Enemy(15)
    .Weapon = 1
    .Ammo = WeaponData(Enemy(15).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(15).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 3
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 3
    .TargetX = 600
    .TargetY = 200
    .XCoord = 600
    .YCoord = 200
End With
With Enemy(16)
    .Weapon = 5
    .Ammo = WeaponData(Enemy(16).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(16).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 5
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 3
    .TargetX = 1100
    .TargetY = 200
    .XCoord = 1100
    .YCoord = 200
End With
With Enemy(17)
    .Weapon = 7
    .Ammo = WeaponData(Enemy(17).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(17).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 5
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2500
    .TargetY = 200
    .XCoord = 2500
    .YCoord = 200
End With
With Enemy(18)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(18).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(18).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2500
    .TargetY = 600
    .XCoord = 2500
    .YCoord = 600
End With

TriggerCount = 1
ReDim Trigger(1 To TriggerCount)
With Trigger(1)
    .Link = 19
    .XCoord = 800
    .YCoord = 2100
    .Height = 200
    .Width = 200
End With
End Sub
Public Sub LoadLevel6()
'all weapons carry over from last level
With Player
    .CoolDownTime = 0
    .Dead = False
    .Deaths = 0
    .Firing = False
    .HP = PlayerMaxHealth * PlayerBonus
    .MoveBack = False
    .MoveForward = False
    .PickUpItem = False
    .Reloading = False
    .ReloadTime = 0
    .RespawnTime = 0
    .Score = 0
    .Shield = PlayerMaxShield * PlayerBonus
    .StrafeLeft = False
    .StrafeRight = False
    .Switching = False
    .TargetX = 0
    .TargetY = 0
    .Theta = PI
    .TurnLeft = False
    .TurnRight = False
    .XCoord = 2300
    .YCoord = 800
End With
Dim i
'=================
'load map graphics
'=================
BackgroundDC = GenerateDC(App.Path & "\Graphics\Background3.bmp")
BackBuffDC = GenerateDC(App.Path & "\Graphics\Background3.bmp")
For i = 1 To 2
WallDC(i) = GenerateDC(App.Path & "\Graphics\Wall" & i & "_3.bmp")
Next i
'1
Call GenerateWall(0, 0, 4500, 500, 1)
'2
Call GenerateWall(500, 0, 1300, 700, 1)
'3
Call GenerateWall(1200, 700, 100, 500, 1)
'4
Call GenerateWall(1700, 0, 900, 300, 1)
'5
Call GenerateWall(2200, 0, 500, 800, 1)
'6
Call GenerateWall(2200, 500, 200, 200, 1)
'7
Call GenerateWall(2700, 500, 800, 300, 1)
'8
Call GenerateWall(3000, 0, 4500, 500, 1)
'9
Call GenerateWall(1700, 900, 400, 700, 1)
'10
Call GenerateWall(2400, 1100, 200, 100, 1)
'11
Call GenerateWall(2600, 1100, 200, 100, 1)
'12
Call GenerateWall(1900, 1600, 200, 200, 1)
'13
Call GenerateWall(2200, 1800, 1000, 200, 1)
'14
Call GenerateWall(2400, 1800, 400, 200, 1)
'15
Call GenerateWall(2800, 1800, 200, 200, 1)
'16
Call GenerateWall(1700, 2800, 400, 700, 1)
'17
Call GenerateWall(2600, 3000, 200, 400, 1)
'18
Call GenerateWall(1700, 3200, 400, 200, 1)
'19
Call GenerateWall(1200, 3550, 50, 500, 2)
'20
Call GenerateWall(1200, 3800, 50, 500, 2)
'21
Call GenerateWall(1700, 3800, 500, 1300, 1)
'22
Call GenerateWall(0, 3800, 500, 1200, 1)
'23
Call GenerateWall(1000, 2500, 1100, 200, 1)
'24
Call GenerateWall(700, 2800, 200, 300, 1)
'25
Call GenerateWall(1000, 1900, 200, 200, 1)
'26
Call GenerateWall(1300, 2100, 1100, 300, 1)
'27
Call GenerateWall(1000, 1300, 300, 200, 1)
'28
Call GenerateWall(1200, 0, 500, 500, 0)
'29
Call GenerateWall(2000, 0, 900, 200, 0)
'30
Call GenerateWall(1700, 1300, 1500, 50, 2)


Holocount = 2
ReDim Holoscreen(1 To Holocount)
With Holoscreen(1)
    .Goal = False
    .GraphicsDC = 2
    .Text = "Good to see you managed to fight your way into the lower levels.  You will be resupplied once you reach the base, so don't be afraid to set your trusty rusty laser pistol on 'fully automatic'."
    .XCoord = 2300
    .YCoord = 800
End With
With Holoscreen(2)
    .Goal = True
    .GraphicsDC = 3
    .Text = "Level complete."
    .XCoord = 1450
    .YCoord = 600
End With
    
EnemyCount = 19
ReDim Enemy(1 To EnemyCount)
With Enemy(1)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(1).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(1).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1850
    .TargetY = 1550
    .XCoord = 1850
    .YCoord = 1550
End With
With Enemy(2)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(2).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(2).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2900
    .TargetY = 1400
    .XCoord = 2900
    .YCoord = 1400
End With
With Enemy(3)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(3).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(3).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2100
    .TargetY = 2700
    .XCoord = 2100
    .YCoord = 2700
End With
With Enemy(4)
    .Weapon = 7
    .Ammo = WeaponData(Enemy(4).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(4).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 4
    .Stance = 4
    .TargetX = 1900
    .TargetY = 2500
    .XCoord = 1900
    .YCoord = 2500
End With
With Enemy(5)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(5).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(5).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2500
    .TargetY = 2300
    .XCoord = 2500
    .YCoord = 2300
End With
With Enemy(6)
    .Weapon = 5
    .Ammo = WeaponData(Enemy(6).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(6).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 5
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2500
    .TargetY = 2900
    .XCoord = 2500
    .YCoord = 2900
End With
With Enemy(7)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(7).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(7).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2000
    .TargetY = 3300
    .XCoord = 2000
    .YCoord = 3300
End With
With Enemy(8)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(8).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(8).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2900
    .TargetY = 3300
    .XCoord = 2900
    .YCoord = 3300
End With
With Enemy(9)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(9).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(9).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1800
    .TargetY = 3700
    .XCoord = 1800
    .YCoord = 3700
End With
With Enemy(10)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(10).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(10).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1300
    .TargetY = 3400
    .XCoord = 1300
    .YCoord = 3400
End With
With Enemy(11)
    .Weapon = 5
    .Ammo = WeaponData(Enemy(11).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(11).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 5
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1550
    .TargetY = 3350
    .XCoord = 1550
    .YCoord = 3350
End With
With Enemy(12)
    .Weapon = 1
    .Ammo = WeaponData(Enemy(12).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(12).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 3
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 3
    .TargetX = 900
    .TargetY = 3100
    .XCoord = 900
    .YCoord = 3100
End With
With Enemy(13)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(13).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(13).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 950
    .TargetY = 2750
    .XCoord = 950
    .YCoord = 2750
End With
With Enemy(14)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(14).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(14).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 950
    .TargetY = 2050
    .XCoord = 950
    .YCoord = 2050
End With
With Enemy(15)
    .Weapon = 1
    .Ammo = WeaponData(Enemy(15).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(15).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 3
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 3
    .TargetX = 950
    .TargetY = 1350
    .XCoord = 950
    .YCoord = 1350
End With
With Enemy(16)
    .Weapon = 5
    .Ammo = WeaponData(Enemy(16).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(16).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 5
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 3
    .TargetX = 1300
    .TargetY = 600
    .XCoord = 1300
    .YCoord = 600
End With
With Enemy(17)
    .Weapon = 7
    .Ammo = WeaponData(Enemy(17).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(17).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 5
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1400
    .TargetY = 600
    .XCoord = 1400
    .YCoord = 600
End With
With Enemy(18)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(18).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(18).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1500
    .TargetY = 600
    .XCoord = 1500
    .YCoord = 600
End With
With Enemy(19)
    .Weapon = 1
    .Ammo = WeaponData(Enemy(19).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(19).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 3
    .HP = 100 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1600
    .TargetY = 600
    .XCoord = 1600
    .YCoord = 600
End With

TriggerCount = 1
ReDim Trigger(1 To TriggerCount)
With Trigger(1)
    .Link = 3
    .XCoord = 1200
    .YCoord = 800
    .Height = 2000
    .Width = 500
End With
End Sub
Public Sub LoadLevel7()
'all weapons carry over from last level
With Player
    .CoolDownTime = 0
    .Dead = False
    .Deaths = 0
    .Firing = False
    .HP = PlayerMaxHealth * PlayerBonus
    .MoveBack = False
    .MoveForward = False
    .PickUpItem = False
    .Reloading = False
    .ReloadTime = 0
    .RespawnTime = 0
    .Score = 0
    .Shield = PlayerMaxShield * PlayerBonus
    .StrafeLeft = False
    .StrafeRight = False
    .Switching = False
    .TargetX = 0
    .TargetY = 0
    .Theta = PI
    .TurnLeft = False
    .TurnRight = False
    .XCoord = 1950
    .YCoord = 2500
End With
Dim i
'=================
'load map graphics
'=================
BackgroundDC = GenerateDC(App.Path & "\Graphics\Background3.bmp")
BackBuffDC = GenerateDC(App.Path & "\Graphics\Background3.bmp")
For i = 1 To 2
WallDC(i) = GenerateDC(App.Path & "\Graphics\Wall" & i & "_3.bmp")
Next i
'1
Call GenerateWall(0, 0, 500, 3200, 1)
'2
Call GenerateWall(0, 500, 4200, 500, 1)
'3
Call GenerateWall(3000, 500, 4200, 500, 1)
'4
Call GenerateWall(500, 3700, 500, 2500, 1)
'5
Call GenerateWall(1300, 500, 400, 700, 1)
'6
Call GenerateWall(2600, 500, 800, 400, 1)
'7
Call GenerateWall(500, 1300, 600, 500, 1)
'8
Call GenerateWall(1500, 1300, 600, 800, 1)
'9
Call GenerateWall(1200, 2300, 600, 400, 1)
'10
Call GenerateWall(1200, 3100, 600, 600, 1)
'11
Call GenerateWall(2300, 2400, 1300, 700, 1)

Holocount = 2
ReDim Holoscreen(1 To Holocount)
With Holoscreen(1)
    .Goal = False
    .GraphicsDC = 2
    .Text = "Welcome to the advanced training level.  Here you will test and study the capabilities of your weapons.  You should examine whether they are most effective at long or short range, how much shots spread, how much damage they do, and what their maximum range is.  If you run out of drones to destroy, click the exit button to replay the level."
    .XCoord = 1950
    .YCoord = 2500
End With
With Holoscreen(2)
    .Goal = True
    .GraphicsDC = 3
    .Text = "Level complete."
    .XCoord = 2900
    .YCoord = 1850
End With
WeaponCount = 9
ReDim Weapon(1 To WeaponCount)
With Weapon(1)
    .WeaponType = 1
    .Ammo = WeaponData(Weapon(1).WeaponType).MaxAmmo
    .ClipAmmo = WeaponData(Weapon(1).WeaponType).ClipSize
    .Deleted = False
    .XCoord = 1800
    .YCoord = 2300
End With
With Weapon(2)
    .WeaponType = 2
    .Ammo = WeaponData(Weapon(2).WeaponType).MaxAmmo
    .ClipAmmo = WeaponData(Weapon(2).WeaponType).ClipSize
    .Deleted = False
    .XCoord = 2000
    .YCoord = 2400
End With
With Weapon(3)
    .WeaponType = 2
    .Ammo = WeaponData(Weapon(3).WeaponType).MaxAmmo
    .ClipAmmo = WeaponData(Weapon(3).WeaponType).ClipSize
    .Deleted = False
    .XCoord = 2100
    .YCoord = 2700
End With
With Weapon(4)
    .WeaponType = 3
    .Ammo = WeaponData(Weapon(4).WeaponType).MaxAmmo
    .ClipAmmo = WeaponData(Weapon(4).WeaponType).ClipSize
    .Deleted = False
    .XCoord = 1800
    .YCoord = 2800
End With
With Weapon(5)
    .WeaponType = 4
    .Ammo = WeaponData(Weapon(5).WeaponType).MaxAmmo
    .ClipAmmo = WeaponData(Weapon(5).WeaponType).ClipSize
    .Deleted = False
    .XCoord = 2600
    .YCoord = 2100
End With
With Weapon(6)
    .WeaponType = 5
    .Ammo = WeaponData(Weapon(6).WeaponType).MaxAmmo
    .ClipAmmo = WeaponData(Weapon(6).WeaponType).ClipSize
    .Deleted = False
    .XCoord = 1400
    .YCoord = 2100
End With
With Weapon(7)
    .WeaponType = 6
    .Ammo = WeaponData(Weapon(7).WeaponType).MaxAmmo
    .ClipAmmo = WeaponData(Weapon(7).WeaponType).ClipSize
    .Deleted = False
    .XCoord = 1400
    .YCoord = 3000
End With
With Weapon(8)
    .WeaponType = 7
    .Ammo = WeaponData(Weapon(8).WeaponType).MaxAmmo
    .ClipAmmo = WeaponData(Weapon(8).WeaponType).ClipSize
    .Deleted = False
    .XCoord = 900
    .YCoord = 2700
End With
With Weapon(9)
    .WeaponType = 2
    .Ammo = WeaponData(Weapon(9).WeaponType).MaxAmmo
    .ClipAmmo = WeaponData(Weapon(9).WeaponType).ClipSize
    .Deleted = False
    .XCoord = 1100
    .YCoord = 1400
End With
EnemyCount = 15
ReDim Enemy(1 To EnemyCount)
With Enemy(1)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 1200
    .TargetY = 700
    .XCoord = 1200
    .YCoord = 700
End With
With Enemy(2)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 900
    .TargetY = 700
    .XCoord = 900
    .YCoord = 700
End With
With Enemy(3)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 700
    .TargetY = 900
    .XCoord = 700
    .YCoord = 900
End With
With Enemy(4)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 700
    .TargetY = 1200
    .XCoord = 700
    .YCoord = 1200
End With
With Enemy(5)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 2900
    .TargetY = 1400
    .XCoord = 2900
    .YCoord = 1400
End With
With Enemy(6)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 2900
    .TargetY = 2300
    .XCoord = 2900
    .YCoord = 2300
End With
With Enemy(7)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 2300
    .TargetY = 700
    .XCoord = 2300
    .YCoord = 700
End With
With Enemy(8)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 1900
    .TargetY = 3600
    .XCoord = 1900
    .YCoord = 3600
End With
With Enemy(9)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 2200
    .TargetY = 3600
    .XCoord = 2200
    .YCoord = 3600
End With
With Enemy(10)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 2050
    .TargetY = 3350
    .XCoord = 2050
    .YCoord = 3350
End With
With Enemy(11)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 600
    .TargetY = 2000
    .XCoord = 600
    .YCoord = 2000
End With
With Enemy(12)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 900
    .TargetY = 2000
    .XCoord = 900
    .YCoord = 2000
End With
With Enemy(13)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 600
    .TargetY = 3600
    .XCoord = 600
    .YCoord = 3600
End With
With Enemy(14)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 900
    .TargetY = 3600
    .XCoord = 900
    .YCoord = 3600
End With
With Enemy(15)
    .Weapon = 2
    .Ammo = 0
    .ClipAmmo = 0
    .Dead = False
    .GraphicsDC = 1
    .HP = 100 * EnemyBonus
    .MoveSpeed = 0
    .Stance = 1
    .TargetX = 2400
    .TargetY = 1800
    .XCoord = 2400
    .YCoord = 1800
End With
TriggerCount = 0
End Sub


Public Sub LoadLevel8()
With Player
    .CoolDownTime = 0
    .Dead = False
    .Deaths = 0
    .Firing = False
    .HP = PlayerMaxHealth * PlayerBonus
    .MoveBack = False
    .MoveForward = False
    .PickUpItem = False
    .Reloading = False
    .ReloadTime = 0
    .RespawnTime = 0
    .Score = 0
    .Shield = PlayerMaxShield * PlayerBonus
    .StrafeLeft = False
    .StrafeRight = False
    .Switching = False
    .TargetX = 0
    .TargetY = 0
    .Theta = PI
    .TurnLeft = False
    .TurnRight = False
    .XCoord = 2500
    .YCoord = 3600
    .Weapon(1) = 1
    .Weapon(2) = 7
    .Ammo(1) = 72
    .ClipAmmo(1) = 12
    .Ammo(2) = 240
    .ClipAmmo(2) = 60
End With
Dim i
'=================
'load map graphics
'=================
BackgroundDC = GenerateDC(App.Path & "\Graphics\Background4.bmp")
BackBuffDC = GenerateDC(App.Path & "\Graphics\Background4.bmp")
For i = 1 To 2
WallDC(i) = GenerateDC(App.Path & "\Graphics\Wall" & i & "_3.bmp")
Next i
'1
Call GenerateWall(0, 0, 500, 3500, 1)
'2
Call GenerateWall(0, 500, 4200, 500, 1)
'3
Call GenerateWall(3000, 500, 4200, 500, 1)
'4
Call GenerateWall(500, 3700, 500, 2500, 1)
'5
Call GenerateWall(500, 1700, 600, 200, 1)
'6
Call GenerateWall(1200, 800, 500, 500, 1)
'7
Call GenerateWall(1000, 1300, 800, 900, 1)
'8
Call GenerateWall(1200, 2100, 700, 1000, 1)
'9
Call GenerateWall(1000, 2800, 900, 900, 1)
'10
Call GenerateWall(2400, 800, 300, 300, 1)
'11
Call GenerateWall(2800, 1900, 500, 200, 1)
'12
Call GenerateWall(2300, 3200, 300, 400, 1)

Holocount = 2
ReDim Holoscreen(1 To Holocount)
With Holoscreen(1)
    .Goal = False
    .GraphicsDC = 2
    .Text = "The enemy has invaded the south quarter.  Be careful as it appears they have some sort of new short-range weapon."
    .XCoord = 2500
    .YCoord = 3600
End With
With Holoscreen(2)
    .Goal = True
    .GraphicsDC = 3
    .Text = "Level complete."
    .XCoord = 750
    .YCoord = 3600
End With
    
EnemyCount = 12
ReDim Enemy(1 To EnemyCount)
With Enemy(1)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(1).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(1).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2900
    .TargetY = 3400
    .XCoord = 2900
    .YCoord = 3400
End With
With Enemy(2)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(2).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(2).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2500
    .TargetY = 3100
    .XCoord = 2500
    .YCoord = 3100
End With
With Enemy(3)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(3).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(3).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2200
    .TargetY = 2900
    .XCoord = 2200
    .YCoord = 2900
End With
With Enemy(4)
    .Weapon = 8
    .Ammo = WeaponData(Enemy(4).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(4).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 4
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2300
    .TargetY = 2000
    .XCoord = 2300
    .YCoord = 2000
End With
With Enemy(5)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(5).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(5).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2900
    .TargetY = 1800
    .XCoord = 2900
    .YCoord = 1800
End With
With Enemy(6)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(6).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(6).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1800
    .TargetY = 1100
    .XCoord = 1800
    .YCoord = 1100
End With
With Enemy(7)
    .Weapon = 8
    .Ammo = WeaponData(Enemy(7).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(7).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 4
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2650
    .TargetY = 650
    .XCoord = 2650
    .YCoord = 650
End With
With Enemy(8)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(8).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(8).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 150 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1800
    .TargetY = 1100
    .XCoord = 1800
    .YCoord = 1100
End With
With Enemy(9)
    .Weapon = 3
    .Ammo = WeaponData(Enemy(9).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(9).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 150 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 600
    .TargetY = 1600
    .XCoord = 600
    .YCoord = 1600
End With
With Enemy(10)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(10).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(10).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1000
    .TargetY = 2700
    .XCoord = 1000
    .YCoord = 2700
End With
With Enemy(11)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(11).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(11).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 600
    .TargetY = 2400
    .XCoord = 600
    .YCoord = 2400
End With
With Enemy(12)
    .Weapon = 8
    .Ammo = WeaponData(Enemy(12).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(12).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 4
    .HP = 150 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 900
    .TargetY = 3600
    .XCoord = 900
    .YCoord = 3600
End With
WeaponCount = 0
TriggerCount = 0
End Sub

Public Sub LoadLevel9()
'all weapons carry over from last level
With Player
    .CoolDownTime = 0
    .Dead = False
    .Deaths = 0
    .Firing = False
    .HP = PlayerMaxHealth * PlayerBonus
    .MoveBack = False
    .MoveForward = False
    .PickUpItem = False
    .Reloading = False
    .ReloadTime = 0
    .RespawnTime = 0
    .Score = 0
    .Shield = PlayerMaxShield * PlayerBonus
    .StrafeLeft = False
    .StrafeRight = False
    .Switching = False
    .TargetX = 0
    .TargetY = 0
    .Theta = PI
    .TurnLeft = False
    .TurnRight = False
    .XCoord = 850
    .YCoord = 3600
End With
Dim i
'=================
'load map graphics
'=================
BackgroundDC = GenerateDC(App.Path & "\Graphics\Background3.bmp")
BackBuffDC = GenerateDC(App.Path & "\Graphics\Background3.bmp")
For i = 1 To 2
WallDC(i) = GenerateDC(App.Path & "\Graphics\Wall" & i & "_3.bmp")
Next i
'1
Call GenerateWall(0, 0, 500, 3200, 1)
'2
Call GenerateWall(0, 500, 4200, 500, 1)
'3
Call GenerateWall(3000, 500, 4200, 500, 1)
'4
Call GenerateWall(500, 3700, 500, 2500, 1)
'5
Call GenerateWall(1300, 500, 400, 700, 1)
'6
Call GenerateWall(2600, 500, 800, 400, 1)
'7
Call GenerateWall(500, 1300, 600, 500, 1)
'8
Call GenerateWall(1500, 1300, 600, 800, 1)
'9
Call GenerateWall(1200, 2300, 600, 400, 1)
'10
Call GenerateWall(1200, 3100, 600, 600, 1)
'11
Call GenerateWall(2300, 2400, 1100, 700, 1)
'12
Call GenerateWall(2300, 3500, 200, 500, 1)


Holocount = 3
ReDim Holoscreen(1 To Holocount)
With Holoscreen(1)
    .Goal = False
    .GraphicsDC = 2
    .Text = "Welcome to the advanced training level.  Here you will test and study the capabilities of your weapons.  You should examine whether they are most effective at long or short range, how much shots spread, how much damage they do, and what their maximum range is.  If you run out of drones to destroy, click the exit button to replay the level."
    .XCoord = 850
    .YCoord = 3600
End With
With Holoscreen(2)
    .Goal = True
    .GraphicsDC = 3
    .Text = "Level complete."
    .XCoord = 2900
    .YCoord = 3600
End With
With Holoscreen(3)
    .Goal = False
    .GraphicsDC = 2
    .Text = "Gate opened."
    .XCoord = 2300
    .YCoord = 700
End With
WeaponCount = 0
EnemyCount = 10
ReDim Enemy(1 To EnemyCount)
With Enemy(1)
    .Weapon = 8
    .Ammo = WeaponData(Enemy(1).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(1).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 4
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1400
    .TargetY = 3000
    .XCoord = 1400
    .YCoord = 3000
End With
With Enemy(2)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(2).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(2).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1400
    .TargetY = 2200
    .XCoord = 1400
    .YCoord = 2200
End With
With Enemy(3)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(3).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(3).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 700
    .TargetY = 2000
    .XCoord = 1800
    .YCoord = 2000
End With
With Enemy(4)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(4).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(4).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1800
    .TargetY = 3550
    .XCoord = 2900
    .YCoord = 3550
End With
With Enemy(5)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(5).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(5).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2900
    .TargetY = 1400
    .XCoord = 2900
    .YCoord = 1400
End With
With Enemy(6)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(6).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(6).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2400
    .TargetY = 1400
    .XCoord = 2400
    .YCoord = 1400
End With
With Enemy(7)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(7).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(7).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 1
    .HP = 150 * EnemyBonus
    .MoveSpeed = 2
    .Stance = 3
    .TargetX = 2300
    .TargetY = 700
    .XCoord = 2300
    .YCoord = 700
End With
With Enemy(8)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(1).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(1).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1900
    .TargetY = 2200
    .XCoord = 1900
    .YCoord = 3600
End With
With Enemy(9)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(1).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(1).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 2200
    .TargetY = 2200
    .XCoord = 2200
    .YCoord = 3600
End With
With Enemy(10)
    .Weapon = 6
    .Ammo = WeaponData(Enemy(1).Weapon).StartAmmo
    .ClipAmmo = WeaponData(Enemy(1).Weapon).ClipSize
    .Dead = False
    .GraphicsDC = 2
    .HP = 120 * EnemyBonus
    .MoveSpeed = 3
    .Stance = 4
    .TargetX = 1950
    .TargetY = 3350
    .XCoord = 2050
    .YCoord = 3350
End With

TriggerCount = 1
ReDim Trigger(1 To TriggerCount)
With Trigger(1)
    .XCoord = 2200
    .YCoord = 600
    .Height = 200
    .Width = 200
    .Link = 12
End With
End Sub
