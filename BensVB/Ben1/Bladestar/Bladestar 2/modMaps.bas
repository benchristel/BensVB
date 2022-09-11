Attribute VB_Name = "modMaps"

Public Sub GenerateMap1() 'oracle
Dim i
'=================
'load map graphics
'=================
BackgroundDC = GenerateDC(App.Path & "\Graphics\Background1.bmp")
For i = 1 To 2
BackBuffDC(i) = GenerateDC(App.Path & "\Graphics\Background1.bmp")
Next i
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
'==================
'Player Spawnpoints
'==================
PlayerSpawnCount = 5
WeaponSpawnCount = 7
ReDim PlayerSpawn(1 To 5)
ReDim WeaponSpawn(1 To 7)
With PlayerSpawn(1)
    .XCoord = 680
    .YCoord = 540
End With
With PlayerSpawn(2)
    .XCoord = 1120
    .YCoord = 1520
End With
With PlayerSpawn(3)
    .XCoord = 1920
    .YCoord = 1520
End With
With PlayerSpawn(4)
    .XCoord = 240
    .YCoord = 1440
End With
With PlayerSpawn(5)
    .XCoord = 1680
    .YCoord = 2680
End With
With WeaponSpawn(1)
    .Frequency = 25
    .TimeLeft = 0
    .WeaponType = VariantWeapon(2)
    .XCoord = 480
    .YCoord = 1840
End With
With WeaponSpawn(2)
    .Frequency = 20
    .TimeLeft = 0
    .WeaponType = VariantWeapon(3)
    .XCoord = 1960
    .YCoord = 320
End With
With WeaponSpawn(3)
    .Frequency = 25
    .TimeLeft = 0
    .WeaponType = VariantWeapon(4)
    .XCoord = 2320
    .YCoord = 1360
End With
With WeaponSpawn(4)
    .Frequency = 35
    .TimeLeft = 0
    .WeaponType = VariantWeapon(5)
    .XCoord = 1120
    .YCoord = 2320
End With
With WeaponSpawn(5)
    .Frequency = 35
    .TimeLeft = 0
    .WeaponType = VariantWeapon(6)
    .XCoord = 2320
    .YCoord = 2680
End With
With WeaponSpawn(6)
    .Frequency = 35
    .TimeLeft = 0
    .WeaponType = VariantWeapon(5)
    .XCoord = 800
    .YCoord = 760
End With
With WeaponSpawn(7)
    .Frequency = 35
    .TimeLeft = 0
    .WeaponType = VariantWeapon(9)
    .XCoord = 1600
    .YCoord = 1520
End With

With Player(1)
    .XCoord = 680
    .YCoord = 540
    .Weapon = StartWeapon
If WeaponData(Player(1).Weapon).StartAmmo - WeaponData(Player(1).Weapon).ClipSize >= 0 Then
    .Ammo = WeaponData(Player(1).Weapon).StartAmmo - WeaponData(Player(1).Weapon).ClipSize
    .ClipAmmo = WeaponData(Player(1).Weapon).ClipSize
Else
    .Ammo = 0
    .ClipAmmo = WeaponData(Player(1).Weapon).StartAmmo
End If
End With
With Player(2)
    .XCoord = 1680
    .YCoord = 2680
    .Weapon = StartWeapon
If WeaponData(Player(2).Weapon).StartAmmo - WeaponData(Player(2).Weapon).ClipSize >= 0 Then
    .Ammo = WeaponData(Player(2).Weapon).StartAmmo - WeaponData(Player(2).Weapon).ClipSize
    .ClipAmmo = WeaponData(Player(2).Weapon).ClipSize
Else
    .Ammo = 0
    .ClipAmmo = WeaponData(Player(2).Weapon).StartAmmo
End If
End With
End Sub

Public Sub GenerateWall(x As Double, y As Double, Height As Double, Width As Double, WallType As Integer)
If WallType <> 1 And WallType <> 2 Then Exit Sub
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

Public Sub GenerateMap2() 'giza
Dim i
'=================
'load map graphics
'=================
BackgroundDC = GenerateDC(App.Path & "\Graphics\Background2.bmp")
For i = 1 To 2
BackBuffDC(i) = GenerateDC(App.Path & "\Graphics\Background2.bmp")
Next i
For i = 1 To 2
WallDC(i) = GenerateDC(App.Path & "\Graphics\Wall" & i & "_2.bmp")
Next i
'1
Call GenerateWall(0, 0, 1300, 900, 1)
'2
Call GenerateWall(900, 0, 600, 900, 1)
'3
Call GenerateWall(1800, 0, 400, 600, 1)
'4
Call GenerateWall(2400, 0, 3800, 400, 1)
'5
Call GenerateWall(1100, 600, 200, 200, 1)
'6
Call GenerateWall(1600, 600, 200, 200, 1)
'7
Call GenerateWall(1900, 600, 200, 50, 1)
'8
Call GenerateWall(1950, 600, 50, 300, 1)
'9
Call GenerateWall(2250, 600, 200, 50, 1)
'10
Call GenerateWall(1900, 900, 50, 400, 2)
'11
Call GenerateWall(1100, 900, 400, 200, 1)
'12
Call GenerateWall(1200, 1100, 300, 200, 1)
'13
Call GenerateWall(1600, 900, 200, 200, 1)
'14
Call GenerateWall(1500, 1100, 200, 300, 1)
'15
Call GenerateWall(1600, 1300, 100, 200, 2)
'16
Call GenerateWall(400, 1300, 100, 400, 1)
'17
Call GenerateWall(1600, 1400, 300, 200, 1)
'18
Call GenerateWall(0, 1300, 1500, 400, 1)
'19
Call GenerateWall(600, 1500, 300, 200, 1)
'20
Call GenerateWall(800, 1600, 200, 100, 1)
'21
Call GenerateWall(600, 1800, 100, 100, 1)
'22
Call GenerateWall(900, 1600, 200, 100, 2)
'23
Call GenerateWall(1000, 1600, 200, 200, 1)
'24
Call GenerateWall(1200, 1500, 500, 200, 1)
'25
Call GenerateWall(1100, 1800, 100, 100, 1)
'26
Call GenerateWall(1600, 1700, 100, 200, 2)
'27
Call GenerateWall(1600, 1800, 200, 400, 1)
'28
Call GenerateWall(1200, 2000, 100, 200, 2)
'29
Call GenerateWall(2200, 1800, 900, 200, 1)
'30
Call GenerateWall(600, 2100, 100, 100, 1)
'31
Call GenerateWall(600, 2200, 300, 200, 1)
'32
Call GenerateWall(1100, 2100, 300, 300, 1)
'33
Call GenerateWall(1000, 2200, 400, 200, 1)
'34
Call GenerateWall(1500, 2200, 200, 400, 1)
'35
Call GenerateWall(2000, 2200, 200, 200, 1)
'36
Call GenerateWall(1700, 2400, 200, 200, 1)
'37
Call GenerateWall(1400, 2200, 200, 100, 2)
'38
Call GenerateWall(600, 2600, 100, 200, 1)
'39
Call GenerateWall(600, 2700, 100, 200, 2)
'40
Call GenerateWall(1000, 2700, 200, 200, 1)
'41
Call GenerateWall(1700, 2700, 200, 700, 1)
'42
Call GenerateWall(1000, 2900, 200, 400, 1)
'43
Call GenerateWall(1500, 2900, 200, 100, 1)
'44
Call GenerateWall(1600, 2900, 900, 800, 1)
'45
Call GenerateWall(0, 2800, 1000, 800, 1)
'46
Call GenerateWall(800, 3300, 500, 600, 1)
'47
Call GenerateWall(1400, 3400, 50, 200, 2)
'48
Call GenerateWall(1400, 2900, 200, 100, 2)
'49
Call GenerateWall(1400, 1100, 200, 100, 2)
'==================
'Player Spawnpoints
'==================
PlayerSpawnCount = 5
WeaponSpawnCount = 7
ReDim PlayerSpawn(1 To 5)
ReDim WeaponSpawn(1 To 7)
With PlayerSpawn(1)
    .XCoord = 1000
    .YCoord = 1450
End With
With PlayerSpawn(2)
    .XCoord = 2050
    .YCoord = 2550
End With
With PlayerSpawn(3)
    .XCoord = 1500
    .YCoord = 3200
End With
With PlayerSpawn(4)
    .XCoord = 500
    .YCoord = 2600
End With
With PlayerSpawn(5)
    .XCoord = 2100
    .YCoord = 500
End With
With WeaponSpawn(1)
    .Frequency = 100
    .TimeLeft = 20
    .WeaponType = VariantWeapon(7)
    .XCoord = 2100
    .YCoord = 800
End With
With WeaponSpawn(2)
    .Frequency = 35
    .TimeLeft = 0
    .WeaponType = VariantWeapon(6)
    .XCoord = 2100
    .YCoord = 2100
End With
With WeaponSpawn(3)
    .Frequency = 35
    .TimeLeft = 0
    .WeaponType = VariantWeapon(4)
    .XCoord = 1000
    .YCoord = 700
End With
With WeaponSpawn(4)
    .Frequency = 35
    .TimeLeft = 0
    .WeaponType = VariantWeapon(2)
    .XCoord = 1300
    .YCoord = 2500
End With
With WeaponSpawn(5)
    .Frequency = 35
    .TimeLeft = 0
    .WeaponType = VariantWeapon(3)
    .XCoord = 900
    .YCoord = 2000
End With
With WeaponSpawn(6)
    .Frequency = 35
    .TimeLeft = 0
    .WeaponType = VariantWeapon(8)
    .XCoord = 500
    .YCoord = 2700
End With
With WeaponSpawn(7)
    .Frequency = 35
    .TimeLeft = 0
    .WeaponType = VariantWeapon(9)
    .XCoord = 1500
    .YCoord = 3300
End With

With Player(1)
    .XCoord = 1000
    .YCoord = 1450
    .Weapon = StartWeapon
If WeaponData(Player(1).Weapon).StartAmmo - WeaponData(Player(1).Weapon).ClipSize >= 0 Then
    .Ammo = WeaponData(Player(1).Weapon).StartAmmo - WeaponData(Player(1).Weapon).ClipSize
    .ClipAmmo = WeaponData(Player(1).Weapon).ClipSize
Else
    .Ammo = 0
    .ClipAmmo = WeaponData(Player(1).Weapon).StartAmmo
End If
End With
With Player(2)
    .XCoord = 2050
    .YCoord = 2550
    .Weapon = StartWeapon
If WeaponData(Player(2).Weapon).StartAmmo - WeaponData(Player(2).Weapon).ClipSize >= 0 Then
    .Ammo = WeaponData(Player(2).Weapon).StartAmmo - WeaponData(Player(2).Weapon).ClipSize
    .ClipAmmo = WeaponData(Player(2).Weapon).ClipSize
Else
    .Ammo = 0
    .ClipAmmo = WeaponData(Player(2).Weapon).StartAmmo
End If
End With
End Sub

Public Sub GenerateMap3() 'Quicksilver
Dim i
'=================
'load map graphics
'=================
BackgroundDC = GenerateDC(App.Path & "\Graphics\Background1.bmp")
For i = 1 To 2
BackBuffDC(i) = GenerateDC(App.Path & "\Graphics\Background1.bmp")
Next i
For i = 1 To 2
WallDC(i) = GenerateDC(App.Path & "\Graphics\Wall" & i & "_1.bmp")
Next i
'1
Call GenerateWall(0, 1600, 650, 600, 1)
'2
Call GenerateWall(600, 800, 1200, 500, 1)
'3
Call GenerateWall(1100, 700, 500, 650, 1)
'4
Call GenerateWall(1750, 600, 450, 500, 1)
'5
Call GenerateWall(2250, 700, 500, 850, 1)
'6
Call GenerateWall(3100, 800, 650, 600, 1)
'7
Call GenerateWall(3250, 1450, 400, 1050, 1)
'8
Call GenerateWall(3900, 1850, 500, 400, 1)
'9
Call GenerateWall(3250, 2350, 950, 1050, 1)
'10
Call GenerateWall(900, 3200, 400, 2350, 1)
'11
Call GenerateWall(0, 2900, 400, 900, 1)
'12
Call GenerateWall(0, 2250, 650, 400, 1)
'13
Call GenerateWall(850, 2800, 100, 50, 1)
'14
Call GenerateWall(850, 2500, 100, 50, 1)
'15
Call GenerateWall(850, 2200, 100, 50, 1)
'16
Call GenerateWall(700, 2200, 50, 150, 1)
'17
Call GenerateWall(900, 2200, 400, 600, 2)
'18
Call GenerateWall(1300, 1600, 400, 200, 1)
'19
Call GenerateWall(1300, 1400, 200, 450, 1)
'20
Call GenerateWall(1750, 1550, 50, 150, 1)
'21
Call GenerateWall(1900, 1550, 50, 200, 2)
'22
Call GenerateWall(2100, 1550, 50, 150, 1)
'23
Call GenerateWall(2250, 1400, 200, 50, 1)
'24
Call GenerateWall(2300, 1400, 200, 400, 2)
'25
Call GenerateWall(2700, 1400, 200, 50, 1)
'26
Call GenerateWall(2750, 1400, 50, 150, 1)
'27
Call GenerateWall(2700, 1800, 200, 50, 1)
'28
Call GenerateWall(2750, 1950, 50, 150, 1)
'29
Call GenerateWall(3100, 1950, 300, 400, 1)
'30
Call GenerateWall(2750, 2200, 50, 150, 1)
'31
Call GenerateWall(2700, 2200, 200, 50, 1)
'32
Call GenerateWall(2700, 2600, 200, 50, 1)
'33
Call GenerateWall(2750, 2750, 50, 500, 1)
'34
Call GenerateWall(2500, 2200, 600, 200, 2)
'35
Call GenerateWall(2500, 2800, 400, 750, 2)
'36
Call GenerateWall(2100, 2900, 100, 100, 1)
'37
Call GenerateWall(1100, 2600, 400, 800, 1)
'38
Call GenerateWall(1750, 1800, 50, 150, 1)
'39
Call GenerateWall(2100, 1800, 50, 150, 1)
'40
Call GenerateWall(2500, 1600, 400, 200, 2)
'41
Call GenerateWall(1700, 1800, 200, 50, 1)
'42
Call GenerateWall(2250, 1800, 200, 50, 1)
'43
Call GenerateWall(1700, 2000, 200, 50, 2)
'44
Call GenerateWall(2250, 2000, 200, 50, 2)
'45
Call GenerateWall(1700, 2200, 200, 50, 2)
'46
Call GenerateWall(2250, 2200, 200, 50, 2)
'47
Call GenerateWall(1750, 2350, 50, 150, 2)
'48
Call GenerateWall(2100, 2350, 50, 150, 2)

'==================
'Player Spawnpoints
'==================
PlayerSpawnCount = 6
WeaponSpawnCount = 7
ReDim PlayerSpawn(1 To 6)
ReDim WeaponSpawn(1 To 7)
With PlayerSpawn(1)
    .XCoord = 2400
    .YCoord = 3100
End With
With PlayerSpawn(2)
    .XCoord = 1800
    .YCoord = 1500
End With
With PlayerSpawn(3)
    .XCoord = 3400
    .YCoord = 1900
End With
With PlayerSpawn(4)
    .XCoord = 3100
    .YCoord = 2600
End With
With PlayerSpawn(5)
    .XCoord = 3400
    .YCoord = 1900
End With
With PlayerSpawn(6)
    .XCoord = 800
    .YCoord = 2100
End With
With WeaponSpawn(1)
    .Frequency = 35
    .TimeLeft = 0
    .WeaponType = VariantWeapon(5)
    .XCoord = 800
    .YCoord = 2700
End With
With WeaponSpawn(2)
    .Frequency = 35
    .TimeLeft = 0
    .WeaponType = VariantWeapon(4)
    .XCoord = 1200
    .YCoord = 1300
End With
With WeaponSpawn(3)
    .Frequency = 35
    .TimeLeft = 0
    .WeaponType = VariantWeapon(3)
    .XCoord = 3000
    .YCoord = 1300
End With
With WeaponSpawn(4)
    .Frequency = 35
    .TimeLeft = 0
    .WeaponType = VariantWeapon(2)
    .XCoord = 1400
    .YCoord = 2100
End With
With WeaponSpawn(5)
    .Frequency = 35
    .TimeLeft = 0
    .WeaponType = VariantWeapon(9)
    .XCoord = 3000
    .YCoord = 2100
End With
With WeaponSpawn(6)
    .Frequency = 100
    .TimeLeft = 20
    .WeaponType = VariantWeapon(7)
    .XCoord = 1850
    .YCoord = 1950
End With
'With WeaponSpawn(7)
'    .Frequency = 60
'    .TimeLeft = 0
'    .WeaponType = VariantWeapon(9)
'    .XCoord = 1500
'    .YCoord = 3100
'End With
With WeaponSpawn(7)
    .Frequency = 100
    .TimeLeft = 0
    .WeaponType = VariantWeapon(10)
    .XCoord = 3800
    .YCoord = 2100
End With

With Player(1)
    .XCoord = 2400
    .YCoord = 3100
    .Weapon = StartWeapon
If WeaponData(Player(1).Weapon).StartAmmo - WeaponData(Player(1).Weapon).ClipSize >= 0 Then
    .Ammo = WeaponData(Player(1).Weapon).StartAmmo - WeaponData(Player(1).Weapon).ClipSize
    .ClipAmmo = WeaponData(Player(1).Weapon).ClipSize
Else
    .Ammo = 0
    .ClipAmmo = WeaponData(Player(1).Weapon).StartAmmo
End If
End With
With Player(2)
    .XCoord = 1800
    .YCoord = 1500
    .Weapon = StartWeapon
If WeaponData(Player(2).Weapon).StartAmmo - WeaponData(Player(2).Weapon).ClipSize >= 0 Then
    .Ammo = WeaponData(Player(2).Weapon).StartAmmo - WeaponData(Player(2).Weapon).ClipSize
    .ClipAmmo = WeaponData(Player(2).Weapon).ClipSize
Else
    .Ammo = 0
    .ClipAmmo = WeaponData(Player(2).Weapon).StartAmmo
End If
End With
End Sub
