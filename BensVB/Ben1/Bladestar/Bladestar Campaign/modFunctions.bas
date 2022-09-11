Attribute VB_Name = "modFunctions"
Const PI = 3.1416
Const KEY_TOGGLED As Integer = &H1
Const KEY_PRESSED As Integer = &H1000

Public Function Distance(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
Distance = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
End Function

Public Function Overlap(X1 As Double, Y1 As Double, Height1 As Double, Width1 As Double, X2 As Double, Y2 As Double, Height2 As Double, Width2 As Double) As Boolean
If X1 >= X2 - Width1 And X1 <= X2 + Width2 And Y1 >= Y2 - Height1 And Y1 <= Y2 + Height2 Then
Overlap = True
Else
Overlap = False
End If
End Function

Public Function FindLine(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, Length) As Slope
Dim proportion, pointdistance
pointdistance = Distance(X1, Y1, X2, Y2)
If pointdistance = 0 Then
FindLine.Rise = 0
FindLine.Run = 0
Exit Function
End If
proportion = Length / pointdistance
FindLine.Run = proportion * (X2 - X1)
FindLine.Rise = proportion * (Y2 - Y1)
End Function
Public Sub UpdateKeys()
Dim i, tempweapon As Weapon, tempweapontype As Integer, tempammo As Integer, tempclipammo As Integer
Static rdown(1 To 2) As Boolean, sdown(1 To 2) As Boolean
'===================
'Player 1 Keystrokes
'===================
If Player.Dead = True Then Exit Sub
If (GetKeyState(vbKeyS) And KEY_PRESSED) Then
    Player.MoveBack = True
    Player.Sniping = False
Else
    Player.MoveBack = False
End If
If (GetKeyState(vbKeyW) And KEY_PRESSED) Then
    Player.MoveForward = True
    Player.Sniping = False
Else
    Player.MoveForward = False
End If
'=====
'A KEY
'=====
If (GetKeyState(vbKeyA) And KEY_PRESSED) Then
    Select Case AccountData(CurrentAccount).Layout
        Case Is = 2
            If AccountData(CurrentAccount).Inverted = False Then
                Player.TurnLeft = True
            Else
                Player.TurnRight = True
            End If
        Case Else
            If AccountData(CurrentAccount).Inverted = False Then
                Player.StrafeLeft = True
            Else
                Player.StrafeRight = True
            End If
    End Select
Else
    Select Case AccountData(CurrentAccount).Layout
        Case Is = 2
            If AccountData(CurrentAccount).Inverted = False Then
                Player.TurnLeft = False
            Else
                Player.TurnRight = False
            End If
        Case Else
            If AccountData(CurrentAccount).Inverted = False Then
                Player.StrafeLeft = False
            Else
                Player.StrafeRight = False
            End If
    End Select
End If
'
'
'
If (GetKeyState(vbKeyD) And KEY_PRESSED) Then
    Select Case AccountData(CurrentAccount).Layout
        Case Is = 2
            If AccountData(CurrentAccount).Inverted = False Then
                Player.TurnRight = True
            Else
                Player.TurnLeft = True
            End If
        Case Else
            If AccountData(CurrentAccount).Inverted = False Then
                Player.StrafeRight = True
            Else
                Player.StrafeLeft = True
            End If
    End Select
Else
    Select Case AccountData(CurrentAccount).Layout
        Case Is = 2
            If AccountData(CurrentAccount).Inverted = False Then
                Player.TurnRight = False
            Else
                Player.TurnLeft = False
            End If
        Case Else
            If AccountData(CurrentAccount).Inverted = False Then
                Player.StrafeRight = False
            Else
                Player.StrafeLeft = False
            End If
    End Select
End If
'============
'Q and E KEYS
'============
If AccountData(CurrentAccount).Layout = 2 Then
If AccountData(CurrentAccount).Inverted = False Then
    If (GetKeyState(vbKeyQ) And KEY_PRESSED) Then
        Player.StrafeLeft = True
    Else
        Player.StrafeLeft = False
    End If
    If (GetKeyState(vbKeyE) And KEY_PRESSED) Then
        Player.StrafeRight = True
    Else
        Player.StrafeRight = False
    End If
Else
    If (GetKeyState(vbKeyQ) And KEY_PRESSED) Then
        Player.StrafeRight = True
    Else
        Player.StrafeRight = False
    End If
    If (GetKeyState(vbKeyE) And KEY_PRESSED) Then
        Player.StrafeLeft = True
    Else
        Player.StrafeLeft = False
    End If
End If
End If
'===========
'ACTION KEYS
'===========
If (GetKeyState(vbKeyShift) And KEY_PRESSED) Then
If Player.Switching = False Then
    Player.Switching = True 'switching disables shooting and reloading until both fire and reload buttons are released.
    Player.Reloading = False
    frmMain.lblPlayerStatus.Caption = ""
    tempammo = Player.Ammo(1)
    tempclipammo = Player.ClipAmmo(1)
    tempweapontype = Player.Weapon(1)
    Player.Ammo(1) = Player.Ammo(2)
    Player.ClipAmmo(1) = Player.ClipAmmo(2)
    Player.Weapon(1) = Player.Weapon(2)
    Player.Ammo(2) = tempammo
    Player.ClipAmmo(2) = tempclipammo
    Player.Weapon(2) = tempweapontype
    Player.Reloading = False
    If WeaponData(Player.Weapon(1)).Reloads = True Then
        frmMain.lblClipAmmo.Visible = True
    Else
        frmMain.lblClipAmmo.Visible = False
    End If
    frmMain.lblClipAmmo.Caption = Player.ClipAmmo(1)
    frmMain.lblAmmo.Caption = Player.Ammo(1)
    frmMain.lblWeaponName(1).Caption = WeaponData(Player.Weapon(1)).Name
    frmMain.lblWeaponName(2).Caption = WeaponData(Player.Weapon(2)).Name
End If
If (Player.ClipAmmo(1) = 0 And Player.Ammo(1) = 0) Then
    Call Message("Out of ammo.  Press [SHIFT] to switch to " & WeaponData(Player.Weapon(2)).Name & ".", -2)
ElseIf (Player.ClipAmmo(1) > 0 Or Player.Ammo(1) > 0) And PromptTime = -2 Then
    frmMain.lblPrompt.Caption = ""
    PromptTime = 0
End If
Else
Player.Switching = False
End If
If GetKeyState(vbKeyLButton) < 0 Then
        If WeaponData(Player.Weapon(1)).SemiAuto = True And Player.Firing = False And Player.CoolDownTime <= 0 Then
            Call PlayerFire
            Player.Firing = True
            Player.CoolDownTime = WeaponData(Player.Weapon(1)).CoolDown
        Else
            Player.Firing = True
            frmMain.lblClipAmmo.Caption = Player.ClipAmmo(1)
            Player.Reloading = False
            frmMain.lblPlayerStatus.Caption = ""
        End If
Else
            Player.Firing = False
            If (GetKeyState(vbKeyP) And KEY_PRESSED) = False Then Switching = False
End If
'
'
'
If (GetKeyState(vbKeyRButton) And KEY_PRESSED) And Player.Firing = False Then 'keypad + key was pressed
If Player.PickUpItem = False Then
 For i = WeaponMin To WeaponCount
    If Weapon(i).Deleted = False Then
    If Distance(Weapon(i).XCoord, Weapon(i).YCoord, Player.XCoord, Player.YCoord) <= 100 And Distance(Weapon(i).XCoord, Weapon(i).YCoord, Player.TargetX + Player.XCoord, Player.TargetY + Player.YCoord) <= 20 Then
    'For j = 1 To 2
        'If Player.Weapon(j) = Weapon(i).WeaponType And WeaponData(Player.Weapon(j)).PickUpAmmo = True Then
            'If Weapon(i).Ammo < WeaponData(Player.Weapon(j)).MaxAmmo - Player.Ammo(j) Then
                'Player.Ammo(j) = Player.Ammo(j) + Weapon(i).Ammo
                'Weapon(i).Ammo = 0
            'Else
                'Weapon(i).Ammo = Weapon(i).Ammo - (WeaponData(Player.Weapon(j)).MaxAmmo - Player.Ammo(j))
                'Player.Ammo(j) = WeaponData(Player.Weapon(j)).MaxAmmo
            'End If
            'frmMain.lblAmmo.Caption = Player.Ammo(1)
            'Player.PickUpItem = True
            'GoTo DoNotReload
        'End If
    'Next j
        WeaponCount = WeaponCount + 1
        ReDim Preserve Weapon(1 To WeaponCount)
        With Weapon(WeaponCount)
            .Ammo = Player.Ammo(1)
            .ClipAmmo = Player.ClipAmmo(1)
            .ClipSize = WeaponData(Player.Weapon(1)).ClipSize
            .Deleted = False
            .DespawnEnabled = True
            .Name = WeaponData(Player.Weapon(1)).Name
            .TimeToDespawn = 40
            .WeaponType = Player.Weapon(1)
            .XCoord = Player.XCoord
            .YCoord = Player.YCoord
        End With
        With Player
            .Ammo(1) = Weapon(i).Ammo
            .ClipAmmo(1) = Weapon(i).ClipAmmo
            .Weapon(1) = Weapon(i).WeaponType
            .Reloading = False
        End With
        frmMain.lblPlayerStatus.Caption = ""
        Weapon(i).Deleted = True
        frmMain.lblAmmo.Caption = Player.Ammo(1)
        frmMain.lblClipAmmo.Caption = Player.ClipAmmo(1)
            If WeaponData(Player.Weapon(1)).Reloads = True Then
                frmMain.lblClipAmmo.Visible = True
            Else
                frmMain.lblClipAmmo.Visible = False
            End If
        frmMain.lblWeaponName(1).Caption = WeaponData(Player.Weapon(1)).Name
        Player.PickUpItem = True
        GoTo DoNotReload
    End If
    End If
    Next i
End If
    If Player.Reloading = False And Player.PickUpItem = False Then
    Player.Reloading = True
    frmMain.lblPlayerStatus.Caption = "Reloading..."
    Player.ReloadTime = WeaponData(Player.Weapon(1)).ReloadTime
    End If
    Player.Firing = False
DoNotReload:
Else
    Player.PickUpItem = False
End If
'frmMain.lblPlayerStatus.Caption = GetKeyState(vbKeyLButton)
End Sub
Public Sub UpdateObjects()
Dim i, j As Integer, k, tempslope(1 To 5) As Slope, tempslope2 As Slope, transferammo As Integer, mousemovedist As Single
Dim MoveX As Boolean, MoveY As Boolean, tempdistance As Double
'>>>Update shots and check collisions
For i = ShotMin To ShotCount
If Shot(i).Deleted = False Then
For k = 1 To Shot(i).FrameSteps
Shot(i).XCoord = Shot(i).XCoord + Shot(i).XVector
Shot(i).YCoord = Shot(i).YCoord + Shot(i).YVector
Shot(i).Distance = Shot(i).Distance + Shot(i).Speed
Shot(i).Lifespan = Shot(i).Lifespan - 1
If Shot(i).Alignment = 2 And Player.Dead = False Then
        If Distance(Shot(i).XCoord, Shot(i).YCoord, Player.XCoord, Player.YCoord) <= Shot(i).Radius + 20 Then
            Shot(i).Deleted = True
            Shot(i).Damage = Shot(i).Damage + Shot(i).Distance * Shot(i).DamageMultiplier
                If Shot(i).Damage >= Player.Shield Then
                    Player.HP = Player.HP - Shot(i).Damage + Player.Shield
                    Player.Shield = 0
                    If Player.HP <= 0 Then
                        Player.HP = 0
                    End If
                Else
                    Player.Shield = Player.Shield - Shot(i).Damage
                End If
                If Shot(i).ExplodeDamage > 0 Then
                    Call Explode(Shot(i).XCoord, Shot(i).YCoord, Shot(i).ExplodeDamage, Shot(i).ExplodeRadius)
                End If
                Player.RechargeTime = 250 * EnemyBonus
        End If
End If
For j = 1 To EnemyCount
If Shot(i).Alignment = 1 And Enemy(j).Dead = False Then
        If Distance(Shot(i).XCoord, Shot(i).YCoord, Enemy(j).XCoord, Enemy(j).YCoord) <= Shot(i).Radius + 20 Then
            Shot(i).Deleted = True
            Shot(i).Damage = Shot(i).Damage + Shot(i).Distance * Shot(i).DamageMultiplier
                Enemy(j).HP = Enemy(j).HP - Shot(i).Damage
                Select Case Enemy(j).Stance
                    Case Is = 1
                    Enemy(j).Stance = 3
                    Case Is = 2
                    Enemy(j).Stance = 4
                End Select
                If Shot(i).ExplodeDamage > 0 Then
                    Call Explode(Shot(i).XCoord, Shot(i).YCoord, Shot(i).ExplodeDamage, Shot(i).ExplodeRadius)
                End If
        End If
End If
Next j
For j = 1 To WallCount
    If Wall(j).Type = 1 Then
        If Shot(i).Bounce = False And Overlap(Shot(i).XCoord, Shot(i).YCoord, 0, 0, Wall(j).XCoord, Wall(j).YCoord, Wall(j).Height, Wall(j).Width) = True Then
            Shot(i).Deleted = True
            If Shot(i).ExplodeDamage > 0 Then
                Call Explode(Shot(i).XCoord, Shot(i).YCoord, Shot(i).ExplodeDamage, Shot(i).ExplodeRadius)
                'Call CheckExplosion(i)
            End If
        End If
        If Shot(i).Bounce = True And Overlap(Shot(i).XCoord + Shot(i).XVector, Shot(i).YCoord + Shot(i).YVector, 0, 0, Wall(j).XCoord, Wall(j).YCoord, Wall(j).Height, Wall(j).Width) = True Then
            If Shot(i).XCoord < Wall(j).XCoord Or Shot(i).XCoord > Wall(j).XCoord + Wall(j).Width Then
                Shot(i).XVector = Shot(i).XVector * -1
            Else
                Shot(i).YVector = Shot(i).YVector * -1
            End If
        End If
    End If
Next j
If Shot(i).Lifespan <= 0 Then
Shot(i).Deleted = True
    If Shot(i).ExplodeDamage > 0 Then
        Call Explode(Shot(i).XCoord, Shot(i).YCoord, Shot(i).ExplodeDamage, Shot(i).ExplodeRadius)
        'Call CheckExplosion(i)
    End If
End If
Next k
ElseIf i = ShotMin Then ShotMin = ShotMin + 1
If ShotMin > ShotCount Then
    ShotMin = 1
    ShotCount = 0
    ReDim Shot(1 To 1)
End If
End If
Next i
'>>>Update player positions
If AccountData(CurrentAccount).Layout <> 3 Then
    If Player.TurnLeft = True Then Player.Theta = Player.Theta + 0.02
    If Player.TurnRight = True Then Player.Theta = Player.Theta - 0.02
    If Player.Theta > 2 Then Player.Theta = Player.Theta - 2
    If Player.Theta < 0 Then Player.Theta = Player.Theta + 2
    Player.MoveX = 100 * Sin(Player.Theta * PI)
    Player.MoveY = 100 * Cos(Player.Theta * PI)
End If
If Player.MoveForward = True Then tempslope(1) = FindLine(0, 0, Player.MoveX, Player.MoveY, 4 + WeaponData(Player.Weapon(1)).SpeedBonus)
If Player.StrafeLeft = True Then tempslope(2) = FindLine(0, 0, Player.MoveY, -Player.MoveX, 4 + WeaponData(Player.Weapon(1)).SpeedBonus)
If Player.StrafeRight = True Then tempslope(3) = FindLine(0, 0, -Player.MoveY, Player.MoveX, 4 + WeaponData(Player.Weapon(1)).SpeedBonus)
If Player.MoveBack = True Then tempslope(4) = FindLine(0, 0, -Player.MoveX, -Player.MoveY, 4 + WeaponData(Player.Weapon(1)).SpeedBonus)
With tempslope(5)
    .Rise = tempslope(1).Rise + tempslope(2).Rise + tempslope(3).Rise + tempslope(4).Rise
    .Run = tempslope(1).Run + tempslope(2).Run + tempslope(3).Run + tempslope(4).Run
End With
If tempslope(5).Rise = 0 And tempslope(5).Run = 0 Then GoTo nocheck 'if player is not moving, don't look for collisions
'check to see if new position would overlap with a wall
MoveX = True
MoveY = True
For j = 1 To WallCount
If Wall(j).Type < 0 Then GoTo checknextwall
If MoveX = True Then
    If Overlap(Player.XCoord - 21 + tempslope(5).Run, Player.YCoord - 21, 40, 40, Wall(j).XCoord, Wall(j).YCoord, Wall(j).Height, Wall(j).Width) = True Then
    MoveX = False
    End If
End If
If MoveY = True Then
    If Overlap(Player.XCoord - 21, Player.YCoord + tempslope(5).Rise - 21, 40, 40, Wall(j).XCoord, Wall(j).YCoord, Wall(j).Height, Wall(j).Width) = True Then
    MoveY = False
    End If
End If
checknextwall:
Next j
If MoveX = True Then Player.XCoord = Player.XCoord + tempslope(5).Run
If MoveY = True Then Player.YCoord = Player.YCoord + tempslope(5).Rise
nocheck:
'>>>check for ammo to pick up
For j = WeaponMin To WeaponCount
    For k = 1 To 2
    If WeaponData(Weapon(j).WeaponType).Despawns = True And Weapon(j).Ammo = 0 And Weapon(j).ClipAmmo = 0 Then Weapon(j).Deleted = True
    If Weapon(j).Deleted = False And Weapon(j).WeaponType = Player.Weapon(k) And Player.Dead = False And WeaponData(Player.Weapon(k)).PickUpAmmo = True Then
        If Overlap(Player.XCoord, Player.YCoord, 0, 0, Weapon(j).XCoord - 20, Weapon(j).YCoord - 20, 40, 40) = True Then
            If Weapon(j).Ammo < WeaponData(Player.Weapon(k)).MaxAmmo - Player.Ammo(k) Then
                If Weapon(j).Ammo > 1 Then Call Message("Picked up " & Weapon(j).Ammo & " " & WeaponData(Player.Weapon(k)).AmmoName(2) & ".", 180)
                If Weapon(j).Ammo = 1 Then Call Message("Picked up " & Weapon(j).Ammo & " " & WeaponData(Player.Weapon(k)).AmmoName(1) & ".", 180)
                Player.Ammo(k) = Player.Ammo(k) + Weapon(j).Ammo
                Weapon(j).Ammo = 0
            Else
                If WeaponData(Player.Weapon(k)).MaxAmmo - Player.Ammo(k) > 1 Then Call Message("Picked up " & (WeaponData(Player.Weapon(k)).MaxAmmo - Player.Ammo(k)) & " " & WeaponData(Player.Weapon(k)).AmmoName(2) & ".", 180)
                If WeaponData(Player.Weapon(k)).MaxAmmo - Player.Ammo(k) = 1 Then Call Message("Picked up " & (WeaponData(Player.Weapon(k)).MaxAmmo - Player.Ammo(k)) & " " & WeaponData(Player.Weapon(k)).AmmoName(1) & ".", 180)
                Weapon(j).Ammo = Weapon(j).Ammo - (WeaponData(Player.Weapon(k)).MaxAmmo - Player.Ammo(k))
                Player.Ammo(k) = WeaponData(Player.Weapon(k)).MaxAmmo
            End If
            frmMain.lblAmmo.Caption = Player.Ammo(1)
        End If
    End If
    Next k
Next j
'>>>shooting
If Player.CoolDownTime <= 0 And Player.Firing = True Then
    Player.CoolDownTime = WeaponData(Player.Weapon(1)).CoolDown
    If WeaponData(Player.Weapon(1)).SemiAuto = False Then
    Call PlayerFire
    End If
Else
    Player.CoolDownTime = Player.CoolDownTime - 1
End If
'>>>melee
If Player.Firing = False And Player.Dead = False And Player.CoolDownTime <= 0 And Player.Reloading = False Then
mousemovedist = Distance(Player.TargetX, Player.TargetY, Player.LastTargetX, Player.LastTargetY)
For j = 1 To EnemyCount
    If Distance(Enemy(j).XCoord, Enemy(j).YCoord, Player.XCoord, Player.YCoord) <= 100 And Enemy(j).VisualContact = True And mousemovedist > 1 Then
        If Distance(Enemy(j).XCoord, Enemy(j).YCoord, Player.TargetX + Player.XCoord, Player.TargetY + Player.YCoord) <= 20 Then
        Enemy(j).HP = Enemy(j).HP - WeaponData(Player.Weapon(1)).MeleeDamage * mousemovedist / 16
        End If
    End If
Next j
End If
'>>>reloading
If Player.Reloading = True And Player.ReloadTime = 0 Then
    If Player.Ammo(1) > WeaponData(Player.Weapon(1)).ClipSize - Player.ClipAmmo(1) Then
        Player.Ammo(1) = Player.Ammo(1) - WeaponData(Player.Weapon(1)).ClipSize + Player.ClipAmmo(1)
        Player.ClipAmmo(1) = WeaponData(Player.Weapon(1)).ClipSize
    Else
        Player.ClipAmmo(1) = Player.ClipAmmo(1) + Player.Ammo(1)
        Player.Ammo(1) = 0
    End If
    frmMain.lblPlayerStatus.Caption = ""
    frmMain.lblAmmo.Caption = Player.Ammo(1)
    frmMain.lblClipAmmo.Caption = Player.ClipAmmo(1)
    Player.Reloading = False
End If
If Player.Reloading = True And Player.ReloadTime > 0 Then
    Player.ReloadTime = Player.ReloadTime - 1
End If
'>>>shield recharge
If Player.RechargeTime > 0 And Player.Dead = False Then Player.RechargeTime = Player.RechargeTime - 1
If Player.RechargeTime = 0 And Player.Shield < PlayerMaxShield * PlayerBonus Then Player.Shield = Player.Shield + 1
'check for victory + update holoscreen text
If GameOver = False Then
For i = 1 To Holocount
If Distance(Holoscreen(i).XCoord, Holoscreen(i).YCoord, Player.XCoord, Player.YCoord) <= 30 Then
    If Holoscreen(i).Goal = True Then
        GameOver = True
        victory = True
        GameOverCountdown = 250
    End If
frmMain.lblInfo.Caption = Holoscreen(i).Text
End If
Next i
End If
'>>>map triggers
For j = 1 To TriggerCount
If Overlap(Player.XCoord, Player.YCoord, 0, 0, Trigger(j).XCoord, Trigger(j).YCoord, Trigger(j).Height, Trigger(j).Width) = True Then Wall(Trigger(j).Link).Type = -1
Next j
'reset tempslopes for next player
'For j = 1 To 5
'tempslope(j).Rise = 0
'tempslope(j).Run = 0
'Next j
'>>>Check for player deaths
    If Player.HP = 0 And Player.Dead = False Then
    With Player
        .Dead = True
        .Firing = False
        .MoveBack = False
        .MoveForward = False
        .TurnLeft = False
        .TurnRight = False
        .StrafeLeft = False
        .StrafeRight = False
    End With
        GameOver = True
        GameOverCountdown = 250
    End If
'==================
'Enemy Calculations
'==================
For j = 1 To EnemyCount
If Enemy(j).HP <= 0 Then
    If Enemy(j).Dead = False Then
        WeaponCount = WeaponCount + 1
        ReDim Preserve Weapon(1 To WeaponCount)
        With Weapon(WeaponCount)
            .Ammo = Enemy(j).Ammo
            .ClipAmmo = Enemy(j).ClipAmmo
            .Deleted = False
            .WeaponType = Enemy(j).Weapon
            .XCoord = Enemy(j).XCoord
            .YCoord = Enemy(j).YCoord
        End With
        Call EnemyDeath(Enemy(j).XCoord, Enemy(j).YCoord)
    End If
    Enemy(j).Dead = True
    GoTo checknext
End If
If Enemy(j).Stance < 3 Then GoTo checkonlyradar
If Enemy(j).VisualContact = True Then
    Enemy(j).TargetX = Player.XCoord
    Enemy(j).TargetY = Player.YCoord
    Select Case Distance(Enemy(j).XCoord, Enemy(j).YCoord, Enemy(j).TargetX, Enemy(j).TargetY)
    Case Is > WeaponData(Enemy(j).Weapon).ShotSpeed * WeaponData(Enemy(j).Weapon).ShotLifespan * 0.85
        Enemy(j).MoveForward = 1
    Case Is < WeaponData(Enemy(j).Weapon).ShotSpeed * WeaponData(Enemy(j).Weapon).ShotLifespan * 0.2
        Enemy(j).MoveForward = -1
    Case Else
        Enemy(j).MoveForward = 0
    End Select
    tempdistance = Distance(Enemy(j).XCoord, Enemy(j).YCoord, Enemy(j).TargetX, Enemy(j).TargetY)
    If tempdistance <= WeaponData(Enemy(j).Weapon).ShotSpeed * WeaponData(Enemy(j).Weapon).ShotLifespan Or (WeaponData(Enemy(j).Weapon).Arc > 0 And tempdistance <= WeaponData(Enemy(j).Weapon).Arc * 450) Then
        If WeaponData(Enemy(j).Weapon).Reloads = True Then
            If Enemy(j).ClipAmmo > 0 Or (WeaponData(Enemy(j).Weapon).Reloads = False And Enemy(j).Ammo > 0) Then
                If Enemy(j).CoolDownTime = 0 Then
                    Call EnemyFire(j)
                    Enemy(j).CoolDownTime = WeaponData(Enemy(j).Weapon).CoolDown
                    If WeaponData(Enemy(j).Weapon).SemiAuto = True Then Enemy(j).CoolDownTime = Enemy(j).CoolDownTime + 10
                Else
                    Enemy(j).CoolDownTime = Enemy(j).CoolDownTime - 1
                End If
            Else
            'reload if enemy is out of ammo
            If Enemy(j).ReloadTime = 0 Then
                If Enemy(j).Reloading = False Then
                Enemy(j).Reloading = True
                Enemy(j).ReloadTime = WeaponData(Enemy(j).Weapon).ReloadTime
                Else
                Enemy(j).Reloading = False
                Enemy(j).ClipAmmo = WeaponData(Enemy(j).Weapon).ClipSize
                Enemy(j).Ammo = Enemy(j).Ammo - WeaponData(Enemy(j).Weapon).ClipSize
                If Enemy(j).Ammo < 0 Then Enemy(j).Ammo = 0
                End If
            Else
                Enemy(j).ReloadTime = Enemy(j).ReloadTime - 1
            End If
            End If
        Else
            If Enemy(j).Ammo > 0 Then
                If Enemy(j).CoolDownTime = 0 Then
                    Call EnemyFire(j)
                    Enemy(j).CoolDownTime = WeaponData(Enemy(j).Weapon).CoolDown
                Else
                    Enemy(j).CoolDownTime = Enemy(j).CoolDownTime - 1
                End If
            End If
        End If
    End If
Else 'if there is no visual contact
    If Distance(Enemy(j).XCoord, Enemy(j).YCoord, Enemy(j).TargetX, Enemy(j).TargetY) > Enemy(j).MoveSpeed * 2 Then Enemy(j).MoveForward = 1
End If
    tempslope(1) = FindLine(Enemy(j).XCoord, Enemy(j).YCoord, Enemy(j).TargetX, Enemy(j).TargetY, Enemy(j).MoveSpeed * Enemy(j).MoveForward)
    'check for wall intersection
MoveX = True
MoveY = True
For k = 1 To WallCount
If Wall(k).Type < 0 Then GoTo checknextwall2
If MoveX = True Then
    If Overlap(Enemy(j).XCoord - 21 + tempslope(1).Run, Enemy(j).YCoord - 21, 40, 40, Wall(k).XCoord, Wall(k).YCoord, Wall(k).Height, Wall(k).Width) = True Then
    MoveX = False
    End If
End If
If MoveY = True Then
    If Overlap(Enemy(j).XCoord - 21, Enemy(j).YCoord + tempslope(1).Rise - 21, 40, 40, Wall(k).XCoord, Wall(k).YCoord, Wall(k).Height, Wall(k).Width) = True Then
    MoveY = False
    End If
End If
checknextwall2:
Next k
'move enemy
If Enemy(j).Stance = 4 Or Enemy(j).VisualContact = True Then
    If MoveX = True Then Enemy(j).XCoord = Enemy(j).XCoord + tempslope(1).Run
    If MoveY = True Then Enemy(j).YCoord = Enemy(j).YCoord + tempslope(1).Rise
End If
'>>>check for visual contact between players
checkonlyradar:
If Player.Dead = False And Enemy(j).Dead = False Then
    Enemy(j).VisualContact = True
Else
    Enemy(j).VisualContact = False
    GoTo checknext
End If
For i = 1 To WallCount
    If Wall(i).Type = 1 Then
        If IntersectHorizLine(Player.XCoord, Player.YCoord, Enemy(j).XCoord, Enemy(j).YCoord, Wall(i).XCoord, Wall(i).YCoord, Wall(i).Width) = True Then Enemy(j).VisualContact = False
        If IntersectVertLine(Player.XCoord, Player.YCoord, Enemy(j).XCoord, Enemy(j).YCoord, Wall(i).XCoord, Wall(i).YCoord, Wall(i).Height) = True Then Enemy(j).VisualContact = False
        If IntersectVertLine(Player.XCoord, Player.YCoord, Enemy(j).XCoord, Enemy(j).YCoord, Wall(i).XCoord + Wall(i).Width, Wall(i).YCoord, Wall(i).Height) = True Then Enemy(j).VisualContact = False
    End If
Next i
If Enemy(j).VisualContact = True Then
    tempslope2 = FindLine(Player.XCoord, Player.YCoord, Enemy(j).XCoord, Enemy(j).YCoord, 100)
    Enemy(j).RadarBlipX = tempslope2.Run
    Enemy(j).RadarBlipY = tempslope2.Rise
    tempslope2 = FindLine(Player.XCoord, Player.YCoord, Enemy(j).XCoord, Enemy(j).YCoord, 0.2 * Distance(Enemy(j).XCoord, Enemy(j).YCoord, Player.XCoord, Player.YCoord))
    Enemy(j).RadarBlipX = tempslope2.Run
    Enemy(j).RadarBlipY = tempslope2.Rise
End If
checknext:
Next j
'==========
'animations
'==========
animations:
For i = ExplosionMin To ExplosionCount
If Explosion(i).Deleted = False Then
    With Explosion(i)
        .Lifespan = Explosion(i).Lifespan - 1
        .XCoord = Explosion(i).XCoord + Explosion(i).XVector
        .YCoord = Explosion(i).YCoord + Explosion(i).YVector
    End With
    If Explosion(i).Lifespan <= 0 Then Explosion(i).Deleted = True
ElseIf i = ExplosionMin Then ExplosionMin = ExplosionMin + 1
If ExplosionMin > ExplosionCount Then
    ExplosionMin = 1
    ExplosionCount = 0
    ReDim Explosion(1 To 1)
End If
End If
Next i
'Game over check
If GameOver = True Then
If GameOverCountdown = 0 Then
Terminated = True
Else
GameOverCountdown = GameOverCountdown - 1
End If
End If
'==================
'Message Management
'==================
If PromptTime > 0 Then
    PromptTime = PromptTime - 1
ElseIf PromptTime = 0 Then
    frmMain.lblPrompt.Caption = ""
End If

If (Player.ClipAmmo(1) = 0 And Player.Ammo(1) = 0 And PromptTime <> -2) Then
    Call Message("Out of ammo.  Press [SHIFT] to switch to " & WeaponData(Player.Weapon(2)).Name & ".", -2)
ElseIf (Player.ClipAmmo(1) > 0 Or Player.Ammo(1) > 0) And PromptTime = -2 Then
    frmMain.lblPrompt.Caption = ""
    PromptTime = 0
End If
If Player.Ammo(1) > 0 And Player.ClipAmmo(1) = 0 And WeaponData(Player.Weapon(1)).Reloads = True And PromptTime <> -1 Then
    Call Message("Right-click to reload.", -1)
ElseIf Player.ClipAmmo(1) > 0 And PromptTime = -1 Then
    frmMain.lblPrompt.Caption = ""
    PromptTime = 0
End If
End Sub
    
Public Sub PlayerFire()

Dim tempslope As Slope, xspread As Double, yspread As Double
Select Case WeaponData(Player.Weapon(1)).Reloads
Case Is = True
    If Player.ClipAmmo(1) = 0 Then Exit Sub
    Player.ClipAmmo(1) = Player.ClipAmmo(1) - 1
Case Is = False
    If Player.Ammo(1) = 0 Then Exit Sub
    Player.Ammo(1) = Player.Ammo(1) - 1
End Select
ShotCount = ShotCount + 1
ReDim Preserve Shot(1 To ShotCount)
xspread = Rnd * (WeaponData(Player.Weapon(1)).ShotSpread * 2) - WeaponData(Player.Weapon(1)).ShotSpread
yspread = Rnd * (WeaponData(Player.Weapon(1)).ShotSpread * 2) - WeaponData(Player.Weapon(1)).ShotSpread
tempslope = FindLine(0, 0, Player.TargetX + xspread, Player.TargetY + yspread, WeaponData(Player.Weapon(1)).ShotSpeed)
With Shot(ShotCount)
    .Alignment = 1
    .Bounce = WeaponData(Player.Weapon(1)).Bounce
    .Damage = WeaponData(Player.Weapon(1)).Damage
    .DamageMultiplier = WeaponData(Player.Weapon(1)).DamageMultiplier
    .Deleted = False
    .Distance = 0
    .ExplodeDamage = WeaponData(Player.Weapon(1)).ExplodeDamage
    .ExplodeRadius = WeaponData(Player.Weapon(1)).ExplodeRadius
    .FrameSteps = WeaponData(Player.Weapon(1)).FrameSteps
    .GraphicsIndex = WeaponData(Player.Weapon(1)).ShotGraphicsIndex
If WeaponData(Player.Weapon(1)).Arc = 0 Then
    .Lifespan = WeaponData(Player.Weapon(1)).ShotLifespan
Else
    .Lifespan = (WeaponData(Player.Weapon(1)).Arc * Distance(0, 0, Player.TargetX, Player.TargetY)) / WeaponData(Player.Weapon(1)).ShotSpeed
End If
    .Radius = WeaponData(Player.Weapon(1)).ShotRadius
    .Visible = WeaponData(Player.Weapon(1)).ShotVisible
    .Speed = WeaponData(Player.Weapon(1)).ShotSpeed
    .XCoord = Player.XCoord
    .YCoord = Player.YCoord
    .XVector = tempslope.Run
    .YVector = tempslope.Rise
End With
frmMain.lblClipAmmo.Caption = Player.ClipAmmo(1)
frmMain.lblAmmo.Caption = Player.Ammo(1)
End Sub

Public Sub Explode(XCoord As Double, YCoord As Double, Damage As Integer, Radius As Integer)
Dim i, targetdistance, tempslope As Slope, tempdistance
For i = 1 To EnemyCount
targetdistance = Distance(XCoord, YCoord, Enemy(i).XCoord, Enemy(i).YCoord)
If Enemy(i).Dead = False And targetdistance < Radius Then
Enemy(i).HP = Enemy(i).HP - Damage * (Radius - targetdistance) / Radius
    Select Case Enemy(i).Stance
        Case Is = 1
        Enemy(i).Stance = 3
        Case Is = 2
        Enemy(i).Stance = 4
    End Select
End If
Next i
targetdistance = Distance(XCoord, YCoord, Player.XCoord, Player.YCoord)
If Player.Dead = False And targetdistance < Radius Then
                If Damage * (Radius - targetdistance) / Radius >= Player.Shield Then
                    Player.HP = Player.HP - Damage * (Radius - targetdistance) / Radius
                    Player.Shield = 0
                    If Player.HP <= 0 Then
                        Player.HP = 0
                    End If
                Else
                    Player.Shield = Player.Shield - Damage * (Radius - targetdistance) / Radius
                End If
                Player.RechargeTime = 250 * EnemyBonus
End If
ExplosionCount = ExplosionCount + 15
ReDim Preserve Explosion(1 To ExplosionCount)
For i = ExplosionCount - 14 To ExplosionCount
With Explosion(i)
    .GraphicsIndex = Int(Rnd * 1 + 1)
    .Lifespan = Int(Rnd * 40 + 110)
    .Deleted = False
    .XCoord = XCoord + Int(Rnd * 80 - 40)
    .YCoord = YCoord + Int(Rnd * 80 - 40)
    tempdistance = Distance(XCoord, YCoord, Explosion(i).XCoord, Explosion(i).YCoord)
tempslope = FindLine(XCoord, YCoord, Explosion(i).XCoord, Explosion(i).YCoord, tempdistance ^ 2 / 3000)
    .XVector = tempslope.Run
    .YVector = tempslope.Rise
End With
Next i
ExplosionCount = ExplosionCount + 6
ReDim Preserve Explosion(1 To ExplosionCount)
For i = ExplosionCount - 5 To ExplosionCount
With Explosion(i)
    .GraphicsIndex = 0
    .Lifespan = Int(Rnd * 4 + 6)
    .Deleted = False
    .XCoord = XCoord + Int(Rnd * 20 - 10)
    .YCoord = YCoord + Int(Rnd * 20 - 10)
    .XVector = 0
    .YVector = 0
End With
Next i

End Sub
Public Function IntersectHorizLine(Point1X, Point1Y, Point2X, Point2Y, LineX, lineY, LineLength) As Boolean
Dim IntersectX
If Point1Y >= lineY And Point2Y >= lineY Then
IntersectHorizLine = False
Exit Function
End If
If Point1Y <= lineY And Point2Y <= lineY Then
IntersectHorizLine = False
Exit Function
End If
If Point1X = Point2X And (Point1X <= LineX Or Point1X >= LineX + LineLength) Then
IntersectHorizLine = False
Exit Function
End If
IntersectX = Point1X + ((lineY - Point1Y) * (Point2X - Point1X)) / (Point2Y - Point1Y)
If IntersectX > LineX And IntersectX < LineX + LineLength Then
IntersectHorizLine = True
Else
IntersectHorizLine = False
End If

End Function

Public Function IntersectVertLine(Point1X, Point1Y, Point2X, Point2Y, LineX, lineY, LineLength) As Boolean
Dim IntersectY
If Point1X >= LineX And Point2X >= LineX Then
IntersectVertLine = False
Exit Function
End If
If Point1X <= LineX And Point2X <= LineX Then
IntersectVertLine = False
Exit Function
End If
If Point1X = LineX Then
IntersectVertLine = False
Exit Function
End If
If Point1Y = Point2Y And (Point1Y <= lineY Or Point1Y >= lineY + LineLength) Then
IntersectVertLine = False
Exit Function
End If
IntersectY = Point1Y + ((LineX - Point1X) * (Point2Y - Point1Y)) / (Point2X - Point1X)
If IntersectY >= lineY And IntersectY <= lineY + LineLength Then
IntersectVertLine = True
Else
IntersectVertLine = False
End If
End Function

Public Sub EnemyFire(Index As Integer)
Dim tempslope As Slope, xspread As Double, yspread As Double
Select Case WeaponData(Enemy(Index).Weapon).Reloads
Case Is = True
    If Enemy(Index).ClipAmmo = 0 Then Exit Sub
    Enemy(Index).ClipAmmo = Enemy(Index).ClipAmmo - 1
Case Is = False
    If Enemy(Index).Ammo = 0 Then Exit Sub
    Enemy(Index).Ammo = Enemy(Index).Ammo - 1
End Select
ShotCount = ShotCount + 1
ReDim Preserve Shot(1 To ShotCount)
xspread = Rnd * (WeaponData(Enemy(Index).Weapon).ShotSpread * 2) - WeaponData(Enemy(Index).Weapon).ShotSpread
yspread = Rnd * (WeaponData(Enemy(Index).Weapon).ShotSpread * 2) - WeaponData(Enemy(Index).Weapon).ShotSpread
tempslope = FindLine(Enemy(Index).XCoord, Enemy(Index).YCoord, Enemy(Index).TargetX + xspread, Enemy(Index).TargetY + yspread, WeaponData(Enemy(Index).Weapon).ShotSpeed)
With Shot(ShotCount)
    .Alignment = 2
    .Bounce = WeaponData(Enemy(Index).Weapon).Bounce
    .Damage = WeaponData(Enemy(Index).Weapon).Damage
    .DamageMultiplier = WeaponData(Enemy(Index).Weapon).DamageMultiplier
    .Deleted = False
    .Distance = 0
    .ExplodeDamage = WeaponData(Enemy(Index).Weapon).ExplodeDamage
    .ExplodeRadius = WeaponData(Enemy(Index).Weapon).ExplodeRadius
    .FrameSteps = WeaponData(Enemy(Index).Weapon).FrameSteps
    .GraphicsIndex = WeaponData(Enemy(Index).Weapon).ShotGraphicsIndex
If WeaponData(Enemy(Index).Weapon).Arc = 0 Then
    .Lifespan = WeaponData(Enemy(Index).Weapon).ShotLifespan
Else
    .Lifespan = (WeaponData(Enemy(Index).Weapon).Arc * Distance(Enemy(Index).XCoord, Enemy(Index).YCoord, Enemy(Index).TargetX, Enemy(Index).TargetY)) / WeaponData(Enemy(Index).Weapon).ShotSpeed
End If
    .Radius = WeaponData(Enemy(Index).Weapon).ShotRadius
    .Visible = WeaponData(Enemy(Index).Weapon).ShotVisible
    .Speed = WeaponData(Enemy(Index).Weapon).ShotSpeed
    .XCoord = Enemy(Index).XCoord
    .YCoord = Enemy(Index).YCoord
    .XVector = tempslope.Run
    .YVector = tempslope.Rise
End With
End Sub

Public Sub Message(Text As String, Duration As Integer)
    PromptTime = Duration
    frmMain.lblPrompt.Caption = Text
End Sub

Public Sub EnemyDeath(XCoord As Double, YCoord As Double)
Dim i, tempslope As Slope
ShotCount = ShotCount + 6
ReDim Preserve Shot(1 To ShotCount)
For i = ShotCount - 5 To ShotCount
With Shot(i)
    .Alignment = 1
    .Bounce = True
    .Damage = 0
    .DamageMultiplier = 0
    .Deleted = False
    .FrameSteps = 1
    .GraphicsIndex = 6
    .Lifespan = 18
    .Radius = 0
    .Speed = 15
    .Visible = True
    .WallPiercing = False
    .XCoord = XCoord - 20 + Int(Rnd * 40 + 1)
    .YCoord = YCoord - 20 + Int(Rnd * 40 + 1)
tempslope = FindLine(XCoord, YCoord, Shot(i).XCoord, Shot(i).YCoord, 5)
    .XVector = tempslope.Run
    .YVector = tempslope.Rise
End With
Next i
End Sub
