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
If Player(1).Dead = True Then GoTo player2
If (GetKeyState(vbKeyK) And KEY_PRESSED) Then
    Player(1).MoveBack = True
    Player(1).Sniping = False
Else
    Player(1).MoveBack = False
End If
If (GetKeyState(vbKeyI) And KEY_PRESSED) Then
    Player(1).MoveForward = True
    Player(1).Sniping = False
Else
    Player(1).MoveForward = False
End If
If (GetKeyState(vbKeyJ) And KEY_PRESSED) Then
    Player(1).TurnLeft = True
Else
    Player(1).TurnLeft = False
End If
If (GetKeyState(vbKeyL) And KEY_PRESSED) Then
    Player(1).TurnRight = True
Else
    Player(1).TurnRight = False
End If
If (GetKeyState(vbKeyU) And KEY_PRESSED) Then
    Player(1).StrafeLeft = True
Else
    Player(1).StrafeLeft = False
End If
If (GetKeyState(vbKeyO) And KEY_PRESSED) Then
    Player(1).StrafeRight = True
Else
    Player(1).StrafeRight = False
End If
If (GetKeyState(vbKeySpace) And KEY_PRESSED) Then
    If rdown(1) = True And Player(1).Switching = False Then ' if fire button is pressed after reload button, switch weapons
        Player(1).Switching = True 'switching disables shooting and reloading until both fire and reload buttons are released.
        Player(1).Reloading = False
        frmMain.lblPlayerStatus(1).Caption = ""
        tempammo = Player(1).Ammo(1)
        tempclipammo = Player(1).ClipAmmo(1)
        tempweapontype = Player(1).Weapon(1)
        Player(1).Ammo(1) = Player(1).Ammo(2)
        Player(1).ClipAmmo(1) = Player(1).ClipAmmo(2)
        Player(1).Weapon(1) = Player(1).Weapon(2)
        Player(1).Ammo(2) = tempammo
        Player(1).ClipAmmo(2) = tempclipammo
        Player(1).Weapon(2) = tempweapontype
        Player(1).Reloading = False
        frmMain.lblClipAmmo(1).Caption = Player(1).ClipAmmo(1)
        frmMain.lblAmmo(1).Caption = Player(1).Ammo(1)
        frmMain.lblWeaponName(1).Caption = WeaponData(Player(1).Weapon(1)).Name
        frmMain.lblWeapon2Name(1).Caption = WeaponData(Player(1).Weapon(2)).Name
    ElseIf Player(1).Switching = False Then
        If WeaponData(Player(1).Weapon(1)).SemiAuto = True And Player(1).Firing = False And Player(1).CoolDownTime <= 0 Then
        If Distance(Player(1).XCoord, Player(1).YCoord, Player(2).XCoord, Player(2).YCoord) > 40 Or WeaponData(Player(1).Weapon(1)).Melee = False Then
            Call Fire(1)
            Player(1).CoolDownTime = WeaponData(Player(1).Weapon(1)).CoolDown
        End If
        End If
            Player(1).Firing = True
            frmMain.lblClipAmmo(1).Caption = Player(1).ClipAmmo(1)
            Player(1).Reloading = False
            frmMain.lblPlayerStatus(1).Caption = ""
    End If
Else
            Player(1).Firing = False
            If (GetKeyState(vbKeyP) And KEY_PRESSED) = False Then Switching = False
End If
'
'
'
If (GetKeyState(vbKeyP) And KEY_PRESSED) And Player(1).Firing = False Then 'keypad + key was pressed
    rdown(1) = True
ElseIf rdown(1) = True And Player(1).Firing = False Then
rdown(1) = False
'If (GetKeyState(vbKeySpace) And KEY_PRESSED) = False Then Player(1).Switching = False
If Player(1).Switching = True Then GoTo DoNotReload1
If Player(1).PickUpItem = False Then
    For i = WeaponMin To WeaponCount
    If Weapon(i).Deleted = False Then
    If Distance(Weapon(i).XCoord, Weapon(i).YCoord, Player(1).XCoord, Player(1).YCoord) <= 20 Then
    If Player(1).Ammo(1) > 0 Or Player(1).ClipAmmo(1) > 0 Then
        WeaponCount = WeaponCount + 1
        ReDim Preserve Weapon(1 To WeaponCount)
        With Weapon(WeaponCount)
            .Ammo = Player(1).Ammo(1)
            .ClipAmmo = Player(1).ClipAmmo(1)
            .ClipSize = WeaponData(Player(1).Weapon(1)).ClipSize
            .Deleted = False
            .DespawnEnabled = True
            .Name = WeaponData(Player(1).Weapon(1)).Name
            .TimeToDespawn = 40
            .WeaponType = Player(1).Weapon(1)
            .XCoord = Player(1).XCoord
            .YCoord = Player(1).YCoord
        End With
    End If
        With Player(1)
            .Ammo(1) = Weapon(i).Ammo
            .ClipAmmo(1) = Weapon(i).ClipAmmo
            .Weapon(1) = Weapon(i).WeaponType
            .Reloading = False
        End With
        frmMain.lblPlayerStatus(1).Caption = ""
        Weapon(i).Deleted = True
        frmMain.lblAmmo(1).Caption = Player(1).Ammo(1)
        frmMain.lblClipAmmo(1).Caption = Player(1).ClipAmmo(1)
        frmMain.lblWeaponName(1).Caption = WeaponData(Player(1).Weapon(1)).Name
        Player(1).PickUpItem = True
    End If
    End If
    Next i
End If
    If Player(1).PickUpItem = False And Player(1).Reloading = False Then
    Player(1).Reloading = True
    frmMain.lblPlayerStatus(1).Caption = "Reloading..."
    Player(1).ReloadTime = WeaponData(Player(1).Weapon(1)).ReloadTime
    End If
    Player(1).Firing = False
Else
    Player(1).PickUpItem = False
DoNotReload1:
If (GetKeyState(vbKeySpace) And KEY_PRESSED) = False Then Player(1).Switching = False
End If
'
'
'
If (GetKeyState(vbKeyC) And KEY_PRESSED) Then
If Player(1).SnipeToggle = False Then
    Player(1).SnipeToggle = True
    If Player(1).Sniping = False Then
        Player(1).Sniping = True
    Else
        Player(1).Sniping = False
    End If
End If
Else
Player(1).SnipeToggle = False
End If
        
'===================
'Player 2 Keystrokes
'===================
player2:
If Player(2).Dead = True Then Exit Sub
If (GetKeyState(vbKeyNumpad5) And KEY_PRESSED) Then
    Player(2).MoveBack = True
    Player(2).Sniping = False
Else
    Player(2).MoveBack = False
End If
If (GetKeyState(vbKeyNumpad8) And KEY_PRESSED) Then
    Player(2).MoveForward = True
    Player(2).Sniping = False
Else
    Player(2).MoveForward = False
End If
If (GetKeyState(vbKeyNumpad4) And KEY_PRESSED) Then
    Player(2).TurnLeft = True
Else
    Player(2).TurnLeft = False
End If
If (GetKeyState(vbKeyNumpad6) And KEY_PRESSED) Then
    Player(2).TurnRight = True
Else
    Player(2).TurnRight = False
End If
If (GetKeyState(vbKeyNumpad7) And KEY_PRESSED) Then
    Player(2).StrafeLeft = True
Else
    Player(2).StrafeLeft = False
End If
If (GetKeyState(vbKeyNumpad9) And KEY_PRESSED) Then
    Player(2).StrafeRight = True
Else
    Player(2).StrafeRight = False
End If
If (GetKeyState(vbKeyNumpad0) And KEY_PRESSED) Then
    If rdown(2) = True And Player(2).Switching = False Then ' if fire is pressed after reload button is pressed then switch weapons
        Player(2).Switching = True
        Player(2).Reloading = False
        frmMain.lblPlayerStatus(2).Caption = ""
        tempammo = Player(2).Ammo(1)
        tempclipammo = Player(2).ClipAmmo(1)
        tempweapontype = Player(2).Weapon(1)
        Player(2).Ammo(1) = Player(2).Ammo(2)
        Player(2).ClipAmmo(1) = Player(2).ClipAmmo(2)
        Player(2).Weapon(1) = Player(2).Weapon(2)
        Player(2).Ammo(2) = tempammo
        Player(2).ClipAmmo(2) = tempclipammo
        Player(2).Weapon(2) = tempweapontype
        frmMain.lblClipAmmo(2).Caption = Player(2).ClipAmmo(1)
        frmMain.lblAmmo(2).Caption = Player(2).Ammo(1)
        frmMain.lblWeaponName(2).Caption = WeaponData(Player(2).Weapon(1)).Name
        frmMain.lblWeapon2Name(2).Caption = WeaponData(Player(2).Weapon(2)).Name
    ElseIf Player(2).Switching = False Then
        If WeaponData(Player(2).Weapon(1)).SemiAuto = True And Player(2).Firing = False And Player(2).CoolDownTime <= 0 Then
        If Distance(Player(2).XCoord, Player(2).YCoord, Player(1).XCoord, Player(1).YCoord) > 40 Or WeaponData(Player(2).Weapon(1)).Melee = False Then
            Call Fire(2)
            Player(2).CoolDownTime = WeaponData(Player(2).Weapon(1)).CoolDown
        End If
        End If
            Player(2).Firing = True
            frmMain.lblClipAmmo(2).Caption = Player(2).ClipAmmo(1)
            Player(2).Reloading = False
            frmMain.lblPlayerStatus(2).Caption = ""
    End If
Else
            Player(2).Firing = False
End If
'
'
'
If (GetKeyState(107) And KEY_PRESSED) And Player(2).Firing = False Then 'keypad + key was pressed
    rdown(2) = True
ElseIf rdown(2) = True And Player(1).Firing = False Then
rdown(2) = False
If Player(2).Switching = True Then GoTo DoNotReload2
If Player(2).PickUpItem = False Then
    For i = WeaponMin To WeaponCount
    If Weapon(i).Deleted = False Then
    If Distance(Weapon(i).XCoord, Weapon(i).YCoord, Player(2).XCoord, Player(2).YCoord) <= 20 Then
    If Player(2).Ammo(1) > 0 Or Player(2).ClipAmmo(1) > 0 Then
        WeaponCount = WeaponCount + 1
        ReDim Preserve Weapon(1 To WeaponCount)
        With Weapon(WeaponCount)
            .Ammo = Player(2).Ammo(1)
            .ClipAmmo = Player(2).ClipAmmo(1)
            .ClipSize = WeaponData(Player(2).Weapon(1)).ClipSize
            .Deleted = False
            .DespawnEnabled = True
            .Name = WeaponData(Player(2).Weapon(1)).Name
            .TimeToDespawn = 40
            .WeaponType = Player(2).Weapon(1)
            .XCoord = Player(2).XCoord
            .YCoord = Player(2).YCoord
        End With
    End If
        With Player(2)
            .Ammo(1) = Weapon(i).Ammo
            .ClipAmmo(1) = Weapon(i).ClipAmmo
            .Weapon(1) = Weapon(i).WeaponType
            .Reloading = False
        End With
        Weapon(i).Deleted = True
        frmMain.lblPlayerStatus(2).Caption = ""
        frmMain.lblAmmo(2).Caption = Player(2).Ammo(1)
        frmMain.lblClipAmmo(2).Caption = Player(2).ClipAmmo(1)
        frmMain.lblWeaponName(2).Caption = WeaponData(Player(2).Weapon(1)).Name
        Player(2).PickUpItem = True
    End If
    End If
    Next i
End If
    If Player(2).PickUpItem = False And Player(2).Reloading = False Then
    Player(2).Reloading = True
    frmMain.lblPlayerStatus(2).Caption = "Reloading..."
    Player(2).ReloadTime = WeaponData(Player(2).Weapon(1)).ReloadTime
    End If
    Player(2).Firing = False
Else
    Player(2).PickUpItem = False
DoNotReload2:
If (GetKeyState(vbKeyNumpad0) And KEY_PRESSED) = False Then Player(2).Switching = False
End If
'
'
If (GetKeyState(vbKeyNumpad1) And KEY_PRESSED) Then
If Player(2).SnipeToggle = False Then
    Player(2).SnipeToggle = True
    If Player(2).Sniping = False Then
        Player(2).Sniping = True
    Else
        Player(2).Sniping = False
    End If
End If
Else
Player(2).SnipeToggle = False
End If

'For i = 5 To 120
'If GetKeyState(i) < 0 Then
'MsgBox "you pressed " & i
'End If
'Next i
End Sub

Public Sub UpdateObjects()
Dim i, j, tempslope(1 To 5) As Slope, tempslope2 As Slope
Dim moveX(1 To 2) As Boolean, moveY(1 To 2) As Boolean
'>>>Update shots and check collisions
For i = ShotMin To ShotCount
If Shot(i).Deleted = False Then
Shot(i).XCoord = Shot(i).XCoord + Shot(i).XVector
Shot(i).YCoord = Shot(i).YCoord + Shot(i).YVector
Shot(i).Distance = Shot(i).Distance + Shot(i).Speed
Shot(i).Lifespan = Shot(i).Lifespan - 1
For j = 1 To 2
If j <> Shot(i).Alignment And Player(j).Dead = False Then
        If Distance(Shot(i).XCoord, Shot(i).YCoord, Player(j).XCoord, Player(j).YCoord) <= Shot(i).Radius + 20 Then
            Shot(i).Deleted = True
            Shot(i).Damage = Shot(i).Damage + Shot(i).Distance * Shot(i).DamageMultiplier
            If Shot(i).Damage >= Player(j).Shield Then
                Player(j).HP = Player(j).HP - Shot(i).Damage + Player(j).Shield
                Player(j).Shield = 0
                If Player(j).HP <= 0 Then
                    Player(j).HP = 0
                End If
            Else
                Player(j).Shield = Player(j).Shield - Shot(i).Damage
            End If
            If Shot(i).Explosive = True Then
            Call Explode(Shot(i).XCoord, Shot(i).YCoord)
            Call CheckExplosion(i)
            End If
        End If
End If
Next j
For j = 1 To WallCount
    If Wall(j).Type = 1 Then
        If Shot(i).Bounce = False And Overlap(Shot(i).XCoord, Shot(i).YCoord, 0, 0, Wall(j).XCoord, Wall(j).YCoord, Wall(j).Height, Wall(j).Width) = True Then
            Shot(i).Deleted = True
            If Shot(i).Explosive = True Then
                Call Explode(Shot(i).XCoord, Shot(i).YCoord)
                Call CheckExplosion(i)
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
If Shot(i).Lifespan = 0 Then
Shot(i).Deleted = True
    If Shot(i).Explosive = True Then
        Call Explode(Shot(i).XCoord, Shot(i).YCoord)
        Call CheckExplosion(i)
        End If
    End If
End If
If Shot(i).Deleted = True And i = ShotMin Then ShotMin = ShotMin + 1
Next i
'>>>Update player positions
For i = 1 To 2
If Player(i).TurnLeft = True And Player(i).Sniping = False Then Player(i).Theta = Player(i).Theta + 0.02
If Player(i).TurnLeft = True And Player(i).Sniping = True Then Player(i).Theta = Player(i).Theta + 0.005
If Player(i).TurnRight = True And Player(i).Sniping = False Then Player(i).Theta = Player(i).Theta - 0.02
If Player(i).TurnRight = True And Player(i).Sniping = True Then Player(i).Theta = Player(i).Theta - 0.005
If Player(i).Theta > 2 Then Player(i).Theta = Player(i).Theta - 2
If Player(i).Theta < 0 Then Player(i).Theta = Player(i).Theta + 2
If Player(i).Sniping = False Then
    Player(i).TargetX = 100 * Sin(Player(i).Theta * PI)
    Player(i).TargetY = 100 * Cos(Player(i).Theta * PI)
Else
    Player(i).TargetX = 200 * Sin(Player(i).Theta * PI)
    Player(i).TargetY = 200 * Cos(Player(i).Theta * PI)
End If
If Player(i).MoveForward = True Then tempslope(1) = FindLine(0, 0, Player(i).TargetX, Player(i).TargetY, 3 + WeaponData(Player(i).Weapon(1)).SpeedBonus)
If Player(i).StrafeLeft = True Then tempslope(2) = FindLine(0, 0, Player(i).TargetY, -Player(i).TargetX, 3 + WeaponData(Player(i).Weapon(1)).SpeedBonus)
If Player(i).StrafeRight = True Then tempslope(3) = FindLine(0, 0, -Player(i).TargetY, Player(i).TargetX, 3 + WeaponData(Player(i).Weapon(1)).SpeedBonus)
If Player(i).MoveBack = True Then tempslope(4) = FindLine(0, 0, -Player(i).TargetX, -Player(i).TargetY, 3 + WeaponData(Player(i).Weapon(1)).SpeedBonus)
With tempslope(5)
    .Rise = tempslope(1).Rise + tempslope(2).Rise + tempslope(3).Rise + tempslope(4).Rise
    .Run = tempslope(1).Run + tempslope(2).Run + tempslope(3).Run + tempslope(4).Run
End With
If tempslope(5).Rise = 0 And tempslope(5).Run = 0 Then GoTo nocheck 'if player is not moving, don't look for collisions
'check to see if new position would overlap with a wall
moveX(i) = True
moveY(i) = True
For j = 1 To WallCount
If moveX(i) = True Then
    If Overlap(Player(i).XCoord - 21 + tempslope(5).Run, Player(i).YCoord - 21, 40, 40, Wall(j).XCoord, Wall(j).YCoord, Wall(j).Height, Wall(j).Width) = True Then
    moveX(i) = False
    End If
End If
If moveY(i) = True Then
    If Overlap(Player(i).XCoord - 21, Player(i).YCoord + tempslope(5).Rise - 21, 40, 40, Wall(j).XCoord, Wall(j).YCoord, Wall(j).Height, Wall(j).Width) = True Then
    moveY(i) = False
    End If
End If
Next j
If moveX(i) = True Then Player(i).XCoord = Player(i).XCoord + tempslope(5).Run
If moveY(i) = True Then Player(i).YCoord = Player(i).YCoord + tempslope(5).Rise
nocheck:
'>>>check for ammo to pick up
For j = WeaponMin To WeaponCount
    For k = 1 To 2
    If Weapon(j).Deleted = False And Weapon(j).WeaponType = Player(i).Weapon(k) And Player(i).Dead = False Then
        If Overlap(Player(i).XCoord, Player(i).YCoord, 0, 0, Weapon(j).XCoord - 20, Weapon(j).YCoord - 20, 40, 40) = True Then
            Player(i).Ammo(k) = Player(i).Ammo(k) + Weapon(j).Ammo + Weapon(j).ClipAmmo
            Weapon(j).Deleted = True
            frmMain.lblAmmo(i).Caption = Player(i).Ammo(1)
        End If
    End If
    Next k
Next j
'>>>shooting
If Distance(Player(1).XCoord, Player(1).YCoord, Player(2).XCoord, Player(2).YCoord) > 40 Or WeaponData(Player(i).Weapon(1)).Melee = False Then
    'shooting
If Player(i).CoolDownTime <= 0 And Player(i).Firing = True Then
    Player(i).CoolDownTime = WeaponData(Player(i).Weapon(1)).CoolDown
    If WeaponData(Player(i).Weapon(1)).SemiAuto = False Then
    Call Fire(i)
    End If
Else
    Player(i).CoolDownTime = Player(i).CoolDownTime - 1
End If
ElseIf WeaponData(Player(i).Weapon(1)).Melee = True Then
    'melee
If Player(i).CoolDownTime <= 0 And Player(i).Firing = True Then
    Player(i).CoolDownTime = 100
    If Player(i).Firing = True Then
        If i = 1 Then
        If Player(2).Dead = False Then
            If WeaponData(Player(1).Weapon(1)).MeleeDamage >= Player(2).Shield Then
                Player(2).HP = Player(2).HP - WeaponData(Player(1).Weapon(1)).MeleeDamage + Player(2).Shield
                Player(2).Shield = 0
                If Player(2).HP <= 0 Then
                    Player(2).HP = 0
                End If
            Else
                Player(2).Shield = Player(2).Shield - WeaponData(Player(1).Weapon(1)).MeleeDamage
            End If
        End If
        Else
        If Player(1).Dead = False Then ' i=2
            If WeaponData(Player(2).Weapon(1)).MeleeDamage >= Player(1).Shield Then
                Player(1).HP = Player(1).HP - WeaponData(Player(2).Weapon(1)).MeleeDamage + Player(1).Shield
                Player(1).Shield = 0
                If Player(1).HP <= 0 Then
                    Player(1).HP = 0
                End If
            Else
                Player(1).Shield = Player(1).Shield - WeaponData(Player(2).Weapon(1)).MeleeDamage
            End If
        End If
        End If
    End If
Else
    Player(i).CoolDownTime = Player(i).CoolDownTime - 1
End If
End If
'>>>reloading
If Player(i).Reloading = True And Player(i).ReloadTime = 0 Then
    If Player(i).Ammo(1) > WeaponData(Player(i).Weapon(1)).ClipSize - Player(i).ClipAmmo(1) Then
        Player(i).Ammo(1) = Player(i).Ammo(1) - WeaponData(Player(i).Weapon(1)).ClipSize + Player(i).ClipAmmo(1)
        Player(i).ClipAmmo(1) = WeaponData(Player(i).Weapon(1)).ClipSize
    Else
        Player(i).ClipAmmo(1) = Player(i).ClipAmmo(1) + Player(i).Ammo(1)
        Player(i).Ammo(1) = 0
    End If
    frmMain.lblPlayerStatus(i).Caption = ""
    frmMain.lblAmmo(i).Caption = Player(i).Ammo(1)
    frmMain.lblClipAmmo(i).Caption = Player(i).ClipAmmo(1)
    Player(i).Reloading = False
End If
If Player(i).Reloading = True And Player(i).ReloadTime > 0 Then
    Player(i).ReloadTime = Player(i).ReloadTime - 1
End If
'reset tempslopes for next player
For j = 1 To 5
tempslope(j).Rise = 0
tempslope(j).Run = 0
Next j
Next i
'>>>Check for player deaths
For i = 1 To 2
    If Player(i).HP = 0 Then
        Call KillPlayer(i)
    End If
Next i
'>>>check for visual contact between players
If Player(1).Dead = False And Player(2).Dead = False Then
VisualContact = True
Else
VisualContact = False
GoTo animations
End If
For i = 1 To WallCount
    If Wall(i).Type = 1 Then
        If IntersectHorizLine(Player(1).XCoord, Player(1).YCoord, Player(2).XCoord, Player(2).YCoord, Wall(i).XCoord, Wall(i).YCoord, Wall(i).Width) = True Then VisualContact = False
        If IntersectVertLine(Player(1).XCoord, Player(1).YCoord, Player(2).XCoord, Player(2).YCoord, Wall(i).XCoord, Wall(i).YCoord, Wall(i).Height) = True Then VisualContact = False
        If IntersectVertLine(Player(1).XCoord, Player(1).YCoord, Player(2).XCoord, Player(2).YCoord, Wall(i).XCoord + Wall(i).Width, Wall(i).YCoord, Wall(i).Height) = True Then VisualContact = False
    End If
Next i
If VisualContact = True Then
If Player(1).Sniping = False Then
    tempslope2 = FindLine(Player(1).XCoord, Player(1).YCoord, Player(2).XCoord, Player(2).YCoord, 100)
Else
    tempslope2 = FindLine(Player(1).XCoord, Player(1).YCoord, Player(2).XCoord, Player(2).YCoord, 200)
End If
Player(1).RadarBlipX = tempslope2.Run
Player(1).RadarBlipY = tempslope2.Rise
If Player(2).Sniping = False Then
    tempslope2 = FindLine(Player(2).XCoord, Player(2).YCoord, Player(1).XCoord, Player(1).YCoord, 100)
Else
    tempslope2 = FindLine(Player(2).XCoord, Player(2).YCoord, Player(1).XCoord, Player(1).YCoord, 200)
End If
Player(2).RadarBlipX = tempslope2.Run
Player(2).RadarBlipY = tempslope2.Rise
End If
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
End If
Next i
End Sub

Public Sub UpdateSpawns()
Dim i, j, RespawnIndex As Integer ', farthestdistance As Double, tempdistance As Double, opponentindex As Integer
If WeaponSpawnCount > 0 Then
    For i = 1 To WeaponSpawnCount
    If WeaponSpawn(i).LastSpawnedIndex > 0 Then
        If Weapon(WeaponSpawn(i).LastSpawnedIndex).Deleted = True Then WeaponSpawn(i).Enabled = True
    Else
        WeaponSpawn(i).Enabled = True
    End If
    If WeaponSpawn(i).Enabled = True Then WeaponSpawn(i).TimeLeft = WeaponSpawn(i).TimeLeft - 1
        If WeaponSpawn(i).TimeLeft <= 0 Then
        WeaponSpawn(i).TimeLeft = WeaponSpawn(i).Frequency
        WeaponCount = WeaponCount + 1
        ReDim Preserve Weapon(1 To WeaponCount)
        If RandomWeapons = True Then WeaponSpawn(i).WeaponType = Int(Rnd * WeaponDataCount + 1)
        With Weapon(WeaponCount)
            .WeaponType = WeaponSpawn(i).WeaponType
            .Ammo = WeaponData(Weapon(WeaponCount).WeaponType).StartAmmo
            .ClipSize = WeaponData(WeaponSpawn(i).WeaponType).ClipSize
            .Deleted = False
            .Name = WeaponData(WeaponSpawn(i).WeaponType).Name
'            .SemiAuto = WeaponData(WeaponSpawn(i).WeaponType).SemiAuto
            .TimeToDespawn = 0
            .DespawnEnabled = False
            .WeaponType = WeaponSpawn(i).WeaponType
            .XCoord = WeaponSpawn(i).XCoord
            .YCoord = WeaponSpawn(i).YCoord
            If Weapon(WeaponCount).Ammo >= Weapon(WeaponCount).ClipSize Then
            .Ammo = Weapon(WeaponCount).Ammo - Weapon(WeaponCount).ClipSize
            .ClipAmmo = Weapon(WeaponCount).ClipSize
            Else
            .Ammo = 0
            .ClipAmmo = WeaponData(WeaponSpawn(i).WeaponType).StartAmmo
            End If
        End With
        WeaponSpawn(i).LastSpawnedIndex = WeaponCount
        WeaponSpawn(i).Enabled = False
        End If
    Next i
End If
If WeaponCount >= WeaponMin Then
    For i = WeaponMin To WeaponCount
        If Weapon(i).Deleted = False And Weapon(i).DespawnEnabled = True Then
            If Weapon(i).TimeToDespawn = 0 Then
            Weapon(i).Deleted = True
            If i = WeaponMin Then WeaponMin = WeaponMin + 1
            Else
            Weapon(i).TimeToDespawn = Weapon(i).TimeToDespawn - 1
            End If
        End If
    Next i
End If
For i = 1 To 2
    If Player(i).Dead = True Then
        If Player(i).RespawnTime > 0 Then
            Player(i).RespawnTime = Player(i).RespawnTime - 1
            frmMain.lblPlayerStatus(i).Caption = "Respawning in " & Player(i).RespawnTime & "..."
        Else
        'find spawnpoint furthest away from opponent
'        If i = 1 Then
'        opponentindex = 2
'        Else
'        opponentindex = 1
'        End If
'            For j = 1 To PlayerSpawnCount
'            tempdistance = Distance(Player(opponentindex).XCoord, Player(opponentindex).YCoord, PlayerSpawn(j).XCoord, PlayerSpawn(j).YCoord)
'            If tempdistance > farthestdistance Then
'            RespawnIndex = j
'            farthestdistance = tempdistance
'            End If
'            Next j
            RespawnIndex = Int(Rnd * PlayerSpawnCount + 1)
            With Player(i)
            .Dead = False
            .XCoord = PlayerSpawn(RespawnIndex).XCoord
            .YCoord = PlayerSpawn(RespawnIndex).YCoord
            End With
            frmMain.lblPlayerStatus(i).Caption = ""
        End If
    Else
        'regenerate player health
        Player(i).HP = Player(i).HP + 2
        If Player(i).HP > 200 Then Player(i).HP = 200
    End If
Next i
If GameOver = True Then
    UnloadCountdown = UnloadCountdown - 1
    If UnloadCountdown = 0 Then Terminated = True
End If
End Sub

Public Sub Fire(Index)
Dim tempslope As Slope, xspread As Double, yspread As Double
If Player(Index).ClipAmmo(1) > 0 Then
Player(Index).ClipAmmo(1) = Player(Index).ClipAmmo(1) - 1
ShotCount = ShotCount + 1
ReDim Preserve Shot(1 To ShotCount)
xspread = Rnd * (WeaponData(Player(Index).Weapon(1)).ShotSpread * 2) - WeaponData(Player(Index).Weapon(1)).ShotSpread
yspread = Rnd * (WeaponData(Player(Index).Weapon(1)).ShotSpread * 2) - WeaponData(Player(Index).Weapon(1)).ShotSpread
tempslope = FindLine(0, 0, Player(Index).TargetX + xspread, Player(Index).TargetY + yspread, WeaponData(Player(Index).Weapon(1)).ShotSpeed)
With Shot(ShotCount)
    .Alignment = Index
    .Bounce = WeaponData(Player(Index).Weapon(1)).Bounce
    .Damage = WeaponData(Player(Index).Weapon(1)).Damage
    .DamageMultiplier = WeaponData(Player(Index).Weapon(1)).DamageMultiplier
    .Deleted = False
    .Distance = 0
    .Explosive = WeaponData(Player(Index).Weapon(1)).ShotExplosive
    .GraphicsIndex = WeaponData(Player(Index).Weapon(1)).ShotGraphicsIndex
    .Lifespan = WeaponData(Player(Index).Weapon(1)).ShotLifespan
    .Radius = WeaponData(Player(Index).Weapon(1)).ShotRadius
    .Visible = WeaponData(Player(Index).Weapon(1)).ShotVisible
    .Speed = WeaponData(Player(Index).Weapon(1)).ShotSpeed
    .XCoord = Player(Index).XCoord
    .YCoord = Player(Index).YCoord
    .XVector = tempslope.Run
    .YVector = tempslope.Rise
End With
frmMain.lblClipAmmo(Index).Caption = Player(Index).ClipAmmo(1)
End If
End Sub

Public Sub KillPlayer(Index)
Dim i, j
For i = 1 To 2
If Player(Index).Ammo(i) > 0 Or Player(Index).ClipAmmo(i) > 0 Then
        WeaponCount = WeaponCount + 1
        ReDim Preserve Weapon(1 To WeaponCount)
        With Weapon(WeaponCount)
            .Ammo = Player(Index).Ammo(i)
            .ClipAmmo = Player(Index).ClipAmmo(i)
            .ClipSize = WeaponData(Player(Index).Weapon(i)).ClipSize
            .Deleted = False
            .Name = WeaponData(Player(Index).Weapon(i)).Name
'            .SemiAuto = WeaponData(Player(Index).weapon(1)).SemiAuto
            .TimeToDespawn = 35
            .DespawnEnabled = True
            .WeaponType = Player(Index).Weapon(i)
            .XCoord = Player(Index).XCoord
            .YCoord = Player(Index).YCoord
        End With
End If
Next i
With Player(Index)
If RandomWeapons = False Then
.Weapon(1) = StartWeapon(1)
.Weapon(2) = StartWeapon(2)
Else
.Weapon(1) = Int(Rnd * WeaponDataCount + 1)
Do Until Player(Index).Weapon(2) <> Player(Index).Weapon(1)
.Weapon(2) = Int(Rnd * WeaponDataCount + 1)
Loop
End If
For i = 1 To 2
If WeaponData(Player(Index).Weapon(i)).StartAmmo - WeaponData(Player(Index).Weapon(i)).ClipSize >= 0 Then
    .Ammo(i) = WeaponData(Player(Index).Weapon(i)).StartAmmo - WeaponData(Player(Index).Weapon(i)).ClipSize
    .ClipAmmo(i) = WeaponData(Player(Index).Weapon(i)).ClipSize
Else
    .Ammo(i) = 0
    .ClipAmmo(i) = WeaponData(Player(Index).Weapon(i)).StartAmmo
End If
Next i
.CoolDownTime = 0
.Deaths = Player(Index).Deaths + 1
.Firing = False
.HP = 200
.Shield = 100
.MoveBack = False
.MoveForward = False
.Reloading = False
.ReloadTime = 0
.StrafeLeft = False
.StrafeRight = False
.TurnLeft = False
.TurnRight = False
.Dead = True
.RespawnTime = PlayerRespawnTime
End With
Player(1).Score = Player(2).Deaths
Player(2).Score = Player(1).Deaths
For i = 1 To 2
If i <> Index Then
Player(i).Shield = 100
End If
frmMain.lblScore(i) = Player(i).Score
If Player(i).Score = ScoreToWin(i) Then
GameOver = True
Select Case ScoreMethod
Case Is = 1
    PlayerRecord(Player(i).RecordIndex).Wins = PlayerRecord(Player(i).RecordIndex).Wins + 1
    PlayerRecord(Player(Index).RecordIndex).Losses = PlayerRecord(Player(Index).RecordIndex).Losses + 1
Case Is = 2
For j = 1 To 2
    PlayerRecord(Player(j).RecordIndex).Wins = PlayerRecord(Player(j).RecordIndex).Wins + Player(j).Score
    PlayerRecord(Player(j).RecordIndex).Losses = PlayerRecord(Player(j).RecordIndex).Losses + Player(j).Deaths
Next j
End Select
Call UpdateLabels
UnloadCountdown = 5
End If
Next i
frmMain.lblAmmo(Index).Caption = Player(Index).Ammo(1)
frmMain.lblClipAmmo(Index).Caption = Player(Index).ClipAmmo(1)
frmMain.lblWeaponName(Index).Caption = WeaponData(Player(Index).Weapon(1)).Name
frmMain.lblWeapon2Name(Index).Caption = WeaponData(Player(Index).Weapon(2)).Name
End Sub

Public Sub ClearPlayers()
Dim i
For i = 1 To 2
With Player(i)
    .Ammo(1) = 0
    .Ammo(2) = 0
    .ClipAmmo(1) = 0
    .ClipAmmo(2) = 0
    .CoolDownTime = 0
    .Dead = False
    .Deaths = 0
    .Firing = False
    .HP = 0
    .MoveBack = False
    .MoveForward = False
    .PickUpItem = False
    .Reloading = False
    .ReloadTime = 0
    .RespawnTime = 0
    .Score = 0
    .Shield = 0
    .SnipeToggle = False
    .Sniping = False
    .StrafeLeft = False
    .StrafeRight = False
    .TargetX = 0
    .TargetY = 0
    .Theta = 0
    .TurnLeft = False
    .TurnRight = False
    .Weapon(1) = 0
    .Weapon(2) = 0
    .XCoord = 0
    .YCoord = 0
End With
Next i
End Sub

Public Sub Explode(x As Double, y As Double)
Dim i, j, tempslope As Slope
ExplosionCount = ExplosionCount + 10
ReDim Preserve Explosion(1 To ExplosionCount)
For i = ExplosionCount - 9 To ExplosionCount
With Explosion(i)
    .GraphicsIndex = Int(Rnd * 1 + 1)
    .Lifespan = Int(Rnd * 100 + 100)
    .Deleted = False
    .XCoord = x + Int(Rnd * 80 - 40)
    .YCoord = y + Int(Rnd * 80 - 40)
tempslope = FindLine(x, y, Explosion(i).XCoord, Explosion(i).YCoord, 0.2)
    .XVector = tempslope.Run
    .YVector = tempslope.Rise
End With
Next i
End Sub

Public Function IntersectHorizLine(Point1X, Point1Y, Point2X, Point2Y, LineX, lineY, LineLength) As Boolean
Dim IntersectX
If Point1Y > lineY And Point2Y > lineY Then
IntersectHorizLine = False
Exit Function
End If
If Point1Y < lineY And Point2Y < lineY Then
IntersectHorizLine = False
Exit Function
End If
If Point1X = Point2X Then
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
If Point1X > LineX And Point2X > LineX Then
IntersectVertLine = False
Exit Function
End If
If Point1X < LineX And Point2X < LineX Then
IntersectVertLine = False
Exit Function
End If
If Point1Y = Point2Y Then
IntersectVertLine = False
Exit Function
End If
IntersectY = Point1Y + ((LineX - Point1X) * (Point2Y - Point1Y)) / (Point2X - Point1X)
If IntersectY > lineY And IntersectY < lineY + LineLength Then
IntersectVertLine = True
Else
IntersectVertLine = False
End If
End Function

Public Sub CheckExplosion(Index)
If Distance(Shot(Index).XCoord, Shot(Index).YCoord, Player(Shot(Index).Alignment).XCoord, Player(Shot(Index).Alignment).YCoord) <= Shot(Index).Radius + 20 And Player(Shot(Index).Alignment).Dead = False Then
    If Shot(Index).Damage >= Player(Shot(Index).Alignment).Shield Then
        Player(Shot(Index).Alignment).HP = Player(Shot(Index).Alignment).HP - Shot(Index).Damage + Player(Shot(Index).Alignment).Shield
        Player(Shot(Index).Alignment).Shield = 0
        If Player(Shot(Index).Alignment).HP <= 0 Then
            Player(Shot(Index).Alignment).HP = 0
            KillPlayer (Shot(Index).Alignment)
        End If
    Else
        Player(Shot(Index).Alignment).Shield = Player(Shot(Index).Alignment).Shield - Shot(Index).Damage
    End If
End If
End Sub
Public Sub UpdateLabels()
Dim i
For i = 0 To 9
If Player(1).RecordIndex = i Then
    frmSetup.lblPlayer1(i).BorderStyle = 1
    frmSetup.lblPlayer2(i).Enabled = False
ElseIf Player(2).RecordIndex <> i Then
    frmSetup.lblPlayer1(i).BorderStyle = 0
    frmSetup.lblPlayer1(i).Enabled = True
    frmSetup.lblPlayer2(i).BorderStyle = 0
    frmSetup.lblPlayer2(i).Enabled = True
End If
If Player(2).RecordIndex = i Then
    frmSetup.lblPlayer2(i).BorderStyle = 1
    frmSetup.lblPlayer1(i).Enabled = False
ElseIf Player(1).RecordIndex <> i Then
    frmSetup.lblPlayer2(i).BorderStyle = 0
    frmSetup.lblPlayer2(i).Enabled = True
    frmSetup.lblPlayer1(i).BorderStyle = 0
    frmSetup.lblPlayer1(i).Enabled = True
End If
If PlayerRecord(i).Name = "" Then PlayerRecord(i).Name = "Player" & i + 1
frmSetup.lblPlayer1(i).Caption = PlayerRecord(i).Name & "     " & PlayerRecord(i).Wins & "/" & PlayerRecord(i).Losses
frmSetup.lblPlayer2(i).Caption = PlayerRecord(i).Name & "     " & PlayerRecord(i).Wins & "/" & PlayerRecord(i).Losses
Next i
End Sub

