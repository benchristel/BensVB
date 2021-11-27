Attribute VB_Name = "modFunctions"
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

Public Function FindLine(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, Length) As Slope
Dim proportion, pointdistance
pointdistance = Distance((X1), (Y1), (X2), (Y2))
If pointdistance = 0 Then
FindLine.Rise = 0
FindLine.Run = 0
Exit Function
End If
proportion = Length / pointdistance
FindLine.Run = proportion * (X2 - X1)
FindLine.Rise = proportion * (Y2 - Y1)
End Function

Public Sub UpdateObjects()
Dim i, j, tempdamage As Integer, random
'==================
'===Update Shots===
'==================
For i = ShotMin To ShotCount
    Shot(i).YVector = Shot(i).YVector + 0.1
    Shot(i).XCoord = Shot(i).XCoord + Shot(i).XVector
    Shot(i).YCoord = Shot(i).YCoord + Shot(i).YVector
    For j = EnemyMin To EnemyCount
    If Shot(i).Deleted = False And Enemy(j).Dead = False And Overlap(Shot(i).XCoord, Shot(i).YCoord, 0, 0, Enemy(j).XCoord - 20, Enemy(j).YCoord - 40, 40, 40) = True Then
        tempdamage = Shot(i).Damage
        Shot(i).Damage = Shot(i).Damage - Enemy(j).Health
        Enemy(j).Health = Enemy(j).Health - tempdamage
        If Shot(i).Damage <= 0 Then Shot(i).Deleted = True
        If Enemy(j).Health <= 0 Then
            Enemy(j).Dead = True
            Player.Kills = Player.Kills + 1
            Enemy(j).DeadYVector = -6
            Player.Points = Player.Points + Enemy(j).Points
        End If
    End If
    Next j
    If Shot(i).YCoord > 480 Then Shot(i).Deleted = True
    If Shot(i).Deleted = True And i = ShotMin Then ShotMin = ShotMin + 1
Next i
'====================
'===Update Enemies===
'====================
For i = EnemyMin To EnemyCount
    If Enemy(i).Dead = False Then
        Enemy(i).XCoord = Enemy(i).XCoord + Enemy(i).Speed
        If Enemy(i).XCoord > 600 Then
            Enemy(i).Dead = True
            Player.Health = Player.Health - Enemy(i).Damage
            Player.Damage = Player.Damage - Enemy(i).PowerDrain
        End If
    Else
        Enemy(i).DeadYVector = Enemy(i).DeadYVector + 0.2
        Enemy(i).XCoord = Enemy(i).XCoord - Enemy(i).Speed
        Enemy(i).YCoord = Enemy(i).YCoord + Enemy(i).DeadYVector
        If Enemy(i).YCoord > 600 Then Enemy(i).Deleted = True
        If Enemy(i).Deleted = True And i = EnemyMin Then EnemyMin = EnemyMin + 1
    End If
Next i
'===================
'===Spawn Enemies===
'===================
random = Rnd * 1000
If random <= SpawnThreshold And random <= 25 Then
    Call SpawnEnemy
    'If SpawnThreshold < 10 Then
        SpawnThreshold = SpawnThreshold + 0.03
    'End If
End If
'===================
'===Spawn Friends===
'===================
If Rnd * 1000 <= 0.1 Then
    Call SpawnFriend
End If
'===========================
'===Update Points Counter===
'===========================
If Player.Health <= 0 Then
Terminated = True
MsgBox "You have been defeated.", , "System Error"
End If
If frmMain.lblScore.Caption <> Player.Points Then frmMain.lblScore.Caption = Player.Points
If frmMain.lblHealth.Caption <> Player.Health Then frmMain.lblHealth.Caption = Player.Health
If frmMain.lblKills.Caption <> Player.Kills Then frmMain.lblKills.Caption = Player.Kills
End Sub

Public Sub SpawnEnemy()
Dim enemytype
EnemyCount = EnemyCount + 1
ReDim Preserve Enemy(1 To EnemyCount)
enemytype = Rnd * SpawnThreshold
Select Case enemytype
Case Is > 7
enemytype = 3
Case Is > 6
enemytype = 2
Case Is > 5.6
enemytype = 3
Case Is > 5
enemytype = 6
Case Is > 3.6
enemytype = 1
Case Is > 2.6
enemytype = 3
Case Is > 1.8
enemytype = 2
Case Else
enemytype = 1
End Select
With Enemy(EnemyCount)
    .Damage = EnemyData(enemytype).Damage
    .Dead = False
    .DeadYVector = 0
    .Deleted = False
    .GraphicsDC = EnemyData(enemytype).GraphicsDC
    .Health = EnemyData(enemytype).Health
    .Name = EnemyData(enemytype).Name
    .Speed = EnemyData(enemytype).Speed
    .Points = EnemyData(enemytype).Points
    .PowerDrain = EnemyData(enemytype).PowerDrain
    .XCoord = -40
    .YCoord = 400 - EnemyData(enemytype).Altitude
End With
End Sub

Public Sub PlayerFire()
Dim tempslope As Slope
ShotCount = ShotCount + 1
ReDim Preserve Shot(1 To ShotCount)
tempslope = FindLine(Player.DrawX, Player.DrawY, Player.TargetX, Player.TargetY, Distance((Player.TargetX), (Player.TargetY), (Player.DrawX), (Player.DrawY)) / 50)
With Shot(ShotCount)
    .Damage = Player.Damage
    .Deleted = False
    .XCoord = 630
    .YCoord = 150
    .XVector = tempslope.Run
    .YVector = tempslope.Rise
End With
End Sub

Public Sub BlitObjects()
Dim i
'blit background to buffer
BitBlt BackBuffDC, 0, 0, 640, 400, BackgroundDC, 0, 0, vbSrcCopy
'blit shots to buffer
For i = ShotMin To ShotCount
BitBlt BackBuffDC, Shot(i).XCoord - 2.5, Shot(i).YCoord - 2.5, 5, 5, ShotMaskDC, 0, 0, vbSrcAnd
BitBlt BackBuffDC, Shot(i).XCoord - 2.5, Shot(i).YCoord - 2.5, 5, 5, ShotDC, 0, 0, vbSrcPaint
Next i
'blit enemies to buffer
For i = EnemyMin To EnemyCount
BitBlt BackBuffDC, Enemy(i).XCoord - 20, Enemy(i).YCoord - 40, 40, 40, EnemyMaskDC(Enemy(i).GraphicsDC), 0, 0, vbSrcAnd
BitBlt BackBuffDC, Enemy(i).XCoord - 20, Enemy(i).YCoord - 40, 40, 40, EnemyDC(Enemy(i).GraphicsDC), 0, 0, vbSrcPaint
Next i
'blit buffer to screen
BitBlt frmMain.picScreen.hdc, 0, 0, 640, 400, BackBuffDC, 0, 0, vbSrcCopy
End Sub

Public Sub SpawnFriend()
Dim enemytype As Integer
EnemyCount = EnemyCount + 1
ReDim Preserve Enemy(1 To EnemyCount)
enemytype = Int(Rnd * 100 + 1)
Select Case enemytype
Case Is < 30 + (Player.Damage - 5) * 4
enemytype = 4
Case Else
enemytype = 5
End Select
With Enemy(EnemyCount)
    .Damage = EnemyData(enemytype).Damage
    .Dead = False
    .DeadYVector = 0
    .Deleted = False
    .GraphicsDC = EnemyData(enemytype).GraphicsDC
    .Health = EnemyData(enemytype).Health
    .Name = EnemyData(enemytype).Name
    .Speed = EnemyData(enemytype).Speed
    .Points = EnemyData(enemytype).Points
    .PowerDrain = EnemyData(enemytype).PowerDrain
    .XCoord = -40
    .YCoord = 400 - EnemyData(enemytype).Altitude
End With
End Sub
