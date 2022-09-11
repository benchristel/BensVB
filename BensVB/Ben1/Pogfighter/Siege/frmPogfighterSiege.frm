VERSION 5.00
Begin VB.Form frmPogfighter 
   BackColor       =   &H00000000&
   Caption         =   "Pogfighter"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   Icon            =   "frmPogfighterSiege.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   8265
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCommand 
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   -1000
      Width           =   375
   End
   Begin VB.Timer tmrTime 
      Interval        =   100
      Left            =   660
      Top             =   60
   End
   Begin VB.Image imgNMissile 
      Height          =   480
      Left            =   3660
      Picture         =   "frmPogfighterSiege.frx":08CA
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMExplode 
      Height          =   480
      Index           =   2
      Left            =   3060
      Picture         =   "frmPogfighterSiege.frx":1194
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMExplode 
      Height          =   480
      Index           =   1
      Left            =   2580
      Picture         =   "frmPogfighterSiege.frx":1A5E
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMExplode 
      Height          =   480
      Index           =   0
      Left            =   2100
      Picture         =   "frmPogfighterSiege.frx":2328
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      Height          =   315
      Left            =   60
      Top             =   7860
      Width           =   4635
   End
   Begin VB.Image imgMissile 
      Height          =   480
      Index           =   0
      Left            =   1680
      Picture         =   "frmPogfighterSiege.frx":2BF2
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblPTorpedoCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   5400
      TabIndex        =   3
      Top             =   60
      Width           =   435
   End
   Begin VB.Image imgPTorpedo 
      Height          =   480
      Left            =   4860
      Picture         =   "frmPogfighterSiege.frx":34BC
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgAmmoPod 
      Height          =   480
      Left            =   3840
      Picture         =   "frmPogfighterSiege.frx":3D86
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPod 
      Height          =   480
      Index           =   0
      Left            =   3360
      Picture         =   "frmPogfighterSiege.frx":4650
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode 
      Height          =   480
      Index           =   9
      Left            =   2220
      Picture         =   "frmPogfighterSiege.frx":4F1A
      Top             =   1020
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode 
      Height          =   480
      Index           =   8
      Left            =   1680
      Picture         =   "frmPogfighterSiege.frx":57E4
      Top             =   1020
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode 
      Height          =   480
      Index           =   7
      Left            =   1140
      Picture         =   "frmPogfighterSiege.frx":60AE
      Top             =   1020
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode 
      Height          =   480
      Index           =   6
      Left            =   600
      Picture         =   "frmPogfighterSiege.frx":6978
      Top             =   1020
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode 
      Height          =   480
      Index           =   5
      Left            =   60
      Picture         =   "frmPogfighterSiege.frx":7242
      Top             =   1020
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode 
      Height          =   480
      Index           =   4
      Left            =   2340
      Picture         =   "frmPogfighterSiege.frx":7B0C
      Top             =   540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode 
      Height          =   480
      Index           =   3
      Left            =   1680
      Picture         =   "frmPogfighterSiege.frx":83D6
      Top             =   540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode 
      Height          =   480
      Index           =   2
      Left            =   1140
      Picture         =   "frmPogfighterSiege.frx":8CA0
      Top             =   540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode 
      Height          =   480
      Index           =   1
      Left            =   600
      Picture         =   "frmPogfighterSiege.frx":956A
      Top             =   540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode 
      Height          =   480
      Index           =   0
      Left            =   60
      Picture         =   "frmPogfighterSiege.frx":9E34
      Top             =   540
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   3240
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label lblAmmo 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   4380
      TabIndex        =   2
      Top             =   7500
      Width           =   315
   End
   Begin VB.Shape shpAmmo 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   1320
      Top             =   7500
      Width           =   3015
   End
   Begin VB.Label lblKills 
      BackStyle       =   0  'Transparent
      Caption         =   "Kills: 0 "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   7500
      Width           =   1395
   End
   Begin VB.Shape shpTimer 
      BorderColor     =   &H00808000&
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   60
      Top             =   7860
      Width           =   15
   End
   Begin VB.Label lblOverlay 
      BackStyle       =   0  'Transparent
      Height          =   7275
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4755
   End
   Begin VB.Image imgEnemy 
      Height          =   480
      Index           =   0
      Left            =   1140
      Picture         =   "frmPogfighterSiege.frx":A6FE
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLaser 
      Height          =   480
      Index           =   0
      Left            =   5220
      Picture         =   "frmPogfighterSiege.frx":AFC8
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPlayer 
      Height          =   480
      Left            =   2160
      Picture         =   "frmPogfighterSiege.frx":BF8A
      Top             =   6780
      Width           =   480
   End
End
Attribute VB_Name = "frmPogfighter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TargetY(), Lasers, LaserMin, Enemies, EnemyMin, _
TimeElapsed, TotalTime, Level, EnemyProb, EnemyMove, Kills, Ammo
Dim KillsThisLevel, Score, Totalscore, Money, Pods, PodMin, Bonus, KillsThisLaser, EExplode()
Dim ActiveWeapon, Missiles, MissileMin, MTargetY(), MExplode(), PTAmmo, NMAmmo

Private Sub lbloverlay_Click()
Select Case ActiveWeapon
Case Is = 1
Lasers = Lasers + 1
ReDim Preserve TargetY(1 To Lasers + 1)
Load imgLaser(Lasers)
With imgLaser(Lasers)
.Top = imgPlayer.Top
.Left = imgPlayer.Left
.Visible = True
End With
Ammo = Ammo - 1
If Ammo = -1 Then Ammo = 0
shpAmmo.Width = Ammo / 20 * 3015
Case Is = 2, 3, 4
Missiles = Missiles + 1
ReDim Preserve MTargetY(1 To Missiles + 1)
ReDim Preserve MExplode(1 To Missiles + 1)
Load imgMissile(Missiles)
With imgMissile(Missiles)
.Top = imgPlayer.Top
.Left = imgPlayer.Left
.Visible = True
End With
Select Case ActiveWeapon
Case Is = 2
imgMissile(Missiles).Picture = imgPTorpedo.Picture
'Case Is = 3
'imgMissile(Missiles).Picture = imgNMissile.Picture
'Case Is = 4
'imgMissile(Missiles).Picture = imgWGTorpedo.Picture
Case Is = 3
imgMissile(Missiles).Picture = imgNMissile.Picture
End Select
End Select
End Sub

Private Sub Form_Load()
ReDim TargetY(1 To 1)
ReDim MTargetY(1 To 1)
ReDim MExplode(1 To 1)
Level = InputBox("Hello!  Please enter the level number you want to start on.", "Welcome to Pogfighter")
If Int(Val(Level)) <= 0 Then
Do Until Int(Val(Level)) > 0
Level = InputBox("Please enter a positive number!", "Welcome to Pogfighter")
Loop
End If
Randomize
LaserMin = 1
Lasers = 0
EnemyMin = 1
Enemies = 0
TotalTime = Level * 300
TotalTime = TotalTime + 600
Ammo = 20
Pods = 0
PodMin = 1
Missiles = 0
MissileMin = 1
Kills = 0
ActiveWeapon = 1
EnemyProb = 50 * 0.9 ^ Level
EnemyMove = Level * 2
EnemyMove = EnemyMove + 15
End Sub

Private Sub lblOverlay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
imgPlayer.Left = x - 240
TargetY(Lasers + 1) = y
MTargetY(Missiles + 1) = y
End Sub

Private Sub tmrTime_Timer()
Dim i, x, response
TimeElapsed = TimeElapsed + 1
'''
'''===<<<THIS IS THE CODE FOR WHAT HAPPENS AT THE END OF EACH LEVEL>>>===
'''
shpTimer.Width = TimeElapsed / TotalTime * 4635
If TimeElapsed = TotalTime Then
If Bonus > 0 Then
Bonus = Bonus * 25
Score = Score + Bonus
End If
Money = Money + Score
Totalscore = Totalscore + Score
MsgBox "Level " & Level & " Completed.  You made " & KillsThisLevel & " kills.  Your multiple kill bonus is " & Bonus & ".  Your score for this level is " & Score & " points. ", , "Completion of Level"
response = MsgBox("Your recent bounty has increased your funds to $" & Money & ".  One or more upgrades may be available to you at this time.  Do you want to upgrade your systems?", vbYesNo, "Completion of Level.")
If response = 6 Then ' YES
tmrTime.Enabled = False
Load frmUpgrade
End If
Level = Level + 1
TotalTime = Level * 300
TotalTime = TotalTime + 600
TimeElapsed = 0
shpTimer.Width = 0
KillsThisLevel = 0
Score = 0
Bonus = 0
EnemyProb = 50 * 0.9 ^ Level
EnemyMove = Level * 2
EnemyMove = EnemyMove + 15
shpAmmo.Width = 3015
ReDim TargetY(1 To 1)
ReDim MTargetY(1 To 1)
ReDim MExplode(1 To 1)
For i = EnemyMin To Enemies
Unload imgEnemy(i)
Next i
EnemyMin = 1
Enemies = 0
For i = LaserMin To Lasers
Unload imgLaser(i)
Next i
LaserMin = 1
Lasers = 0
For i = PodMin To Pods
Unload imgPod(i)
Next i
PodMin = 1
Pods = 0
For i = MissileMin To Missiles
Unload imgMissile(i)
Next i
MissileMin = 1
Missiles = 0
Ammo = 20
End If
'''
'''===<<<THIS IS THE CODE FOR LASER MOVES & HIT CALCULATION>>>===
'''
If Lasers < LaserMin Then GoTo 1
For i = LaserMin To Lasers
imgLaser(i).Top = imgLaser(i).Top - 300
If imgLaser(i).Top <= TargetY(i) Then
KillsThisLaser = -1
 For x = EnemyMin To Enemies
If EExplode(x) = False Then
If imgLaser(i).Left > imgEnemy(x).Left - 480 And _
imgLaser(i).Left < imgEnemy(x).Left + 480 Then
If imgLaser(i).Top < imgEnemy(x).Top + 480 And _
imgLaser(i).Top > imgEnemy(x).Top - 480 Then
EExplode(x) = True
Kills = Kills + 1
KillsThisLaser = KillsThisLaser + 1
KillsThisLevel = KillsThisLevel + 1
Score = Score + 15
'Load imgPod(Pods)
'With imgPod(Pods)
'.Top = imgEnemy(x).Top
'.Left = imgEnemy(x).Left
'.Visible = True
'End With
If Kills Mod 10 = 0 Then
Pods = Pods + 1
Load imgPod(Pods)
With imgPod(Pods)
.Top = imgEnemy(x).Top
.Left = imgEnemy(x).Left
.Visible = True
.Picture = imgAmmoPod.Picture
End With
End If
End If
End If
End If
Next x
If imgLaser(i).Visible = True Then Bonus = Bonus + KillsThisLaser
imgLaser(i).Visible = False
End If
If imgLaser(i).Visible = False And i = LaserMin Then
LaserMin = LaserMin + 1
Unload imgLaser(i)
End If
Next i
1:
'''
'''===<<<THIS IS THE CODE FOR ENEMY GENERATION & MOVEMENT>>>===
'''
If Int(Rnd * EnemyProb + 1) = 1 Then
Enemies = Enemies + 1
Load imgEnemy(Enemies)
With imgEnemy(Enemies)
.Top = -1000
.Left = Int(Rnd * 4276)
.Visible = True
End With
ReDim Preserve EExplode(EnemyMin To Enemies)
End If
If Enemies < EnemyMin Then GoTo 2
For i = EnemyMin To Enemies
If EExplode(i) = True Then
Call EnemyExplode(i)
GoTo Next_i
End If
imgEnemy(i).Top = imgEnemy(i).Top + EnemyMove
If imgEnemy(i).Visible = False And i = EnemyMin Then
EnemyMin = EnemyMin + 1
Unload imgEnemy(i)
End If
Next_i:
Next i
2:
'''
'''===<<<THIS IS THE CODE FOR POD MOVEMENT & COLLECTION DETECTION>>>===
'''
If Pods < PodMin Then GoTo 3
For i = PodMin To Pods
imgPod(i).Top = imgPod(i).Top + 60
    If imgPod(i).Left > imgPlayer.Left - 480 And _
    imgPod(i).Left < imgPlayer.Left + 480 Then
        If imgPod(i).Top < imgPlayer.Top + 480 And _
        imgPod(i).Top > imgPlayer.Top - 480 Then
                If imgPod(i).Picture = imgAmmoPod.Picture Then
                Ammo = 20
                shpAmmo.Width = 3015
                End If
        imgPod(i).Visible = False
        End If
    End If
    If imgPod(i).Visible = False And i = PodMin Then
    PodMin = PodMin + 1
Unload imgPod(i)
End If
Next i
3:
'''
'''===<<<THIS IS THE CODE FOR MISSILES>>>===
'''
    If Missiles < MissileMin Then GoTo 4
        For i = MissileMin To Missiles
            Select Case imgMissile(i).Picture
            Case Is = imgPTorpedo.Picture
                If MExplode(i) = False Then imgMissile(i).Top = imgMissile(i).Top - 300
                If imgMissile(i).Top <= MTargetY(i) Then
                imgMissile(i).Top = MTargetY(i)
                MExplode(i) = True
                GoTo MissileExplode
                End If
            Case Is = imgNMissile.Picture
                If imgMissile(i).Top > -1000 Then
                imgMissile(i).Top = imgMissile(i).Top - 300
                Else
                imgMissile(i).Visible = False
                End If
            End Select
MissileExplode:
            For x = EnemyMin To Enemies
                If EExplode(x) = False And imgMissile(i).Visible = True Then
                    If imgMissile(i).Left > imgEnemy(x).Left - 1020 And _
                    imgMissile(i).Left < imgEnemy(x).Left + 1020 Then
                        If imgMissile(i).Top < imgEnemy(x).Top + 1020 And _
                        imgMissile(i).Top > imgEnemy(x).Top - 1020 Then
                        EExplode(x) = True
                        Kills = Kills + 1
                        KillsThisLevel = KillsThisLevel + 1
                        Score = Score + 15
                        'Load imgPod(Pods)
                        'With imgPod(Pods)
                        '.Top = imgEnemy(x).Top
                        '.Left = imgEnemy(x).Left
                        '.Visible = True
                        'End With
                            If Kills Mod 10 = 0 Then
                            Pods = Pods + 1
                            Load imgPod(Pods)
                                With imgPod(Pods)
                                .Top = imgEnemy(x).Top
                                .Left = imgEnemy(x).Left
                                .Visible = True
                                .Picture = imgAmmoPod.Picture
                                End With
                            End If
                        End If
                    End If
                End If
            Next x
        If imgMissile(i).Picture <> imgNMissile.Picture Then
        Select Case imgMissile(i).Picture
        Case Is = imgPTorpedo.Picture
        imgMissile(i).Picture = imgMExplode(0).Picture
        Case Is = imgMExplode(0).Picture
        imgMissile(i).Picture = imgMExplode(2).Picture
        Case Is = imgMExplode(1).Picture
        imgMissile(i).Picture = imgMExplode(3).Picture
        Case Is = imgMExplode(2).Picture
        imgMissile(i).Visible = False
        End Select
        End If
        If imgMissile(i).Visible = False And i = MissileMin Then
        MissileMin = MissileMin + 1
        Unload imgMissile(i)
        End If
        Next i

4:
lblKills.Caption = "Kills: " & Kills
lblAmmo.Caption = Ammo
End Sub

Private Sub EnemyExplode(EnemyNo)
Select Case imgEnemy(EnemyNo).Picture
Case Is = imgEExplode(1).Picture
imgEnemy(EnemyNo).Picture = imgEExplode(2).Picture
Case Is = imgEExplode(2).Picture
imgEnemy(EnemyNo).Picture = imgEExplode(3).Picture
Case Is = imgEExplode(3).Picture
imgEnemy(EnemyNo).Picture = imgEExplode(4).Picture
Case Is = imgEExplode(4).Picture
imgEnemy(EnemyNo).Picture = imgEExplode(5).Picture
Case Is = imgEExplode(5).Picture
imgEnemy(EnemyNo).Picture = imgEExplode(6).Picture
Case Is = imgEExplode(6).Picture
imgEnemy(EnemyNo).Picture = imgEExplode(7).Picture
Case Is = imgEExplode(7).Picture
imgEnemy(EnemyNo).Picture = imgEExplode(8).Picture
Case Is = imgEExplode(8).Picture
imgEnemy(EnemyNo).Picture = imgEExplode(9).Picture
Case Is = imgEExplode(9).Picture
imgEnemy(EnemyNo).Visible = False
Case Is = imgEnemy(0).Picture
imgEnemy(EnemyNo).Picture = imgEExplode(1).Picture
End Select
End Sub

Private Sub txtCommand_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = 192
ActiveWeapon = 1
Case Is = 49
ActiveWeapon = 2
Case Is = 50
ActiveWeapon = 3
End Select
End Sub
