VERSION 5.00
Begin VB.Form frmPogfighter 
   BackColor       =   &H80000007&
   Caption         =   "Pogfighter: Anadon"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   Icon            =   "frmPogfighter.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmPogfighter.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   4365
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrEnemyFire 
      Interval        =   700
      Left            =   120
      Top             =   660
   End
   Begin VB.Timer tmrMain 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.Image imgLaserBlast 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   660
      Picture         =   "frmPogfighter.frx":1194
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDFighterExp3 
      Height          =   480
      Left            =   4500
      Picture         =   "frmPogfighter.frx":1A5E
      Top             =   1860
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDFighterExp2 
      Height          =   480
      Left            =   3960
      Picture         =   "frmPogfighter.frx":2728
      Top             =   1860
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDFighterExp1 
      Height          =   480
      Left            =   3420
      Picture         =   "frmPogfighter.frx":33F2
      Top             =   1860
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDFighterFire 
      Height          =   480
      Left            =   2940
      Picture         =   "frmPogfighter.frx":40BC
      Top             =   1860
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDFighter 
      Height          =   480
      Left            =   2460
      Picture         =   "frmPogfighter.frx":4D86
      Top             =   1860
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgShield 
      Height          =   480
      Left            =   1860
      Picture         =   "frmPogfighter.frx":5A50
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblKills 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Destroyed-0"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   18
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   675
      Left            =   9420
      TabIndex        =   1
      Top             =   540
      Width           =   5715
   End
   Begin VB.Image imgFHit4 
      Height          =   480
      Left            =   3480
      Picture         =   "frmPogfighter.frx":671A
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFHit3 
      Height          =   480
      Left            =   2940
      Picture         =   "frmPogfighter.frx":73E4
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFHit2 
      Height          =   480
      Left            =   2400
      Picture         =   "frmPogfighter.frx":80AE
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFHit1 
      Height          =   480
      Left            =   1860
      Picture         =   "frmPogfighter.frx":8D78
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgELaserBlast 
      Height          =   480
      Index           =   0
      Left            =   660
      Picture         =   "frmPogfighter.frx":9A42
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEnemyFire 
      Height          =   480
      Left            =   2880
      Picture         =   "frmPogfighter.frx":A70C
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode5 
      Height          =   480
      Left            =   4020
      Picture         =   "frmPogfighter.frx":B3D6
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode4 
      Height          =   480
      Left            =   3480
      Picture         =   "frmPogfighter.frx":BCA0
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode3 
      Height          =   480
      Left            =   2940
      Picture         =   "frmPogfighter.frx":C56A
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode2 
      Height          =   480
      Left            =   2400
      Picture         =   "frmPogfighter.frx":CE34
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode1 
      Height          =   480
      Left            =   1800
      Picture         =   "frmPogfighter.frx":D6FE
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblEnergy 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Energy-5000"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   18
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   675
      Left            =   9420
      TabIndex        =   0
      Top             =   120
      Width           =   5715
   End
   Begin VB.Image imgEnemy 
      Height          =   480
      Index           =   0
      Left            =   1200
      Picture         =   "frmPogfighter.frx":DFC8
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPlayerFire2 
      Height          =   480
      Left            =   2160
      Picture         =   "frmPogfighter.frx":E892
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPlayerFire1 
      Height          =   480
      Left            =   1680
      Picture         =   "frmPogfighter.frx":F15C
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPlayer 
      Height          =   480
      Left            =   1200
      Picture         =   "frmPogfighter.frx":FA26
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmPogfighter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim firing As Boolean, FLasers, FlaserMin, FlaserCount, ELasers, ElaserMin, ElaserCount
Dim PlayerX, PlayerY
Dim Level
Dim Enemies, Enemymin, EDamage(1 To 10), EnNo, Enemycount, EFire As Boolean, LatMove(1 To 5), EnemiesOff
Dim Energy, Kills
Dim Hit As Boolean, Shield As Boolean
Dim score

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 83 Then
Shield = True
frmPogfighter.MouseIcon = imgShield.Picture
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 83 Then
Shield = False
frmPogfighter.MouseIcon = imgPlayer.Picture
End If
End Sub

Private Sub Form_Load()
Dim i
'initiate variables
FlaserMin = 1
ElaserMin = 1
Level = 5
Energy = 1500
Randomize
For i = 1 To 10
EDamage(i) = 0
Next i
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this fires the laser cannon
firing = True
Load imgLaserBlast(FLasers + 1)
FLasers = FLasers + 1
FlaserCount = FlaserCount + 1
imgLaserBlast(FLasers).Left = PlayerX - 240
imgLaserBlast(FLasers).Top = PlayerY - 240
imgLaserBlast(FLasers).Visible = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PlayerX = X
PlayerY = Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
firing = False
End Sub

Private Sub tmrEnemyFire_Timer()
For EnNo = 1 To Enemies
EFire = True
Next EnNo
End Sub

Private Sub tmrMain_Timer()
Dim LasNo
If Shield = False Then
    If firing = True And Hit = False Then
    'set right picture
    If frmPogfighter.MouseIcon = imgPlayer.Picture Or _
        frmPogfighter.MouseIcon = imgPlayerFire2.Picture Then
        frmPogfighter.MouseIcon = imgPlayerFire1.Picture
        Else
        frmPogfighter.MouseIcon = imgPlayerFire2.Picture
    End If
    Load imgLaserBlast(FLasers + 1)
    FLasers = FLasers + 1 'raises the maximum no. of lasers
    FlaserCount = FlaserCount + 1 'counts laser blasts
    imgLaserBlast(FLasers).Left = PlayerX - 240
    imgLaserBlast(FLasers).Top = PlayerY - 240
    imgLaserBlast(FLasers).Visible = True
    Energy = Energy - 5
    End If
End If
If firing = False And Hit = False And Shield = False Then frmPogfighter.MouseIcon = imgPlayer.Picture

For LasNo = FlaserMin To FLasers 'check if laser has hit anything
    imgLaserBlast(LasNo).Top = imgLaserBlast(LasNo).Top - 400
        If imgLaserBlast(LasNo).Top < -735 And LasNo = FlaserMin Then
        Unload imgLaserBlast(LasNo)
        If LasNo < FLasers Then LasNo = LasNo + 1
        FlaserMin = FlaserMin + 1
        FlaserCount = FlaserCount - 1
        If FlaserCount = 0 Then GoTo 1 'if the last laser has been unloaded
        If FlaserMin > FLasers Then FlaserMin = FLasers
        End If
    For EnNo = 1 To Enemies
        If imgLaserBlast(LasNo).Top > imgEnemy(EnNo).Top And imgLaserBlast(LasNo).Top < imgEnemy(EnNo).Top + 480 Then
            If imgLaserBlast(LasNo).Left > imgEnemy(EnNo).Left - 480 And imgLaserBlast(LasNo).Left < imgEnemy(EnNo).Left + 480 Then
            EDamage(EnNo) = EDamage(EnNo) + 1
            End If
        End If
    Next EnNo
Next LasNo
1:
If Enemycount = 0 Then GoTo 2 'avoid error
2:
If Int(Rnd * 100 + 1) <= Level And Enemycount < 3 Then
Load imgEnemy(Enemies + 1)
Enemies = Enemies + 1
Enemycount = Enemycount + 1
With imgEnemy(Enemies)
    .Top = -735
    .Left = Int(Rnd * 15000)
    .Visible = True
End With
LatMove(Enemies) = "Left"
End If
Energy = Energy - 1
'
'
'
For EnNo = 1 To Enemies
Select Case LatMove(EnNo)
Case Is = "Left"
imgEnemy(EnNo).Left = imgEnemy(EnNo).Left - 100
Case Is = "Right"
imgEnemy(EnNo).Left = imgEnemy(EnNo).Left + 100
End Select
Select Case imgEnemy(EnNo).Left
Case Is <= 0
LatMove(EnNo) = "Right"
Case Is >= 15000
LatMove(EnNo) = "Left"
End Select
If EFire = True And EDamage(EnNo) < 10 Then
imgEnemy(EnNo).Picture = imgEnemyFire.Picture
ELasers = ELasers + 1 'raises the maximum no. of lasers
ElaserCount = ElaserCount + 1 'counts laser blasts
Load imgELaserBlast(ELasers)
imgELaserBlast(ELasers).Left = imgEnemy(EnNo).Left
imgELaserBlast(ELasers).Top = imgEnemy(EnNo).Top
imgELaserBlast(ELasers).Visible = True
End If
If EFire = False And EDamage(EnNo) < 10 Then imgEnemy(EnNo).Picture = imgEnemy(0).Picture
imgEnemy(EnNo).Top = imgEnemy(EnNo).Top + 100
If Enemycount = 0 Then GoTo 4
    If EDamage(EnNo) > 10 And EnNo >= 1 Then
    Call EDestroy
    End If
    If imgEnemy(EnNo).Top > 12000 And EnemiesOff < 20 Then
    With imgEnemy(EnNo)
        .Top = -735
    .Left = Int(Rnd * 15000)
    End With
    EnemiesOff = EnemiesOff + 1
    End If
Next EnNo
EFire = False
For LasNo = ElaserMin To ELasers 'check if laser has hit anything and move laser
    imgELaserBlast(LasNo).Top = imgELaserBlast(LasNo).Top + 400
        If imgELaserBlast(LasNo).Top > 12000 And LasNo = ElaserMin Then
        Unload imgELaserBlast(LasNo)
        If LasNo < ELasers Then LasNo = LasNo + 1
        ElaserMin = ElaserMin + 1
        ElaserCount = ElaserCount - 1
        If ElaserCount = 0 Then GoTo 3 'if the last laser has been unloaded
        If ElaserMin > ELasers Then ElaserMin = ELasers
        End If
'    For EnNo = 1 To Enemies
        If imgELaserBlast(LasNo).Top > PlayerY And imgELaserBlast(LasNo).Top < PlayerY + 480 And Shield = False Then
            If imgELaserBlast(LasNo).Left > PlayerX - 480 And imgELaserBlast(LasNo).Left < PlayerX + 480 Then
            Hit = True
            Energy = Energy - 1000
            End If
        End If
'    Next EnNo
3:
Next LasNo
4:
If Hit = True Then Call FHit
    If EnemiesOff >= 20 Then
        If Kills >= 18 Then
        score = 1000
        score = score * Kills / 20
        score = score * Energy / 100
        score = score * Kills * 10 / FLasers
        lblEnergy.Caption = Energy
        MsgBox "Your mission was successful." & vbLf _
            & "You destroyed " & Kills & " out of 20 targets" & vbLf _
            & "You finished the game with " & Energy & " E.U." & vbLf _
            & "You scored a total of " & Kills * 10 & " hits with " & FLasers & " shots." & vbLf _
            & "This gives you a score of " & Int(score) & " points.", , "Victory!"
            End
        Else
            MsgBox "You did not satisfy the mission victory requirements.  You lose.", , "Uh-oh!"
            End
        End If
    End If
    If Energy < 0 Then
lblEnergy.Caption = "System Failure!"
lblEnergy.ForeColor = RGB(255, 0, 0)
Else
lblEnergy.Caption = "Energy-" & Energy
End If
End Sub

Public Sub EDestroy()
Select Case imgEnemy(EnNo).Picture
    Case Is = imgEExplode1.Picture
    imgEnemy(EnNo).Picture = imgEExplode2.Picture
    Case Is = imgEExplode2.Picture
    imgEnemy(EnNo).Picture = imgEExplode3.Picture
    Case Is = imgEExplode3.Picture
    imgEnemy(EnNo).Picture = imgEExplode4.Picture
    Case Is = imgEExplode4.Picture
    imgEnemy(EnNo).Picture = imgEExplode5.Picture
    Case Is = imgEExplode5.Picture
    'reinitiate enemy fighter
    If EnemiesOff < 20 Then
    imgEnemy(EnNo).Top = -735
    EnemiesOff = EnemiesOff + 1
    imgEnemy(EnNo).Left = Int(Rnd * 15000)
    imgEnemy(EnNo).Picture = imgEnemy(0).Picture
    EDamage(EnNo) = 0
    Energy = Energy + 500
    Kills = Kills + 1
    lblKills.Caption = "Destroyed-" & Kills
    Else
    imgEnemy(EnNo).Top = 12000
    End If
    Case Else
    imgEnemy(EnNo).Picture = imgEExplode1.Picture
    
End Select
End Sub

Public Sub FHit()
    Select Case frmPogfighter.MouseIcon
    Case Is = imgFHit1.Picture
    frmPogfighter.MouseIcon = imgFHit2.Picture
    Case Is = imgFHit2.Picture
    frmPogfighter.MouseIcon = imgFHit3.Picture
    Case Is = imgFHit3.Picture
    frmPogfighter.MouseIcon = imgFHit4.Picture
    Case Is = imgFHit4.Picture
    frmPogfighter.MouseIcon = imgPlayer.Picture
    Hit = False
    Case Else
    frmPogfighter.MouseIcon = imgFHit1.Picture
    End Select
End Sub
