VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   15285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   1019
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1280
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picScreen 
      Height          =   14655
      Left            =   120
      ScaleHeight     =   973
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   973
      TabIndex        =   0
      Top             =   120
      Width           =   14655
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   14880
      Width           =   18975
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "||| Exit |||"
      BeginProperty Font 
         Name            =   "Weltron Urban"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   615
      Left            =   15720
      TabIndex        =   7
      Top             =   14160
      Width           =   2535
   End
   Begin VB.Label lblPlayerStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   14880
      TabIndex        =   6
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   6135
      Left            =   14880
      TabIndex        =   5
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label lblClipAmmo 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Left            =   15960
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblAmmo 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Left            =   14880
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblWeaponName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   375
      Index           =   2
      Left            =   14880
      TabIndex        =   2
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label lblWeaponName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Index           =   1
      Left            =   14880
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(10, 30, 10)
Randomize
'================
'Set graphics dcs
'================
PlayerGraphicsDC = GenerateDC(App.Path & "\Graphics\Player1.bmp")
PlayerMaskDC = GenerateDC(App.Path & "\Graphics\PlayerMask1.bmp")
For i = 1 To 6
EnemyDC(i) = GenerateDC(App.Path & "\Graphics\Enemy" & i & ".bmp")
EnemyMaskDC(i) = GenerateDC(App.Path & "\Graphics\EnemyMask1.bmp")
Next i
'BackBuffDC = GenerateDC(App.Path & "\Graphics\Background1.bmp")
'BackgroundDC = GenerateDC(App.Path & "\Graphics\Background1.bmp")
For i = 1 To 12
WeaponGraphicsDC(i) = GenerateDC(App.Path & "\Graphics\Item" & i & ".bmp")
WeaponMaskDC(i) = GenerateDC(App.Path & "\Graphics\ItemMask" & i & ".bmp")
Next i
SightDC = GenerateDC(App.Path & "\Graphics\Sight.bmp")
SightMaskDC = GenerateDC(App.Path & "\Graphics\SightMask.bmp")
For i = 1 To 2
RadarBlipDC(i) = GenerateDC(App.Path & "\Graphics\Blip" & i & ".bmp")
Next i
RadarBlipMaskDC = GenerateDC(App.Path & "\Graphics\BlipMask.bmp")
For i = 1 To 6
ShotGraphicsDC(i) = GenerateDC(App.Path & "\Graphics\Shot" & i & ".bmp")
ShotMaskDC(i) = GenerateDC(App.Path & "\Graphics\ShotMask" & i & ".bmp")
Next i
For i = 0 To 1
ExplosionGraphicsDC(i) = GenerateDC(App.Path & "\Graphics\Explosion" & i & ".bmp")
ExplosionMaskDC(i) = GenerateDC(App.Path & "\Graphics\ExplosionMask" & i & ".bmp")
Next i
For i = 1 To 3
ObjectGraphicsDC(i) = GenerateDC(App.Path & "\Graphics\Object" & i & ".bmp")
ObjectMaskDC(i) = GenerateDC(App.Path & "\Graphics\ObjectMask" & i & ".bmp")
Next i
ShieldBarDC = GenerateDC(App.Path & "\Graphics\ShieldBar.bmp")
HealthBarDC = GenerateDC(App.Path & "\Graphics\HealthBar.bmp")
GameOverDC = GenerateDC(App.Path & "\Graphics\GameOver.bmp")
'=================
'variable settings
'=================
WeaponMin = 1
ShotMin = 1
ShotCount = 0
ExplosionMin = 1
ExplosionCount = 0
If AccountData(CurrentAccount).Layout = 1 Then Player.Theta = 1
'====================
'Load variant and map
'====================
lblWeaponName(1).Caption = WeaponData(Player.Weapon(1)).Name
lblWeaponName(2).Caption = WeaponData(Player.Weapon(2)).Name
lblAmmo.Caption = Player.Ammo(1)
lblClipAmmo.Caption = Player.ClipAmmo(1)
    If WeaponData(Player.Weapon(1)).Reloads = True Then
        frmMain.lblClipAmmo.Visible = True
    Else
        frmMain.lblClipAmmo.Visible = False
    End If
Terminated = False
GameOverCountdown = 300
Call GameLoop
End Sub
Private Sub GameLoop()
Dim FrameCount As Integer, i
Dim CurrentTick As Long
Dim LastTick As Long
Const FrameDifference As Long = 10
Const FPS = 100
Me.Show
FrameCount = 0
Do
    If Terminated = True Then
        'EndIt
        Exit Do
    End If
    
    CurrentTick = GetTickCount()
       
    If CurrentTick - LastTick > FrameDifference Then
        UpdateKeys
        UpdateObjects
        BlitObjects
        LastTick = CurrentTick
        If FrameCount < FPS Then
        FrameCount = FrameCount + 1
        Else
        FrameCount = 0
       'insert actions to be performed every second here
        End If
        DoEvents
    
    Else
        
        DoEvents
        'Sleep 2
    
    End If
Loop
'frmSetup.Visible = True
If victory = True And Player.Dead = False Then
AccountData(CurrentAccount).Level = AccountData(CurrentAccount).Level + 1
For i = 1 To 2
AccountData(CurrentAccount).Weapon(i) = Player.Weapon(i)
AccountData(CurrentAccount).Ammo(i) = Player.Ammo(i)
AccountData(CurrentAccount).ClipAmmo(i) = Player.ClipAmmo(i)
Next i
End If
GameOver = False
ReDim Shot(1 To 1)
WallCount = 0
WeaponCount = 0
WeaponMin = 0
WeaponSpawnCount = 0
PlayerSpawnCount = 0
ShotMin = 0
ShotCount = 0
ExplosionMin = 0
ExplosionCount = 0
'Call ClearPlayers
DeleteGeneratedDC (PlayerGraphicsDC)
DeleteGeneratedDC (PlayerMaskDC)
DeleteGeneratedDC (BackBuffDC)
DeleteGeneratedDC (BackgroundDC)
DeleteGeneratedDC (SightDC)
DeleteGeneratedDC (SightMaskDC)
DeleteGeneratedDC (HealthBarDC)
DeleteGeneratedDC (ShieldBarDC)
For i = 1 To 1
DeleteGeneratedDC (ExplosionGraphicsDC(i))
DeleteGeneratedDC (ExplosionMaskDC(i))
Next i
For i = 1 To 5
DeleteGeneratedDC (ShotGraphicsDC(i))
DeleteGeneratedDC (ShotMaskDC(i))
Next i
For i = 1 To 2
DeleteGeneratedDC (WallDC(i))
Next i
For i = 1 To 9
DeleteGeneratedDC (WeaponGraphicsDC(i))
DeleteGeneratedDC (WeaponMaskDC(i))
Next i
frmMain.Visible = False
Unload frmMain
Set frmMain = Nothing
End Sub


Public Sub BlitObjects()
Dim i, j
'>>> bitblt graphix to backbuffers
'>>> background
BitBlt BackBuffDC, 0, 0, 975, 975, BackgroundDC, 0, 0, vbSrcCopy
'>>> walls
For j = 1 To WallCount
If Wall(j).Type > 0 Then BitBlt BackBuffDC, Wall(j).XCoord - Player.XCoord + picScreen.ScaleWidth / 2, Wall(j).YCoord - Player.YCoord + picScreen.ScaleHeight / 2, Wall(j).Width, Wall(j).Height, WallDC(Wall(j).Type), Wall(j).XCoord - Player.XCoord + picScreen.ScaleWidth / 2, Wall(j).YCoord - Player.YCoord + picScreen.ScaleHeight / 2, vbSrcCopy
Next j
'>>> holoscreen
For j = 1 To Holocount
BitBlt BackBuffDC, Holoscreen(j).XCoord - 30 + picScreen.ScaleWidth / 2 - Player.XCoord, Holoscreen(j).YCoord - 30 + picScreen.ScaleHeight / 2 - Player.YCoord, 60, 60, ObjectMaskDC(Holoscreen(j).GraphicsDC), 0, 0, vbSrcAnd
BitBlt BackBuffDC, Holoscreen(j).XCoord - 30 + picScreen.ScaleWidth / 2 - Player.XCoord, Holoscreen(j).YCoord - 30 + picScreen.ScaleHeight / 2 - Player.YCoord, 60, 60, ObjectGraphicsDC(Holoscreen(j).GraphicsDC), 0, 0, vbSrcPaint
Next j
'>>> weapons
For j = WeaponMin To WeaponCount
If Weapon(j).Deleted = False Then
BitBlt BackBuffDC, picScreen.ScaleWidth / 2 - 20 + Weapon(j).XCoord - Player.XCoord, picScreen.ScaleHeight / 2 - 20 + Weapon(j).YCoord - Player.YCoord, 40, 40, WeaponMaskDC(WeaponData(Weapon(j).WeaponType).GraphicsIndex), 0, 0, vbSrcAnd
BitBlt BackBuffDC, picScreen.ScaleWidth / 2 - 20 + Weapon(j).XCoord - Player.XCoord, picScreen.ScaleHeight / 2 - 20 + Weapon(j).YCoord - Player.YCoord, 40, 40, WeaponGraphicsDC(WeaponData(Weapon(j).WeaponType).GraphicsIndex), 0, 0, vbSrcPaint
End If
Next j
'>>> player stuff
If Player.Dead = False Then
    BitBlt BackBuffDC, picScreen.ScaleWidth / 2 - 20, picScreen.ScaleHeight / 2 - 20, 40, 40, PlayerMaskDC, 0, 0, vbSrcAnd
    BitBlt BackBuffDC, picScreen.ScaleWidth / 2 - 20, picScreen.ScaleHeight / 2 - 20, 40, 40, PlayerGraphicsDC, 0, 0, vbSrcPaint
End If
For j = 1 To EnemyCount
If Enemy(j).VisualContact = True And Enemy(j).Dead = False Then
        BitBlt BackBuffDC, picScreen.ScaleWidth / 2 + Enemy(j).RadarBlipX - 8, picScreen.ScaleHeight / 2 + Enemy(j).RadarBlipY - 8, 16, 16, RadarBlipMaskDC, 0, 0, vbSrcAnd
    If Distance(Player.XCoord, Player.YCoord, Enemy(j).XCoord, Enemy(j).YCoord) <= WeaponData(Player.Weapon(1)).ShotSpeed * WeaponData(Player.Weapon(1)).ShotLifespan Then
        BitBlt BackBuffDC, picScreen.ScaleWidth / 2 + Enemy(j).RadarBlipX - 8, picScreen.ScaleHeight / 2 + Enemy(j).RadarBlipY - 8, 16, 16, RadarBlipDC(2), 0, 0, vbSrcPaint
    Else
        BitBlt BackBuffDC, picScreen.ScaleWidth / 2 + Enemy(j).RadarBlipX - 8, picScreen.ScaleHeight / 2 + Enemy(j).RadarBlipY - 8, 16, 16, RadarBlipDC(1), 0, 0, vbSrcPaint
    End If
End If
If Player.Dead = False And AccountData(CurrentAccount).Layout = 2 Then
    BitBlt BackBuffDC, picScreen.ScaleWidth / 2 + Player.MoveX - 8, picScreen.ScaleHeight / 2 + Player.MoveY - 8, 16, 16, SightMaskDC, 0, 0, vbSrcAnd
    BitBlt BackBuffDC, picScreen.ScaleWidth / 2 + Player.MoveX - 8, picScreen.ScaleHeight / 2 + Player.MoveY - 8, 16, 16, SightDC, 0, 0, vbSrcPaint
End If
If Enemy(j).Dead = False Then
    BitBlt BackBuffDC, Enemy(j).XCoord - Player.XCoord + picScreen.ScaleWidth / 2 - 20, Enemy(j).YCoord - Player.YCoord + picScreen.ScaleHeight / 2 - 20, 40, 40, EnemyMaskDC(Enemy(j).GraphicsDC), 0, 0, vbSrcAnd
    BitBlt BackBuffDC, Enemy(j).XCoord - Player.XCoord + picScreen.ScaleWidth / 2 - 20, Enemy(j).YCoord - Player.YCoord + picScreen.ScaleHeight / 2 - 20, 40, 40, EnemyDC(Enemy(j).GraphicsDC), 0, 0, vbSrcPaint
End If
Next j
'>>> explosions
For j = ExplosionMin To ExplosionCount
    If Explosion(j).Deleted = False Then
        BitBlt BackBuffDC, picScreen.ScaleWidth / 2 + Explosion(j).XCoord - 40 - Player.XCoord, picScreen.ScaleHeight / 2 + Explosion(j).YCoord - 40 - Player.YCoord, 80, 80, ExplosionMaskDC(Explosion(j).GraphicsIndex), 0, 0, vbSrcAnd
        BitBlt BackBuffDC, picScreen.ScaleWidth / 2 + Explosion(j).XCoord - 40 - Player.XCoord, picScreen.ScaleHeight / 2 + Explosion(j).YCoord - 40 - Player.YCoord, 80, 80, ExplosionGraphicsDC(Explosion(j).GraphicsIndex), 0, 0, vbSrcPaint
    End If
Next j
'>>> shots
For j = ShotMin To ShotCount
If Shot(j).Deleted = False And Shot(j).Visible = True Then
BitBlt BackBuffDC, Shot(j).XCoord + picScreen.ScaleWidth / 2 - 15 - Player.XCoord, Shot(j).YCoord + picScreen.ScaleHeight / 2 - 15 - Player.YCoord, 30, 30, ShotMaskDC(Shot(j).GraphicsIndex), 0, 0, vbSrcAnd
BitBlt BackBuffDC, Shot(j).XCoord + picScreen.ScaleWidth / 2 - 15 - Player.XCoord, Shot(j).YCoord + picScreen.ScaleHeight / 2 - 15 - Player.YCoord, 30, 30, ShotGraphicsDC(Shot(j).GraphicsIndex), 0, 0, vbSrcPaint
End If
Next j
'>>> healthbars
BitBlt BackBuffDC, 825, 25, Player.HP / (PlayerMaxHealth * PlayerBonus) * 125, 25, HealthBarDC, 0, 0, vbSrcCopy
BitBlt BackBuffDC, 825, 25, Player.Shield / (PlayerMaxShield * PlayerBonus) * 125, 25, ShieldBarDC, 0, 0, vbSrcCopy
'>>> GAME OVER screen
'If GameOver = True Then BitBlt BackBuffDC, 0, 0, 625, 625, GameOverDC, 0, 0, vbSrcPaint
'bitblt backbuffers to screen
BitBlt picScreen.hdc, 0, 0, 975, 975, BackBuffDC, 0, 0, vbSrcCopy
End Sub

Private Sub lblExit_Click()
victory = False
Terminated = True
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Player.LastTargetX = Player.TargetX
Player.LastTargetY = Player.TargetY
Player.TargetX = x - picScreen.ScaleWidth / 2
Player.TargetY = y - picScreen.ScaleHeight / 2
If AccountData(CurrentAccount).Layout = 3 Then
Player.MoveX = x - picScreen.ScaleWidth / 2
Player.MoveY = y - picScreen.ScaleHeight / 2
End If
End Sub
