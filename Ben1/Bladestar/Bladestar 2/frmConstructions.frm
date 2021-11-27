VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   15240
   ClientLeft      =   60
   ClientTop       =   195
   ClientWidth     =   19080
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmConstructions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1016
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1272
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrGameLoop 
      Interval        =   10
      Left            =   1200
      Top             =   12960
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   9375
      Index           =   2
      Left            =   9600
      ScaleHeight     =   623
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   623
      TabIndex        =   2
      Top             =   600
      Width           =   9375
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   9375
      Index           =   1
      Left            =   120
      ScaleHeight     =   623
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   623
      TabIndex        =   1
      Top             =   600
      Width           =   9375
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   18855
   End
   Begin VB.Label lblWeaponName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   12
      Top             =   10680
      Width           =   2295
   End
   Begin VB.Label lblWeaponName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Index           =   2
      Left            =   9960
      TabIndex        =   11
      Top             =   10680
      Width           =   2295
   End
   Begin VB.Label lblPlayerStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   3720
      TabIndex        =   10
      Top             =   10080
      Width           =   4575
   End
   Begin VB.Label lblPlayerStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   2
      Left            =   13200
      TabIndex        =   9
      Top             =   10080
      Width           =   4575
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Index           =   2
      Left            =   17880
      TabIndex        =   8
      Top             =   10080
      Width           =   1095
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Index           =   1
      Left            =   8400
      TabIndex        =   7
      Top             =   10080
      Width           =   1095
   End
   Begin VB.Label lblClipAmmo 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Index           =   2
      Left            =   11175
      TabIndex        =   6
      Top             =   10080
      Width           =   1095
   End
   Begin VB.Label lblClipAmmo 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Index           =   1
      Left            =   1695
      TabIndex        =   5
      Top             =   10080
      Width           =   1095
   End
   Begin VB.Label lblAmmo 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Index           =   1
      Left            =   495
      TabIndex        =   4
      Top             =   10080
      Width           =   1095
   End
   Begin VB.Label lblAmmo 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Index           =   2
      Left            =   9975
      TabIndex        =   3
      Top             =   10080
      Width           =   1095
   End
   Begin VB.Label lblEXIT 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "||| ABORT |||"
      BeginProperty Font 
         Name            =   "Weltron Urban"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Left            =   16440
      TabIndex        =   0
      Top             =   14520
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(0, 40, 80)
Randomize
'================
'Set graphics dcs
'================
For i = 1 To 2
PlayerGraphicsDC(i) = GenerateDC(App.Path & "\Graphics\Player" & i & ".bmp")
PlayerMaskDC(i) = GenerateDC(App.Path & "\Graphics\PlayerMask" & i & ".bmp")
Next i
For i = 1 To 9
WeaponGraphicsDC(i) = GenerateDC(App.Path & "\Graphics\Item" & i & ".bmp")
WeaponMaskDC(i) = GenerateDC(App.Path & "\Graphics\ItemMask" & i & ".bmp")
Next i
SightDC = GenerateDC(App.Path & "\Graphics\Sight.bmp")
SightMaskDC = GenerateDC(App.Path & "\Graphics\SightMask.bmp")
For i = 1 To 2
RadarBlipDC(i) = GenerateDC(App.Path & "\Graphics\Blip" & i & ".bmp")
Next i
RadarBlipMaskDC = GenerateDC(App.Path & "\Graphics\BlipMask.bmp")
For i = 1 To 4
ShotGraphicsDC(i) = GenerateDC(App.Path & "\Graphics\Shot" & i & ".bmp")
ShotMaskDC(i) = GenerateDC(App.Path & "\Graphics\ShotMask" & i & ".bmp")
Next i
For i = 1 To 1
ExplosionGraphicsDC(i) = GenerateDC(App.Path & "\Graphics\Explosion" & i & ".bmp")
ExplosionMaskDC(i) = GenerateDC(App.Path & "\Graphics\ExplosionMask" & i & ".bmp")
Next i
ShieldBarDC = GenerateDC(App.Path & "\Graphics\ShieldBar.bmp")
HealthBarDC = GenerateDC(App.Path & "\Graphics\HealthBar.bmp")
GameOverDC = GenerateDC(App.Path & "\Graphics\GameOver.bmp")
'=================
'variable settings
'=================
WeaponMin = 1
WeaponCount = 0
ShotMin = 1
ShotCount = 0
ExplosionMin = 1
ExplosionCount = 0
For i = 1 To 2
With Player(i)
    .Shield = 100
    .HP = 200
End With
Next i
lblHeader.Caption = PlayerRecord(Player(1).RecordIndex).Name & " vs. " & PlayerRecord(Player(2).RecordIndex).Name
'====================
'Load variant and map
'====================
Select Case RuleVariant
    Case Is = 0
Call InitializeWeaponsNormal
    Case Is = 1
Call InitializeWeaponsGrenades
    Case Is = 2
Call InitializeWeaponsSwords
    Case Is = 3
Call InitializeWeaponsHardcoreNormal
End Select
If RandomWeapons = True Then StartWeapon = Int(Rnd * WeaponDataCount + 1)
Select Case MapSelected
    Case Is = 0
Call GenerateMap1
    Case Is = 1
Call GenerateMap2
    Case Is = 2
Call GenerateMap3
End Select
For i = 1 To 2
lblWeaponName(i).Caption = WeaponData(Player(i).Weapon).Name
lblAmmo(i).Caption = Player(i).Ammo
lblClipAmmo(i).Caption = Player(i).ClipAmmo
Next i
Call UpdateSpawns
Terminated = False
Call GameLoop
End Sub

Private Sub lblExit_Click()
Dim i
If MsgBox("Are you sure you want to end the game?" & vbLf & "This game will not be scored on the tournament records.", vbYesNo, "Exit Game") = vbYes Then
GameOver = True
UnloadCountdown = 5
lblEXIT.Enabled = False
End If
End Sub

Private Sub BlitObjects()
Dim i, j
'>>> bitblt graphix to backbuffers
For i = 1 To 2
'>>> background
BitBlt BackBuffDC(i), 0, 0, 625, 625, BackgroundDC, 0, 0, vbSrcCopy
'>>> walls
For j = 1 To WallCount
BitBlt BackBuffDC(i), Wall(j).XCoord - Player(i).XCoord + picScreen(i).ScaleWidth / 2, Wall(j).YCoord - Player(i).YCoord + picScreen(i).ScaleHeight / 2, Wall(j).Width, Wall(j).Height, WallDC(Wall(j).Type), Wall(j).XCoord - Player(i).XCoord + 312.5, Wall(j).YCoord - Player(i).YCoord + 312.5, vbSrcCopy
Next j
'>>> weapons
For j = WeaponMin To WeaponCount
If Weapon(j).Deleted = False Then
BitBlt BackBuffDC(i), picScreen(i).ScaleWidth / 2 - 20 + Weapon(j).XCoord - Player(i).XCoord, picScreen(i).ScaleHeight / 2 - 20 + Weapon(j).YCoord - Player(i).YCoord, 40, 40, WeaponMaskDC(WeaponData(Weapon(j).WeaponType).GraphicsIndex), 0, 0, vbSrcAnd
BitBlt BackBuffDC(i), picScreen(i).ScaleWidth / 2 - 20 + Weapon(j).XCoord - Player(i).XCoord, picScreen(i).ScaleHeight / 2 - 20 + Weapon(j).YCoord - Player(i).YCoord, 40, 40, WeaponGraphicsDC(WeaponData(Weapon(j).WeaponType).GraphicsIndex), 0, 0, vbSrcPaint
End If
Next j
'>>> player stuff
For j = 1 To 2
    If Player(j).Dead = False Then
        BitBlt BackBuffDC(i), picScreen(i).ScaleWidth / 2 - 20 + Player(j).XCoord - Player(i).XCoord, picScreen(i).ScaleHeight / 2 - 20 + Player(j).YCoord - Player(i).YCoord, 40, 40, PlayerMaskDC(j), 0, 0, vbSrcAnd
        BitBlt BackBuffDC(i), picScreen(i).ScaleWidth / 2 - 20 + Player(j).XCoord - Player(i).XCoord, picScreen(i).ScaleHeight / 2 - 20 + Player(j).YCoord - Player(i).YCoord, 40, 40, PlayerGraphicsDC(j), 0, 0, vbSrcPaint
    End If
Next j
If Player(i).Dead = False Then
BitBlt BackBuffDC(i), picScreen(i).ScaleWidth / 2 + Player(i).TargetX - 8, picScreen(i).ScaleHeight / 2 + Player(i).TargetY - 8, 16, 16, SightMaskDC, 0, 0, vbSrcAnd
BitBlt BackBuffDC(i), picScreen(i).ScaleWidth / 2 + Player(i).TargetX - 8, picScreen(i).ScaleHeight / 2 + Player(i).TargetY - 8, 16, 16, SightDC, 0, 0, vbSrcPaint
If VisualContact = True Then
        BitBlt BackBuffDC(i), picScreen(i).ScaleWidth / 2 + Player(i).RadarBlipX - 8, picScreen(i).ScaleHeight / 2 + Player(i).RadarBlipY - 8, 16, 16, RadarBlipMaskDC, 0, 0, vbSrcAnd
    If Distance(Player(1).XCoord, Player(1).YCoord, Player(2).XCoord, Player(2).YCoord) <= WeaponData(Player(i).Weapon).ShotSpeed * WeaponData(Player(i).Weapon).ShotLifespan Then
        BitBlt BackBuffDC(i), picScreen(i).ScaleWidth / 2 + Player(i).RadarBlipX - 8, picScreen(i).ScaleHeight / 2 + Player(i).RadarBlipY - 8, 16, 16, RadarBlipDC(2), 0, 0, vbSrcPaint
    Else
        BitBlt BackBuffDC(i), picScreen(i).ScaleWidth / 2 + Player(i).RadarBlipX - 8, picScreen(i).ScaleHeight / 2 + Player(i).RadarBlipY - 8, 16, 16, RadarBlipDC(1), 0, 0, vbSrcPaint
    End If
End If
End If
'>>> explosions
For j = ExplosionMin To ExplosionCount
    If Explosion(j).Deleted = False Then
        BitBlt BackBuffDC(i), picScreen(i).ScaleWidth / 2 + Explosion(j).XCoord - 40 - Player(i).XCoord, picScreen(i).ScaleHeight / 2 + Explosion(j).YCoord - 40 - Player(i).YCoord, 80, 80, ExplosionMaskDC(Explosion(j).GraphicsIndex), 0, 0, vbSrcAnd
        BitBlt BackBuffDC(i), picScreen(i).ScaleWidth / 2 + Explosion(j).XCoord - 40 - Player(i).XCoord, picScreen(i).ScaleHeight / 2 + Explosion(j).YCoord - 40 - Player(i).YCoord, 80, 80, ExplosionGraphicsDC(Explosion(j).GraphicsIndex), 0, 0, vbSrcPaint
    End If
Next j
'>>> shots
For j = ShotMin To ShotCount
If Shot(j).Deleted = False And Shot(j).Visible = True Then
BitBlt BackBuffDC(i), Shot(j).XCoord + picScreen(i).ScaleWidth / 2 - 15 - Player(i).XCoord, Shot(j).YCoord + picScreen(i).ScaleHeight / 2 - 15 - Player(i).YCoord, 30, 30, ShotMaskDC(Shot(j).GraphicsIndex), 0, 0, vbSrcAnd
BitBlt BackBuffDC(i), Shot(j).XCoord + picScreen(i).ScaleWidth / 2 - 15 - Player(i).XCoord, Shot(j).YCoord + picScreen(i).ScaleHeight / 2 - 15 - Player(i).YCoord, 30, 30, ShotGraphicsDC(Shot(j).GraphicsIndex), 0, 0, vbSrcPaint
End If
Next j
'>>> healthbars
BitBlt BackBuffDC(i), 25, picScreen(i).ScaleHeight - 50, Player(i).HP / 200 * 125, 25, HealthBarDC, 0, 0, vbSrcCopy
BitBlt BackBuffDC(i), 25, picScreen(i).ScaleHeight - 50, Player(i).Shield / 100 * 125, 25, ShieldBarDC, 0, 0, vbSrcCopy
'>>> GAME OVER screen
If GameOver = True Then BitBlt BackBuffDC(i), 0, 0, 625, 625, GameOverDC, 0, 0, vbSrcPaint
'bitblt backbuffers to screen
BitBlt picScreen(i).hdc, 0, 0, 625, 625, BackBuffDC(i), 0, 0, vbSrcCopy
Next i
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
        'Me.Refresh
        LastTick = CurrentTick
        If FrameCount < FPS Then
        FrameCount = FrameCount + 1
        Else
        FrameCount = 0
        UpdateSpawns
        End If
        DoEvents
    
    Else
        
        DoEvents
        'Sleep 2
    
    End If
Loop
frmSetup.Visible = True
GameOver = False
WallCount = 0
WeaponCount = 0
WeaponMin = 0
WeaponSpawnCount = 0
PlayerSpawnCount = 0
ShotMin = 0
ShotCount = 0
ExplosionMin = 0
ExplosionCount = 0
Call ClearPlayers
For i = 1 To 2
DeleteGeneratedDC (PlayerGraphicsDC(i))
DeleteGeneratedDC (PlayerMaskDC(i))
DeleteGeneratedDC (BackBuffDC(i))
Next i
DeleteGeneratedDC (BackgroundDC)
DeleteGeneratedDC (SightDC)
DeleteGeneratedDC (SightMaskDC)
DeleteGeneratedDC (HealthBarDC)
DeleteGeneratedDC (ShieldBarDC)
For i = 1 To 1
DeleteGeneratedDC (ExplosionGraphicsDC(i))
DeleteGeneratedDC (ExplosionMaskDC(i))
Next i
For i = 1 To 1
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

